# -*- coding: utf-8 -*-
"""
SEO診断ツール — Streamlit Webアプリ

Search ConsoleのExcel/CSVをアップロードし、
ユーザー自身のGemini APIキーで改善処方箋を生成する。
"""

import io
import logging
import re
from datetime import datetime

import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from core.sc_parser import parse_sc_file
from core.analyzer import analyze_sc_data
from core.gemini_client import generate_prescription

logging.basicConfig(level=logging.INFO)

# ──────────────────────────────────────
# ページ設定
# ──────────────────────────────────────
st.set_page_config(
    page_title="SEO診断ツール",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
section[data-testid="stSidebar"] { min-width: 340px; }
.stMetric label { font-size: 0.85rem; }
.privacy-box {
    background: #1e2a1e;
    border: 1px solid #2d5a2d;
    border-radius: 6px;
    padding: 12px 16px;
    font-size: 0.82rem;
    color: #90c990;
}
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────
# ユーティリティ
# ──────────────────────────────────────

def _extract_url_from_filename(name: str) -> str:
    """SCエクスポートのファイル名からサイトURLを抽出する"""
    stem = re.sub(r'\.(xlsx|csv)$', '', name, flags=re.IGNORECASE)
    m = re.match(r'^(https?___[^_]+(?:_[^-][^_]*)*?)(?=_-|$)', stem)
    if m:
        return m.group(1).replace('___', '://')
    return ""


def _build_excel(analysis: dict, prescription: str, site_url: str) -> bytes:
    """診断結果をExcel形式で生成する"""
    wb = Workbook()

    # ── スタイル定義 ──
    hdr_font  = Font(bold=True, color="FFFFFF", size=11)
    hdr_fill  = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin      = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin"),
    )
    wrap = Alignment(wrap_text=True, vertical="top")

    # ════════════════════════════════════
    # シート①：概況
    # ════════════════════════════════════
    ws1 = wb.active
    ws1.title = "概況"

    ws1.column_dimensions["A"].width = 22
    ws1.column_dimensions["B"].width = 20

    ws1["A1"] = "SEO診断レポート"
    ws1["A1"].font = Font(bold=True, size=14)
    ws1.merge_cells("A1:B1")

    rows = [
        ("サイトURL",       site_url),
        ("診断日時",        datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("分析キーワード数", f"{analysis['total_kw']:,} 件"),
        ("総クリック数",    f"{analysis['total_clicks']:,}"),
        ("総表示回数",      f"{analysis['total_impressions']:,}"),
        ("全体CTR",         f"{analysis['avg_ctr']:.1%}"),
        ("加重平均順位",    f"{analysis['avg_position']:.1f} 位"),
    ]
    for i, (label, val) in enumerate(rows, start=3):
        ws1[f"A{i}"] = label
        ws1[f"A{i}"].font = Font(bold=True)
        ws1[f"B{i}"] = val

    ws1["A11"] = "順位帯分布"
    ws1["A11"].font = Font(bold=True)
    for j, (label, cnt) in enumerate(analysis["pos_distribution"].items(), start=12):
        ws1[f"A{j}"] = label
        ws1[f"B{j}"] = f"{cnt} 件"

    # ════════════════════════════════════
    # シート②：改善処方箋
    # ════════════════════════════════════
    ws2 = wb.create_sheet("改善処方箋")
    ws2.column_dimensions["A"].width = 100
    ws2.row_dimensions[1].height = 24

    ws2["A1"] = "改善処方箋（AI生成）"
    ws2["A1"].font = Font(bold=True, size=13)

    # Markdownを除去してプレーンテキストで格納
    plain = re.sub(r'\*{1,3}([^\*]+)\*{1,3}', r'\1', prescription)
    plain = re.sub(r'^#{1,6}\s*', '', plain, flags=re.MULTILINE)
    plain = re.sub(r'^---+$', '─' * 40, plain, flags=re.MULTILINE)

    ws2["A2"] = plain
    ws2["A2"].alignment = wrap
    ws2.row_dimensions[2].height = max(400, plain.count('\n') * 15)

    # ════════════════════════════════════
    # シート③：改善対象キーワード
    # ════════════════════════════════════
    ws3 = wb.create_sheet("改善対象キーワード")

    headers = ["#", "キーワード", "順位", "表示回数", "クリック数", "CTR", "CTR評価"]
    col_widths = [5, 40, 8, 12, 12, 10, 14]

    for col, (h, w) in enumerate(zip(headers, col_widths), start=1):
        cell = ws3.cell(row=1, column=col, value=h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = hdr_align
        cell.border = thin
        ws3.column_dimensions[get_column_letter(col)].width = w

    for i, kw in enumerate(analysis["top_keywords"], start=1):
        row = i + 1
        values = [
            i,
            kw["query"],
            kw["position"],
            kw["impressions"],
            kw["clicks"],
            kw["ctr_pct"],
            kw["ctr_status"],
        ]
        for col, val in enumerate(values, start=1):
            cell = ws3.cell(row=row, column=col, value=val)
            cell.border = thin
            cell.alignment = Alignment(vertical="center")

    ws3.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _build_html(analysis: dict, prescription: str, site_url: str) -> str:
    """診断結果をHTML形式で生成する"""
    import html as html_mod

    dist_rows = "".join(
        f"<tr><td>{k}</td><td>{v}件</td></tr>"
        for k, v in analysis["pos_distribution"].items()
    )

    kw_rows = "".join(
        f"<tr><td>{i}</td><td>{html_mod.escape(kw['query'])}</td>"
        f"<td>{kw['position']}</td><td>{kw['impressions']:,}</td>"
        f"<td>{kw['clicks']}</td><td>{kw['ctr_pct']}</td>"
        f"<td>{kw['ctr_status']}</td></tr>"
        for i, kw in enumerate(analysis["top_keywords"], 1)
    )

    # MarkdownをHTMLに簡易変換
    import re as _re
    md = prescription
    md = _re.sub(r'^## (.+)$', r'<h2>\1</h2>', md, flags=_re.MULTILINE)
    md = _re.sub(r'^### (.+)$', r'<h3>\1</h3>', md, flags=_re.MULTILINE)
    md = _re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', md)
    md = _re.sub(r'\*(.+?)\*', r'<em>\1</em>', md)
    md = _re.sub(r'^---+$', '<hr>', md, flags=_re.MULTILINE)
    md = _re.sub(r'^\*   (.+)$', r'<li>\1</li>', md, flags=_re.MULTILINE)
    md = md.replace('\n', '<br>\n')

    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SEO診断レポート — {html_mod.escape(site_url)}</title>
<style>
  body {{ font-family: 'Hiragino Sans', 'Meiryo', sans-serif; max-width: 960px; margin: 0 auto; padding: 24px; color: #222; }}
  h1 {{ color: #2F5496; border-bottom: 2px solid #2F5496; padding-bottom: 8px; }}
  h2 {{ color: #c0392b; margin-top: 32px; }}
  h3 {{ color: #2980b9; }}
  table {{ border-collapse: collapse; width: 100%; margin: 16px 0; }}
  th {{ background: #2F5496; color: #fff; padding: 8px 12px; text-align: left; }}
  td {{ border: 1px solid #ddd; padding: 6px 12px; }}
  tr:nth-child(even) {{ background: #f5f8ff; }}
  .summary-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin: 16px 0; }}
  .metric {{ background: #f0f4ff; border-radius: 8px; padding: 12px; text-align: center; }}
  .metric-value {{ font-size: 1.6em; font-weight: bold; color: #2F5496; }}
  .metric-label {{ font-size: 0.82em; color: #666; }}
  .prescription {{ background: #fff9f0; border-left: 4px solid #e67e22; padding: 16px 24px; border-radius: 4px; }}
  .footer {{ font-size: 0.8em; color: #999; margin-top: 40px; border-top: 1px solid #eee; padding-top: 12px; }}
</style>
</head>
<body>
<h1>🏥 SEO診断レポート</h1>
<p><strong>サイト：</strong>{html_mod.escape(site_url)}　／　<strong>診断日時：</strong>{now}</p>

<h2>📊 サイト概況</h2>
<div class="summary-grid">
  <div class="metric"><div class="metric-value">{analysis['total_kw']:,}</div><div class="metric-label">分析キーワード数</div></div>
  <div class="metric"><div class="metric-value">{analysis['avg_position']:.1f}位</div><div class="metric-label">平均掲載順位</div></div>
  <div class="metric"><div class="metric-value">{analysis['avg_ctr']:.1%}</div><div class="metric-label">全体CTR</div></div>
  <div class="metric"><div class="metric-value">{analysis['total_clicks']:,}</div><div class="metric-label">総クリック数</div></div>
</div>

<h3>順位帯分布</h3>
<table><tr><th>順位帯</th><th>件数</th></tr>{dist_rows}</table>

<h2>💊 改善処方箋</h2>
<div class="prescription">{md}</div>

<h2>📋 改善対象キーワード一覧</h2>
<table>
  <tr><th>#</th><th>キーワード</th><th>順位</th><th>表示回数</th><th>クリック数</th><th>CTR</th><th>CTR評価</th></tr>
  {kw_rows}
</table>

<div class="footer">※ この処方箋はGemini AIが生成したものです。実施前にご自身で内容をご確認ください。</div>
</body>
</html>"""


# ──────────────────────────────────────
# サイドバー
# ──────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 設定")

    gemini_api_key = st.text_input(
        "Gemini APIキー",
        type="password",
        placeholder="AIza...",
        help="Google AI Studio（aistudio.google.com）で無料取得できます",
    )

    if not gemini_api_key:
        with st.expander("🔑 APIキーの取得方法（無料）"):
            st.markdown("""
1. [Google AI Studio](https://aistudio.google.com/) を開く
2. 右上 **「Get API key」** をクリック
3. **「APIキーを作成」** → キーをコピー
4. 上の欄に貼り付ける

**無料枠の目安**
- 1日 1,500リクエストまで
- 1分 15リクエストまで
- 診断1回 ≒ 1リクエスト → **ほぼ無制限で使用可能**
""")

    st.markdown("""
<div class="privacy-box">
🔒 <strong>プライバシーについて</strong><br>
・APIキーはこのブラウザセッション内のみで使用されます<br>
・当サービスのサーバーへの保存・送信は一切行いません<br>
・アップロードしたキーワードデータはGemini API（Google）に送信されます<br>
・セッション終了後にデータは破棄されます
</div>
""", unsafe_allow_html=True)

    st.divider()

    st.subheader("📥 データの準備方法")
    st.markdown("""
**Search Consoleからのエクスポート手順：**

1. [Google Search Console](https://search.google.com/search-console) を開く
2. 左メニュー **「検索パフォーマンス」** をクリック
3. 期間を **「3か月」** に設定
4. 右上 **「エクスポート」** ボタンをクリック
5. **「Excelとしてダウンロード」** または **「CSVをダウンロード」** を選択

> ✅ Excel (.xlsx) と CSV (.csv) の両方に対応しています
> ✅ Googleスプレッドシートで開いた場合は「ファイル→ダウンロード→CSV」でエクスポートしてください
""")


# ──────────────────────────────────────
# メインエリア
# ──────────────────────────────────────
st.title("🏥 SEO診断ツール")
st.caption("Search ConsoleのファイルをアップロードするだけでAIが改善処方箋を生成します（Excel・CSV対応）")

st.divider()

col_upload, col_url = st.columns([3, 2])

with col_upload:
    uploaded_file = st.file_uploader(
        "📂 Search Consoleファイルをアップロード",
        type=["csv", "xlsx"],
        help="Excel (.xlsx) または CSV (.csv) に対応しています",
        label_visibility="visible",
    )
    if uploaded_file:
        st.caption(f"✅ ファイル受信: {uploaded_file.name}")

with col_url:
    auto_url = _extract_url_from_filename(uploaded_file.name) if uploaded_file else ""
    site_url = st.text_input(
        "🌐 サイトURL（任意）",
        value=auto_url,
        placeholder="https://example-blog.com",
        help="ファイル名から自動取得します。変更も可能です。",
    )

st.divider()

# ──────────────────────────────────────
# 診断実行
# ──────────────────────────────────────
ready = uploaded_file is not None and bool(gemini_api_key)

if ready:
    if st.button("🔍 診断を開始する", type="primary", use_container_width=True):

        # 1. ファイル解析
        with st.spinner("📊 ファイルを解析中..."):
            df, error_msg = parse_sc_file(uploaded_file)

        if error_msg:
            st.error(f"❌ {error_msg}")
            st.stop()

        if df is None or df.empty:
            st.error(
                "❌ データが空でした。\n\n"
                "Search Console の「検索パフォーマンス」→「検索結果」画面で"
                "エクスポートしたファイルを使用してください。"
            )
            st.stop()

        # 2. データ分析
        with st.spinner("🔬 データを分析中..."):
            analysis = analyze_sc_data(df)

        # 概況表示
        st.subheader("📊 サイト概況")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("分析キーワード数",  f"{analysis['total_kw']:,} 件")
        m2.metric("平均掲載順位",      f"{analysis['avg_position']:.1f} 位")
        m3.metric("全体CTR",           f"{analysis['avg_ctr']:.1%}")
        m4.metric("総クリック数",      f"{analysis['total_clicks']:,}")

        with st.expander("順位帯の内訳を見る"):
            dist = analysis["pos_distribution"]
            cols = st.columns(len(dist))
            for col, (label, count) in zip(cols, dist.items()):
                col.metric(label, f"{count} 件")

        st.divider()

        # 3. AI処方箋生成
        with st.spinner("💊 AIが処方箋を作成中です...（30秒〜1分かかります）"):
            prescription = generate_prescription(
                api_key=gemini_api_key,
                analysis=analysis,
                site_url=site_url or "（URL未入力）",
            )

        st.subheader("💊 改善処方箋")
        st.markdown(prescription)

        st.divider()

        # ダウンロードボタン群
        display_url = site_url or "seo_report"
        safe_url = re.sub(r'[^a-zA-Z0-9._-]', '_', display_url)
        date_str = datetime.now().strftime("%Y%m%d")
        base_name = f"seo_report_{safe_url}_{date_str}"

        dl1, dl2 = st.columns(2)

        with dl1:
            excel_bytes = _build_excel(analysis, prescription, site_url or "（URL未入力）")
            st.download_button(
                label="📊 Excelでダウンロード",
                data=excel_bytes,
                file_name=f"{base_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with dl2:
            html_bytes = _build_html(analysis, prescription, site_url or "（URL未入力）").encode("utf-8")
            st.download_button(
                label="🌐 HTMLでダウンロード",
                data=html_bytes,
                file_name=f"{base_name}.html",
                mime="text/html",
                use_container_width=True,
            )

        st.caption(
            "💡 ExcelファイルはGoogleスプレッドシートにアップロードして開くこともできます。"
            "　※ この処方箋はGemini AIが生成したものです。実施前にご自身で内容をご確認ください。"
        )

# ──────────────────────────────────────
# 未入力時のガイド
# ──────────────────────────────────────
elif uploaded_file is None and not gemini_api_key:
    st.info("👈 左のサイドバーにGemini APIキーを入力し、ファイルをアップロードしてください")

elif uploaded_file is None:
    st.info("📂 Search Consoleのファイルをアップロードしてください（Excel または CSV）")

elif not gemini_api_key:
    st.warning("🔑 左のサイドバーにGemini APIキーを入力してください（無料で取得できます）")

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
    stem = re.sub(r'\.(xlsx|csv)$', '', name, flags=re.IGNORECASE)
    m = re.match(r'^(https?___[^_]+(?:_[^-][^_]*)*?)(?=_-|$)', stem)
    if m:
        return m.group(1).replace('___', '://')
    return ""


def _apply_header(ws, headers, col_widths, row=1):
    hdr_font  = Font(bold=True, color="FFFFFF", size=10)
    hdr_fill  = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin"),
    )
    for col, (h, w) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = hdr_align
        cell.border = thin
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[row].height = 28


def _build_excel(analysis: dict, prescription: str, site_url: str) -> bytes:
    wb = Workbook()
    thin = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin"),
    )
    cell_align = Alignment(vertical="center", wrap_text=False)
    wrap_align = Alignment(vertical="top", wrap_text=True)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    # ════════════════════════════════
    # シート①：概要
    # ════════════════════════════════
    ws_summary = wb.active
    ws_summary.title = "概要"
    ws_summary.column_dimensions["A"].width = 24
    ws_summary.column_dimensions["B"].width = 30

    ws_summary["A1"] = f"SEO Doctor 診断概要 — {site_url}"
    ws_summary["A1"].font = Font(bold=True, size=13)
    ws_summary.merge_cells("A1:B1")
    ws_summary.row_dimensions[1].height = 22

    info_rows = [
        ("■ 基本情報", ""),
        ("対象サイト", site_url),
        ("診断モード", "ワンクリック診断（Search Console）"),
        ("実行日時", now_str),
        ("SC分析期間", "エクスポート期間（通常3か月）"),
        ("", ""),
        ("■ 検出件数", ""),
        ("分析キーワード数", f"{analysis['total_kw']:,} 件"),
        ("リライト候補", f"{sum(1 for k in analysis['all_keywords'] if k['position'] <= 20):,} 件（1〜20位）"),
        ("新規記事提案", f"{sum(1 for k in analysis['all_keywords'] if 20 < k['position'] <= 50):,} 件（21〜50位）"),
        ("", ""),
        ("■ パフォーマンス概況", ""),
        ("総クリック数", f"{analysis['total_clicks']:,}"),
        ("総表示回数", f"{analysis['total_impressions']:,}"),
        ("全体CTR", f"{analysis['avg_ctr']:.1%}"),
        ("加重平均掲載順位", f"{analysis['avg_position']:.1f} 位"),
        ("", ""),
        ("■ 順位帯分布", ""),
    ]
    for i, (lbl, val) in enumerate(info_rows, start=2):
        ws_summary[f"A{i}"] = lbl
        ws_summary[f"B{i}"] = val
        if lbl.startswith("■"):
            ws_summary[f"A{i}"].font = Font(bold=True, color="2F5496")

    row = len(info_rows) + 2
    for label, cnt in analysis["pos_distribution"].items():
        ws_summary[f"A{row}"] = label
        ws_summary[f"B{row}"] = f"{cnt} 件"
        row += 1

    row += 1
    ws_summary[f"A{row}"] = "※ 各タブを順番に確認してください。まず「📋処方箋」→「① リライト候補」→「② 新規記事提案」の順で取り組んでください。"
    ws_summary[f"A{row}"].font = Font(italic=True, color="666666", size=9)
    ws_summary.merge_cells(f"A{row}:B{row}")

    # ════════════════════════════════
    # シート②：使い方
    # ════════════════════════════════
    ws_howto = wb.create_sheet("使い方")
    ws_howto.column_dimensions["A"].width = 90
    howto_lines = [
        f"SEO Doctor — {site_url}",
        "",
        "【このファイルの使い方】",
        "",
        "■ 📋処方箋タブ",
        "  AIが生成した改善処方箋です。優先度の高い順に並んでいます。",
        "  「着手状況」列をプルダウンで「対応中」「対応済み」に変更しながら進めてください。",
        "",
        "■ ① リライト候補タブ",
        "  現在1〜20位に表示されているキーワードのうち、改善余地の大きいものを優先順で並べています。",
        "  CTR差分がマイナス（平均より低い）ほど、タイトル・メタディスクリプションの改善効果が期待できます。",
        "",
        "■ ② 新規記事提案タブ",
        "  21〜50位のキーワードです。1ページ目に引き上げるための新規記事や既存記事の強化候補です。",
        "  表示回数が多いほど、記事を書いた際のインパクトが大きくなります。",
        "",
        "【Google スプレッドシートで開く方法】",
        "  1. Google ドライブを開く（drive.google.com）",
        "  2. このファイルをドラッグ＆ドロップでアップロード",
        "  3. ファイルをダブルクリック → 「Googleスプレッドシートで開く」",
        "",
        "※ この診断書はGemini AIが生成したものです。実施前にご自身で内容をご確認ください。",
        f"※ 生成日時: {now_str}",
    ]
    for i, line in enumerate(howto_lines, start=1):
        ws_howto[f"A{i}"] = line
        if line.startswith("■") or line.startswith("【"):
            ws_howto[f"A{i}"].font = Font(bold=True)
    ws_howto[f"A1"].font = Font(bold=True, size=13)

    # ════════════════════════════════
    # シート③：処方箋
    # ════════════════════════════════
    ws_rx = wb.create_sheet("📋処方箋")
    ws_rx.column_dimensions["A"].width = 12
    ws_rx.column_dimensions["B"].width = 90

    ws_rx["A1"] = "着手状況"
    ws_rx["A1"].font = Font(bold=True, color="FFFFFF")
    ws_rx["A1"].fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    ws_rx["A1"].alignment = Alignment(horizontal="center")
    ws_rx["B1"] = "AI改善処方箋（Gemini生成）"
    ws_rx["B1"].font = Font(bold=True, color="FFFFFF")
    ws_rx["B1"].fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    ws_rx.row_dimensions[1].height = 24

    # Markdownを除去してプレーンテキストに
    plain = re.sub(r'\*{1,3}([^\*]+)\*{1,3}', r'\1', prescription)
    plain = re.sub(r'^#{1,6}\s*', '', plain, flags=re.MULTILINE)
    plain = re.sub(r'^---+$', '─' * 60, plain, flags=re.MULTILINE)

    # セクションごとに行分割して書き込む
    sections = plain.split('\n')
    data_row = 2

    level_fills = {
        "🔴": PatternFill(start_color="FFE0E0", end_color="FFE0E0", fill_type="solid"),
        "🟡": PatternFill(start_color="FFFDE0", end_color="FFFDE0", fill_type="solid"),
        "🟢": PatternFill(start_color="E0FFE8", end_color="E0FFE8", fill_type="solid"),
        "💡": PatternFill(start_color="E0F0FF", end_color="E0F0FF", fill_type="solid"),
    }

    current_fill = None
    for line in sections:
        if not line.strip():
            continue
        for emoji, fill in level_fills.items():
            if emoji in line:
                current_fill = fill
                break
        ws_rx.cell(row=data_row, column=1, value="未着手")
        ws_rx.cell(row=data_row, column=1).alignment = Alignment(horizontal="center", vertical="top")
        cell = ws_rx.cell(row=data_row, column=2, value=line.strip())
        cell.alignment = wrap_align
        if current_fill:
            ws_rx.cell(row=data_row, column=1).fill = current_fill
            cell.fill = current_fill
        ws_rx.row_dimensions[data_row].height = max(15, len(line) // 8 * 13)
        data_row += 1

    ws_rx.freeze_panes = "A2"

    # ════════════════════════════════
    # シート④：リライト候補（1〜20位）
    # ════════════════════════════════
    ws_rw = wb.create_sheet("① リライト候補")

    rw_headers = [
        "着手状況", "順位", "クエリ", "表示回数", "クリック数",
        "CTR", "サイト平均CTR", "CTR差分", "CTR評価", "順位帯診断",
    ]
    rw_widths = [12, 8, 40, 12, 12, 10, 14, 12, 14, 22]
    _apply_header(ws_rw, rw_headers, rw_widths)

    rw_data = [k for k in analysis["all_keywords"] if k["position"] <= 20]
    low_fill  = PatternFill(start_color="FFE0E0", end_color="FFE0E0", fill_type="solid")
    mid_fill  = PatternFill(start_color="FFFDE0", end_color="FFFDE0", fill_type="solid")

    for i, kw in enumerate(rw_data, start=2):
        row_vals = [
            "未着手",
            kw["position"],
            kw["query"],
            kw["impressions"],
            kw["clicks"],
            kw["ctr_pct"],
            kw["avg_ctr_pct"],
            kw["ctr_diff_pct"],
            kw["ctr_status"],
            kw["pos_band"],
        ]
        fill = low_fill if kw["ctr_status"] == "かなり低め" else (
               mid_fill if kw["ctr_status"] == "やや低め" else None)
        for col, val in enumerate(row_vals, start=1):
            cell = ws_rw.cell(row=i, column=col, value=val)
            cell.border = thin
            cell.alignment = cell_align
            if fill and col in (6, 7, 8, 9):
                cell.fill = fill

    ws_rw.freeze_panes = "A2"

    # ════════════════════════════════
    # シート⑤：新規記事提案（21〜50位）
    # ════════════════════════════════
    ws_new = wb.create_sheet("② 新規記事提案")

    new_headers = [
        "着手状況", "順位", "クエリ", "表示回数", "クリック数",
        "CTR", "サイト平均CTR", "CTR差分", "CTR評価", "順位帯診断",
    ]
    new_widths = [12, 8, 40, 12, 12, 10, 14, 12, 14, 22]
    _apply_header(ws_new, new_headers, new_widths)

    new_data = [k for k in analysis["all_keywords"] if 20 < k["position"] <= 50]
    for i, kw in enumerate(new_data, start=2):
        row_vals = [
            "未着手",
            kw["position"],
            kw["query"],
            kw["impressions"],
            kw["clicks"],
            kw["ctr_pct"],
            kw["avg_ctr_pct"],
            kw["ctr_diff_pct"],
            kw["ctr_status"],
            kw["pos_band"],
        ]
        for col, val in enumerate(row_vals, start=1):
            cell = ws_new.cell(row=i, column=col, value=val)
            cell.border = thin
            cell.alignment = cell_align

    ws_new.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _build_html(analysis: dict, prescription: str, site_url: str) -> str:
    import html as html_mod

    dist_rows = "".join(
        f"<tr><td>{k}</td><td>{v}件</td></tr>"
        for k, v in analysis["pos_distribution"].items()
    )

    kw_rows = "".join(
        f"<tr><td>{i}</td><td>{html_mod.escape(kw['query'])}</td>"
        f"<td>{kw['position']}</td><td>{kw['impressions']:,}</td>"
        f"<td>{kw['clicks']}</td><td>{kw['ctr_pct']}</td>"
        f"<td>{kw['ctr_diff_pct']}</td><td>{kw['ctr_status']}</td></tr>"
        for i, kw in enumerate(analysis["all_keywords"][:50], 1)
    )

    md = prescription
    md = re.sub(r'^## (.+)$', r'<h2>\1</h2>', md, flags=re.MULTILINE)
    md = re.sub(r'^### (.+)$', r'<h3>\1</h3>', md, flags=re.MULTILINE)
    md = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', md)
    md = re.sub(r'\*(.+?)\*', r'<em>\1</em>', md)
    md = re.sub(r'^---+$', '<hr>', md, flags=re.MULTILINE)
    md = re.sub(r'^\*   (.+)$', r'<li>\1</li>', md, flags=re.MULTILINE)
    md = md.replace('\n', '<br>\n')

    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SEO診断レポート — {html_mod.escape(site_url)}</title>
<style>
  body {{ font-family: 'Hiragino Sans', 'Meiryo', sans-serif; max-width: 1000px; margin: 0 auto; padding: 24px; color: #222; }}
  h1 {{ color: #2F5496; border-bottom: 3px solid #2F5496; padding-bottom: 8px; }}
  h2 {{ color: #c0392b; margin-top: 36px; border-left: 4px solid #c0392b; padding-left: 10px; }}
  h3 {{ color: #2980b9; }}
  table {{ border-collapse: collapse; width: 100%; margin: 16px 0; font-size: 0.9em; }}
  th {{ background: #2F5496; color: #fff; padding: 8px 12px; text-align: left; }}
  td {{ border: 1px solid #ddd; padding: 6px 10px; }}
  tr:nth-child(even) {{ background: #f5f8ff; }}
  .summary-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin: 16px 0; }}
  .metric {{ background: #f0f4ff; border-radius: 8px; padding: 12px; text-align: center; }}
  .metric-value {{ font-size: 1.6em; font-weight: bold; color: #2F5496; }}
  .metric-label {{ font-size: 0.82em; color: #666; }}
  .prescription {{ background: #fff9f0; border-left: 4px solid #e67e22; padding: 16px 24px; border-radius: 4px; line-height: 1.8; }}
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

<h2>📋 改善対象キーワード一覧（上位50件）</h2>
<table>
  <tr><th>#</th><th>キーワード</th><th>順位</th><th>表示回数</th><th>クリック数</th><th>CTR</th><th>CTR差分</th><th>CTR評価</th></tr>
  {kw_rows}
</table>

<div class="footer">
  ※ この処方箋はGemini AIが生成したものです。実施前にご自身で内容をご確認ください。<br>
  ※ Excelファイルはリライト候補・新規記事提案の全キーワードを含んでいます。
</div>
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

        with st.spinner("📊 ファイルを解析中..."):
            df, error_msg = parse_sc_file(uploaded_file)

        if error_msg:
            st.error(f"❌ {error_msg}")
            st.stop()

        if df is None or df.empty:
            st.error("❌ データが空でした。Search Console の「検索パフォーマンス」でエクスポートしたファイルを使用してください。")
            st.stop()

        with st.spinner("🔬 データを分析中..."):
            analysis = analyze_sc_data(df)

        with st.spinner("💊 AIが処方箋を作成中です...（30秒〜1分かかります）"):
            prescription = generate_prescription(
                api_key=gemini_api_key,
                analysis=analysis,
                site_url=site_url or "（URL未入力）",
            )

        # session_stateに保存（ダウンロードボタンが消えないようにするため）
        st.session_state["analysis"]     = analysis
        st.session_state["prescription"] = prescription
        st.session_state["site_url"]     = site_url or "（URL未入力）"

# ──────────────────────────────────────
# 診断結果の表示（session_stateから読む）
# ──────────────────────────────────────
if "prescription" in st.session_state:
    analysis     = st.session_state["analysis"]
    prescription = st.session_state["prescription"]
    display_url  = st.session_state["site_url"]

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

    st.subheader("💊 改善処方箋")
    st.markdown(prescription)

    st.divider()

    # ダウンロードボタン（session_stateのデータを使うため消えない）
    date_str  = datetime.now().strftime("%Y%m%d")
    safe_url  = re.sub(r'[^a-zA-Z0-9._-]', '_', display_url)[:40]
    base_name = f"SEO診断書_{safe_url}_{date_str}"

    excel_bytes = _build_excel(analysis, prescription, display_url)
    html_bytes  = _build_html(analysis, prescription, display_url).encode("utf-8")

    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            label="📊 Excelでダウンロード（5シート）",
            data=excel_bytes,
            file_name=f"{base_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with dl2:
        st.download_button(
            label="🌐 HTMLでダウンロード",
            data=html_bytes,
            file_name=f"{base_name}.html",
            mime="text/html",
            use_container_width=True,
        )

    st.info(
        "💡 **Googleスプレッドシートで開く方法：**　"
        "ExcelをダウンロードしてGoogleドライブにアップロード → ダブルクリック → 「Googleスプレッドシートで開く」\n\n"
        "※ この処方箋はGemini AIが生成したものです。実施前にご自身で内容をご確認ください。"
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

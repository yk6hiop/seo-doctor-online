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
from core.serp import analyze_serp, analyze_serp_pages

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
    ws_summary[f"A{row}"] = "※ 「📋処方箋」タブをご確認ください。AIが生成した改善提案が優先度順に記載されています。"
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
        "  AIが生成した改善処方箋です。優先度の高い順（🔴今すぐ → 🟡今月中 → 🟢中長期）に並んでいます。",
        "  「着手状況」列を「対応中」「対応済み」に変更しながら進めてください。",
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

    plain = re.sub(r'\*{1,3}([^\*]+)\*{1,3}', r'\1', prescription)
    plain = re.sub(r'^#{1,6}\s*', '', plain, flags=re.MULTILINE)
    plain = re.sub(r'^---+$', '─' * 60, plain, flags=re.MULTILINE)

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

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _parse_prescription_sections(prescription: str) -> list:
    """AI処方箋Markdownをセクションごとのdictリストに変換する"""
    section_colors = {
        "🔴": "#8b1a1a",
        "🟡": "#7a5a00",
        "🟢": "#1a5a1a",
        "💡": "#1a3a5c",
    }
    sections = []
    current = None

    for line in prescription.split('\n'):
        if line.startswith('## '):
            if current:
                sections.append(current)
            title = line[3:].strip()
            color = "#2F5496"
            for emoji, c in section_colors.items():
                if emoji in title:
                    color = c
                    break
            current = {"title": title, "color": color, "items": [], "body_lines": []}
        elif line.startswith('### ') and current is not None:
            current["items"].append({"subtitle": line[4:].strip(), "lines": []})
        elif current is not None:
            if current["items"]:
                current["items"][-1]["lines"].append(line)
            else:
                current["body_lines"].append(line)

    if current:
        sections.append(current)
    return sections


def _md_inline(text: str) -> str:
    """インラインMarkdown（**太字**、*斜体*）をHTMLに変換"""
    import html as h
    text = h.escape(text)
    text = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', text)
    text = re.sub(r'\*(.+?)\*', r'<em>\1</em>', text)
    return text


def _find_serp_for_subtitle(subtitle: str, serp_data: dict) -> list | None:
    """処方箋の`### subtitle`に対応するSERPデータを検索する（近似マッチ）"""
    if not serp_data:
        return None
    # 鍵括弧を除去して比較
    clean = re.sub(r'[「」【】\d\.\s]', '', subtitle).lower()
    for kw, results in serp_data.items():
        kw_clean = re.sub(r'\s', '', kw).lower()
        if kw_clean in clean or clean in kw_clean:
            return results
    return None


def _render_rival_titles(serp_results: list) -> str:
    """ライバルタイトル比較セクションのHTMLを生成する（公式/ブランドを除外）"""
    rivals = [r for r in serp_results if r.get("site_type") != "公式/ブランド"]
    if not rivals:
        return '<p class="no-rival-note">※ 比較対象となる一般サイトが見つかりませんでした。</p>'
    items = ""
    for r in rivals[:7]:
        items += (
            f'<li>'
            f'<span class="rival-num-tag">{r["rank"]}位</span>'
            f'<span class="rival-title-text">{_md_inline(r["title"][:80])}</span>'
            f'<span class="rival-domain-note">（{r["domain"]}）</span>'
            f'</li>'
        )
    return f'<ul class="rival-titles-list">{items}</ul>'


def _render_rival_body(serp_results: list) -> str:
    """ライバル本文の例セクションのHTMLを生成する"""
    rivals = [r for r in serp_results
              if r.get("site_type") != "公式/ブランド" and r.get("h2_with_snippet")]
    if not rivals:
        return (
            '<div class="no-rival-note">'
            '※ 本文データが取得できませんでした（ページアクセス制限の可能性があります）'
            '</div>'
        )
    html_parts = []
    for r in rivals[:3]:
        h2_items = ""
        for h2_item in r["h2_with_snippet"][:5]:
            snippet_text = h2_item.get("snippet", "")
            snippet_html = (
                f'<p class="rival-snippet">{_md_inline(snippet_text[:200])}</p>'
                if snippet_text else ""
            )
            h2_items += (
                f'<div class="rival-topic">{_md_inline(h2_item["h2"][:60])}</div>'
                f'{snippet_html}'
            )
        html_parts.append(
            f'<div class="rival-block">'
            f'<div class="rival-domain">'
            f'<a href="{r["url"]}" target="_blank" rel="noopener">{r["domain"]}</a>'
            f'（{r["rank"]}位）'
            f'</div>'
            f'{h2_items}'
            f'</div>'
        )
    return "".join(html_parts)


def _build_html(analysis: dict, prescription: str, site_url: str, serp_data: dict | None = None) -> str:
    import html as html_mod

    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    serp_data = serp_data or {}

    dist_rows = "".join(
        f"<tr><td>{k}</td><td>{v}件</td></tr>"
        for k, v in analysis["pos_distribution"].items()
    )

    sections = _parse_prescription_sections(prescription)
    cards_html = ""
    for sec in sections:
        color = sec["color"]
        items_html = ""
        for item in sec["items"]:
            # 処方箋テキスト部分
            lines_html = "".join(
                f"<p>{_md_inline(l.lstrip('*- '))}</p>" if l.strip() else ""
                for l in item["lines"]
            )

            # このキーワードに対応するSERPデータを検索
            rival_results = _find_serp_for_subtitle(item["subtitle"], serp_data)

            if rival_results:
                # 2カラムレイアウト（改善案 ＋ ライバルタイトル比較）
                rival_titles_html = _render_rival_titles(rival_results)
                rival_body_html = _render_rival_body(rival_results)
                items_html += f"""
<div class="rx-item">
  <div class="rx-item-title">{_md_inline(item['subtitle'])}</div>
  <div class="rx-item-grid">
    <div class="rx-section">
      <div class="section-title">✅ 改善案</div>
      <div class="rx-item-body">{lines_html}</div>
    </div>
    <div class="rx-section">
      <div class="section-title">🔤 ライバルタイトル比較 <span class="filter-note">※公式サイト・店舗ページは除外</span></div>
      {rival_titles_html}
    </div>
  </div>
  <div class="rx-rival-body">
    <div class="section-title">📄 ライバル本文の例（参考）</div>
    {rival_body_html}
  </div>
</div>"""
            else:
                # SERPデータなし（既存スタイル）
                items_html += f"""
<div class="rx-item">
  <div class="rx-item-title">{_md_inline(item['subtitle'])}</div>
  <div class="rx-item-body">{lines_html}</div>
</div>"""

        body_html = "".join(
            f"<p>{_md_inline(l.lstrip('*- '))}</p>" if l.strip() else ""
            for l in sec["body_lines"]
        )
        cards_html += f"""
<div class="card">
  <div class="card-header" style="background:{color};">
    <span class="card-title">{html_mod.escape(sec['title'])}</span>
  </div>
  <div class="card-body">
    {body_html}
    {items_html}
  </div>
</div>"""

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SEO処方箋 詳細レポート</title>
<style>
  body {{
    font-family: "Hiragino Kaku Gothic ProN", "Meiryo", sans-serif;
    background: #f5f7fa;
    color: #222;
    margin: 0;
    padding: 20px;
    font-size: 14px;
    line-height: 1.7;
  }}
  h1 {{ font-size: 20px; color: #1a3a5c; border-bottom: 3px solid #2d7dd2; padding-bottom: 8px; margin-bottom: 6px; }}
  .subtitle {{ color: #666; font-size: 13px; margin-bottom: 24px; }}
  .summary-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin: 20px 0; }}
  .metric {{ background: #fff; border-radius: 8px; padding: 14px; text-align: center; box-shadow: 0 2px 6px rgba(0,0,0,0.08); }}
  .metric-value {{ font-size: 1.6em; font-weight: bold; color: #1a3a5c; }}
  .metric-label {{ font-size: 0.82em; color: #666; margin-top: 4px; }}
  .dist-table {{ border-collapse: collapse; margin: 12px 0 24px; font-size: 13px; }}
  .dist-table th {{ background: #2d7dd2; color: #fff; padding: 6px 16px; text-align: left; }}
  .dist-table td {{ border: 1px solid #ddd; padding: 5px 16px; }}
  .dist-table tr:nth-child(even) {{ background: #f5f8ff; }}
  .card {{ background: #fff; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.09); margin-bottom: 24px; overflow: hidden; }}
  .card-header {{ color: #fff; padding: 12px 20px; }}
  .card-title {{ font-size: 16px; font-weight: bold; }}
  .card-body {{ padding: 16px 20px; }}
  .card-body > p {{ margin: 6px 0; }}
  /* キーワードカード */
  .rx-item {{ margin-bottom: 20px; border-radius: 6px; border: 1px solid #dce6f4; overflow: hidden; }}
  .rx-item-title {{ font-size: 14px; font-weight: bold; color: #fff; background: #2d7dd2; margin: 0; padding: 9px 16px; }}
  /* SERPあり: 2カラムグリッド */
  .rx-item-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 0; }}
  .rx-section {{ padding: 14px 16px; }}
  .rx-section + .rx-section {{ border-left: 1px solid #e0eaf5; }}
  .rx-item-body p {{ margin: 4px 0; font-size: 13.5px; color: #333; }}
  /* SERPなし: 既存スタイル */
  .rx-item > .rx-item-body {{ padding: 12px 16px; background: #f8fbff; }}
  .rx-item > .rx-item-body p {{ margin: 4px 0; font-size: 13.5px; color: #333; }}
  /* セクションタイトル */
  .section-title {{ font-size: 13px; font-weight: bold; color: #2d7dd2; margin-bottom: 8px; padding-bottom: 4px; border-bottom: 1px solid #d0e4f7; letter-spacing: 0.03em; }}
  .filter-note {{ font-size: 11px; font-weight: normal; color: #999; margin-left: 6px; }}
  /* ライバルタイトルリスト */
  .rival-titles-list {{ list-style: none; padding: 0; margin: 0; }}
  .rival-titles-list li {{ margin-bottom: 6px; font-size: 13px; padding: 5px 8px; background: #f9f9f9; border: 1px solid #e8e8e8; border-radius: 4px; }}
  .rival-num-tag {{ display: inline-block; background: #2d7dd2; color: #fff; font-weight: bold; font-size: 11px; padding: 1px 7px; border-radius: 10px; margin-right: 6px; }}
  .rival-title-text {{ color: #334; }}
  .rival-domain-note {{ font-size: 11px; color: #888; margin-left: 4px; }}
  /* ライバル本文の例 */
  .rx-rival-body {{ padding: 14px 16px; background: #fdfaf3; border-top: 1px solid #e0eaf5; }}
  .rival-block {{ margin-bottom: 16px; }}
  .rival-domain {{ font-size: 11.5px; color: #888; margin-bottom: 6px; }}
  .rival-domain a {{ color: #2d7dd2; text-decoration: none; }}
  .rival-domain a:hover {{ text-decoration: underline; }}
  .rival-topic {{ font-size: 12px; font-weight: bold; background: #fff3cd; border-left: 3px solid #e6a817; padding: 4px 8px; margin-bottom: 4px; border-radius: 0 4px 4px 0; }}
  .rival-snippet {{ font-size: 11.5px; color: #555; background: #fffaf0; border-left: 2px solid #e6a817; padding: 4px 8px; margin-bottom: 6px; border-radius: 0 3px 3px 0; }}
  .no-rival-note {{ color: #999; font-size: 12px; font-style: italic; }}
  .footer {{ font-size: 0.8em; color: #999; margin-top: 40px; border-top: 1px solid #eee; padding-top: 12px; }}
  @media (max-width: 800px) {{
    .rx-item-grid {{ grid-template-columns: 1fr; }}
    .rx-section + .rx-section {{ border-left: none; border-top: 1px solid #e0eaf5; }}
    .summary-grid {{ grid-template-columns: repeat(2, 1fr); }}
  }}
</style>
</head>
<body>
<h1>📋 SEO処方箋 詳細レポート</h1>
<p class="subtitle">対象サイト：{html_mod.escape(site_url)} ／ 診断日時：{now}</p>

<div class="summary-grid">
  <div class="metric"><div class="metric-value">{analysis['total_kw']:,}</div><div class="metric-label">分析キーワード数</div></div>
  <div class="metric"><div class="metric-value">{analysis['avg_position']:.1f}位</div><div class="metric-label">平均掲載順位</div></div>
  <div class="metric"><div class="metric-value">{analysis['avg_ctr']:.1%}</div><div class="metric-label">全体CTR</div></div>
  <div class="metric"><div class="metric-value">{analysis['total_clicks']:,}</div><div class="metric-label">総クリック数</div></div>
</div>

<table class="dist-table">
  <tr><th>順位帯</th><th>件数</th></tr>{dist_rows}
</table>

{cards_html}

<div class="footer">
  ※ この処方箋はGemini AIが生成したものです。実施前にご自身で内容をご確認ください。
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

        # ── SERP スクレイピング ──
        serp_data: dict = {}
        target_kws = [kw["query"] for kw in analysis["top_keywords"][:10]]

        serp_progress = st.progress(0, text="🔍 ライバルサイトを調査中...")
        for i, kw in enumerate(target_kws):
            serp_progress.progress(
                (i + 1) / len(target_kws),
                text=f"🔍 ライバル調査中 ({i+1}/{len(target_kws)})：「{kw}」",
            )
            results = analyze_serp(kw)
            analyze_serp_pages(results, max_pages=5)
            serp_data[kw] = results
        serp_progress.empty()

        with st.spinner("💊 AIが処方箋を作成中です...（30秒〜2分かかります）"):
            prescription = generate_prescription(
                api_key=gemini_api_key,
                analysis=analysis,
                site_url=site_url or "（URL未入力）",
                serp_data=serp_data,
            )

        st.session_state["analysis"]     = analysis
        st.session_state["prescription"] = prescription
        st.session_state["site_url"]     = site_url or "（URL未入力）"
        st.session_state["serp_data"]    = serp_data

# ──────────────────────────────────────
# 診断結果の表示（session_stateから読む）
# ──────────────────────────────────────
if "prescription" in st.session_state:
    analysis     = st.session_state["analysis"]
    prescription = st.session_state["prescription"]
    display_url  = st.session_state["site_url"]
    serp_data    = st.session_state.get("serp_data", {})

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

    date_str  = datetime.now().strftime("%Y%m%d")
    safe_url  = re.sub(r'[^a-zA-Z0-9._-]', '_', display_url)[:40]
    base_name = f"SEO診断書_{safe_url}_{date_str}"

    excel_bytes = _build_excel(analysis, prescription, display_url)
    html_bytes  = _build_html(analysis, prescription, display_url, serp_data).encode("utf-8")

    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            label="📊 Excelでダウンロード（3シート）",
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

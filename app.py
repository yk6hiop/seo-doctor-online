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
from core.gemini_client import generate_keyword_prescriptions
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


def _h(text: str) -> str:
    """HTMLエスケープ"""
    import html
    return html.escape(str(text or ""))


# ──────────────────────────────────────
# HTML 出力
# ──────────────────────────────────────

_LEVEL_COLOR = {
    "Lv5": "#8b1a1a",
    "Lv4": "#7a3900",
    "Lv3": "#6b5a00",
    "Lv2": "#1a5c2d",
    "Lv1": "#3a4a55",
}
_DEFAULT_COLOR = "#1a3a5c"

_EVIDENCE_CLASS = {
    "title": "ai-evidence-title",
    "serp":  "ai-evidence-serp",
    "rank":  "ai-evidence-rank",
    "ctr":   "ai-evidence-ctr",
}
_EVIDENCE_TAG = {
    "title": "■ タイトル",
    "serp":  "▶ ライバル比較",
    "rank":  "★ 現在順位・CTR",
    "ctr":   "★ CTR",
}
_TODO_EMOJI = {
    "タイトル改善":  "🔤",
    "本文深掘り":    "📝",
    "UX改善":        "🎨",
    "内部リンク":    "🔗",
    "構造化データ":  "🏷️",
}


def _render_actions_html(actions: list[dict]) -> str:
    """改善案アクションリストをHTMLに変換する"""
    if not actions:
        return '<div class="ai-actions-raw">（改善案データなし）</div>'

    parts = ['<div class="ai-actions">']
    for act in actions:
        ev_type = act.get("evidence_type", "")
        css_cls = _EVIDENCE_CLASS.get(ev_type, "ai-evidence-other")
        ev_tag  = _EVIDENCE_TAG.get(ev_type, "💡 ポイント")
        marker  = act.get("marker", "◆")
        todo    = act.get("todo", "")
        evidence= act.get("evidence", "")
        action  = act.get("action", "")

        parts.append(f'<div class="ai-action-item {css_cls}">')
        if todo:
            parts.append(
                f'<div class="ai-action-category">'
                f'<span class="ai-action-marker">{_h(marker)}</span>'
                f'{_h(todo)}'
                f'</div>'
            )
        if evidence:
            parts.append(
                f'<div class="ai-evidence">'
                f'<span class="ai-ev-tag">{_h(ev_tag)}</span> '
                f'{_h(evidence)}'
                f'</div>'
            )
        if action:
            parts.append(f'<div class="ai-action">→ {_h(action)}</div>')
        parts.append('</div>')
    parts.append('</div>')
    return "".join(parts)


def _render_rival_titles_html(serp_results: list, own_position: float) -> str:
    """ライバルタイトル比較セクションを生成する（公式/ブランドを除外）"""
    rivals = [r for r in serp_results if r.get("site_type") != "公式/ブランド"]
    if not rivals:
        return '<p class="no-rival">比較対象となる一般サイトが見つかりませんでした。</p>'

    items = []
    for r in rivals[:7]:
        items.append(
            f'<li>'
            f'<span class="rival-num-tag">SERP{r["rank"]}位</span>'
            f'{_h(r["title"][:80])}'
            f'</li>'
        )

    pos_int = round(own_position)
    items.append(
        f'<li class="own-position-row">'
        f'<strong>あなたの記事: SERP{pos_int}位</strong>'
        f'</li>'
    )

    return f'<ul class="rival-titles-list">{"".join(items)}</ul>'


def _render_rival_body_html(serp_results: list) -> str:
    """ライバル本文の例セクションを生成する"""
    rivals = [r for r in serp_results
              if r.get("site_type") != "公式/ブランド" and r.get("h2_with_snippet")]
    if not rivals:
        return (
            '<div class="no-rival-unanalyzed">'
            '<strong>⚠️ ライバル本文データを取得できませんでした</strong><br>'
            'ページアクセス制限またはタイムアウトの可能性があります。'
            '</div>'
        )

    parts = []
    for r in rivals[:3]:
        h2_items_html = ""
        for h2_item in r["h2_with_snippet"][:5]:
            snippet = h2_item.get("snippet", "")
            snippet_html = (
                f'<div class="rival-summary">{_h(snippet[:200])}</div>'
                if snippet else ""
            )
            h2_items_html += (
                f'<div class="rival-topic">▼ 共通トピック「{_h(h2_item["h2"][:60])}」</div>'
                f'<div class="rival-note">このトピックに関する説明・比較が記事に不足している可能性があります。</div>'
                f'{snippet_html}'
            )
        parts.append(
            f'<div class="rival-block">'
            f'<div class="rival-domain">'
            f'<a href="{_h(r["url"])}" target="_blank" rel="noopener">{_h(r["domain"])}</a>'
            f'（SERP{r["rank"]}位）より'
            f'</div>'
            f'{h2_items_html}'
            f'</div>'
        )
    return "".join(parts)


def _build_html(
    analysis: dict,
    prescriptions: dict,
    serp_data: dict,
    site_url: str,
) -> str:
    """ローカル版と同じカード形式のHTMLレポートを生成する"""
    import html as html_mod

    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    keywords  = prescriptions.get("keywords", [])
    site_summary = prescriptions.get("site_summary", "")

    dist_rows = "".join(
        f"<tr><td>{k}</td><td>{v}件</td></tr>"
        for k, v in analysis["pos_distribution"].items()
    )

    # ── キーワードカードを生成 ──
    cards_html = ""
    for card_num, kw_rx in enumerate(keywords, 1):
        query     = kw_rx.get("query", "")
        level     = kw_rx.get("level", "Lv3")
        pri_label = kw_rx.get("priority_label", "🟢中長期で取り組む")
        actions   = kw_rx.get("actions", [])

        header_color = _LEVEL_COLOR.get(level, _DEFAULT_COLOR)

        # SCデータから順位・CTRを取得
        kw_sc = next(
            (k for k in analysis["top_keywords"] if k["query"] == query),
            None,
        )
        position    = kw_sc["position"] if kw_sc else 0.0
        impressions = kw_sc["impressions"] if kw_sc else 0
        clicks      = kw_sc["clicks"] if kw_sc else 0
        ctr_pct     = kw_sc["ctr_pct"] if kw_sc else "─"
        ctr_status  = kw_sc["ctr_status"] if kw_sc else "─"

        # 🎯 やること バッジ
        todos = list(dict.fromkeys(a.get("todo", "") for a in actions if a.get("todo")))
        todo_badges = "".join(
            f'<span class="todo-badge todo-{t.replace("改善","").replace("掘り","").replace("リンク","")}">'
            f'{_TODO_EMOJI.get(t, "📌")} {_h(t)}</span>'
            for t in todos
        )

        # SERPデータ
        serp_results = serp_data.get(query, [])

        actions_html       = _render_actions_html(actions)
        rival_titles_html  = _render_rival_titles_html(serp_results, position) if serp_results else '<p class="no-rival">（SERP未取得）</p>'
        rival_body_html    = _render_rival_body_html(serp_results) if serp_results else '<div class="no-rival-unanalyzed">（SERP未取得）</div>'

        cards_html += f"""
<div class="card">
  <div class="card-header" style="background:{header_color};">
    <span class="card-num">#{card_num}</span>
    <span class="card-kw">{_h(query)}</span>
    <span class="card-rank">現在順位: {position:.1f}位</span>
  </div>
  <div class="action-conclusion">
    <strong>🎯 やること：</strong>{todo_badges if todo_badges else '（分類なし）'}
    <span class="priority-badge">{_h(pri_label)}</span>
  </div>
  <div class="card-body">
    <div class="section">
      <div class="section-title">✅ 改善案（具体的なアクション）</div>
      {actions_html}
    </div>
    <div class="section">
      <div class="section-title">🔤 ライバルタイトル比較 <span class="filter-note">※公式サイト・店舗ページは除外</span></div>
      {rival_titles_html}
    </div>
  </div>
  <div class="rival-body-section">
    <div class="section-title">📄 ライバル本文の例（参考）</div>
    {rival_body_html}
  </div>
  <div class="meta-row">
    <span class="meta-item"><strong>表示回数</strong> {impressions:,}回</span>
    <span class="meta-item"><strong>クリック数</strong> {clicks:,}回</span>
    <span class="meta-item"><strong>CTR</strong> {ctr_pct}</span>
    <span class="meta-item"><strong>サイト平均比</strong> {_h(ctr_status)}</span>
  </div>
</div>"""

    summary_html = (
        f'<div class="site-summary"><h2>💡 サイト全体への提言</h2><p>{_h(site_summary)}</p></div>'
        if site_summary else ""
    )

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SEO処方箋 詳細レポート</title>
<style>
  body {{ font-family: "Hiragino Kaku Gothic ProN","Meiryo",sans-serif; background:#f5f7fa; color:#222; margin:0; padding:20px; font-size:14px; line-height:1.7; }}
  h1 {{ font-size:20px; color:#1a3a5c; border-bottom:3px solid #2d7dd2; padding-bottom:8px; margin-bottom:6px; }}
  h2 {{ font-size:16px; color:#1a3a5c; margin-top:0; }}
  .subtitle {{ color:#666; font-size:13px; margin-bottom:24px; }}
  .summary-grid {{ display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin:20px 0; }}
  .metric {{ background:#fff; border-radius:8px; padding:14px; text-align:center; box-shadow:0 2px 6px rgba(0,0,0,.08); }}
  .metric-value {{ font-size:1.6em; font-weight:bold; color:#1a3a5c; }}
  .metric-label {{ font-size:.82em; color:#666; margin-top:4px; }}
  .dist-table {{ border-collapse:collapse; margin:12px 0 24px; font-size:13px; }}
  .dist-table th {{ background:#2d7dd2; color:#fff; padding:6px 16px; text-align:left; }}
  .dist-table td {{ border:1px solid #ddd; padding:5px 16px; }}
  .dist-table tr:nth-child(even) {{ background:#f5f8ff; }}
  .site-summary {{ background:#fff; border-radius:8px; padding:16px 20px; margin-bottom:24px; box-shadow:0 2px 6px rgba(0,0,0,.08); border-left:4px solid #1a3a5c; }}
  /* カード */
  .card {{ background:#fff; border-radius:8px; box-shadow:0 2px 8px rgba(0,0,0,.09); margin-bottom:28px; overflow:hidden; }}
  .card-header {{ color:#fff; padding:12px 20px; display:flex; align-items:baseline; gap:12px; flex-wrap:wrap; }}
  .card-num {{ font-size:22px; font-weight:bold; opacity:.7; min-width:36px; }}
  .card-kw {{ font-size:16px; font-weight:bold; flex:1; }}
  .card-rank {{ font-size:12px; opacity:.85; background:rgba(255,255,255,.18); padding:2px 8px; border-radius:12px; white-space:nowrap; }}
  /* やること */
  .action-conclusion {{ background:#fff8dc; border-bottom:1px solid #e6d97a; padding:10px 20px; font-size:13px; color:#5a4a00; display:flex; align-items:center; gap:8px; flex-wrap:wrap; }}
  .todo-badge {{ display:inline-block; padding:2px 10px; border-radius:12px; font-size:12px; font-weight:bold; background:#e6e6e6; color:#333; margin-right:4px; }}
  .priority-badge {{ margin-left:auto; font-size:12px; color:#888; }}
  /* ボディ2カラム */
  .card-body {{ display:grid; grid-template-columns:1fr 1fr; gap:0; }}
  .section {{ padding:18px 20px; }}
  .section+.section {{ border-left:1px solid #eee; }}
  .section-title {{ font-size:13px; font-weight:bold; color:#2d7dd2; margin-bottom:10px; padding-bottom:4px; border-bottom:1px solid #d0e4f7; letter-spacing:.03em; }}
  .filter-note {{ font-size:11px; font-weight:normal; color:#999; margin-left:6px; }}
  /* 改善案 */
  .ai-actions {{ margin:4px 0 0; font-size:13.5px; line-height:1.7; }}
  .ai-action-item {{ margin-bottom:10px; padding:10px 12px; border-radius:6px; border-left:4px solid #ccc; background:#fafbfc; }}
  .ai-action-item.ai-evidence-serp  {{ border-left-color:#d97706; background:#fdf6ec; }}
  .ai-action-item.ai-evidence-rank  {{ border-left-color:#7a3aa6; background:#f7f1fa; }}
  .ai-action-item.ai-evidence-title {{ border-left-color:#7c8400; background:#f8f8e8; }}
  .ai-action-item.ai-evidence-ctr   {{ border-left-color:#c05c00; background:#fff3e8; }}
  .ai-action-item.ai-evidence-other {{ border-left-color:#4a7a9b; background:#f0f7fc; }}
  .ai-action-category {{ font-size:14px; font-weight:700; color:#1a3a5c; margin-bottom:6px; padding-bottom:4px; border-bottom:2px solid #d0e4f7; }}
  .ai-action-marker {{ display:inline-block; margin-right:6px; font-size:16px; color:#c75c00; background:#fff8e0; border:2px solid #e6a817; border-radius:50%; width:24px; height:24px; text-align:center; line-height:20px; font-weight:bold; vertical-align:middle; }}
  .ai-ev-tag {{ display:inline-block; margin-right:6px; padding:1px 8px; border-radius:3px; background:#fff; border:1px solid currentColor; font-size:11.5px; font-weight:600; line-height:1.5; }}
  .ai-evidence {{ color:#2d3a4a; margin-bottom:5px; font-weight:500; }}
  .ai-action {{ color:#1a4a3a; padding-left:4px; }}
  .ai-actions-raw {{ white-space:pre-wrap; font-size:12.5px; background:#f7f7f7; padding:10px; border-radius:4px; }}
  /* ライバルタイトル */
  .rival-titles-list {{ list-style:none; padding:0; margin:0; }}
  .rival-titles-list li {{ margin-bottom:5px; font-size:12.5px; padding:4px 8px; background:#f9f9f9; border:1px solid #e8e8e8; border-radius:4px; }}
  .rival-titles-list li.own-position-row {{ margin-top:8px; padding-top:8px; border-top:1px dashed #c0c0c0; background:#fff8c4; border-color:#e6a817; }}
  .rival-num-tag {{ display:inline-block; background:#2d7dd2; color:#fff; font-weight:bold; font-size:11px; padding:1px 6px; border-radius:10px; margin-right:5px; letter-spacing:.03em; }}
  /* ライバル本文 */
  .rival-body-section {{ padding:18px 20px; background:#fdfaf3; border-top:1px solid #eee; }}
  .rival-block {{ margin-bottom:16px; border-left:3px solid #e6a817; padding-left:12px; }}
  .rival-domain {{ font-size:11px; color:#888; margin-bottom:6px; }}
  .rival-domain a {{ color:#2d7dd2; text-decoration:none; }}
  .rival-domain a:hover {{ text-decoration:underline; }}
  .rival-topic {{ font-size:12px; font-weight:bold; background:#fff3cd; padding:4px 8px; margin-bottom:3px; border-radius:0 4px 4px 0; }}
  .rival-note {{ font-size:11.5px; color:#7a5a00; font-style:italic; margin-bottom:4px; padding-left:4px; }}
  .rival-summary {{ font-size:11.5px; color:#555; background:#fffaf0; border-left:2px solid #e6a817; padding:4px 8px; margin-bottom:8px; border-radius:0 3px 3px 0; }}
  .no-rival {{ color:#999; font-size:12px; font-style:italic; padding:8px 0; }}
  .no-rival-unanalyzed {{ color:#7a5c00; background:#fffbe6; border:1px solid #f0d060; border-radius:6px; padding:12px 14px; font-size:12.5px; line-height:1.8; }}
  .no-rival-unanalyzed strong {{ display:block; margin-bottom:4px; font-size:13px; }}
  /* メタ行 */
  .meta-row {{ background:#f5f7fa; border-top:1px solid #eee; padding:10px 20px; font-size:12px; color:#666; display:flex; gap:20px; flex-wrap:wrap; }}
  .meta-item strong {{ color:#444; margin-right:4px; }}
  /* フッター */
  .footer {{ font-size:.8em; color:#999; margin-top:40px; border-top:1px solid #eee; padding-top:12px; }}
  @media(max-width:800px) {{
    .card-body {{ grid-template-columns:1fr; }}
    .section+.section {{ border-left:none; border-top:1px solid #eee; }}
    .summary-grid {{ grid-template-columns:repeat(2,1fr); }}
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

{summary_html}

{cards_html}

<div class="footer">
  ※ この処方箋はGemini AIが生成したものです。実施前にご自身で内容をご確認ください。
</div>
</body>
</html>"""


# ──────────────────────────────────────
# Excel 出力
# ──────────────────────────────────────

def _build_excel(
    analysis: dict,
    prescriptions: dict,
    serp_data: dict,
    site_url: str,
) -> bytes:
    wb = Workbook()
    thin = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin"),
    )
    wrap_align = Alignment(vertical="top", wrap_text=True)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    keywords = prescriptions.get("keywords", [])

    _LEVEL_FILL = {
        "Lv5": PatternFill(start_color="FFD0D0", end_color="FFD0D0", fill_type="solid"),
        "Lv4": PatternFill(start_color="FFE4C8", end_color="FFE4C8", fill_type="solid"),
        "Lv3": PatternFill(start_color="FFF9C8", end_color="FFF9C8", fill_type="solid"),
    }
    _HDR_FONT  = Font(bold=True, color="FFFFFF", size=10)
    _HDR_FILL  = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    _HDR_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def _write_header(ws, headers_widths):
        for col, (hdr, w) in enumerate(headers_widths, start=1):
            c = ws.cell(row=1, column=col, value=hdr)
            c.font = _HDR_FONT
            c.fill = _HDR_FILL
            c.alignment = _HDR_ALIGN
            c.border = thin
            ws.column_dimensions[get_column_letter(col)].width = w
        ws.row_dimensions[1].height = 28

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

    info_rows = [
        ("■ 基本情報", ""),
        ("対象サイト", site_url),
        ("実行日時", now_str),
        ("", ""),
        ("■ パフォーマンス概況", ""),
        ("分析キーワード数", f"{analysis['total_kw']:,} 件"),
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
        "  各キーワードに対して「レベル」「やること」「改善案」「ライバルタイトル比較」「ライバル本文H2」が入っています。",
        "  「着手状況」列を「対応中」→「対応済み」に更新しながら進めてください。",
        "",
        "【Google スプレッドシートで開く方法】",
        "  1. Google ドライブを開く（drive.google.com）",
        "  2. このファイルをドラッグ＆ドロップでアップロード",
        "  3. ダブルクリック → 「Googleスプレッドシートで開く」",
        "",
        f"※ 生成日時: {now_str}",
    ]
    for i, line in enumerate(howto_lines, start=1):
        ws_howto[f"A{i}"] = line
        if line.startswith("■") or line.startswith("【"):
            ws_howto[f"A{i}"].font = Font(bold=True)
    ws_howto["A1"].font = Font(bold=True, size=13)

    # ════════════════════════════════
    # シート③：処方箋（キーワード別）
    # ════════════════════════════════
    ws_rx = wb.create_sheet("📋処方箋")
    headers_widths = [
        ("着手状況",             10),
        ("レベル",               8),
        ("優先度",               18),
        ("クエリ（キーワード）", 28),
        ("現在順位",             9),
        ("表示回数",             9),
        ("クリック数",           9),
        ("CTR",                  8),
        ("サイト平均比",         12),
        ("やること",             22),
        ("改善案（具体的なアクション）", 55),
        ("ライバルタイトル比較", 45),
        ("ライバル本文の例（H2）", 45),
    ]
    _write_header(ws_rx, headers_widths)

    data_row = 2
    for kw_rx in keywords:
        query     = kw_rx.get("query", "")
        level     = kw_rx.get("level", "Lv3")
        pri_label = kw_rx.get("priority_label", "")
        actions   = kw_rx.get("actions", [])

        kw_sc = next(
            (k for k in analysis["top_keywords"] if k["query"] == query),
            None,
        )
        position    = kw_sc["position"] if kw_sc else 0.0
        impressions = kw_sc["impressions"] if kw_sc else 0
        clicks      = kw_sc["clicks"] if kw_sc else 0
        ctr_pct     = kw_sc["ctr_pct"] if kw_sc else ""
        ctr_status  = kw_sc["ctr_status"] if kw_sc else ""

        # やること（ユニークなtodo値を列挙）
        todos   = list(dict.fromkeys(a.get("todo", "") for a in actions if a.get("todo")))
        todo_txt = " / ".join(todos)

        # 改善案テキスト
        action_lines = []
        for a in actions:
            mk   = a.get("marker", "◆")
            todo = a.get("todo", "")
            ev   = a.get("evidence", "")
            act  = a.get("action", "")
            action_lines.append(f"◆ {todo} {mk}")
            if ev:
                action_lines.append(f"  根拠: {ev}")
            if act:
                action_lines.append(f"  → {act}")
        action_txt = "\n".join(action_lines)

        # ライバルタイトル
        serp_results = serp_data.get(query, [])
        rival_general = [r for r in serp_results if r.get("site_type") != "公式/ブランド"]
        rival_title_txt = "\n".join(
            f"SERP{r['rank']}位: {r['title'][:70]}"
            for r in rival_general[:5]
        )

        # ライバル本文H2
        rival_h2_lines = []
        for r in rival_general[:3]:
            for h2_item in r.get("h2_with_snippet", [])[:4]:
                rival_h2_lines.append(f"[{r['rank']}位] {h2_item['h2'][:40]}")
        rival_h2_txt = "\n".join(rival_h2_lines)

        row_fill = _LEVEL_FILL.get(level)
        row_values = [
            "未着手", level, pri_label, query,
            position, impressions, clicks, ctr_pct, ctr_status,
            todo_txt, action_txt, rival_title_txt, rival_h2_txt,
        ]
        for col, val in enumerate(row_values, start=1):
            c = ws_rx.cell(row=data_row, column=col, value=val)
            c.alignment = wrap_align
            c.border = thin
            if row_fill:
                c.fill = row_fill

        ws_rx.row_dimensions[data_row].height = max(20, len(action_txt.splitlines()) * 14)
        data_row += 1

    ws_rx.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


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

    with st.expander("🔑 APIキーの取得方法（Gmailがあれば無料）", expanded=not bool(gemini_api_key)):
        st.markdown("""
**Gmailアカウント（Googleアカウント）があれば無料で取得できます。**

**取得手順（約3分）**

1. **[Google AI Studio](https://aistudio.google.com/app/apikey)** を開く
2. Googleアカウントでログイン
3. 左メニューまたは画面中央の **「APIキーを作成」** をクリック
4. 「新しいプロジェクトでAPIキーを作成」を選択
5. 生成されたキー（`AIza`で始まる文字列）を **今すぐコピーして保存**
6. 上の入力欄に貼り付ける

---

⚠️ **重要：キーは作成時の1回しか表示されません**

キーを閉じた後は同じキーを再表示することができません。
メモ帳やパスワードマネージャーに保存してください。
**紛失した場合は同じ手順で新しいキーを再発行できます**（古いキーは削除してください）。

---

**無料枠の目安（2026年5月時点）**
- 1日 1,500リクエストまで
- 1分 15リクエストまで
- 診断1回 ≒ 1リクエスト → **1日に何度でも使用可能**
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

if not gemini_api_key:
    st.info(
        "**はじめての方へ** ── Gmailアカウントがあれば **無料** で使えます。\n\n"
        "左のサイドバー「🔑 APIキーの取得方法」を開いて、Gemini APIキーを取得してください。"
        "取得は約3分、完全無料です。",
        icon="👈",
    )

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
            results = analyze_serp(kw, timeout=10)
            analyze_serp_pages(results, max_pages=3, timeout=5)
            serp_data[kw] = results
        serp_progress.empty()

        # ── Gemini JSON処方箋生成 ──
        with st.spinner("💊 AIが処方箋を作成中です...（30秒〜2分かかります）"):
            prescriptions = generate_keyword_prescriptions(
                api_key    = gemini_api_key,
                top_keywords = analysis["top_keywords"][:10],
                serp_data  = serp_data,
                avg_ctr    = analysis["avg_ctr"],
                site_url   = site_url or "（URL未入力）",
            )

        if prescriptions.get("error"):
            err = prescriptions["error"]
            if "APIキー" in err or "INVALID" in err.upper() or "PERMISSION" in err.upper():
                st.error("⚠️ APIキーが無効です。Google AI Studioでキーを確認してください。")
            elif "クォータ" in err:
                st.error("⚠️ APIの利用制限に達しました。しばらく待ってから再試行してください。")
            else:
                st.error(f"⚠️ 処方箋の生成に失敗しました: {err}")
            st.stop()

        if not prescriptions.get("keywords"):
            st.error("⚠️ AIからの応答を解析できませんでした。もう一度お試しください。")
            st.stop()

        st.session_state["analysis"]      = analysis
        st.session_state["prescriptions"] = prescriptions
        st.session_state["serp_data"]     = serp_data
        st.session_state["site_url"]      = site_url or "（URL未入力）"

# ──────────────────────────────────────
# 診断結果の表示（session_stateから読む）
# ──────────────────────────────────────
if "prescriptions" in st.session_state:
    analysis      = st.session_state["analysis"]
    prescriptions = st.session_state["prescriptions"]
    display_url   = st.session_state["site_url"]
    serp_data     = st.session_state.get("serp_data", {})

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

    if prescriptions.get("site_summary"):
        st.info(f"💡 **サイト全体への提言：** {prescriptions['site_summary']}")

    st.divider()
    st.subheader("💊 改善処方箋（キーワード別）")

    for kw_rx in prescriptions.get("keywords", []):
        query = kw_rx.get("query", "")
        level = kw_rx.get("level", "Lv3")
        todos = list(dict.fromkeys(a.get("todo", "") for a in kw_rx.get("actions", []) if a.get("todo")))
        kw_sc = next((k for k in analysis["top_keywords"] if k["query"] == query), None)
        pos = kw_sc["position"] if kw_sc else 0

        with st.expander(
            f"{kw_rx.get('priority_label','')[:2]} 「{query}」 — {level} / {pos:.1f}位 / やること: {' / '.join(todos)}",
            expanded=(level == "Lv5"),
        ):
            for act in kw_rx.get("actions", []):
                st.markdown(
                    f"**{act.get('marker','')} {act.get('todo','')}**  \n"
                    f"根拠: {act.get('evidence','')}  \n"
                    f"→ {act.get('action','')}"
                )
            serp_results = serp_data.get(query, [])
            rivals = [r for r in serp_results if r.get("site_type") != "公式/ブランド"]
            if rivals:
                st.markdown("**🔤 ライバルタイトル比較**")
                for r in rivals[:5]:
                    st.markdown(f"- SERP{r['rank']}位: {r['title'][:80]}")

    st.divider()

    date_str  = datetime.now().strftime("%Y%m%d")
    safe_url  = re.sub(r'[^a-zA-Z0-9._-]', '_', display_url)[:40]
    base_name = f"SEO診断書_{safe_url}_{date_str}"

    excel_bytes = _build_excel(analysis, prescriptions, serp_data, display_url)
    html_bytes  = _build_html(analysis, prescriptions, serp_data, display_url).encode("utf-8")

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

# -*- coding: utf-8 -*-
"""
SEO診断ツール — Streamlit Webアプリ

Search ConsoleのExcel/CSVをアップロードし、
ユーザー自身のGemini APIキーで改善処方箋を生成する。
"""

import logging
import streamlit as st

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

    # APIキー取得ガイド
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

    # プライバシー説明（重要）
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

    # エクスポート手順
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

def _extract_url_from_filename(name: str) -> str:
    """SCエクスポートのファイル名からサイトURLを抽出する"""
    import re
    stem = re.sub(r'\.(xlsx|csv)$', '', name, flags=re.IGNORECASE)
    # ファイル名形式: https___example.com_-Performance-on-Search-...
    # "_-" がドメインと残りの区切り
    m = re.match(r'^(https?___[^_]+(?:_[^-][^_]*)*?)(?=_-|$)', stem)
    if m:
        return m.group(1).replace('___', '://')
    return ""

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

        st.download_button(
            label="📄 処方箋をダウンロード（Markdownファイル）",
            data=prescription,
            file_name="seo_prescription.md",
            mime="text/markdown",
            use_container_width=True,
        )

        st.caption(
            "※ この処方箋はGemini AIが生成したものです。"
            "実施前にご自身で内容をご確認ください。"
        )

# ──────────────────────────────────────
# 未入力時のガイド
# ──────────────────────────────────────
elif uploaded_file is None and not gemini_api_key:
    st.info("👈 左のサイドバーからGemini APIキーを入力し、ファイルをアップロードしてください")

elif uploaded_file is None:
    st.info("📂 Search ConsoleのファイルをアップロードしてType（Excel または CSV）")

elif not gemini_api_key:
    st.warning("🔑 サイドバーからGemini APIキーを入力してください（無料で取得できます）")

# -*- coding: utf-8 -*-
"""
Gemini API クライアント

ユーザーが持参したAPIキーを使って処方箋テキストを生成する。
セキュリティ:
  - APIキーはリクエスト時のみ使用。ログに書き出さない。
  - プロンプトの長さ上限を設ける（巨大プロンプトによる誤作動防止）。
"""

import logging
import re
from google import genai

logger = logging.getLogger(__name__)


def _sanitize_url(url: str) -> str:
    """サイトURLをサニタイズ（プロンプトインジェクション対策）"""
    url = url.strip()[:200]
    url = re.sub(r"[\x00-\x1f\x7f\n\r]", "", url)
    # httpsかhttpで始まらないURLは安全のためマスク
    if url and not re.match(r"^https?://", url, re.IGNORECASE):
        return "（URL省略）"
    return url or "（URL未入力）"

_MODEL          = "gemini-2.0-flash"
_MAX_PROMPT_LEN = 20_000   # プロンプト文字数上限


def _redact_key(text: str, api_key: str) -> str:
    """エラーメッセージからAPIキーを除去する"""
    if api_key and len(api_key) > 8:
        return text.replace(api_key, "AIza***[REDACTED]***")
    return text


def _build_prompt(analysis: dict, site_url: str) -> str:
    """Gemini送信用プロンプトを構築する"""
    dist = analysis["pos_distribution"]
    dist_text = "　".join(f"{k} {v}件" for k, v in dist.items())

    kw_lines = []
    for i, kw in enumerate(analysis["top_keywords"], 1):
        kw_lines.append(
            f"{i:2}. 「{kw['query']}」"
            f"　{kw['position']}位"
            f"　表示{kw['impressions']:,}回"
            f"　クリック{kw['clicks']}回"
            f"　CTR {kw['ctr_pct']}（{kw['ctr_status']}）"
        )

    prompt = f"""あなたは日本の上位SEOコンサルタントです。
以下のSearch Consoleデータを分析し、サイトオーナーが今すぐ実行できる具体的な改善処方箋を作成してください。

【サイト情報】
URL: {site_url}
分析キーワード数: {analysis['total_kw']:,}件
総クリック数: {analysis['total_clicks']:,}
総表示回数: {analysis['total_impressions']:,}
全体CTR: {analysis['avg_ctr']:.1%}
加重平均掲載順位: {analysis['avg_position']:.1f}位
順位帯分布: {dist_text}

【改善余地の大きいキーワード（優先度順・上位{len(kw_lines)}件）】
{chr(10).join(kw_lines)}

---

以下の形式で処方箋を出力してください。
- 各提案は「キーワード名」を必ず明記すること
- 「現状の課題」と「具体的な改善アクション」を含めること
- 専門用語は最小限にし、サイトオーナーが自分で実行できる粒度で書くこと
- 数字や根拠を使って説得力を持たせること

## 🔴 今すぐ取り組む（今週中）
（表示回数は多いがCTRが低いキーワード、あと一歩で1ページ目のキーワードを中心に3〜5件）

## 🟡 今月中に取り組む
（コンテンツ強化・内部リンク改善が効くキーワードを3〜5件）

## 🟢 中長期で取り組む
（現在圏外〜2ページ目だが将来性のあるキーワードを2〜3件）

## 💡 サイト全体への提言
（データ全体から読み取れるサイトの課題と方向性を2〜3点）
"""

    # プロンプト長の安全チェック
    if len(prompt) > _MAX_PROMPT_LEN:
        logger.warning("プロンプトが上限を超えたため切り詰めます")
        prompt = prompt[:_MAX_PROMPT_LEN]

    return prompt


def generate_prescription(api_key: str, analysis: dict, site_url: str) -> str:
    """
    Gemini APIで処方箋テキストを生成する。

    Args:
        api_key: ユーザーのGemini APIキー（ログに絶対書かない）
        analysis: analyze_sc_data()の戻り値
        site_url: サイトURL（プロンプト表示用）

    Returns:
        Markdownフォーマットの処方箋テキスト
    """
    try:
        client = genai.Client(api_key=api_key)
        prompt = _build_prompt(analysis, _sanitize_url(site_url))
        response = client.models.generate_content(
            model=_MODEL,
            contents=prompt,
        )
        return response.text

    except Exception as e:
        raw_msg = str(e)
        safe_msg = _redact_key(raw_msg, api_key)   # APIキーをマスク
        logger.error("Gemini API エラー: %s", safe_msg)

        msg = raw_msg.upper()
        if "API_KEY" in msg or "INVALID" in msg or "PERMISSION" in msg:
            return (
                "⚠️ **APIキーが無効です**\n\n"
                "Gemini APIキーを確認してください。\n"
                "👉 [Google AI Studio](https://aistudio.google.com/) でキーを確認・再発行できます。"
            )
        if "QUOTA" in msg or "RATE" in msg or "RESOURCE_EXHAUSTED" in msg:
            return (
                "⚠️ **APIの利用制限に達しました**\n\n"
                "しばらく待ってから再度お試しください。\n"
                "無料枠の上限: 1日1,500リクエスト / 1分15リクエスト"
            )
        return (
            f"⚠️ **処方箋の生成に失敗しました**\n\n"
            f"エラー内容: {safe_msg}\n\n"
            "APIキーが正しいか、インターネット接続を確認してください。"
        )

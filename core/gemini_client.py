# -*- coding: utf-8 -*-
"""
Gemini API クライアント

ユーザーが持参したAPIキーを使って処方箋テキストを生成する。
セキュリティ:
  - APIキーはリクエスト時のみ使用。ログに書き出さない。
  - プロンプトの長さ上限を設ける（巨大プロンプトによる誤作動防止）。
"""

import json
import logging
import re
from google import genai

logger = logging.getLogger(__name__)


def _sanitize_url(url: str) -> str:
    """サイトURLをサニタイズ（プロンプトインジェクション対策）"""
    url = url.strip()[:200]
    url = re.sub(r"[\x00-\x1f\x7f\n\r]", "", url)
    if url and not re.match(r"^https?://", url, re.IGNORECASE):
        return "（URL省略）"
    return url or "（URL未入力）"


_MODEL          = "gemini-3-flash-preview"
_MAX_PROMPT_LEN = 25_000


def _redact_key(text: str, api_key: str) -> str:
    """エラーメッセージからAPIキーを除去する"""
    if api_key and len(api_key) > 8:
        return text.replace(api_key, "AIza***[REDACTED]***")
    return text


def _handle_api_error(e: Exception, api_key: str) -> str:
    """API例外を日本語エラーメッセージに変換する"""
    raw_msg = str(e)
    safe_msg = _redact_key(raw_msg, api_key)
    logger.error("Gemini API エラー: %s", safe_msg)
    msg = raw_msg.upper()
    if "API_KEY" in msg or "INVALID" in msg or "PERMISSION" in msg:
        return "APIキーエラー"
    if "QUOTA" in msg or "RATE" in msg or "RESOURCE_EXHAUSTED" in msg:
        return "クォータ超過"
    return f"エラー: {safe_msg}"


# ──────────────────────────────────────
# キーワード別構造化JSON処方箋
# ──────────────────────────────────────

def _build_structured_prompt(
    top_keywords: list[dict],
    serp_data: dict,
    avg_ctr: float,
    site_url: str,
) -> str:
    """キーワードごとの構造化JSON改善提案プロンプトを構築する"""

    kw_blocks = []
    for i, kw in enumerate(top_keywords, 1):
        q = kw["query"]
        rivals = serp_data.get(q, [])
        rival_general = [r for r in rivals if r.get("site_type") != "公式/ブランド"]

        rival_titles_lines = "\n   ".join(
            f"SERP{r['rank']}位: {r['title'][:70]}"
            for r in rival_general[:5]
        ) or "（データなし）"

        all_h2s = []
        for r in rival_general[:3]:
            for h2_item in r.get("h2_with_snippet", [])[:4]:
                all_h2s.append(h2_item["h2"][:40])
        h2_text = "、".join(all_h2s[:6]) if all_h2s else "（データなし）"

        kw_blocks.append(
            f"{i}. 「{q}」\n"
            f"   順位:{kw['position']}位 / 表示:{kw['impressions']:,}回 / "
            f"クリック:{kw['clicks']}回 / CTR:{kw['ctr_pct']}（サイト平均比:{kw['ctr_status']}）\n"
            f"   ライバル上位タイトル:\n   {rival_titles_lines}\n"
            f"   ライバルH2見出し例: {h2_text}"
        )

    kw_text = "\n\n".join(kw_blocks)

    prompt = f"""あなたは日本の上位SEOコンサルタントです。
以下のSearch ConsoleデータとSERP分析データを元に、各キーワードの改善提案をJSON形式で出力してください。

【絶対遵守】
- JSONのみを返すこと
- ```json などのコードブロック記号は使わない
- 前置き・後書き・説明文は一切不要
- 各キーワードにつき2〜4個のアクションを提案すること

【出力スキーマ（厳守）】
{{
  "site_summary": "サイト全体の課題と方向性（2〜3文）",
  "keywords": [
    {{
      "query": "キーワード名（上のデータと完全一致）",
      "level": "Lv5",
      "priority_label": "🔴今すぐ取り組む",
      "actions": [
        {{
          "todo": "タイトル改善",
          "marker": "★",
          "evidence_type": "title",
          "evidence": "根拠（50字以内・具体的数値含む）",
          "action": "具体的な改善アクション（100字以内）"
        }}
      ]
    }}
  ]
}}

【フィールド定義】
- level: Lv5=最優先（今週中）/ Lv4=高（今月中）/ Lv3=中（中長期）
- priority_label: "🔴今すぐ取り組む" | "🟡今月中に取り組む" | "🟢中長期で取り組む"
- todo: 「タイトル改善」「本文深掘り」「UX改善」「内部リンク」「構造化データ」のいずれか
- evidence_type: "title"（タイトル起因）| "serp"（SERP差分）| "rank"（順位起因）| "ctr"（CTR起因）
- marker: "★"（最重要）| "●"（重要）| "◆"（推奨）

【サイト情報】
URL: {site_url}
サイト全体CTR: {avg_ctr:.1%}

【キーワードデータ（優先度順）】
{kw_text}
"""

    if len(prompt) > _MAX_PROMPT_LEN:
        logger.warning("構造化プロンプトが上限を超えたため切り詰めます")
        prompt = prompt[:_MAX_PROMPT_LEN]

    return prompt


def generate_keyword_prescriptions(
    api_key: str,
    top_keywords: list[dict],
    serp_data: dict,
    avg_ctr: float,
    site_url: str,
) -> dict:
    """
    各キーワードの改善提案をJSON形式で生成する。

    Returns:
        {
            "site_summary": str,
            "keywords": [
                {
                    "query": str,
                    "level": str,
                    "priority_label": str,
                    "actions": [{todo, marker, evidence_type, evidence, action}]
                }, ...
            ],
            "error": str | None,
        }
    """
    empty = {"site_summary": "", "keywords": [], "error": None}

    try:
        client = genai.Client(api_key=api_key)
        prompt = _build_structured_prompt(top_keywords, serp_data, avg_ctr, _sanitize_url(site_url))
        response = client.models.generate_content(model=_MODEL, contents=prompt)
        raw = (response.text or "").strip()

        # コードブロックを除去
        raw = re.sub(r'^```(?:json)?\s*', '', raw)
        raw = re.sub(r'\s*```\s*$', '', raw)
        raw = raw.strip()

        data = json.loads(raw)

        if isinstance(data, dict):
            return {
                "site_summary": data.get("site_summary", ""),
                "keywords": data.get("keywords", []),
                "error": None,
            }
        if isinstance(data, list):
            return {"site_summary": "", "keywords": data, "error": None}

        empty["error"] = "予期しないJSON形式"
        return empty

    except json.JSONDecodeError as e:
        logger.error("JSON解析エラー（生テキスト）: %s", e)
        empty["error"] = f"JSON解析失敗: {e}"
        return empty

    except Exception as e:
        err = _handle_api_error(e, api_key)
        empty["error"] = err
        return empty

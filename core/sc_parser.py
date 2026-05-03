# -*- coding: utf-8 -*-
"""
Search Console ファイルパーサー

対応形式:
  - Excel (.xlsx): 「Excelとしてダウンロード」→ クエリシートを自動読込
  - CSV   (.csv):  「CSVをダウンロード」→ UTF-8 BOM 対応

セキュリティ:
  - ファイルサイズ上限: 20MB
  - クエリ文字列のサニタイズ（プロンプトインジェクション対策）
"""

import io
import logging
import re

import pandas as pd

logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────────
# 定数
# ─────────────────────────────────────────────────
MAX_FILE_SIZE_BYTES = 20 * 1024 * 1024   # 20 MB
MAX_QUERY_LENGTH   = 120                  # クエリ文字数上限（サニタイズ用）

# SC エクスポートのカラム名マッピング（日本語 / 英語 / Excel版）
_COLUMN_MAP: dict[str, str] = {
    # Excel版（「Excelとしてダウンロード」）
    "上位のクエリ": "query",
    "上位のページ": "query",   # ページシートを誤って使った場合も受け付ける
    # CSV版（「CSVをダウンロード」）
    "クエリ":       "query",
    # 共通
    "クリック数":   "clicks",
    "表示回数":     "impressions",
    "CTR":          "ctr",
    "掲載順位":     "position",
    # 英語版
    "Top queries":  "query",
    "Top pages":    "query",
    "Clicks":       "clicks",
    "Impressions":  "impressions",
    "Position":     "position",
}

# プロンプトインジェクションになりやすいパターン
_INJECTION_PATTERN = re.compile(
    r"(ignore\s+(previous|above)|system\s*prompt|you\s+are\s+(now|an?\s+AI)|"
    r"forget\s+(everything|all)|act\s+as|role\s*play|jailbreak)",
    re.IGNORECASE,
)


# ─────────────────────────────────────────────────
# ユーティリティ
# ─────────────────────────────────────────────────

def _sanitize_query(q: str) -> str:
    """
    クエリ文字列を安全化する。
    - 制御文字・改行を除去
    - 長さ上限を適用
    - プロンプトインジェクションパターンを無害化
    """
    if not isinstance(q, str):
        q = str(q)
    q = q.replace("\n", " ").replace("\r", " ").replace("\t", " ")
    q = re.sub(r"[\x00-\x1f\x7f]", "", q)   # 制御文字除去
    q = q[:MAX_QUERY_LENGTH].strip()
    if _INJECTION_PATTERN.search(q):
        q = "[フィルタ済みクエリ]"
    return q


def _parse_ctr(val) -> float:
    """
    CTR を小数に統一する。
      - Excel版: すでに小数 (0.408 等)
      - CSV版:   パーセント文字列 ('40.8%' 等)
    """
    if isinstance(val, float):
        # 0〜1 なら小数、1超ならパーセント換算
        return val if val <= 1.0 else val / 100.0
    s = str(val).strip().replace("%", "").replace(",", ".")
    try:
        v = float(s)
        return v / 100.0 if v > 1.0 else v
    except ValueError:
        return 0.0


def _normalize_df(df: pd.DataFrame) -> pd.DataFrame | None:
    """カラム名正規化・型変換・サニタイズを一括適用"""
    # カラム名マッピング
    rename = {}
    for col in df.columns:
        key = col.strip()
        if key in _COLUMN_MAP:
            rename[col] = _COLUMN_MAP[key]

    if not rename:
        logger.warning("既知のカラムが見つかりません: %s", list(df.columns))
        return None

    df = df.rename(columns=rename)

    # CTR が無い場合は clicks/impressions から計算
    if "ctr" not in df.columns and {"clicks", "impressions"}.issubset(df.columns):
        df["ctr"] = df.apply(
            lambda r: r["clicks"] / r["impressions"] if r["impressions"] > 0 else 0.0,
            axis=1,
        )

    required = ["query", "clicks", "impressions", "ctr", "position"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        logger.warning("カラム不足: %s", missing)
        return None

    # 型変換
    df["clicks"]      = pd.to_numeric(df["clicks"],      errors="coerce").fillna(0).astype(int)
    df["impressions"] = pd.to_numeric(df["impressions"], errors="coerce").fillna(0).astype(int)
    df["ctr"]         = df["ctr"].apply(_parse_ctr)
    df["position"]    = pd.to_numeric(df["position"],    errors="coerce").fillna(100.0)
    df["query"]       = df["query"].apply(_sanitize_query)

    # 無効行を除去
    df = df[(df["impressions"] > 0) & (df["query"] != "") & (df["query"] != "[フィルタ済みクエリ]")]
    df = df.dropna(subset=["position"])

    return df[required].reset_index(drop=True)


# ─────────────────────────────────────────────────
# 公開 API
# ─────────────────────────────────────────────────

def parse_sc_file(uploaded_file) -> tuple[pd.DataFrame | None, str | None]:
    """
    SC エクスポートファイル（Excel / CSV）を読み込む。

    Returns:
        (DataFrame, None)       : 成功
        (None, エラーメッセージ): 失敗
    """
    # ファイルサイズチェック
    content = uploaded_file.read()
    if len(content) > MAX_FILE_SIZE_BYTES:
        return None, f"ファイルサイズが上限（20MB）を超えています（{len(content)//1024//1024}MB）"

    name = getattr(uploaded_file, "name", "")

    # ─── Excel ───
    if name.lower().endswith(".xlsx") or name.lower().endswith(".xls"):
        return _parse_xlsx(content)

    # ─── CSV ───
    return _parse_csv(content)


def _parse_xlsx(content: bytes) -> tuple[pd.DataFrame | None, str | None]:
    """Excelファイルを解析。クエリシートを優先して読み込む"""
    try:
        xl = pd.ExcelFile(io.BytesIO(content))
        sheet_names = xl.sheet_names

        # クエリシートを探す（日本語 / 英語）
        target_sheet = None
        for candidate in ["クエリ", "Queries", "queries"]:
            if candidate in sheet_names:
                target_sheet = candidate
                break

        if target_sheet is None:
            return None, (
                f"Excelファイルに「クエリ」シートが見つかりませんでした。\n"
                f"含まれるシート: {', '.join(sheet_names)}\n\n"
                "Search Consoleの「エクスポート」→「Excelとしてダウンロード」で"
                "取得したファイルを使用してください。"
            )

        df = pd.read_excel(io.BytesIO(content), sheet_name=target_sheet)
        result = _normalize_df(df)
        if result is None or result.empty:
            return None, "クエリシートのデータを読み込めませんでした。"

        logger.info("Excel 読み込み完了（%s）: %d 件", target_sheet, len(result))
        return result, None

    except Exception as e:
        logger.error("Excel パース失敗: %s", e)
        return None, f"Excelファイルの読み込みに失敗しました: {e}"


def _parse_csv(content: bytes) -> tuple[pd.DataFrame | None, str | None]:
    """CSVファイルを解析。UTF-8 BOM 対応"""
    try:
        if content.startswith(b"\xef\xbb\xbf"):
            content = content[3:]

        text = content.decode("utf-8", errors="replace")
        df = pd.read_csv(io.StringIO(text))
        result = _normalize_df(df)
        if result is None or result.empty:
            return None, (
                "CSVのカラム構成を認識できませんでした。\n"
                "Search Consoleの「エクスポート」→「CSVをダウンロード」で"
                "取得したファイルを使用してください。"
            )

        logger.info("CSV 読み込み完了: %d 件", len(result))
        return result, None

    except Exception as e:
        logger.error("CSV パース失敗: %s", e)
        return None, f"CSVファイルの読み込みに失敗しました: {e}"

# -*- coding: utf-8 -*-
"""
SEO分析コア

内部スコアリングロジック。スコア名・分類名はユーザーに一切露出しない。
"""

import logging
import pandas as pd

logger = logging.getLogger(__name__)


def _position_weight(pos: float) -> float:
    """掲載順位による重み（内部用）"""
    if pos <= 3:
        return 1.5
    elif pos <= 10:
        return 2.0
    elif pos <= 20:
        return 1.8
    elif pos <= 50:
        return 0.5
    return 0.1


def _ctr_gap_ratio(ctr: float, avg_ctr: float) -> float:
    """CTRが平均からどれだけ下回っているかの比率（内部用）"""
    if avg_ctr <= 0:
        return 0.0
    return max(0.0, (avg_ctr - ctr) / avg_ctr)


def _opportunity_score(row: pd.Series, avg_ctr: float) -> float:
    """改善機会スコア（内部用・ユーザーに見せない）"""
    impressions = row["impressions"]
    if impressions < 10:
        return 0.0
    gap = _ctr_gap_ratio(row["ctr"], avg_ctr)
    pos_w = _position_weight(row["position"])
    return impressions * gap * pos_w


def _ctr_status(ctr: float, avg_ctr: float) -> str:
    """CTRの平均比較ラベル（汎用表現・内部分類名を使わない）"""
    if avg_ctr <= 0:
        return "データ不足"
    ratio = ctr / avg_ctr
    if ratio >= 1.2:
        return "平均以上"
    elif ratio >= 0.9:
        return "平均並み"
    elif ratio >= 0.6:
        return "やや低め"
    else:
        return "かなり低め"


def analyze_sc_data(df: pd.DataFrame) -> dict:
    """
    SC DataFrameを分析してGemini処方箋生成用のサマリーを返す。

    Returns dict:
        total_kw, total_clicks, total_impressions,
        avg_ctr, avg_position, pos_distribution,
        top_keywords (list of dict, 優先度順)
    """
    total_clicks      = int(df["clicks"].sum())
    total_impressions = int(df["impressions"].sum())
    avg_ctr = total_clicks / total_impressions if total_impressions > 0 else 0.0

    weighted_pos_sum = (df["position"] * df["impressions"]).sum()
    avg_position = (
        weighted_pos_sum / total_impressions if total_impressions > 0 else 0.0
    )

    # 内部スコアリング
    df = df.copy()
    df["_score"] = df.apply(_opportunity_score, axis=1, avg_ctr=avg_ctr)

    # 上位30件（スコア降順）をGemini用データとして構築
    top30 = df.nlargest(30, "_score")

    keyword_entries = []
    for _, row in top30.iterrows():
        keyword_entries.append({
            "query":      row["query"],
            "position":   round(row["position"], 1),
            "impressions": int(row["impressions"]),
            "clicks":     int(row["clicks"]),
            "ctr_pct":    f"{row['ctr']:.1%}",
            "ctr_status": _ctr_status(row["ctr"], avg_ctr),
        })

    # 順位帯の分布
    pos_distribution = {
        "1〜3位":  int((df["position"] <= 3).sum()),
        "4〜10位": int(((df["position"] > 3)  & (df["position"] <= 10)).sum()),
        "11〜20位":int(((df["position"] > 10) & (df["position"] <= 20)).sum()),
        "21〜50位":int(((df["position"] > 20) & (df["position"] <= 50)).sum()),
        "51位以下": int((df["position"] > 50).sum()),
    }

    return {
        "total_kw":          len(df),
        "total_clicks":      total_clicks,
        "total_impressions": total_impressions,
        "avg_ctr":           avg_ctr,
        "avg_position":      avg_position,
        "pos_distribution":  pos_distribution,
        "top_keywords":      keyword_entries,
    }

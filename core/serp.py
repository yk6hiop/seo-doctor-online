# -*- coding: utf-8 -*-
"""
SERP取得・分析モジュール

Yahoo検索でSERP上位結果を取得し、ライバルページのH2見出し・スニペットを抽出する。
ローカル版 analyzer/kw_reverse.py をオンライン版向けに移植。
"""

import json
import logging
import re
from urllib.parse import urlparse, quote_plus, unquote

import requests
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ja,en;q=0.9",
}

_EXCLUDE_DOMAINS = (
    "yahoo.co.jp", "yahoo.com", "google.com", "google.co.jp",
    "chiebukuro.yahoo.co.jp", "detail.chiebukuro.yahoo.co.jp",
    "twitter.com", "x.com", "facebook.com", "instagram.com", "tiktok.com",
    "youtube.com", "nicovideo.jp",
    "wikipedia.org", "ja.wikipedia.org",
    "amazon.co.jp", "amazon.com", "rakuten.co.jp",
    "indeed.com", "mynavi.jp", "rikunabi.com",
    "r.search.yahoo.com",
)

_NOISE_CLASS_KEYWORDS = frozenset({
    "ad", "ads", "advertisement", "banner", "sponsor", "sponsored",
    "promotion", "promo", "cta", "sidebar", "widget", "popup",
    "modal", "cookie", "subscribe", "newsletter", "share",
    "social", "related", "recommend", "author", "profile",
    "breadcrumb", "pager", "pagination", "comment", "reply",
    "order", "booking", "reservation", "cart", "checkout",
    "form", "input", "login", "register", "signup",
    "footer", "header", "nav", "menu", "gnav", "lnav",
})

_OFFICIAL_TITLE_MARKERS = (
    "コーポレートサイト", "公式サイト", "公式ホームページ", "公式HP",
    "株式会社", "会社概要", "コーポレート", "corporate",
)
_SERVICE_PAGE_MARKERS = (
    "初めての方へ", "初めてご利用", "のご案内", "店舗案内",
    "ご予約方法", "予約方法", "予約・アクセス",
    "鑑定料金", "ご利用料金", "料金表", "ご利用案内",
    "会員専用", "ログインページ",
    "プロフィール一覧", "占い師一覧", "先生一覧",
)
_LISTING_DIRECTORY_DOMAINS = (
    "itp.ne.jp", "i-town.jp", "tabelog.com", "hotpepper.jp",
    "gnavi.co.jp", "retty.me", "kakaku.com", "jalan.net",
    "tripadvisor.com", "tripadvisor.jp", "ikyu.com",
    "coconala.com", "minne.com", "creema.jp",
    "ekiten.jp", "epark.jp",
)
_NEWS_MARKERS = ("wikipedia", "itmedia", "impress", "mynavi", "goo.ne.jp", "naver.jp", "matome")


def _is_excluded_domain(domain: str) -> bool:
    d = (domain or "").lower().lstrip(".").replace("www.", "")
    for ex in _EXCLUDE_DOMAINS:
        if d == ex or d.endswith("." + ex):
            return True
    return False


def _is_noise_element(el) -> bool:
    classes = " ".join(el.get("class", [])).lower()
    el_id = (el.get("id") or "").lower()
    combined = classes + " " + el_id
    return any(noise in combined for noise in _NOISE_CLASS_KEYWORDS)


def _is_noise_text(text: str) -> bool:
    if len(text) < 30:
        return True
    valid_chars = sum(
        1 for c in text
        if c.isalpha() or "぀" <= c <= "ヿ" or "一" <= c <= "鿿"
    )
    if valid_chars / max(len(text), 1) < 0.3:
        return True
    _NAV_PHRASES = ["会員登録", "お申込", "申し込み", "ログイン", "新規登録", "今すぐ予約"]
    nav_hit = sum(1 for p in _NAV_PHRASES if p in text)
    if nav_hit >= 2 and len(text) < 200:
        return True
    return False


def _classify_site_type(results: list[dict]) -> list[dict]:
    """SERPのサイト種別（公式/ブランド・メディア・一般）を付与する"""
    from collections import Counter
    dom_counts = Counter(
        r["domain"].lower().replace("www.", "").split("/")[0]
        for r in results
    )
    for r in results:
        d_norm = r["domain"].lower().replace("www.", "").split("/")[0]
        url_lower = (r.get("url") or "").lower()
        title = r.get("title") or ""

        if dom_counts[d_norm] >= 2:
            r["site_type"] = "公式/ブランド"
        elif any(m in title or m.lower() in title.lower() for m in _OFFICIAL_TITLE_MARKERS):
            r["site_type"] = "公式/ブランド"
        elif any(m in title for m in _SERVICE_PAGE_MARKERS):
            r["site_type"] = "公式/ブランド"
        elif any(d_norm == d or d_norm.endswith("." + d) for d in _LISTING_DIRECTORY_DOMAINS):
            r["site_type"] = "公式/ブランド"
        elif any(m in d_norm for m in _NEWS_MARKERS):
            r["site_type"] = "メディア/まとめ"
        else:
            r["site_type"] = "一般"
    return results


def analyze_serp(keyword: str, timeout: int = 15) -> list[dict]:
    """
    キーワードでYahoo検索し、上位10件のページ情報を取得する。

    Returns:
        [{rank, url, title, snippet, domain, site_type}, ...]
        取得失敗時は空リスト。
    """
    query = quote_plus(keyword)
    url = f"https://search.yahoo.co.jp/search?p={query}&ei=UTF-8"

    results = []
    try:
        resp = requests.get(url, headers=_HEADERS, timeout=timeout, allow_redirects=True)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
    except Exception as e:
        logger.warning("Yahoo SERP取得エラー（%s）: %s", keyword, e)
        return results

    seen_urls: set[str] = set()

    _NOISE_TITLE_PREFIXES = (
        "他の人はこちらも", "他の人はこれも", "関連検索ワード",
        "関連する質問", "人々はこちらも", "サポートが必要",
        "ウェブ検索結果", "すべての検索結果",
    )

    def _add_result(href: str, title: str, snippet: str):
        if not href or not href.startswith("http"):
            return
        if href in seen_urls:
            return
        if title and any(title.startswith(p) for p in _NOISE_TITLE_PREFIXES):
            return
        domain = urlparse(href).netloc
        if _is_excluded_domain(domain):
            return
        seen_urls.add(href)
        results.append({
            "rank": len(results) + 1,
            "url": href,
            "title": title or "",
            "snippet": snippet or "",
            "domain": domain,
            "site_type": "一般",
        })

    def _resolve_yahoo_redirect(href: str) -> str:
        if "r.search.yahoo.com" in href:
            m = re.search(r"/RU=([^/]+)", href)
            if m:
                try:
                    return unquote(m.group(1))
                except Exception:
                    pass
        return href

    # 戦略1: Yahoo検索の主要結果コンテナ
    selectors = (
        "div.sw-CardBase",
        "div.Algo",
        "div.w.srg",
        "li.Algo",
        "section.sw-CardBase",
    )
    for sel in selectors:
        cards = soup.select(sel)
        if not cards:
            continue
        for card in cards:
            title_a = card.find("a", href=True)
            if not title_a:
                continue
            href = _resolve_yahoo_redirect(title_a.get("href", ""))
            h3 = card.find(["h3", "h2"])
            title = h3.get_text(strip=True) if h3 else title_a.get_text(strip=True)
            snippet = ""
            for tag in card.find_all(["p", "span", "div"]):
                text = tag.get_text(strip=True)
                if 40 <= len(text) <= 300:
                    snippet = text[:300]
                    break
            _add_result(href, title, snippet)
            if len(results) >= 10:
                break
        if len(results) >= 3:
            break

    # 戦略2: h3配下のaリンク
    if len(results) < 3:
        for h3 in soup.find_all(["h3", "h2"]):
            a = h3.find("a", href=True)
            if not a:
                continue
            href = _resolve_yahoo_redirect(a.get("href", ""))
            _add_result(href, h3.get_text(strip=True), "")
            if len(results) >= 10:
                break

    # 戦略3: 全aリンクフォールバック
    if len(results) < 3:
        for a in soup.find_all("a", href=True):
            href = _resolve_yahoo_redirect(a.get("href", ""))
            if not href.startswith("http"):
                continue
            _add_result(href, a.get_text(strip=True)[:100], "")
            if len(results) >= 10:
                break

    final = results[:10]
    return _classify_site_type(final)


def analyze_serp_pages(
    serp_results: list[dict],
    max_pages: int = 5,
    timeout: int = 10,
) -> list[dict]:
    """
    SERP上位ページの詳細分析（H2見出し+本文スニペット）。

    公式/ブランドサイトを除いた上位 max_pages 件のHTMLを取得してH2を抽出する。
    serp_results を直接変更して返す。
    """
    fetched = 0
    for result in serp_results:
        if result.get("site_type") == "公式/ブランド":
            result["h2_with_snippet"] = []
            continue
        if fetched >= max_pages:
            result["h2_with_snippet"] = []
            continue

        url = result.get("url", "")
        if not url:
            result["h2_with_snippet"] = []
            continue

        try:
            resp = requests.get(url, headers=_HEADERS, timeout=timeout, allow_redirects=True)
            resp.raise_for_status()
            soup = BeautifulSoup(resp.text, "html.parser")

            h2_with_snippet = []
            for h2_el in soup.find_all("h2"):
                h2_text = h2_el.get_text(strip=True)
                if not h2_text:
                    continue
                snippet_chars = []
                snippet_total = 0
                seen_prefix: set[str] = set()

                for sibling in h2_el.find_all_next():
                    if sibling.name in ("h1", "h2"):
                        break
                    if sibling.name in ("script", "style", "nav", "footer",
                                        "aside", "form", "iframe", "noscript"):
                        continue
                    if _is_noise_element(sibling):
                        continue
                    if sibling.name in ("p", "li", "dd", "blockquote"):
                        text = sibling.get_text(separator=" ", strip=True)
                        if _is_noise_text(text):
                            continue
                        key = text[:50]
                        if key in seen_prefix:
                            continue
                        seen_prefix.add(key)
                        snippet_chars.append(text)
                        snippet_total += len(text)
                        if snippet_total >= 500:
                            break

                snippet = " ".join(snippet_chars)[:250]
                h2_with_snippet.append({"h2": h2_text, "snippet": snippet})

            result["h2_with_snippet"] = h2_with_snippet[:10]
            fetched += 1

        except Exception as e:
            logger.warning("ページ分析エラー（%s）: %s", url, e)
            result["h2_with_snippet"] = []

    return serp_results

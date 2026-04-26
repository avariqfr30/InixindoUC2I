import concurrent.futures
import hashlib
import json
import logging
import os
from pathlib import Path
import re
from datetime import datetime
from urllib.parse import urlparse

import diskcache as dc
import requests
from ollama import Client
from pydantic import BaseModel, Field

from config import (
    OLLAMA_HOST,
    OSINT_BASE_QUERIES,
    OSINT_BLOCKED_DOMAINS,
    OSINT_CACHE_DIR,
    OSINT_CACHE_TTL_SECONDS,
    OSINT_DEEP_SCRAPE_MAX_CHARS,
    OSINT_MAX_SIGNALS,
    OSINT_QUERY_WORKERS,
    OSINT_RECENCY,
    OSINT_RESULTS_PER_QUERY,
    OSINT_SEARCH_LANGUAGE,
    OSINT_SEARCH_REGION,
    OSINT_TRUSTED_DOMAINS,
    SERPER_API_KEY,
)

logger = logging.getLogger(__name__)

class InsightSchema(BaseModel):
    insight: str = Field(description="The extracted insight in Indonesian. 'NOT_FOUND' if missing.")

# Initialize disk cache for OSINT (single source of truth: OSINT_CACHE_DIR)
osint_cache = dc.Cache(str(Path(OSINT_CACHE_DIR)))
# ==========================================

class Researcher:
    SERPER_ENDPOINT = "https://google.serper.dev/search"
    INVALID_API_KEYS = {"", "YOUR_SERPER_API_KEY", "masukkan_api_key_serper_anda_disini"}
    MACRO_TRENDS_CACHE_VERSION = "macro_trends_v3"
    NON_CACHEABLE_MESSAGES = {
        "Data OSINT eksternal tidak tersedia (SERPER_API_KEY belum diatur).",
        "Tidak ada tren eksternal yang berhasil dimuat.",
        "Tidak ada sinyal OSINT eksternal yang dapat digunakan untuk benchmark periode ini.",
    }
    QUERY_STOPWORDS = {
        "yang", "dengan", "untuk", "pada", "dari", "dan", "atau", "dalam",
        "this", "that", "from", "with", "into", "latest", "trends",
    }

    @staticmethod
    def _is_enabled():
        return (SERPER_API_KEY or "").strip() not in Researcher.INVALID_API_KEYS

    @staticmethod
    def _normalize_cache_token(value):
        return re.sub(r"\s+", " ", str(value or "").strip().lower())

    @classmethod
    def _normalize_query(cls, query):
        return cls._normalize_cache_token(query)

    @staticmethod
    def _source_domain(url):
        try:
            netloc = urlparse(url).netloc.lower()
            return netloc.replace("www.", "") if netloc else "unknown"
        except Exception:
            return "unknown"

    @staticmethod
    def _normalize_url(url):
        try:
            parsed = urlparse(url)
            domain = parsed.netloc.lower().replace("www.", "")
            path = re.sub(r"/+$", "", parsed.path or "/")
            return f"{domain}{path}"
        except Exception:
            return str(url or "").strip().lower()

    @staticmethod
    def _domain_matches(domain, candidates):
        clean_domain = (domain or "").lower().replace("www.", "")
        return any(
            clean_domain == candidate or clean_domain.endswith(f".{candidate}")
            for candidate in candidates
            if candidate
        )

    @classmethod
    def _source_quality_score(cls, url):
        domain = cls._source_domain(url)
        if cls._domain_matches(domain, OSINT_BLOCKED_DOMAINS):
            return -10
        if cls._domain_matches(domain, OSINT_TRUSTED_DOMAINS):
            return 4
        if domain.endswith(".go.id") or domain.endswith(".ac.id") or domain.endswith(".org"):
            return 2
        if domain.endswith(".edu") or domain.endswith(".gov"):
            return 2
        return 0

    @staticmethod
    def _freshness_score(date_text):
        value = str(date_text or "").strip().lower()
        if not value:
            return 0
        current_year = datetime.now().year
        if re.search(r"\b(jam|hour|minute|menit|hari|day|today|yesterday)\b", value):
            return 2
        if re.search(r"\b(minggu|week|bulan|month)\b", value):
            return 1.5
        year_match = re.search(r"\b(20\d{2})\b", value)
        if year_match:
            year = int(year_match.group(1))
            if year >= current_year - 1:
                return 1
            if year <= current_year - 3:
                return -1
        return 0

    @classmethod
    def _macro_trends_cache_key(cls, timeframe, notes, score_engine_label, llm_model):
        key_payload = {
            "version": cls.MACRO_TRENDS_CACHE_VERSION,
            "timeframe": cls._normalize_cache_token(timeframe),
            "notes": cls._normalize_cache_token(notes),
            "score_engine_label": cls._normalize_cache_token(score_engine_label),
            "llm_model": cls._normalize_cache_token(llm_model),
        }
        key_raw = json.dumps(key_payload, ensure_ascii=False, sort_keys=True)
        key_hash = hashlib.sha256(key_raw.encode("utf-8")).hexdigest()
        return f"osint:macro:{key_hash}"

    @classmethod
    def _is_cacheable_macro_result(cls, result):
        if not isinstance(result, str):
            return False
        stripped = result.strip()
        if not stripped:
            return False
        return stripped not in cls.NON_CACHEABLE_MESSAGES

    @staticmethod
    def fetch_full_markdown(url):
        """Fetches the clean markdown text of any URL using Jina Reader."""
        if not url:
            return ""
        try:
            jina_url = f"https://r.jina.ai/{url}"
            headers = {'User-Agent': 'Mozilla/5.0'}
            response = requests.get(jina_url, headers=headers, timeout=12)
            if response.status_code == 200:
                text = re.sub(r"\n{3,}", "\n\n", response.text).strip()
                return text[:OSINT_DEEP_SCRAPE_MAX_CHARS]
            return ""
        except Exception as e:
            logger.warning("Failed to fetch full markdown for %s: %s", url, e)
            return ""

    @classmethod
    def extract_insight_with_llm(cls, url, extraction_goal):
        """Universal Deep Scraper: Reads a URL and extracts a specific qualitative insight via Pydantic/LLM."""
        markdown_text = cls.fetch_full_markdown(url)
        if not markdown_text:
            return ""
            
        prompt = f"""
        You are an expert business researcher. Read the following source text.
        Your goal is to extract: {extraction_goal}
        
        SOURCE TEXT:
        {markdown_text}
        
        Respond ONLY with a valid JSON object using this schema. If the information is not present, use "NOT_FOUND".
        {{
            "insight": "<concise professional summary in Indonesian>"
        }}
        """
        try:
            llm_model = os.getenv("LLM_MODEL", "gpt-oss:120b-cloud")
            client = Client(host=OLLAMA_HOST)
            res = client.chat(
                model=llm_model,
                messages=[{'role': 'user', 'content': prompt}],
                options={'temperature': 0.0}
            )
            raw_text = res['message']['content']
            match = re.search(r'\{.*\}', raw_text, re.DOTALL)
            parsed_dict = json.loads(match.group(0)) if match else json.loads(raw_text)
            
            # Use Pydantic to strictly validate
            data = InsightSchema.model_validate(parsed_dict)
            
            if "NOT_FOUND" in data.insight.upper() or not data.insight:
                return ""
            return data.insight
        except Exception as e:
            logger.warning("Insight extraction failed for %s: %s", url, e)
            return ""

    @staticmethod
    def _search_serper(query, max_results=OSINT_RESULTS_PER_QUERY):
        normalized_query = Researcher._normalize_query(query)
        if not normalized_query:
            raise ValueError("OSINT query kosong setelah normalisasi.")

        payload = {
            "q": normalized_query,
            "num": max_results,
            "gl": OSINT_SEARCH_REGION,
            "hl": OSINT_SEARCH_LANGUAGE,
        }
        if OSINT_RECENCY:
            payload["tbs"] = OSINT_RECENCY

        headers = {
            "X-API-KEY": SERPER_API_KEY,
            "Content-Type": "application/json",
        }

        response = requests.post(
            Researcher.SERPER_ENDPOINT,
            headers=headers,
            json=payload,
            timeout=10,
        )
        response.raise_for_status()
        return response.json()

    @staticmethod
    def _extract_items(query, payload):
        items = []
        for position, entry in enumerate(payload.get("organic", []), start=1):
            items.append({
                "query": query, "title": entry.get("title", "Tanpa Judul"), "snippet": entry.get("snippet", "").strip(),
                "url": entry.get("link", "").strip(), "date": entry.get("date", "").strip(), "source_type": "organic", "position": position
            })
        for position, entry in enumerate(payload.get("news", []), start=1):
            items.append({
                "query": query, "title": entry.get("title", "Tanpa Judul"), "snippet": entry.get("snippet", "").strip(),
                "url": entry.get("link", "").strip(), "date": entry.get("date", "").strip(), "source_type": "news", "position": position
            })
        return [
            item
            for item in items
            if item["url"] and Researcher._source_quality_score(item["url"]) > -10
        ]

    @staticmethod
    def _deduplicate_items(items):
        seen_keys = set()
        unique_items = []
        for item in items:
            key = Researcher._normalize_url(item["url"]) or f"{item['title']}::{item['snippet']}"
            if key in seen_keys:
                continue
            seen_keys.add(key)
            unique_items.append(item)
        return unique_items

    @staticmethod
    def _score_items(items, context_text):
        keywords = {
            token.lower()
            for token in re.findall(r"[A-Za-z]{4,}", context_text)
            if token.lower() not in Researcher.QUERY_STOPWORDS
        }
        for item in items:
            corpus = f"{item['title']} {item['snippet']}".lower()
            coverage_score = sum(1 for keyword in keywords if keyword in corpus)
            source_quality = Researcher._source_quality_score(item["url"])
            freshness_bonus = Researcher._freshness_score(item.get("date"))
            type_bonus = 1 if item["source_type"] == "news" else 0
            ranking_penalty = (item["position"] - 1) * 0.05
            item["source_quality"] = source_quality
            item["score"] = coverage_score + source_quality + freshness_bonus + type_bonus - ranking_penalty
        return sorted(items, key=lambda value: value["score"], reverse=True)

    @staticmethod
    def _format_osint_brief(items, title):
        if not items:
            return "Tidak ada sinyal OSINT eksternal yang dapat digunakan untuk benchmark periode ini."
        lines = [f"{title}:"]
        for index, item in enumerate(items, start=1):
            date_part = f" | tanggal={item['date']}" if item["date"] else ""
            source = Researcher._source_domain(item["url"])
            quality_part = f" | kualitas_sumber={item.get('source_quality', 0)}"
            lines.append(
                (
                    f"{index}. {item['title']} | {item['snippet']} | "
                    f"sumber={source}{date_part}{quality_part} | url={item['url']}"
                ).strip()
            )
        return "\n".join(lines)

    @staticmethod
    def _run_query_batch(queries, max_signals=OSINT_MAX_SIGNALS):
        normalized_queries = []
        seen_queries = set()
        for query in queries:
            normalized_query = Researcher._normalize_query(query)
            if not normalized_query or normalized_query in seen_queries:
                continue
            seen_queries.add(normalized_query)
            normalized_queries.append(normalized_query)

        if not normalized_queries:
            return []

        collected = []
        max_workers = min(OSINT_QUERY_WORKERS, max(1, len(normalized_queries)))
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as pool:
            future_map = {pool.submit(Researcher._search_serper, query): query for query in normalized_queries}
            for future in concurrent.futures.as_completed(future_map):
                query = future_map[future]
                try:
                    payload = future.result()
                    collected.extend(Researcher._extract_items(query, payload))
                except Exception as exc:
                    logger.warning("OSINT query gagal (%s): %s", query, exc)

        deduplicated = Researcher._deduplicate_items(collected)
        ranked = Researcher._score_items(deduplicated, " ".join(normalized_queries))
        return ranked[:max_signals]

    @classmethod
    def get_macro_trends(cls, timeframe, notes="", score_engine_label="Experience Index"):
        scope = timeframe or "periode terbaru"
        compact_notes = re.sub(r"\s+", " ", notes).strip()
        llm_model = os.getenv("LLM_MODEL", "gpt-oss:120b-cloud")
        cache_key = cls._macro_trends_cache_key(
            timeframe=scope,
            notes=compact_notes,
            score_engine_label=score_engine_label,
            llm_model=llm_model,
        )

        cached_result = osint_cache.get(cache_key)
        if isinstance(cached_result, str) and cached_result.strip():
            return cached_result

        contextual_query = (
            f"benchmark sentimen pelanggan pelatihan dan konsultasi IT Indonesia {scope} {score_engine_label}"
        )
        if compact_notes:
            contextual_query += f" {compact_notes[:140]}"

        queries = [f"{query} {scope}" for query in OSINT_BASE_QUERIES] + [contextual_query]

        if not cls._is_enabled():
            return "Data OSINT eksternal tidak tersedia (SERPER_API_KEY belum diatur)."

        try:
            findings = cls._run_query_batch(queries)

            # --- DEEP SCRAPE THE #1 RESULT ---
            deep_insight_text = ""
            if findings and findings[0].get("url"):
                top_link = findings[0]["url"]
                logger.info("Deep scraping OSINT for macro trends: %s", top_link)
                goal = "What are the latest macro trends, challenges, or benchmarks regarding IT training, consulting, and customer expectations in Indonesia?"
                insight = cls.extract_insight_with_llm(top_link, goal)
                if insight:
                    source = cls._source_domain(top_link)
                    deep_insight_text = f"**Insight Mendalam (via {source}):** {insight}\n\n"

            brief = cls._format_osint_brief(findings, "Sinyal OSINT Makro (Indonesia)")
            result = deep_insight_text + brief
            if cls._is_cacheable_macro_result(result):
                osint_cache.set(cache_key, result, expire=OSINT_CACHE_TTL_SECONDS)
            return result
        except Exception as exc:
            logger.warning("OSINT macro trends failed: %s", exc)
            return "Tidak ada tren eksternal yang berhasil dimuat."

"""App-owned writing planner for feedback intelligence reports."""
from __future__ import annotations

import copy
import hashlib
import json
from typing import Any


def _digest(value: Any) -> str:
    raw = json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


class FeedbackSectionPlanner:
    """Build hidden guidance that turns feedback evidence into management action."""
    CACHE_VERSION = "feedback-section-plan-v2"
    _cache: dict[str, dict[str, Any]] = {}
    _stats = {"hits": 0, "misses": 0}

    @classmethod
    def clear_cache(cls) -> None:
        cls._cache.clear()
        cls._stats = {"hits": 0, "misses": 0}

    @classmethod
    def cache_stats(cls) -> dict[str, int]:
        return {**cls._stats, "items": len(cls._cache)}

    @classmethod
    def _remember(cls, key: str, plan: dict[str, Any]) -> dict[str, Any]:
        if key in cls._cache:
            cls._cache.pop(key, None)
        cls._cache[key] = copy.deepcopy(plan)
        while len(cls._cache) > 256:
            cls._cache.pop(next(iter(cls._cache)))
        return copy.deepcopy(plan)

    def build_plan(self, sections: list[str], context: dict[str, Any] | None = None) -> dict[str, Any]:
        context = context or {}
        document_contract = context.get("document_contract") if isinstance(context.get("document_contract"), dict) else {}
        cache_key = _digest({
            "version": self.CACHE_VERSION,
            "sections": sections,
            "context": context,
        })
        if cache_key in self._cache:
            self._stats["hits"] += 1
            return copy.deepcopy(self._cache[cache_key])
        self._stats["misses"] += 1
        section_list = ", ".join(str(item).strip() for item in sections if str(item).strip())
        timeframe = str(context.get("timeframe_label") or context.get("timeframe") or "periode terpilih").strip()
        sentiment = str(context.get("sentiment") or "all").strip()
        segment = str(context.get("segment") or "all").strip()
        osint_dossier = context.get("osint_dossier") if isinstance(context.get("osint_dossier"), dict) else {}
        osint_cards = [
            {
                "claim": str(card.get("claim", "") or "").strip(),
                "why_it_matters": str(card.get("why_it_matters", "") or "").strip(),
                "source_domain": str(card.get("source_domain", "") or "").strip(),
                "allowed_use": str(card.get("allowed_use", "") or "").strip(),
                "matched_internal_fact": str(card.get("matched_internal_fact", "") or "").strip(),
            }
            for card in (osint_dossier.get("evidence_cards") or [])[:5]
            if isinstance(card, dict) and str(card.get("claim", "") or "").strip()
        ]
        plan = {
            "cache_key": cache_key,
            "data_version": str(context.get("data_version") or ""),
            "use_case": "feedback",
            "reader": "management_team",
            "section_title": section_list or "bagian laporan yang sedang disusun",
            "section_goal": "mengubah respons, rating, tema, dan filter pengguna menjadi insight manajemen yang dapat ditindaklanjuti",
            "user_flow_context": f"feedback_report timeframe={timeframe}; sentiment={sentiment}; segment={segment}",
            "evidence_required": ["rating", "komentar", "tema layanan", "filter periode", "segmentasi pengguna"],
            "protected_facts": [timeframe, f"sentiment={sentiment}", f"segment={segment}"],
            "tone_rules": ["neutral_management_analyst", "evidence_first", "respectful_to_participants_and_delivery_team"],
            "avoid_patterns": [
                "copy-paste section conclusions",
                "unsupported sentiment claims",
                "overstated causal claims",
                "repeated status/implication/intervention labels",
                "paragraphs that all move from problem to generic recommendation",
            ],
            "quality_thresholds": {"max_repeated_openings": 2, "require_action_for_each_recommendation": True},
            "narrative_thesis": context.get("narrative_thesis") or "laporan harus menjelaskan pola pengalaman peserta sebagai keputusan perbaikan layanan",
            "internal_insight_cards": context.get("insight_cards") or [],
            "paragraph_roles": context.get("paragraph_role_rotation") or [
                "temuan",
                "bukti",
                "implikasi",
                "trade-off",
                "tindakan",
            ],
            "osint_dossier_quality": osint_dossier.get("quality") or {},
            "osint_evidence_cards": osint_cards,
            "retrieval_intent": {
                "goal": "find feedback themes, representative comments, and segment evidence for the selected filters",
                "preferred_datasets": ["ClassReport", "ReferenceClassReport"],
                "exclude": ["FinanceInvoice", "ProjectStandards"],
                "preferred_terms": [timeframe, sentiment, segment, *sections[:4]],
            },
            "evidence_ledger": [
                {
                    "claim_role": "feedback finding",
                    "evidence_source": "ClassReport or ReferenceClassReport",
                    "confidence": "depends on response count and theme consistency",
                    "allowed_wording": "apa yang dibuktikan feedback, apa implikasinya, dan apa tindakan manajemen",
                }
            ],
            "rationale_summary": {
                "main_reasoning": "Laporan feedback harus membedakan bukti, implikasi layanan, dan tindakan agar tidak terdengar seperti ringkasan template.",
                "evidence_used": ["rating", "komentar", "tema", "filter pengguna"],
                "caveats": ["hindari klaim sebab-akibat jika data hanya menunjukkan pola atau sinyal"],
            },
            "document_thesis": document_contract.get("document_thesis") or "",
            "chapter_contracts": document_contract.get("chapter_contracts") or [],
            "data_gap_register": document_contract.get("data_gap_register") or [],
            "editorial_contract": document_contract.get("editorial_contract") or {},
            "appendix_manifest": document_contract.get("appendix_manifest") or {},
        }
        return self._remember(cache_key, plan)

    def build_prompt_block_from_plan(self, plan: dict[str, Any]) -> str:
        return (
            "[SECTION_PLANNER] "
            f"[SECTION_PLAN_JSON] {plan} "
            f"Rencanakan bagian laporan feedback: {plan.get('section_title')}. "
            f"Konteks filter: {plan.get('user_flow_context')}. "
            "Untuk setiap bagian, jawab tiga hal sebelum menulis: apa yang dibuktikan feedback, apa implikasinya bagi layanan, "
            "dan apa tindakan manajemen yang paling masuk akal. "
            f"Tesis naratif: {plan.get('narrative_thesis')}. "
            f"Insight internal terkurasi: {plan.get('internal_insight_cards')}. "
            f"Rotasi peran paragraf: {plan.get('paragraph_roles')}. "
            "Gunakan suara analis manajemen yang netral, berbasis bukti, dan menghormati peserta maupun tim pelaksana. "
            "Kurangi pola kalimat pembuka yang sama antar-bab; variasikan struktur antara temuan, bukti, implikasi, trade-off, dan tindakan. "
            "Jangan mengulang label seperti 'Status kesiapan saat ini', 'Implikasi untuk pengambilan keputusan', atau 'Jika tidak ada intervensi' di banyak bagian. "
            f"Retrieval intent: {plan.get('retrieval_intent')}. Evidence ledger: {plan.get('evidence_ledger')}. "
            f"OSINT cards matched to internal filters: {plan.get('osint_evidence_cards')}. "
            f"Rationale ringkas: {plan.get('rationale_summary')}. "
            "Jangan membuat klaim di luar respons, rating, tema, atau konteks yang tersedia."
            f" Tesis dokumen: {plan.get('document_thesis')}. Kontrak bagian: {plan.get('chapter_contracts')}. "
            f"Kesenjangan data: {plan.get('data_gap_register')}. Kontrak editorial: {plan.get('editorial_contract')}."
        )

    def build_prompt_block(self, sections: list[str], context: dict[str, Any] | None = None) -> str:
        return self.build_prompt_block_from_plan(self.build_plan(sections, context))

"""Gold-label vs retrieved-text coverage computation (pure functions).

The 金标↔召回匹配 column of the badcase attribution report is computed
objectively here: the gold annotation is split into segments (JSON chunk
arrays or ``[Chunk N]``-prefixed plain text) and each segment's character
bigram coverage inside the retrieved text is measured.

Verdicts mirror the reference report wording:
- all segments >= 0.8 coverage  -> ``✅ 命中``
- any segment >= 0.4 coverage   -> ``⚠️ 部分命中``
- otherwise                     -> ``❌ 未召回``
"""

from __future__ import annotations

import json
import re
from typing import Any

COVERED_THRESHOLD = 0.8
PARTIAL_THRESHOLD = 0.4

_CHUNK_MARKER_RE = re.compile(r"(?i)\[chunk[ _-]?\w*\]")


def parse_gold_segments(annotation: Any) -> list[str]:
    """Split an annotation value into gold segments.

    Supported forms:
    - JSON array ``[{"chunk_id": "Chunk 3", "source_text": "..."}]``
    - plain text with ``[Chunk N]`` / ``[chunk_xxx]`` markers
    - any other text: treated as a single segment
    """
    if annotation is None:
        return []
    text = str(annotation).strip()
    if not text or text.casefold() in {"nan", "none", "null", "[]"}:
        return []

    parsed: Any = None
    if text.startswith("["):
        try:
            parsed = json.loads(text)
        except (json.JSONDecodeError, ValueError):
            parsed = None

    if isinstance(parsed, list):
        segments = []
        for item in parsed:
            if isinstance(item, dict):
                chunk_id = item.get("chunk_id") or ""
                source = item.get("source_text") or item.get("text") or ""
                prefix = f"[{chunk_id}] " if chunk_id else ""
                if str(source).strip():
                    segments.append(f"{prefix}{source}")
        if segments:
            return segments

    positions = [m.start() for m in _CHUNK_MARKER_RE.finditer(text)]
    if len(positions) > 1:
        segments = []
        for i, pos in enumerate(positions):
            end = positions[i + 1] if i + 1 < len(positions) else len(text)
            segment = text[pos:end].strip()
            if segment:
                segments.append(segment)
        return segments

    return [text]


def normalize(text: str) -> str:
    """Lowercase and collapse whitespace for overlap matching."""
    return re.sub(r"\s+", " ", text.strip().lower())


def _bigrams(text: str) -> set[str]:
    chars = normalize(text)
    if len(chars) < 2:
        return {chars} if chars else set()
    return {chars[i:i + 2] for i in range(len(chars) - 1)}


def _strip_chunk_prefix(text: str) -> str:
    """Remove a leading ``[Chunk N]`` display marker before overlap matching."""
    return _CHUNK_MARKER_RE.sub("", text, count=1).strip()


def segment_coverage(segment: str, retrieval: str) -> float:
    """Fraction of the segment's character bigrams present in retrieval."""
    seg_bigrams = _bigrams(_strip_chunk_prefix(segment))
    if not seg_bigrams:
        return 0.0
    ret_bigrams = _bigrams(retrieval)
    if not ret_bigrams:
        return 0.0
    return len(seg_bigrams & ret_bigrams) / len(seg_bigrams)


def _verdict(per_segment: list[float]) -> str:
    if per_segment and all(rate >= COVERED_THRESHOLD for rate in per_segment):
        return "✅ 命中"
    if any(rate >= PARTIAL_THRESHOLD for rate in per_segment):
        return "⚠️ 部分命中"
    return "❌ 未召回"


def coverage_summary(annotation: Any, retrieval: Any) -> dict[str, Any]:
    """Compute the coverage summary for one badcase row.

    Returns ``{"segments", "covered", "rate", "verdict", "text"}`` where
    ``text`` mirrors the reference wording, e.g.
    ``"覆盖率 45%（9/20）→ ⚠️ 部分命中"``.
    """
    segments = parse_gold_segments(annotation)
    retrieval_text = "" if retrieval is None else str(retrieval)
    if not segments:
        return {
            "segments": 0,
            "covered": 0,
            "rate": 0.0,
            "verdict": "❌ 未召回",
            "text": "无金标标注",
        }
    rates = [segment_coverage(seg, retrieval_text) for seg in segments]
    covered = sum(1 for rate in rates if rate >= COVERED_THRESHOLD)
    overall = sum(rates) / len(rates)
    verdict = _verdict(rates)
    return {
        "segments": len(segments),
        "covered": covered,
        "rate": round(overall, 3),
        "verdict": verdict,
        "text": f"覆盖率 {overall * 100:.0f}%（{covered}/{len(segments)} 段）→ {verdict}",
    }

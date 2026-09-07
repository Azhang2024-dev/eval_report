"""Tests for coverage.py: gold-segment parsing and coverage verdicts."""

import unittest

from scripts.coverage import (
    coverage_summary,
    normalize,
    parse_gold_segments,
    segment_coverage,
)


class ParseGoldSegmentsTest(unittest.TestCase):
    def test_json_array(self):
        annotation = '[{"chunk_id": "Chunk 3", "source_text": "first segment"}, {"chunk_id": "Chunk 2", "source_text": "second segment"}]'
        segments = parse_gold_segments(annotation)
        self.assertEqual(len(segments), 2)
        self.assertTrue(segments[0].startswith("[Chunk 3] first segment"))

    def test_json_missing_source_text(self):
        annotation = '[{"chunk_id": "Chunk 1", "source_text": "keep me"}, {"chunk_id": "Chunk 2"}]'
        segments = parse_gold_segments(annotation)
        self.assertEqual(segments, ["[Chunk 1] keep me"])

    def test_plain_chunk_markers(self):
        annotation = "[Chunk 15] first sentence.\n[Chunk 2] second sentence."
        segments = parse_gold_segments(annotation)
        self.assertEqual(len(segments), 2)

    def test_single_plain_text(self):
        self.assertEqual(parse_gold_segments("just one sentence"), ["just one sentence"])

    def test_empty_values(self):
        for value in (None, "", "nan", "[]"):
            self.assertEqual(parse_gold_segments(value), [])


class CoverageMathTest(unittest.TestCase):
    def test_identical_text_full_coverage(self):
        self.assertEqual(segment_coverage("the quick brown fox", "the quick brown fox jumps"), 1.0)

    def test_disjoint_text_zero_coverage(self):
        self.assertAlmostEqual(segment_coverage("alpha beta gamma", "x y z"), 0.0, places=2)

    def test_partial_overlap(self):
        rate = segment_coverage("the quick brown fox", "the slow brown fox")
        self.assertGreater(rate, 0.4)
        self.assertLess(rate, 0.9)

    def test_normalize_collapses_whitespace(self):
        self.assertEqual(normalize("  Hello\n World "), "hello world")


class CoverageSummaryTest(unittest.TestCase):
    def test_full_hit_verdict(self):
        summary = coverage_summary("the quick brown fox", "the quick brown fox jumps high")
        self.assertEqual(summary["verdict"], "✅ 命中")
        self.assertIn("覆盖率 100%", summary["text"])

    def test_no_retrieval(self):
        summary = coverage_summary("the quick brown fox", "")
        self.assertEqual(summary["verdict"], "❌ 未召回")
        self.assertIn("覆盖率 0%", summary["text"])

    def test_no_annotation(self):
        summary = coverage_summary("", "some retrieval text")
        self.assertEqual(summary["verdict"], "❌ 未召回")
        self.assertEqual(summary["text"], "无金标标注")

    def test_partial_verdict(self):
        annotation = "[Chunk 1] the quick brown fox\n[Chunk 2] completely different topic"
        retrieval = "the quick brown fox runs"
        summary = coverage_summary(annotation, retrieval)
        self.assertIn(summary["verdict"], ("✅ 命中", "⚠️ 部分命中"))

    def test_mixed_segments_partial(self):
        annotation = '[{"chunk_id": "Chunk 1", "source_text": "the quick brown fox"}, {"chunk_id": "Chunk 2", "source_text": "zzz zzz zzz zzz"}]'
        summary = coverage_summary(annotation, "the quick brown fox")
        self.assertEqual(summary["verdict"], "⚠️ 部分命中")
        self.assertEqual(summary["segments"], 2)
        self.assertEqual(summary["covered"], 1)


if __name__ == "__main__":
    unittest.main()

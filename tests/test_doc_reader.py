"""Tests for doc_reader: multi-file original-doc reading, budget truncation,
binary placeholders, and index-based resolution (mirrors kb-qa-eval's
test_review_multifile.py for the report side).
"""

import os
import shutil
import tempfile
import unittest

import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))

from doc_reader import build_index, find_doc, read_docs_with_budget, read_from_index


class _TmpDir:
    """Fixed workspace-local temp dir; avoids OS temp-dir walk stalls under
    AV/sandbox (os.walk can block indefinitely on %TEMP% on this box)."""

    def __init__(self, prefix: str):
        self.path = os.path.join(os.path.dirname(os.path.abspath(__file__)), prefix)

    def __enter__(self):
        shutil.rmtree(self.path, ignore_errors=True)
        os.makedirs(self.path, exist_ok=True)
        return self.path

    def __exit__(self, *exc):
        shutil.rmtree(self.path, ignore_errors=True)
        return False


class ReadDocsWithBudgetTest(unittest.TestCase):
    def setUp(self):
        self.tmp = _TmpDir(".dr_tmp_read")
        self.doc_dir = self.tmp.__enter__()
        for name, body in (("one.md", "doc one"), ("two.txt", "doc two"), ("three.pdf", "binary")):
            with open(os.path.join(self.doc_dir, name), "w", encoding="utf-8") as f:
                f.write(body)

    def tearDown(self):
        self.tmp.__exit__(None, None, None)

    def test_multiple_files_all_included(self):
        text = read_docs_with_budget([self.doc_dir], ["one.md", "two.txt"], budget=1000)
        self.assertIn("one.md", text)
        self.assertIn("two.txt", text)
        self.assertIn("doc one", text)
        self.assertIn("doc two", text)

    def test_missing_file_placeholder(self):
        text = read_docs_with_budget([self.doc_dir], ["one.md", "nope.md"], budget=1000)
        self.assertIn("one.md", text)
        self.assertIn("nope.md: (not found in doc dirs)", text)

    def test_budget_truncates_later_files_but_keeps_headers(self):
        big = ("x" * 1000)
        with open(os.path.join(self.doc_dir, "big.txt"), "w", encoding="utf-8") as f:
            f.write(big)
        text = read_docs_with_budget([self.doc_dir], ["one.md", "big.txt"], budget=100)
        self.assertIn("one.md", text)
        # second header present, content truncated to remaining budget
        self.assertIn("big.txt", text)
        self.assertLessEqual(len(text), 100 + 100)  # budget + some slack

    def test_binary_file_placeholder(self):
        text = read_docs_with_budget([self.doc_dir], ["three.pdf"], budget=1000)
        self.assertIn("[Binary file: three.pdf]", text)

    def test_empty_filenames(self):
        self.assertEqual(read_docs_with_budget([self.doc_dir], [], budget=1000), "")


class ReadFromIndexTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        for name, body in (("a.md", "alpha content"), ("b.txt", "beta content")):
            with open(os.path.join(self.tmp.name, name), "w", encoding="utf-8") as f:
                f.write(body)
        self.index = build_index([self.tmp.name])

    def tearDown(self):
        self.tmp.cleanup()

    def test_all_requested_files_read(self):
        text = read_from_index(self.index, ["a.md", "b.txt"])
        self.assertIn("alpha content", text)
        self.assertIn("beta content", text)
        self.assertIn("a.md", text)
        self.assertIn("b.txt", text)

    def test_stem_fallback_resolves_cross_extension(self):
        # report references x.pdf; index has x.md
        with open(os.path.join(self.tmp.name, "c.md"), "w", encoding="utf-8") as f:
            f.write("c markdown body")
        self.index = build_index([self.tmp.name])
        text = read_from_index(self.index, ["c.pdf"])
        self.assertIn("c markdown body", text)

    def test_missing_marks_not_found(self):
        text = read_from_index(self.index, ["missing.pdf"])
        self.assertIn("missing.pdf: (not found in doc dirs)", text)

    def test_pdf_placeholder_when_not_extractable(self):
        with open(os.path.join(self.tmp.name, "scan.pdf"), "wb") as f:
            f.write(b"\x00\x01\x02 not a real pdf")
        self.index = build_index([self.tmp.name])
        text = read_from_index(self.index, ["scan.pdf"])
        self.assertIn("[Binary file: scan.pdf]", text)

    def test_budget_bounds_total(self):
        text = read_from_index(self.index, ["a.md", "b.txt"], budget=20)
        self.assertLessEqual(len(text), 200)  # headers + 20 chars content + slack
        self.assertIn("alpha", text)  # first file content preserved at least partially


class FindDocTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        with open(os.path.join(self.tmp.name, "x.pdf"), "w", encoding="utf-8") as f:
            f.write("x")
        with open(os.path.join(self.tmp.name, "y.md"), "w", encoding="utf-8") as f:
            f.write("y")

    def tearDown(self):
        self.tmp.cleanup()

    def test_exact_basename_wins(self):
        found = find_doc([self.tmp.name], "x.pdf")
        self.assertEqual(os.path.basename(found), "x.pdf")

    def test_stem_fallback(self):
        found = find_doc([self.tmp.name], "x.md")  # x.pdf exists → stem matches
        self.assertIsNotNone(found)


class RankedInjectionTest(unittest.TestCase):
    """Gold-relevance ranked injection (annotation parameter)."""

    GOLD = "the golden needle answer sentence"

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = self.tmp.name
        # a/b: irrelevant fillers; c: contains the gold sentence
        with open(os.path.join(self.dir, "a.txt"), "w", encoding="utf-8") as f:
            f.write("A" * 300)
        with open(os.path.join(self.dir, "b.txt"), "w", encoding="utf-8") as f:
            f.write("betaUniqueMarker" + "B" * 300)
        with open(os.path.join(self.dir, "c.txt"), "w", encoding="utf-8") as f:
            f.write(self.GOLD)
        self.index = build_index([self.dir])
        self.files = ["a.txt", "b.txt", "c.txt"]

    def tearDown(self):
        self.tmp.cleanup()

    def test_ranking_injects_gold_file_first(self):
        # budget = exactly the gold file + one filler (joiner slack removed),
        # so the remaining irrelevant file must be dropped entirely
        two = read_from_index(self.index, ["a.txt", "c.txt"], budget=100000, annotation=self.GOLD)
        text = read_from_index(self.index, self.files, budget=len(two) - 5, annotation=self.GOLD)
        self.assertIn(self.GOLD, text)  # gold file survives despite being last
        self.assertNotIn("betaUniqueMarker", text)  # an irrelevant file was dropped

    def test_no_annotation_keeps_original_order(self):
        text = read_from_index(self.index, ["a.txt", "b.txt"], budget=100000)
        self.assertLess(text.index("a.txt"), text.index("b.txt"))
        self.assertNotIn("金标覆盖", text)

    def test_gold_nowhere_keeps_original_order(self):
        text = read_from_index(
            self.index, ["a.txt", "b.txt"], budget=100000, annotation="zzz absent text qqq"
        )
        self.assertLess(text.index("a.txt"), text.index("b.txt"))
        self.assertIn("betaUniqueMarker", text)

    def test_coverage_header_label(self):
        text = read_from_index(self.index, self.files, budget=100000, annotation=self.GOLD)
        self.assertIn("金标覆盖", text)
        # the label belongs to the gold file's header
        header_line = next(l for l in text.splitlines() if "c.txt" in l)
        self.assertIn("金标覆盖", header_line)

    def test_truncated_file_keeps_best_paragraphs(self):
        # big doc: gold sentence lives in a late paragraph
        paras = ["intro paragraph filler " + "x" * 130]
        for i in range(1, 10):
            paras.append(f"filler paragraph {i} " + "y" * 128)
        paras[7] = f"paragraph seven carries {self.GOLD} inside " + "z" * 90
        big_path = os.path.join(self.dir, "big.md")
        with open(big_path, "w", encoding="utf-8") as f:
            f.write("\n\n".join(paras))
        index = build_index([self.dir])
        budget = 420
        # control: head truncation (no annotation) loses the late gold paragraph
        head = read_from_index(index, ["big.md"], budget=budget)
        self.assertNotIn(self.GOLD, head)
        # ranked excerpt keeps paragraph 0 plus the gold paragraph
        ranked = read_from_index(index, ["big.md"], budget=budget, annotation=self.GOLD)
        self.assertIn(self.GOLD, ranked)
        self.assertLessEqual(len(ranked), budget + 100)  # budget respected (header slack)

    def test_saturated_giant_file_cannot_swallow_budget(self):
        # giant doc whose paragraphs weakly overlap the gold everywhere
        # (file-level coverage saturates) must not starve the small file
        # that carries the exact gold text
        filler_paras = []
        for i in range(60):
            filler_paras.append(f"section {i} " + " ".join(self.GOLD.split()[i % 5:] * 8) + " q" * 40)
        with open(os.path.join(self.dir, "giant.md"), "w", encoding="utf-8") as f:
            f.write("\n\n".join(filler_paras))
        with open(os.path.join(self.dir, "exact.md"), "w", encoding="utf-8") as f:
            f.write(self.GOLD)
        index = build_index([self.dir])
        # giant first (tie-safe original order), tight budget
        text = read_from_index(
            index, ["giant.md", "exact.md"], budget=8000, annotation=self.GOLD
        )
        self.assertIn("exact.md", text)
        self.assertIn(self.GOLD + "\n", text + "\n")  # exact gold file injected intact
        # giant's excerpt is capped (12 paragraphs + intro), far below budget
        giant_block = text.split("- exact.md")[0]
        self.assertLess(len(giant_block), 8000)


if __name__ == "__main__":
    unittest.main()
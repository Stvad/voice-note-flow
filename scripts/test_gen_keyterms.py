#!/usr/bin/env python3
"""Tests for the keyterm filtering logic.

Run: python3 -m unittest discover -s scripts -p 'test_*.py'

Uses a small inline wordlist rather than /usr/share/dict/words so the
expectations stay readable and the tests do not depend on the host dictionary.
"""
import tempfile
import unittest
from pathlib import Path

from gen_keyterms import (
    build_lists,
    is_generic,
    load_terms_file,
    review_candidates,
)

# Stand-in for /usr/share/dict/words. "roam", "tana" and "tap" are in here on
# purpose: the real dictionary contains them, and they are exactly the
# collisions the filter has to avoid.
WORDS = {
    "clothing", "dance", "practice", "deliberate", "community", "algorithm",
    "swing", "whip", "roam", "tana", "tap", "obsidian", "buy", "to", "notes",
}


class TestIsGeneric(unittest.TestCase):
    def test_ordinary_lowercase_words_are_generic(self):
        for term in ["clothing", "dance", "practice", "community"]:
            self.assertTrue(is_generic(term, WORDS), term)

    def test_multiword_lowercase_phrases_are_generic(self):
        self.assertTrue(is_generic("deliberate practice", WORDS))

    def test_capitalised_terms_are_kept_even_when_in_the_dictionary(self):
        # The whole reason the filter is case-sensitive: web2 contains these.
        for term in ["Roam", "Obsidian", "Tana", "Dance"]:
            self.assertFalse(is_generic(term, WORDS), term)

    def test_acronyms_are_kept(self):
        self.assertFalse(is_generic("TAP", WORDS))

    def test_unknown_lowercase_coinages_are_kept(self):
        for term in ["zettelkasten", "habryka", "westie", "wildbow"]:
            self.assertFalse(is_generic(term, WORDS), term)

    def test_hierarchical_tags_are_never_generic(self):
        # "to/buy" and "wcs/whip" are both entirely lowercase dictionary words,
        # but the slash marks deliberate structure the prompt relies on.
        self.assertTrue(is_generic("buy", WORDS))
        self.assertFalse(is_generic("to/buy", WORDS))
        self.assertFalse(is_generic("wcs/whip", WORDS))

    def test_empty_wordlist_disables_the_filter(self):
        self.assertFalse(is_generic("clothing", set()))

    def test_blank_input(self):
        self.assertFalse(is_generic("   ", WORDS))


class TestReviewCandidates(unittest.TestCase):
    def test_surfaces_capitalised_dictionary_words(self):
        self.assertEqual(review_candidates(["Roam", "Obsidian"], WORDS),
                         ["Roam", "Obsidian"])

    def test_omits_terms_the_filter_already_handles_or_should_not_touch(self):
        aliases = ["clothing", "TAP", "Roam Research", "Murphyjitsu", "to/buy"]
        self.assertEqual(review_candidates(aliases, WORDS), [])


class TestLoadTermsFile(unittest.TestCase):
    def _write(self, text):
        path = Path(tempfile.mkdtemp()) / "terms.txt"
        path.write_text(text, encoding="utf-8")
        return path

    def test_parses_lines_commas_and_comments(self):
        path = self._write(
            "# manual additions\n"
            "Elli Haugen\n"
            "Foo Bar, Baz Qux   # trailing comment\n"
            "\n"
            "   \n"
            "Solo Term\n"
        )
        self.assertEqual(
            load_terms_file(path),
            {"Elli Haugen", "Foo Bar", "Baz Qux", "Solo Term"},
        )


class TestBuildListsPinning(unittest.TestCase):
    def test_pinned_term_survives_the_usage_threshold(self):
        entries = [("Popular", 50), ("Rare Name", 1)]
        link, _, _ = build_lists(entries, threshold=12, max_link_bytes=8800,
                                 pinned=["Rare Name"])
        self.assertIn("Rare Name", link)
        self.assertIn("Popular", link)

    def test_pinned_term_absent_from_the_dump_is_still_added(self):
        entries = [("Popular", 50)]
        link, _, _ = build_lists(entries, threshold=12, max_link_bytes=8800,
                                 pinned=["Brand New Person"])
        self.assertIn("Brand New Person", link)

    def test_size_trim_drops_organic_terms_before_pinned_ones(self):
        entries = [(f"Common Name {i}", 100 - i) for i in range(40)]
        link, _, _ = build_lists(entries, threshold=1, max_link_bytes=120,
                                 pinned=["Common Name 39"])
        self.assertLessEqual(len(",".join(link).encode()), 120)
        self.assertIn("Common Name 39", link, "pinned term must outrank usage")

    def test_effective_threshold_ignores_pinned_terms(self):
        entries = [("Popular", 50), ("Rare Name", 1)]
        _, _, eff = build_lists(entries, threshold=12, max_link_bytes=8800,
                                pinned=["Rare Name"])
        self.assertEqual(eff, 50, "a pinned outlier must not report as the cutoff")


if __name__ == "__main__":
    unittest.main()

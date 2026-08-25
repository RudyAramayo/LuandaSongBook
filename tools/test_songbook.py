#!/usr/bin/env python3
"""Regression checks for songbook source loading and publication links."""

from __future__ import annotations

import unittest

from build_songbook import ROOT, load_entries, stanzas, url_domain


class SongbookLoaderTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.entries = load_entries()

    def test_all_song_and_quadra_opening_lines_are_preserved(self) -> None:
        for entry in self.entries:
            if entry.category == "Rhythms":
                continue
            source_first_line = (ROOT / entry.source_path).read_text(encoding="utf-8-sig").splitlines()[0].strip()
            self.assertEqual(entry.lines[0], source_first_line, entry.source_path)

    def test_capoeira_e_luta_opens_with_complete_chorus(self) -> None:
        entry = next(item for item in self.entries if item.title == "Capoeira e luta de mandingueiro")
        self.assertEqual(
            stanzas(entry.lines)[0],
            [
                "Capoeira e luta de mandingueiro",
                "E luta de nego nagô",
                "Angola que jogou seu pastinha",
                "Regional mestre Bimba criou",
            ],
        )

    def test_unavailable_legacy_domain_is_not_published(self) -> None:
        published_domains = {url_domain(url) for entry in self.entries for url in entry.urls}
        self.assertNotIn("capoeira-music.net", published_domains)
        self.assertEqual(published_domains, {"youtu.be"})

    def test_expected_unique_entry_counts(self) -> None:
        counts = {
            category: sum(entry.category == category for entry in self.entries)
            for category in ("Songs", "Quadras de Bimba", "Rhythms")
        }
        self.assertEqual(counts, {"Songs": 36, "Quadras de Bimba": 7, "Rhythms": 8})


if __name__ == "__main__":
    unittest.main()

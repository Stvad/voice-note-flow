#!/usr/bin/env python3
"""Generate KEYTERMS (Claude linking, comprehensive) and ACOUSTIC_KEYTERMS
(Deepgram boost, focused) from a Roam alias CSV dump.

The CSV is expected to have columns: alias, block_id, usage_count,
distinct_sources, via_property_field, block_total_refs, num_aliases_on_block,
block_types, content_preview.

Heuristic:
  LINK list  : dedupe by block_id, drop pure noise (single-char/symbol
               aliases), drop comma-bearing aliases (would break our
               comma-separated config), drop explicit generics.
               Threshold: usage >= --threshold (default 12). The list is
               then trimmed from the lowest-usage end until its joined
               UTF-8 size fits under --max-link-bytes (default 8800).
               Apps Script PropertiesService caps individual values at 9KB
               and silently drops oversized writes.
  ACOUSTIC   : tight subset for transcription boost — personal names,
               unusual proper nouns, distinctive jargon. Filters out
               two-word "name" candidates whose tokens are all common
               English (those phrases dilute the acoustic boost without
               helping). Cap: 95 (Deepgram Nova-3 limit is 100).

Usage:
  scripts/gen_keyterms.py <input.csv> [--threshold N] [--out-dir DIR]

Writes:
  <out-dir>/keyterms_link.txt
  <out-dir>/keyterms_acoustic.txt

Stats printed to stderr.
"""
import argparse
import csv
import re
import sys
from pathlib import Path

# Always drop
EXCLUDE = {
    "w", "->", ">", "<", "/",
    "algorithm",  # explicitly excluded as too generic
}

BRACKET_NOISE_RE = re.compile(r"\[\[")


def keep_alias(alias: str) -> bool:
    a = alias.strip()
    if not a: return False
    if a in EXCLUDE: return False
    if "," in a: return False  # would break comma-separated config
    if BRACKET_NOISE_RE.search(a): return False
    if len(a) == 1: return False
    if not any(c.isalnum() for c in a): return False
    return True


# Words that, when they make up an entire two-word "name", indicate the phrase
# is just common English and won't benefit from acoustic boost.
COMMON_EN = {
    "san", "francisco", "new", "york", "south", "east", "west", "north",
    "bay", "area", "city", "old", "swing", "modern", "after", "party",
    "open", "source", "vitamin", "fish", "oil", "personal", "capital",
    "deliberate", "practice", "burning", "man", "fiction", "moonlit",
    "moves", "digital", "wellbeing", "google", "maps", "kaiser",
    "permanente", "rationality", "community", "easter", "secular",
    "solstice", "seated", "pull", "down", "chest", "press", "row",
    "shoulder", "behavioral", "interview", "spaced", "repetition",
    "tools", "for", "thought", "evergreen", "notes", "the", "of",
    "and", "to", "in", "on", "at", "a", "an", "is",
    "claude", "code", "fest", "festival", "camp", "forum",
}

# Multi-word phrases the transcriber handles natively — no acoustic value,
# would only dilute boost.
HARD_ACOUSTIC_BLOCKLIST = {
    "SF Bay Area", "San Francisco", "New York", "South Bay", "East Bay",
    "Bay Area", "Burning Man", "Open Source", "Vitamin D", "Fish Oil",
    "Personal Capital", "Deliberate Practice", "Digital Wellbeing",
    "Google Maps", "Kaiser Permanente", "Rationality Community",
    "The After Party", "Fiction T", "Moonlit Moves", "Secular Solstice",
    "Seated Pull Down", "Seated Chest Press", "Seated Row",
    "Seated Shoulder Press", "Simon Willison's Weblog",
    "Modern Swing Forum", "Behavioral Interview", "Spaced Repetition",
    "Tools for thought", "Evergreen notes", "Claude Code",
    "Easter Swing", "Choreo Camp", "Two Left Feet", "SO Swing",
    "Capital Swing", "Cross Flow Festival", "Boogie By the Bay",
    "Monterey Swing Fest", "Halloween Swingthing", "All Star SwingJam",
    "Calypso Pacific", "Partner Dance Adventures", "Prague Fall Season",
    "Wild Wild Westie",
}

# Distinctive single-word or short-phrase proper nouns to always include in
# the acoustic list (regardless of usage), since the transcriber reliably
# mangles them.
DISTINCTIVE_SINGLES_OR_PHRASES = {
    "habryka", "johnswentworth", "Aella", "Raemon", "wildbow",
    "Murphyjitsu", "Beeminder", "Modafinil", "Aeropress", "RescueTime",
    "Karabiner", "Promnesia", "CrowdAnki", "Mexifold", "Penultima",
    "Solaris", "Pollycast", "HPMOR", "Swingtacular", "Roam Toolkit",
    "Roam Research", "Toggl", "Waterpik", "Melatonin", "ChatGPT",
    "Tana", "Dnipro", "Kyiv", "WCS", "CFAR", "LessWrong", "LWCW",
    "IntelliJ", "Hypothes.is", "Manifold.markets", "Readwise.io",
    "Interviewing.io", "Lunchclub", "Oura ring", "Pale Lights",
    "80000hours.org", "conceptually.org", "inkandswitch.com",
    "asteriskmag.com", "SwingLiteracy.com", "Genesis House",
    "Mission City Swing", "Westie Pirates SF", "Jack & Jill O'Rama",
    "Programmable attention", "zettelkasten", "TAP",
}


def is_likely_personal_name(s: str) -> bool:
    toks = s.split()
    if len(toks) < 2 or len(toks) > 4: return False
    if not all(t[:1].isalpha() and t[:1].isupper() for t in toks): return False
    if all(t.lower().strip(".,'`") in COMMON_EN for t in toks): return False
    return True


def build_lists(entries, threshold: int, max_link_bytes: int):
    # Initial filter by usage threshold
    candidates = [(a, u) for (a, u) in entries if u >= threshold]
    # Apps Script PropertiesService caps each value at 9KB (per
    # https://developers.google.com/apps-script/guides/services/quotas)
    # and silently truncates oversized writes. Trim from the bottom
    # (lowest usage first, since `entries` is sorted desc) until we fit.
    def joined_bytes(items):
        return len(",".join(a for a, _ in items).encode("utf-8"))
    while candidates and joined_bytes(candidates) > max_link_bytes:
        candidates.pop()
    link_terms = [a for (a, _) in candidates]
    effective_threshold = candidates[-1][1] if candidates else None

    acoustic, seen = [], set()
    # Pass 1: high-usage personal names that aren't common English phrases
    for alias, usage in entries:
        if usage < 18: break
        if alias in seen or alias in HARD_ACOUSTIC_BLOCKLIST: continue
        if is_likely_personal_name(alias) and len(alias) < 40:
            acoustic.append(alias); seen.add(alias)
    # Pass 2: distinctive proper-noun jargon (regardless of usage)
    for alias in DISTINCTIVE_SINGLES_OR_PHRASES:
        if alias not in seen:
            acoustic.append(alias); seen.add(alias)
    acoustic = acoustic[:95]

    return link_terms, acoustic, effective_threshold


def main():
    p = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    p.add_argument("csv", type=Path, help="Path to the Roam alias CSV dump")
    p.add_argument("--threshold", type=int, default=12,
                   help="Min usage_count for LINK list inclusion (default 12)")
    p.add_argument("--max-link-bytes", type=int, default=8800,
                   help="Max UTF-8 byte size of joined LINK list (default 8800; "
                        "Apps Script per-value cap is 9KB)")
    p.add_argument("--out-dir", type=Path, default=Path.cwd(),
                   help="Output directory (default: cwd)")
    args = p.parse_args()

    by_block = {}
    with args.csv.open(newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            alias = row["alias"].strip()
            try:
                usage = int(row["usage_count"])
            except ValueError:
                continue
            if not keep_alias(alias):
                continue
            block_id = row["block_id"]
            cur = by_block.get(block_id)
            if cur is None or cur[1] < usage:
                by_block[block_id] = (alias, usage)

    entries = sorted(by_block.values(), key=lambda x: -x[1])
    link_terms, acoustic, effective_threshold = build_lists(
        entries, args.threshold, args.max_link_bytes)

    args.out_dir.mkdir(parents=True, exist_ok=True)
    link_path = args.out_dir / "keyterms_link.txt"
    acoustic_path = args.out_dir / "keyterms_acoustic.txt"
    link_path.write_text(",".join(link_terms))
    acoustic_path.write_text(",".join(acoustic))

    eff = f" (effective usage cutoff: {effective_threshold})" if effective_threshold else ""
    print(f"LINK list:     {len(link_terms):>4} entries -> {link_path} "
          f"({link_path.stat().st_size} bytes{eff})", file=sys.stderr)
    print(f"ACOUSTIC list: {len(acoustic):>4} entries -> {acoustic_path} "
          f"({acoustic_path.stat().st_size} bytes)", file=sys.stderr)


if __name__ == "__main__":
    main()

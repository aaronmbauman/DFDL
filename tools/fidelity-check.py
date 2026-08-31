"""Fidelity differential for the DFDL spec conversion.

Compares the normalised text inventory of the MS-Word source against the
inventory of the converted Metanorma output and reports *everything* that does
not account for.  This is deliberately a differential, not a search for a
handful of normative phrases: counting occurrences of, say, "It is a Schema
Definition Error" passes happily while an entire property table is dropped.

Three accountings are produced over the whole of both documents, because they
fail in different ways:

  1. **Unit alignment** - the two inventories are aligned with
     ``difflib.SequenceMatcher``; unmatched runs are reported as MISSING (in
     source, not in target), ADDED (in target, not in source) or CHANGED (a
     run that was rewritten), each with its clause, the heading it sits under
     and its neighbouring units, so it can be found in either document.

     Alignment runs over typographically folded text, because that is what
     finds the *correspondence* between the two documents reliably.  The
     verdict on each aligned pair is then taken on the strict text, so a pair
     the alignment calls equal is still compared character by character.

  2. **Token accounting** - a multiset difference over every token of both
     documents: words *and* punctuation.  This is immune to alignment noise:
     if a table is dropped, its tokens show up here even when the surrounding
     alignment is confused, and material that merely *moved* nets out to zero
     here while showing up as a MISSING/ADDED pair above.

  3. **Character census** - per-character counts of the syntax-significant
     characters on both sides, with deltas.  Unit and token accounting are
     both local; a systematic substitution applied document-wide shows up
     here as a single large asymmetry (``#`` 210 -> 98, ``'`` 2704 -> 149)
     even when no individual difference looks alarming.

Every difference is classified so a reviewer can triage it:

  TYPOGRAPHIC  quote or dash *style* only - ``'`` -> ``’``, ``-`` -> ``–``.
               Correct ISO typesetting.  Warns, never fails.
  REFLOW       the same characters, re-split across a different number of
               units.  Advisory.
  SUBSTITUTION content-bearing characters replaced: ``<=`` -> ``⇐``,
               ``0x55`` -> ``0×55``, ``--`` -> ``—``, a ``#`` dropped out of
               a dfdl:textNumberPattern, a regex left with unbalanced
               parentheses.  Fails.
  STRUCTURAL   text present on one side and absent on the other.  Fails.
  CODE         any difference inside a source block, where the text is the
               normative artefact and nothing about it is cosmetic.  Fails.

Exit status is non-zero on SUBSTITUTION, STRUCTURAL and CODE, and when the
token accounting does not balance.  TYPOGRAPHIC and REFLOW differences are
reported and warned about, but pass.

**Normalisation modes.**  The gate runs strict: only Unicode NFC, whitespace
collapse and zero-width removal, because those are the only things that
genuinely must differ between a .docx and Metanorma XML.  ``--relaxed``
additionally folds smart quotes, dash variants and ellipses on *both* sides
before comparing.  Relaxed is useful for finding structural loss without
typographic noise, and that is all it is for: folding both sides means a
target that replaced 2,555 apostrophes, turned every ``--`` into an em dash
and every ``<=`` into ``⇐`` compares identical to its source.  The count of
units that relaxed normalisation would have called identical, and strict does
not, is printed either way.

Text the renderer generates rather than carries is removed from both sides
before comparing: Word's caption paragraphs carry the "Table 7" that
Metanorma numbers for itself, a reference to a table or a figure carries the
same number in the sentence around it, and Word resolves "on page ii"
against a pagination Metanorma does over again.  A reference to a *clause*
is compared, not removed: both documents number the clauses the same way,
and the inventory writes an ``<xref>`` out as the number of the clause it
points at.

Exit status: 0 when the target accounts for every source character that is
not pure typography, 1 when content differs, 2 on usage errors.

Usage:
    python3 tools/fidelity-check.py SOURCE.docx TARGET.xml
    python3 tools/fidelity-check.py SOURCE.docx TARGET.xml --summary
    python3 tools/fidelity-check.py SOURCE.docx TARGET.xml --relaxed  # advisory

SOURCE/TARGET may also be inventory .tsv files previously written by
tools/text-inventory.py, which is useful for large runs.
"""

from __future__ import annotations

import argparse
import collections
import contextlib
import difflib
import importlib.util
import re
import sys
from dataclasses import dataclass
from pathlib import Path

MAX_LINE = 160


def load_inventory_module():
    """Import the sibling text-inventory.py (its name is not importable)."""
    path = Path(__file__).resolve().with_name("text-inventory.py")
    if not path.exists():
        raise SystemExit(f"cannot find {path}")
    spec = importlib.util.spec_from_file_location("dfdl_text_inventory", path)
    if spec is None or spec.loader is None:
        raise SystemExit(f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    # Do not leave a tools/__pycache__ behind in the checkout.
    previous = sys.dont_write_bytecode
    sys.dont_write_bytecode = True
    try:
        spec.loader.exec_module(module)
    finally:
        sys.dont_write_bytecode = previous
    return module


INV = load_inventory_module()


def clip(text: str, width: int = MAX_LINE) -> str:
    if len(text) <= width:
        return text
    return text[: width - 3] + "..."


# --------------------------------------------------------------------------
# Generated text
# --------------------------------------------------------------------------

# "Section 11.2.1", "Table 7", "Appendix A" and friends.  Word writes the
# label it resolved a cross-reference to and Metanorma writes its own -
# "Section 11" against "Clause 11" - so the label goes from both sides.
REFERENCE_WORDS = "Section|Clause|Subclause|Annex|Appendix|Table|Figure"
REFERENCE_NUMBER = r"(?:[0-9]+|[A-Z])(?:\.[0-9A-Za-z]+)*(?:-[0-9A-Za-z]+)?"
REFERENCE_RE = re.compile(
    rf"\b(?P<word>{REFERENCE_WORDS})s?\s+(?P<number>{REFERENCE_NUMBER})"
    rf"(?![0-9A-Za-z])",
    re.IGNORECASE,
)

# Tables and figures are the two things Word and Metanorma number in
# different sequences, so their numbers go along with the label.  Clause and
# annex numbers are the same in both, and are compared rather than removed.
RENUMBERED_LABELS = {"table", "figure"}

CAPTION_LABEL_RE = re.compile(rf"^(?:Table|Figure)\s+({REFERENCE_NUMBER})[.:]?\s+")

# What is left of "See section 12.1.2." once the number has gone.  An
# AsciiDoc cross-reference may or may not have the word in front of it, so
# the word alone, with nothing after it, is numbering too.  So is a page
# number, which Metanorma paginates for itself.
ORPHAN_REFERENCE_RE = re.compile(
    rf"\b({REFERENCE_WORDS}|page)\s*(?=[,.;:)\]]|$)", re.IGNORECASE
)
PAGE_NUMBER_RE = re.compile(r"\b(page)\s+([0-9]+|[ivxlcdm]+)\b", re.IGNORECASE)

# Headings carry clause titles, not references, and code is compared verbatim.
LITERAL_KINDS = {"code"}


def title_key(title: str) -> str:
    """Normalised form of a heading, for recognising it inside a reference.

    Folded typographically on purpose: a heading whose apostrophe was
    restyled is still the same heading, and Word spells the old styling into
    every cross-reference that names it.  This is a locator, not a verdict;
    the units themselves are still compared strictly.
    """
    return " ".join(INV.typographic_fold(title).casefold().split())


class Generated:
    """Removes renderer-generated numbering and reference text."""

    def __init__(self, units) -> None:
        self.titles: dict[str, str] = {}
        self.clause_numbers: set[str] = set()
        for unit in units:
            if unit.kind.startswith("heading"):
                self.titles.setdefault(unit.section, title_key(unit.text))
                self.clause_numbers.add(unit.section)
                continue
            if unit.kind != "caption":
                continue
            label = CAPTION_LABEL_RE.match(unit.text)
            if label:
                self.titles[label.group(1)] = title_key(unit.text[label.end() :])

    def strip(self, kind: str, text: str) -> str:
        if kind.startswith("heading") or kind in LITERAL_KINDS:
            return text
        if kind == "caption":
            # The caption keeps its own title; only Word's number goes.
            return CAPTION_LABEL_RE.sub("", text).strip()
        stripped = PAGE_NUMBER_RE.sub(r"\1", REFERENCE_RE.sub(self.strip_label, text))
        return " ".join(ORPHAN_REFERENCE_RE.sub("", stripped).split())

    def strip_label(self, match: re.Match) -> str:
        """What is left of a reference once the renderer's label has gone.

        The number stays where the two documents number the thing the same
        way, so that it is compared rather than taken on trust.  A table or a
        figure is renumbered, and so is a reference Word resolved against its
        own numbering of the annexes, and those numbers go with the label.
        """
        word, number = match.group("word").lower(), match.group("number")
        if word in RENUMBERED_LABELS or number not in self.clause_numbers:
            return ""
        return number


def comparable(units, generated: Generated) -> list[str]:
    return [generated.strip(unit.kind, unit.text) for unit in units]


# --------------------------------------------------------------------------
# Accounting
# --------------------------------------------------------------------------


def word_counter(texts) -> collections.Counter:
    counter: collections.Counter = collections.Counter()
    for text in texts:
        counter.update(INV.words(text))
    return counter


def token_counter(texts) -> collections.Counter:
    """Multiset of every token - words *and* punctuation characters.

    The old accounting counted only ``[A-Za-z0-9]`` runs, so ``#``, ``<``,
    ``>``, ``=``, ``|``, ``\\``, ``%``, ``'`` and ``"`` were invisible to it:
    a target that turned ``"#,##,##0"`` into ``",,0"`` balanced perfectly.
    Tokens are typographically folded, so pure ISO restyling still nets out to
    zero and only content-bearing punctuation moves the needle.
    """
    counter: collections.Counter = collections.Counter()
    for text in texts:
        counter.update(INV.tokens(text))
    return counter


def split_counter(counter: collections.Counter):
    """Split a token multiset into its word part and its punctuation part."""
    words: collections.Counter = collections.Counter()
    symbols: collections.Counter = collections.Counter()
    for token, count in counter.items():
        if token[0].isalnum():
            words[token] = count
        else:
            symbols[token] = count
    return words, symbols


def align(source_texts, target_texts):
    """Return the alignment opcodes between two lists of comparable texts."""
    matcher = difflib.SequenceMatcher(None, source_texts, target_texts, autojunk=False)
    return matcher.get_opcodes()


# --------------------------------------------------------------------------
# Classification
# --------------------------------------------------------------------------

TYPOGRAPHIC = "TYPOGRAPHIC"
REFLOW = "REFLOW"
SUBSTITUTION = "SUBSTITUTION"
STRUCTURAL = "STRUCTURAL"
CODE = "CODE"

# Only these fail the build.  Smart quotes in prose are correct ISO
# typesetting, and a paragraph re-split across two units has lost nothing.
FAILING = (SUBSTITUTION, STRUCTURAL, CODE)
CLASSES = (TYPOGRAPHIC, REFLOW, SUBSTITUTION, STRUCTURAL, CODE)

# Glyphs a "prettifier" leaves behind, mapped back to the ASCII they replaced.
# Undoing these on both sides tells a substitution apart from a rewrite: if
# the two sides agree once the glyphs are expanded, the only thing that
# happened is that ASCII source syntax was turned into typographic symbols -
# which for a spec full of ``<=``, ``->`` and ``0x55`` is corruption.
SUBSTITUTIONS = {
    "⇐": "<=",  # leftwards double arrow
    "⇒": "=>",  # rightwards double arrow
    "⇔": "<=>",  # left right double arrow
    "←": "<-",  # leftwards arrow
    "→": "->",  # rightwards arrow
    "↔": "<->",  # left right arrow
    "≤": "<=",  # less-than or equal
    "≥": ">=",  # greater-than or equal
    "≠": "!=",  # not equal
    "×": "x",  # multiplication sign (0x55 -> 0x55)
    "—": "--",  # em dash (-- -> em dash)
    "–": "-",  # en dash
    "…": "...",  # horizontal ellipsis
    "•": "*",  # bullet
}

_SUBSTITUTION_TABLE = {ord(glyph): ascii_ for glyph, ascii_ in SUBSTITUTIONS.items()}


def unsubstitute(text: str) -> str:
    """Expand typographic symbols back into the ASCII they stand in for."""
    return text.translate(_SUBSTITUTION_TABLE)


def canonical(text: str) -> str:
    """Fully de-prettified text: substitutions expanded, then quotes folded."""
    return INV.typographic_fold(unsubstitute(text))


def classify(source_text: str, target_text: str, is_code: bool) -> tuple[str, str]:
    """Classify one difference.  Returns (class, one-line reason)."""
    if source_text == target_text:
        return "", ""
    detail = _classify_detail(source_text, target_text)
    if is_code:
        return CODE, f"inside a source block; {detail[1]}"
    return detail


def _classify_detail(source_text: str, target_text: str) -> tuple[str, str]:
    if INV.typographic_fold(source_text) == INV.typographic_fold(target_text):
        return TYPOGRAPHIC, "quote/dash style only"
    if canonical(source_text) == canonical(target_text):
        return SUBSTITUTION, "ASCII syntax replaced by a typographic symbol"
    src_words = collections.Counter(INV.words(source_text))
    tgt_words = collections.Counter(INV.words(target_text))
    if src_words == tgt_words:
        return SUBSTITUTION, "punctuation differs, words identical"
    return STRUCTURAL, "text present on one side and absent on the other"


@dataclass
class Difference:
    """One classified difference, with everything needed to find it."""

    kind: str
    reason: str
    section: str
    heading: str
    unit_kind: str
    source_text: str
    target_text: str
    hidden_by_relaxed: bool = False

    @property
    def fails(self) -> bool:
        return self.kind in FAILING


def is_code_unit(*units) -> bool:
    return any(unit is not None and unit.kind == "code" for unit in units)


def char_edits(source_text: str, target_text: str, limit: int = 6) -> list[str]:
    """The differing character runs, with a little context, for triage."""
    matcher = difflib.SequenceMatcher(None, source_text, target_text, autojunk=False)
    edits: list[str] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        left = source_text[max(0, i1 - 12) : i1]
        right = source_text[i2 : i2 + 12]
        edits.append(
            f"...{left}[{source_text[i1:i2]}]{right}... -> "
            f"...{left}[{target_text[j1:j2]}]{right}..."
        )
        if len(edits) >= limit:
            edits.append("...")
            break
    return edits


# --------------------------------------------------------------------------
# Reporting
# --------------------------------------------------------------------------


def collect_differences(opcodes, src_units, tgt_units, src_texts, tgt_texts):
    """Every classified difference between the two aligned inventories.

    ``opcodes`` come from the typographically folded alignment; the texts
    passed in are the *strict* ones, so a run the alignment calls ``equal`` is
    still compared character by character here.  That is where the differences
    the old checker reported as identical units live.

    A unit that merely moved - Word puts a table caption after the table,
    Metanorma puts the name first - is told apart from one that vanished by
    looking for its exact text among the units the alignment inserted.
    """
    differences: list[Difference] = []
    deleted: list[int] = []
    inserted: list[int] = []

    def add(src_index, tgt_index, source_text, target_text, hidden, missing=False):
        src_unit = src_units[src_index] if src_index is not None else None
        tgt_unit = tgt_units[tgt_index] if tgt_index is not None else None
        anchor = src_unit or tgt_unit
        code = is_code_unit(src_unit, tgt_unit)
        if missing:
            # A whole unit present on one side only is structural by
            # definition, whatever its text happens to be made of.
            side = "source" if src_index is not None else "target"
            kind = CODE if code else STRUCTURAL
            reason = f"a whole unit exists only in the {side}"
        else:
            kind, reason = classify(source_text, target_text, code)
        if not kind:
            return
        differences.append(
            Difference(
                kind=kind,
                reason=reason,
                section=anchor.section,
                heading=heading_before(
                    src_units, src_index if src_index is not None else 0
                ),
                unit_kind=anchor.kind,
                source_text=source_text,
                target_text=target_text,
                hidden_by_relaxed=hidden,
            )
        )

    for tag, i1, i2, j1, j2 in opcodes:
        if tag == "equal":
            for offset in range(i2 - i1):
                src_index, tgt_index = i1 + offset, j1 + offset
                if src_texts[src_index] != tgt_texts[tgt_index]:
                    add(
                        src_index,
                        tgt_index,
                        src_texts[src_index],
                        tgt_texts[tgt_index],
                        True,
                    )
        elif tag == "delete":
            deleted.extend(range(i1, i2))
        elif tag == "insert":
            inserted.extend(range(j1, j2))
        elif (i2 - i1) == (j2 - j1):
            for offset in range(i2 - i1):
                add(
                    i1 + offset,
                    j1 + offset,
                    src_texts[i1 + offset],
                    tgt_texts[j1 + offset],
                    False,
                )
        else:
            # A run re-split across a different number of units: judge the
            # joined text, so a pure reflow does not read as lost content.
            joined_src = " ".join(src_texts[i1:i2])
            joined_tgt = " ".join(tgt_texts[j1:j2])
            if joined_src == joined_tgt:
                differences.append(
                    Difference(
                        kind=REFLOW,
                        reason=f"{i2 - i1} source unit(s) re-split as {j2 - j1}",
                        section=src_units[i1].section,
                        heading=heading_before(src_units, i1),
                        unit_kind=src_units[i1].kind,
                        source_text=joined_src,
                        target_text=joined_tgt,
                    )
                )
            else:
                add(i1, j1, joined_src, joined_tgt, False)

    # A unit that left one place and reappeared unchanged elsewhere moved; it
    # did not go missing.  Matching is on the exact strict text, and the token
    # accounting balances either way.
    available: dict[str, list[int]] = {}
    for tgt_index in inserted:
        available.setdefault(tgt_texts[tgt_index], []).append(tgt_index)
    matched_targets: set[int] = set()
    for src_index in deleted:
        candidates = available.get(src_texts[src_index])
        if candidates:
            tgt_index = candidates.pop()
            matched_targets.add(tgt_index)
            differences.append(
                Difference(
                    kind=REFLOW,
                    reason="moved elsewhere in the document, text unchanged",
                    section=src_units[src_index].section,
                    heading=heading_before(src_units, src_index),
                    unit_kind=src_units[src_index].kind,
                    source_text=src_texts[src_index],
                    target_text=tgt_texts[tgt_index],
                )
            )
            continue
        add(src_index, None, src_texts[src_index], "", False, missing=True)
    for tgt_index in inserted:
        if tgt_index not in matched_targets:
            add(None, tgt_index, "", tgt_texts[tgt_index], False, missing=True)
    return differences


def count_classes(differences) -> collections.Counter:
    counter: collections.Counter = collections.Counter()
    for difference in differences:
        counter[difference.kind] += 1
    return counter


def print_census(src_texts, tgt_texts, opcodes, out) -> int:
    """Per-character counts of the syntax-significant characters, with deltas.

    Unit and token accounting are both local.  A substitution applied across
    the whole document shows up here, and only here, as one large asymmetry.
    Returns the number of characters whose counts do not match.
    """
    src_counts = INV.census(src_texts)
    tgt_counts = INV.census(tgt_texts)
    print("\nCHARACTER CENSUS over the whole document:", file=out)
    print(
        f"  {'char':<8} {'source':>8} {'target':>8} {'delta':>8}   name",
        file=out,
    )
    mismatched = 0
    for label, chars in INV.CENSUS_GROUPS:
        rows = []
        for char in dict.fromkeys(chars):
            source, target = src_counts[char], tgt_counts[char]
            if not source and not target:
                continue
            delta = target - source
            if delta:
                mismatched += 1
            name = INV.CHAR_NAMES.get(char, "")
            flag = "  <== asymmetric" if delta else ""
            rows.append(
                f"  {char!r:<8} {source:>8} {target:>8} {delta:>+8}   {name}{flag}"
            )
        if rows:
            print(f"  -- {label}", file=out)
            for row in rows:
                print(row, file=out)
    if not mismatched:
        print("  every census character balances", file=out)
    print_delimiter_balance(src_texts, tgt_texts, opcodes, out)
    return mismatched


DELIMITERS = (("(", ")"), ("[", "]"), ("{", "}"))


def unbalanced(text: str) -> list[str]:
    """Delimiter pairs that do not close in ``text``."""
    return [
        f"{opener}{closer}"
        for opener, closer in DELIMITERS
        if text.count(opener) != text.count(closer)
    ]


def print_delimiter_balance(src_texts, tgt_texts, opcodes, out) -> None:
    """Aligned units whose brackets balance in the source but not the target.

    A normative regex or expression that lost a parenthesis is still made of
    the same words, so only a character-level look finds it.
    """
    broken = []
    for tag, i1, i2, j1, j2 in opcodes:
        if tag not in ("equal", "replace") or (i2 - i1) != (j2 - j1):
            continue
        for offset in range(i2 - i1):
            source_text = src_texts[i1 + offset]
            target_text = tgt_texts[j1 + offset]
            pairs = unbalanced(target_text)
            if pairs and not unbalanced(source_text):
                broken.append((target_text, pairs))
    if not broken:
        print(
            "  delimiters: every aligned unit closes its brackets as the "
            "source does",
            file=out,
        )
        return
    print(
        f"  delimiters: {len(broken)} unit(s) unbalanced in the target only:",
        file=out,
    )
    for target_text, pairs in broken[:10]:
        print(f"    {' '.join(pairs)}  {clip(target_text, 110)}", file=out)


def print_classified(
    differences, out, limit: int, gate_hidden: bool, detail: bool = True
) -> None:
    """The classified differences, worst class first."""
    counts = count_classes(differences)
    print("\nCLASSIFIED DIFFERENCES:", file=out)
    for name in CLASSES:
        verdict = "FAILS" if name in FAILING else "warns"
        print(f"  {name:<13} {counts.get(name, 0):>6}   {verdict}", file=out)
    hidden = [d for d in differences if d.hidden_by_relaxed]
    if hidden:
        note = "advisory under --relaxed" if gate_hidden else "counted above"
        print(
            f"\n  {len(hidden)} of these sit in units the *relaxed* alignment "
            f"calls identical\n  ({note}): "
            f"{dict(count_classes(hidden))}",
            file=out,
        )

    if not detail:
        return
    # Worst first, so a truncated report still shows what actually fails.
    rank = {CODE: 0, STRUCTURAL: 1, SUBSTITUTION: 2, TYPOGRAPHIC: 3, REFLOW: 4}
    ranked = sorted(differences, key=lambda d: rank.get(d.kind, 99))
    for shown, difference in enumerate(ranked, start=1):
        if limit and shown > limit:
            print(
                f"\n  ... {len(ranked) - shown + 1} further classified "
                f"difference(s) suppressed (--max-report {limit})",
                file=out,
            )
            break
        advisory = " [relaxed-identical]" if difference.hidden_by_relaxed else ""
        print(
            f"\n  {difference.kind}{advisory} - clause {difference.section} "
            f"({difference.unit_kind}) - {difference.reason}",
            file=out,
        )
        print(f"      under: {clip(difference.heading, 100)}", file=out)
        if difference.kind in (TYPOGRAPHIC, SUBSTITUTION, CODE) or (
            difference.source_text and difference.target_text
        ):
            for edit in char_edits(difference.source_text, difference.target_text):
                print(f"      {clip(edit, 150)}", file=out)
        else:
            print(f"    - {clip(difference.source_text, 150)}", file=out)
            print(f"    + {clip(difference.target_text, 150)}", file=out)


def heading_before(units, index: int) -> str:
    for probe in range(min(index, len(units) - 1), -1, -1):
        unit = units[probe]
        if unit.kind.startswith("heading"):
            return f"{unit.section} {unit.text}"
    return "(front matter)"


def context_lines(units, start: int, stop: int, count: int) -> tuple[list, list]:
    before = units[max(0, start - count) : start]
    after = units[stop : min(len(units), stop + count)]
    return before, after


def span(start: int, stop: int) -> str:
    """Human-readable inventory range; an empty range names the insert point."""
    if stop <= start:
        return f"units none (position {start})"
    if stop - start == 1:
        return f"unit {start}"
    return f"units {start}..{stop - 1}"


def describe_block(units, start, stop, indent="      ") -> list[str]:
    lines = []
    for unit in units[start:stop]:
        lines.append(f"{indent}[{unit.section} {unit.kind}] {clip(unit.text)}")
    return lines


def report_difference(
    number, tag, source, target, i1, i2, j1, j2, context, out
) -> None:
    if tag in ("delete", "replace"):
        anchor_units, anchor_start, anchor_stop = source, i1, i2
    else:
        anchor_units, anchor_start, anchor_stop = target, j1, j2
    label = {
        "delete": "MISSING  (in source, absent from target)",
        "insert": "ADDED    (in target, absent from source)",
        "replace": "CHANGED  (rewritten between source and target)",
    }[tag]
    if anchor_start < len(anchor_units):
        section = anchor_units[anchor_start].section
    else:
        section = "?"
    print(f"\n[{number}] {label}", file=out)
    print(f"      clause {section} | under: {heading_before(source, i1)}", file=out)
    print(
        f"      source {span(i1, i2)} | target {span(j1, j2)}",
        file=out,
    )

    before, after = context_lines(anchor_units, anchor_start, anchor_stop, context)
    for unit in before:
        print(f"    . {clip(unit.text, 110)}", file=out)
    if tag in ("delete", "replace"):
        for line in describe_block(source, i1, i2, indent=""):
            print(f"    - {clip(line, 150)}", file=out)
    if tag in ("insert", "replace"):
        for line in describe_block(target, j1, j2, indent=""):
            print(f"    + {clip(line, 150)}", file=out)
    for unit in after:
        print(f"    . {clip(unit.text, 110)}", file=out)

    if tag == "replace":
        # Tokens, not words: a run whose only change was "#,##,##0" -> ",,0"
        # used to print "same words - reflowed only" and be waved through.
        src_tokens = INV.tokens(" ".join(u.text for u in source[i1:i2]))
        tgt_tokens = INV.tokens(" ".join(u.text for u in target[j1:j2]))
        lost = collections.Counter(src_tokens) - collections.Counter(tgt_tokens)
        gained = collections.Counter(tgt_tokens) - collections.Counter(src_tokens)
        if lost:
            print(
                f"      tokens only in source ({sum(lost.values())}): "
                f"{clip(' '.join(sorted(lost.elements())), 200)}",
                file=out,
            )
        if gained:
            print(
                f"      tokens only in target ({sum(gained.values())}): "
                f"{clip(' '.join(sorted(gained.elements())), 200)}",
                file=out,
            )
        if not lost and not gained:
            print(
                "      (same tokens - reflowed, re-split or restyled only)",
                file=out,
            )


def print_summary(
    src_units, tgt_units, src_texts, tgt_texts, opcodes, out, paths, mode
):
    src_tokens = token_counter(src_texts)
    tgt_tokens = token_counter(tgt_texts)
    lost_words, lost_symbols = split_counter(src_tokens - tgt_tokens)
    gained_words, gained_symbols = split_counter(tgt_tokens - src_tokens)
    src_words, _ = split_counter(src_tokens)
    tgt_words, _ = split_counter(tgt_tokens)

    missing_units = sum(i2 - i1 for tag, i1, i2, _, _ in opcodes if tag == "delete")
    added_units = sum(j2 - j1 for tag, _, _, j1, j2 in opcodes if tag == "insert")
    changed_src = sum(i2 - i1 for tag, i1, i2, _, _ in opcodes if tag == "replace")
    changed_tgt = sum(j2 - j1 for tag, _, _, j1, j2 in opcodes if tag == "replace")
    relaxed_equal = sum(i2 - i1 for tag, i1, i2, _, _ in opcodes if tag == "equal")
    strict_equal = 0
    for tag, i1, i2, j1, _ in opcodes:
        if tag != "equal":
            continue
        for offset in range(i2 - i1):
            if src_texts[i1 + offset] == tgt_texts[j1 + offset]:
                strict_equal += 1

    print("=" * 72, file=out)
    print(f"DFDL conversion fidelity summary  [{mode} normalisation]", file=out)
    print("=" * 72, file=out)
    print(f"  source              : {paths[0]}", file=out)
    print(f"  target              : {paths[1]}", file=out)
    print(f"  source units        : {len(src_units)}", file=out)
    print(f"  target units        : {len(tgt_units)}", file=out)
    print(
        f"  identical (strict)  : {strict_equal}"
        f"{share_of(strict_equal, len(src_units))}  <- the honest number",
        file=out,
    )
    print(
        f"  identical (relaxed) : {relaxed_equal}"
        f"{share_of(relaxed_equal, len(src_units))}  <- typography folded away",
        file=out,
    )
    print(
        f"  hidden by relaxed   : {relaxed_equal - strict_equal} "
        "unit(s) that differ character by character but that relaxed\n"
        "                        normalisation calls identical",
        file=out,
    )
    print(f"  missing units       : {missing_units}", file=out)
    print(f"  added units         : {added_units}", file=out)
    print(
        f"  changed units       : {changed_src} source / {changed_tgt} target",
        file=out,
    )
    print(f"  source words        : {sum(src_words.values())}", file=out)
    print(f"  target words        : {sum(tgt_words.values())}", file=out)
    print(
        f"  words only in src   : {sum(lost_words.values())} "
        f"({len(lost_words)} distinct)  <- content actually lost",
        file=out,
    )
    print(
        f"  words only in tgt   : {sum(gained_words.values())} "
        f"({len(gained_words)} distinct)  <- content actually added",
        file=out,
    )
    print(
        f"  punctuation only src: {sum(lost_symbols.values())} "
        f"({len(lost_symbols)} distinct)  <- syntax characters dropped",
        file=out,
    )
    print(
        f"  punctuation only tgt: {sum(gained_symbols.values())} "
        f"({len(gained_symbols)} distinct)  <- syntax characters introduced",
        file=out,
    )
    print(
        "  note                : punctuation counts are typographically folded, "
        "so a smart\n                        quote costs nothing here and "
        "'--' -> em dash costs one '-'.",
        file=out,
    )
    return {
        "lost": lost_words + lost_symbols,
        "gained": gained_words + gained_symbols,
        "strict_equal": strict_equal,
        "relaxed_equal": relaxed_equal,
    }


def share_of(part: int, whole: int) -> str:
    if not whole:
        return ""
    return f"  ({100.0 * part / whole:5.1f}% of source units)"


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Fidelity differential for the DFDL spec conversion.",
    )
    parser.add_argument(
        "source",
        metavar="SOURCE",
        help="the MS-Word source, or a .tsv inventory of it",
    )
    parser.add_argument(
        "target",
        metavar="TARGET",
        help="the built semantic XML, or a .tsv inventory of it",
    )
    parser.add_argument(
        "--summary",
        action="store_true",
        help="print counts only, no per-difference detail",
    )
    parser.add_argument(
        "--max-report",
        type=int,
        default=40,
        help="maximum number of differences to detail (default 40, 0 = all)",
    )
    parser.add_argument(
        "--context",
        type=int,
        default=1,
        help="units of surrounding context per difference (default 1)",
    )
    parser.add_argument(
        "--no-changed",
        action="store_true",
        help="only detail MISSING/ADDED runs, skip CHANGED ones",
    )
    parser.add_argument(
        "--top-words",
        type=int,
        default=25,
        help="how many unaccounted-for tokens to list (default 25)",
    )
    parser.add_argument(
        "--relaxed",
        action="store_true",
        help=(
            "advisory mode: fold smart quotes, dash variants and ellipses on "
            "both sides before judging a unit, so only structural loss is "
            "gated.  Useful, but it cannot see typographic corruption - do "
            "not use it as the gate."
        ),
    )
    parser.add_argument(
        "--no-census",
        action="store_true",
        help="skip the per-character census",
    )
    parser.add_argument("-o", "--output", help="write the report here")
    args = parser.parse_args(argv)
    mode = "relaxed" if args.relaxed else "strict"

    # Always read strictly.  Relaxed text is strict text with the typographic
    # fold applied, so one read carries both accountings and the two inventories
    # are guaranteed to be unit-for-unit the same.
    try:
        src_units = INV.read_any(args.source, mode="strict")
        tgt_units = INV.read_any(args.target, mode="strict")
    except FileNotFoundError as error:
        print(f"error: {error}", file=sys.stderr)
        return 2

    gen_src = Generated(src_units)
    gen_tgt = Generated(tgt_units)

    src_texts = comparable(src_units, gen_src)
    tgt_texts = comparable(tgt_units, gen_tgt)
    # Alignment runs on folded text - that is what finds the correspondence
    # between the two documents; the verdict below is taken on src_texts.
    src_folded = [INV.typographic_fold(text) for text in src_texts]
    tgt_folded = [INV.typographic_fold(text) for text in tgt_texts]

    with contextlib.ExitStack() as stack:
        out = sys.stdout
        if args.output:
            out = stack.enter_context(open(args.output, "w", encoding="utf-8"))
        opcodes = align(src_folded, tgt_folded)
        stats = print_summary(
            src_units,
            tgt_units,
            src_texts,
            tgt_texts,
            opcodes,
            out,
            (args.source, args.target),
            mode,
        )

        census_delta = 0
        if not args.no_census:
            census_delta = print_census(src_texts, tgt_texts, opcodes, out)

        differences = collect_differences(
            opcodes, src_units, tgt_units, src_texts, tgt_texts
        )
        gated = [
            difference
            for difference in differences
            if not (args.relaxed and difference.hidden_by_relaxed)
        ]
        classified_failures = [d for d in gated if d.fails]
        print_classified(
            differences,
            out,
            args.max_report,
            gate_hidden=args.relaxed,
            detail=not args.summary,
        )

        lost, gained = stats["lost"], stats["gained"]
        print("\nToken accounting over the whole document:", file=out)
        if lost or gained:
            print(
                f"  UNACCOUNTED FOR: {sum(lost.values())} token(s) only in the "
                f"source, {sum(gained.values())} only in the target",
                file=out,
            )
        else:
            print("  every token is accounted for on both sides", file=out)

        if not args.summary:
            shown = 0
            for tag, i1, i2, j1, j2 in opcodes:
                if tag == "equal":
                    continue
                if tag == "replace" and args.no_changed:
                    continue
                shown += 1
                if args.max_report and shown > args.max_report:
                    print(
                        f"\n... further differences suppressed "
                        f"(--max-report {args.max_report}); rerun with "
                        f"--max-report 0",
                        file=out,
                    )
                    break
                report_difference(
                    shown, tag, src_units, tgt_units, i1, i2, j1, j2, args.context, out
                )

            if args.top_words:
                for title, counter in (
                    ("Tokens present in source but not in target", stats["lost"]),
                    ("Tokens present in target but not in source", stats["gained"]),
                ):
                    if not counter:
                        continue
                    print(f"\n{title} (top {args.top_words}):", file=out)
                    for word, count in counter.most_common(args.top_words):
                        print(f"  {count:>5}  {word!r}", file=out)

        counts = count_classes(gated)
        warnings = counts.get(TYPOGRAPHIC, 0) + counts.get(REFLOW, 0)
        print(file=out)
        if warnings:
            print(
                f"WARNING: {counts.get(TYPOGRAPHIC, 0)} typographic and "
                f"{counts.get(REFLOW, 0)} reflow difference(s).  Smart quotes in "
                "prose are\n         correct ISO typesetting; these do not fail "
                "the check.",
                file=out,
            )
        if census_delta:
            print(
                f"WARNING: {census_delta} syntax character(s) do not balance "
                "across the two sides;\n         see the character census above.",
                file=out,
            )
        if classified_failures or lost or gained:
            reasons = []
            if classified_failures:
                detail = ", ".join(
                    f"{counts[name]} {name}" for name in FAILING if counts.get(name)
                )
                reasons.append(f"{len(classified_failures)} difference(s) - {detail}")
            if lost or gained:
                reasons.append("the token accounting does not balance")
            print(f"RESULT: FAIL - {'; '.join(reasons)}.", file=out)
            if args.relaxed:
                print(
                    "        (--relaxed is advisory; the gate is the default "
                    "strict mode.)",
                    file=out,
                )
        else:
            print(
                "RESULT: PASS - the target accounts for every source character "
                f"that is not\n        pure typography"
                f"{' (relaxed)' if args.relaxed else ''}.",
                file=out,
            )

    return 1 if (classified_failures or lost or gained) else 0


if __name__ == "__main__":
    sys.exit(main())

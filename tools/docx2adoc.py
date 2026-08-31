"""Convert the DFDL specification MS-Word source into AsciiDoc for Metanorma.

The specification is authored in Word.  This script reads the .docx package
directly (it is a zip of XML parts) and emits AsciiDoc suitable for Metanorma's
ISO backend, one file per top-level clause.

Clauses are addressed by the specification's own numbering, which the tool
reconstructs from Word's heading numbering: ``front`` for the ten unnumbered
front-matter headings, ``1``..``29`` for the numbered clauses, and ``A``..``G``
for the appendices.

Usage:

    python3 tools/docx2adoc.py --list
    python3 tools/docx2adoc.py --clause 13 --out spec/clauses/13-simple-types.adoc
    python3 tools/docx2adoc.py --clause C --out spec/clauses/C-string-literals.adoc
    python3 tools/docx2adoc.py --all --outdir spec/clauses/

Only the Python standard library is used.

The script deliberately carries no shebang and is not executable: ``make
check`` runs every executable in tools/ as a validator against the built
semantic XML, and this is a conversion utility, not a validator.
"""

from __future__ import annotations

import argparse
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

# --------------------------------------------------------------------------
# OOXML namespaces
# --------------------------------------------------------------------------

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def q(prefix: str, tag: str) -> str:
    """Return a Clark-notation qualified name, e.g. q("w", "p")."""
    return f"{{{NS[prefix]}}}{tag}"


W_P = q("w", "p")
W_TBL = q("w", "tbl")
W_TR = q("w", "tr")
W_TC = q("w", "tc")
W_VAL = q("w", "val")

# --------------------------------------------------------------------------
# Text normalisation
# --------------------------------------------------------------------------

#: Word's typographic characters, mapped to something unambiguous in AsciiDoc.
CHAR_MAP = {
    "\u2018": "'",
    "\u2019": "'",
    "\u201a": "'",
    "\u201b": "'",
    "\u201c": '"',
    "\u201d": '"',
    "\u201e": '"',
    "\u00a0": " ",  # non-breaking space
    "\u202f": " ",  # narrow no-break space
    "\u2011": "-",  # non-breaking hyphen
    "\u00ad": "",  # soft hyphen
    "\ufeff": "",
    "\u2028": " ",
    "\u2029": " ",
}

PRIVATE_USE_RE = re.compile(r"[\ue000-\uf8ff]")

#: Heading 1 titles that open an appendix, e.g. "Appendix C: ...".
APPENDIX_RE = re.compile(r"^\s*(?:Appendix|Annex)\s+([A-Z])\b[.:]?")

#: Key used for the Heading 1 clauses that carry no clause number.
FRONT_MATTER = "front"


def normalise_text(text: str) -> str:
    """Apply the character substitutions that are safe everywhere."""
    for src, dst in CHAR_MAP.items():
        text = text.replace(src, dst)
    return PRIVATE_USE_RE.sub("", text)


def slugify(text: str, fallback: str = "id") -> str:
    """Turn arbitrary text into a valid, readable AsciiDoc identifier."""
    text = normalise_text(text)
    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^A-Za-z0-9]+", "-", text).strip("-").lower()
    text = re.sub(r"-{2,}", "-", text)
    if len(text) > 60:
        text = text[:60].rstrip("-")
    if not text:
        text = fallback
    if not text[0].isalpha():
        text = "x-" + text
    return text


# --------------------------------------------------------------------------
# Package access
# --------------------------------------------------------------------------


class Package:
    """Read-only view of the parts of a .docx that this converter needs."""

    def __init__(self, path: Path):
        self.path = path
        self.zip = zipfile.ZipFile(path)
        self.document = self._parse("word/document.xml")

    def _parse(self, name: str) -> ET.Element:
        with self.zip.open(name) as handle:
            return ET.parse(handle).getroot()


# --------------------------------------------------------------------------
# Document model helpers
# --------------------------------------------------------------------------


def para_style(para) -> str:
    pPr = para.find(q("w", "pPr"))
    if pPr is None:
        return ""
    style = pPr.find(q("w", "pStyle"))
    return style.get(W_VAL) if style is not None else ""


def heading_level(style: str):
    """Return the AsciiDoc nesting level for a Word heading style, else None."""
    match = re.fullmatch(r"Heading([1-9])", style)
    if match:
        return int(match.group(1))
    if style == "AppendixH1":
        return 1
    return None


def word_numbered(para) -> bool:
    """True unless the paragraph explicitly switches its style numbering off.

    The front matter uses the Heading 1 style with ``numId`` 0, which is how
    Word suppresses the clause number.
    """
    pPr = para.find(q("w", "pPr"))
    if pPr is None:
        return True
    numPr = pPr.find(q("w", "numPr"))
    if numPr is None:
        return True
    num_id = numPr.find(q("w", "numId"))
    return not (num_id is not None and num_id.get(W_VAL) == "0")


def raw_text(elem) -> str:
    """All literal text under an element, ignoring deletions."""
    parts = []
    for node in elem.iter():
        if node.tag == q("w", "delText"):
            continue
        if node.tag == q("w", "t"):
            parts.append(node.text or "")
        elif node.tag == q("w", "tab"):
            parts.append(" ")
    return normalise_text("".join(parts))


# --------------------------------------------------------------------------
# Conversion
# --------------------------------------------------------------------------


@dataclass
class Clause:
    """One top-level (Heading 1) clause of the specification.

    ``key`` is the specification's own identifier for the clause: ``front``
    for the unnumbered front matter, ``"1"``..``"29"`` for the numbered
    clauses, and ``"A"``..``"G"`` for the appendices.
    """

    key: str
    ordinal: int  # position within its group, used only for file names
    title: str
    start: int
    end: int = 0
    blocks: list = field(default_factory=list)

    @property
    def is_front(self) -> bool:
        return self.key == FRONT_MATTER

    @property
    def is_annex(self) -> bool:
        return len(self.key) == 1 and self.key.isalpha()


class Converter:
    def __init__(self, package: Package):
        self.pkg = package
        self.body = package.document[0]
        self.clauses = self._split_clauses()
        self.current = None

    # -- setup ------------------------------------------------------------

    def _split_clauses(self) -> list:
        """Cut the body at every Heading 1 and label each piece.

        Word numbers the headings itself, and turns numbering off for the
        front matter, so the specification's own clause numbers have to be
        reconstructed: front matter, then 1..n, then the lettered appendices.
        """
        clauses = []
        counter = 0
        front = 0
        first = next(
            (
                index
                for index, block in enumerate(self.body)
                if block.tag == W_P and heading_level(para_style(block)) == 1
            ),
            0,
        )
        if first:
            # Title page, copyright and abstract sit before the first heading.
            front += 1
            clauses.append(Clause(FRONT_MATTER, front, "Preamble", 0))
        for index, block in enumerate(self.body):
            if block.tag != W_P or heading_level(para_style(block)) != 1:
                continue
            title = raw_text(block).strip()
            letter = APPENDIX_RE.match(title)
            if letter:
                key, ordinal = letter.group(1), 0
            elif not word_numbered(block):
                front += 1
                key, ordinal = FRONT_MATTER, front
            else:
                counter += 1
                key, ordinal = str(counter), counter
            clauses.append(Clause(key, ordinal, title, index))
        blocks = list(self.body)
        for position, clause in enumerate(clauses):
            clause.end = (
                clauses[position + 1].start
                if position + 1 < len(clauses)
                else len(blocks)
            )
            clause.blocks = blocks[clause.start : clause.end]
        return clauses

    def find(self, key: str):
        """All clauses selected by a key: one clause, or the front matter."""
        wanted = key.strip().upper() if len(key.strip()) == 1 else key.strip().lower()
        return [
            clause for clause in self.clauses if clause.key.lower() == wanted.lower()
        ]

    def render_blocks(self, blocks, depth=0) -> list:
        # Style mapping arrives with the next commit; --list needs none of it.
        raise NotImplementedError

    # -- entry points -----------------------------------------------------

    def convert(self, clauses, resolvable=None) -> str:
        """Render one or more clauses, resolving xrefs within ``resolvable``."""
        self.emitted_clauses = set(
            resolvable if resolvable else [clause.key for clause in clauses]
        )
        lines = []
        for clause in clauses:
            self.current = clause
            lines.extend(self.render_blocks(clause.blocks))
            lines.append("")
        return "\n".join(trim_blank_lines(lines)).rstrip() + "\n"

    def clause_filename(self, clause: Clause) -> str:
        title = APPENDIX_RE.sub("", clause.title).strip(" :-") or clause.title
        if clause.is_front:
            return f"front-{clause.ordinal:02d}-{slugify(clause.title)}.adoc"
        if clause.ordinal:
            return f"{clause.ordinal:02d}-{slugify(clause.title)}.adoc"
        return f"{clause.key}-{slugify(title)}.adoc"


def trim_blank_lines(lines) -> list:
    """Collapse runs of blank lines and strip leading/trailing ones."""
    out = []
    for line in lines:
        if not line.strip():
            if not out or not out[-1]:
                continue
            out.append("")
        else:
            out.append(line.rstrip())
    while out and not out[-1]:
        out.pop()
    return out


# --------------------------------------------------------------------------
# Command line
# --------------------------------------------------------------------------

DEFAULT_DOCX = "docs/current/draft-gwdrp-dfdl-v1.2.2-GFD-R-P.240-ISO-23415.docx"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=__doc__.splitlines()[0],
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Clauses are named as the specification names them: 1..29 for the\n"
            "numbered clauses, A..G for the appendices, and 'front' for the\n"
            "unnumbered front matter."
        ),
    )
    parser.add_argument("--docx", default=DEFAULT_DOCX, help="source .docx")
    parser.add_argument(
        "--clause",
        metavar="KEY",
        help="convert one clause: a number (13), an appendix letter (C), " "or 'front'",
    )
    parser.add_argument("--out", help="output file for --clause")
    parser.add_argument("--all", action="store_true", help="convert every clause")
    parser.add_argument("--outdir", help="output directory for --all")
    parser.add_argument(
        "--list", action="store_true", help="list the top-level clauses and exit"
    )
    return parser


def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    docx = Path(args.docx)
    if not docx.exists():
        print(f"no such file: {docx}", file=sys.stderr)
        return 2
    converter = Converter(Package(docx))

    if args.list:
        for clause in converter.clauses:
            print(f"{clause.key:>5}  {clause.title}")
        return 0

    if args.all:
        outdir = Path(args.outdir or "spec/clauses")
        outdir.mkdir(parents=True, exist_ok=True)
        every = [clause.key for clause in converter.clauses]
        for clause in converter.clauses:
            text = converter.convert([clause], resolvable=every)
            path = outdir / converter.clause_filename(clause)
            path.write_text(text, encoding="utf-8")
            print(f"{path} ({text.count(chr(10))} lines)")
        return 0

    if args.clause:
        selected = converter.find(args.clause)
        if not selected:
            known = ", ".join(dict.fromkeys(c.key for c in converter.clauses))
            print(f"unknown clause {args.clause!r}; known: {known}", file=sys.stderr)
            return 2
        text = converter.convert(selected)
        if args.out:
            path = Path(args.out)
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(text, encoding="utf-8")
            print(f"{path} ({text.count(chr(10))} lines)")
        else:
            sys.stdout.write(text)
        return 0

    parser.print_help()
    return 1


if __name__ == "__main__":
    sys.exit(main())

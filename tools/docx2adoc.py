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
# Style tables
# --------------------------------------------------------------------------

#: Paragraph styles dropped outright; Metanorma builds its own table of contents.
DROP_STYLES = {
    "TOC1",
    "TOC2",
    "TOC3",
    "TOC4",
    "TOC5",
    "TOC6",
    "TOC7",
    "TOC8",
    "TOC9",
    "TOCHeading",
    "TableofFigures",
    "Index1",
    "Index2",
    "Index3",
}

#: Paragraph styles rendered as verbatim listing blocks, mapped to a language.
CODE_STYLES = {
    "Codeblock0": "xml",
    "CodeBlock": "xml",
    "Code": "xml",
    "XMLexample": "xml",
    "XMLExcerpt": "xml",
    "HTMLPreformatted": "",
    "dataexample": "",
}

#: Paragraph styles that mark a list item even without explicit numbering.
LIST_STYLES = {"ListParagraph", "ListBullet", "ListNumber", "List", "BulletList"}

#: Paragraph styles that carry the caption of the table or figure beside them.
CAPTION_STYLES = {"Caption", "TableCaption", "FigureCaption", "TableTitle"}

#: Run styles rendered as inline monospace.
CODE_RUN_STYLES = {"CodeCharacter", "SourceText", "CodeblockChar0", "HTMLCode"}

#: Run styles rendered as italic / bold.
ITALIC_RUN_STYLES = {"Emphasis", "XMLExcerptEmphasis"}
BOLD_RUN_STYLES = {"Strong"}

#: Word numbering formats that mean "ordered list".
ORDERED_FORMATS = {
    "decimal",
    "lowerLetter",
    "upperLetter",
    "lowerRoman",
    "upperRoman",
    "ordinal",
}

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

#: Line-initial markers that would otherwise be read as AsciiDoc block syntax.
BLOCK_START_RE = re.compile(
    r"""^(
          \.[^\s.]      # block title
        | [*.-]\ +      # list item
        | =+\ +         # section title
        | //            # comment
        | \|            # table cell
        | \[(?!\[)      # block attribute list
        | :\S+:         # attribute entry
        | \+$           # list continuation
        | ={4,}$|-{4,}$|\.{4,}$|_{4,}$|\*{4,}$   # block delimiter
    )""",
    re.VERBOSE,
)

#: `{word}` looks like an AsciiDoc attribute reference and must be escaped.
ATTR_REF_RE = re.compile(r"\{([A-Za-z0-9_][A-Za-z0-9_-]*)\}")

#: A plus sign on its own, which pairs with the next one into a passthrough
#: and takes the text between them with it.
LONE_PLUS_RE = re.compile(r"(?<![\w+])\+(?![\w+])")

#: Constrained inline formatting pairs that we want rendered literally.
CONSTRAINED_PAIRS = [
    re.compile(r"(?<![\w\\])(\*)(\S(?:[^*\n]*\S)?)(\*)(?!\w)"),
    re.compile(r"(?<![\w\\])(_)(\S(?:[^_\n]*\S)?)(_)(?!\w)"),
]

PRIVATE_USE_RE = re.compile(r"[\ue000-\uf8ff]")

# --------------------------------------------------------------------------
# Protecting technical text
#
# Prose reaches the reader through two rewriting passes, and normative text
# has to survive both.  AsciiDoc reads `#` as highlight markup, `(((` as an
# index term and `*`/`_` as formatting; Metanorma then "smart formats" the
# resulting XML, turning `0x55` into 0\u00d755, ` - ` into an em dash, `<=` into
# \u21d0, `...` into \u2026 and decoding `&apos;` to a bare quote.
#
# Only two contexts escape both passes: Metanorma's straightquotes
# passthrough, which leaves the text plain, and a monospace span whose body
# is an AsciiDoc passthrough, since `tt` is exempt from the smart formatter.
# Neither can be applied wholesale without disabling legitimate formatting,
# so the individual hazardous sequences are found and wrapped one by one.
# --------------------------------------------------------------------------

#: Sequences that one of the two passes would rewrite if left unprotected.
HAZARD_RE = re.compile(
    r"""
      (?P<entity> &(?:[A-Za-z][A-Za-z0-9]{1,9}|\#[0-9]{1,6}|\#[xX][0-9A-Fa-f]{1,5}); )
    | (?P<hash> [^\s"'()\[\]]*\#[^\s"'()\[\]]* )   # AsciiDoc highlight markup
    | (?P<hex> \b0[xX][0-9A-Fa-f]+ )               # 0x55 -> 0\u00d755
    | (?P<arrow> <-{1,2}|<={1,2}|-{1,2}>|={1,2}> )  # -> => <- <= and doubles
    | (?P<ellipsis> \.{3,} )                       # ... -> \u2026
    | (?P<emdash> -{2,} )                          # -- -> \u2014, ---- -> \u2014\u2014
    | (?P<index> \({3}|\){3} )                     # (((term))) index macro
    | (?P<symbol> \w*\((?:[CcRr]|TM|tm)\) )        # (C) -> \u00a9, ICU(R) -> ICU\u00ae
    """,
    re.VERBOSE,
)

#: A URL, where a `#` is a fragment marker that AsciiDoc already respects.
URL_RE = re.compile(r"\b(?:https?|ftp|mailto):[^\s\[\]]+")

#: A spaced hyphen, plus the brackets needed to tell an operator from prose.
ARITH_DASH_RE = re.compile(r"""[()\[\]]|(?<=[\w)\]'"]) (-{1,2}) (?=[\w(\['"])""")

#: The wrappers emitted below, so that later passes can leave them alone.
PROTECTED_RE = re.compile(
    r"pass-format:straightquotes\[(?:\\.|[^\]])*\]" r"|`{1,2}\+{2,3}.*?\+{2,3}`{1,2}"
)


def xml_escape(text: str) -> str:
    """Escape the three characters that may not appear raw in Metanorma XML."""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


#: The curly pair Metanorma would have made of each straight quote.
QUOTE_PAIRS = {"'": ("\u2018", "\u2019"), '"': ("\u201c", "\u201d")}


def needs_monospace(token: str) -> bool:
    """True for text a straightquotes passthrough cannot carry intact."""
    return "&" in token or "|" in token


def protect(token: str) -> str:
    """Wrap one hazardous sequence so that neither pass can rewrite it.

    Entity notation has to survive Metanorma decoding ``&apos;`` back to a
    quote, which only a monospace span does; the same span is the safest
    home for anything holding a cell separator.  Everything else keeps its
    ordinary prose appearance inside a straightquotes passthrough.
    """
    if needs_monospace(token):
        return code_passthrough(token)
    escaped = token.replace("]", "\\]")
    return f"pass-format:straightquotes[{escaped}]"


def code_passthrough(text: str, tight: bool = True) -> str:
    """Render text verbatim as monospace, immune to both rewriting passes."""
    mark = "``" if tight else "`"
    # An entity has to be escaped by hand and passed through untouched:
    # AsciiDoc's own escaping rewrites `&apos;` as a numeric character
    # reference, which is the very notation the text is talking about.
    plussed = text.startswith("+") or text.endswith("+") or "++" in text
    if "&" not in text and not plussed:
        return f"{mark}++{text}++{mark}"
    # Nothing is substituted inside a raw passthrough, so a plus sign that
    # would otherwise run into the delimiter is written as a reference.
    body = xml_escape(text).replace("+", "&#43;")
    return f"{mark}+++{body}+++{mark}"


def hazard_spans(text: str) -> list:
    """Locate every sequence in ``text`` that has to be protected."""
    spans = [match.span() for match in HAZARD_RE.finditer(text) if match.group()]
    spans.extend(arith_dash_spans(text))
    for pattern in CONSTRAINED_PAIRS:
        for match in pattern.finditer(text):
            spans.append(match.span(1))
            spans.append(match.span(3))
    spans.sort()
    urls = [match.span() for match in URL_RE.finditer(text)]
    out = []
    for start, end in spans:
        if out and start < out[-1][1]:
            continue
        # A link macro carries its own quoting; protecting part of one only
        # breaks the macro apart.
        if any(low < end and start < high for low, high in urls):
            continue
        out.append((start, end))
    return out


def arith_dash_spans(text: str) -> list:
    """Spaced hyphens that are subtraction rather than an em dash.

    Metanorma rewrites every ``a - b`` to an em dash, which is right for the
    glossary's "term - definition" lines and wrong inside a formula.  A
    bracketed context, two numeric operands, or an operand that no English
    sentence could hold mark out the arithmetic ones.
    """
    spans = []
    depth = 0
    for match in ARITH_DASH_RE.finditer(text):
        token = match.group()
        if token in "([":
            depth += 1
        elif token in ")]":
            depth = max(0, depth - 1)
        elif depth or is_operator_context(text, match.start(), match.end()):
            spans.append(match.span(1))
    return spans


def is_operator_context(text: str, start: int, end: int) -> bool:
    """True if the operands either side of a spaced hyphen are not prose."""
    before = text[:start].rsplit(" ", 1)[-1]
    after = text[end:].split(" ", 1)[0]
    if before[-1:].isdigit() and after[:1].isdigit():
        return True
    return any(mark in before + after for mark in ("\\", ".."))


#: Heading 1 titles that open an appendix, e.g. "Appendix C: ...".
APPENDIX_RE = re.compile(r"^\s*(?:Appendix|Annex)\s+([A-Z])\b[.:]?")

#: Key used for the Heading 1 clauses that carry no clause number.
FRONT_MATTER = "front"


def normalise_text(text: str) -> str:
    """Apply the character substitutions that are safe everywhere."""
    for src, dst in CHAR_MAP.items():
        text = text.replace(src, dst)
    return PRIVATE_USE_RE.sub("", text)


def escape_plain(text: str) -> str:
    """Neutralise the inline markup that a backslash is enough to disarm."""
    text = text.replace("`", "\\`")
    text = ATTR_REF_RE.sub(r"\\{\1}", text)
    text = LONE_PLUS_RE.sub("{plus}", text)
    return text.replace("<<", "\\<<")


def escape_inline(text: str) -> str:
    """Escape AsciiDoc inline syntax in text that came from plain Word runs.

    Sequences that survive a backslash, or that Metanorma rewrites after
    AsciiDoc has finished, are wrapped in a passthrough instead; the rest of
    the run is left alone so that legitimate formatting still applies.
    """
    out = []
    position = 0
    for start, end in hazard_spans(text):
        quote = text[start - 1 : start]
        # A literal in quotation marks loses the pairing of those marks to
        # the passthrough that now sits between them, so the pair is taken
        # inside it and curled here instead.
        if (
            start > position
            and quote in QUOTE_PAIRS
            and text[end : end + 1] == quote
            and not needs_monospace(text[start:end])
        ):
            opening, closing = QUOTE_PAIRS[quote]
            out.append(escape_plain(text[position : start - 1]))
            out.append(protect(opening + text[start:end] + closing))
            position = end + 1
            continue
        out.append(escape_plain(text[position:start]))
        out.append(protect(text[start:end]))
        position = end
    out.append(escape_plain(text[position:]))
    return "".join(out)


def protect_line_start(line: str) -> str:
    """Stop a paragraph's first characters being read as block syntax."""
    if BLOCK_START_RE.match(line):
        return "{empty}" + line
    return line


def code_span(text: str) -> str:
    """Render text as an inline monospace span.

    Monospace is not verbatim in AsciiDoc: the body is still substituted, so
    a regular expression inside one loses to the index-term macro and a hex
    literal to Metanorma's multiplication sign.  The body therefore always
    goes through a passthrough.
    """
    if not text:
        return ""
    # The unconstrained delimiter fires everywhere, including next to the
    # quotation marks that surround most of the literals in this document.
    return code_passthrough(text)


def unspaced(text: str) -> str:
    """Escape a super/subscript body, whose delimiters cannot span a space."""
    return escape_inline(text).replace(" ", "{nbsp}")


def emphasis(text: str, mark: str, tight: bool = False) -> str:
    """Wrap text in an italic or bold pair, doubled when it abuts a word."""
    pair = mark * 2 if tight else mark
    return f"{pair}{escape_inline(text)}{pair}"


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
        self.numbering = self._parse_optional("word/numbering.xml")

    def _parse(self, name: str) -> ET.Element:
        with self.zip.open(name) as handle:
            return ET.parse(handle).getroot()

    def _parse_optional(self, name: str):
        try:
            return self._parse(name)
        except KeyError:
            return None


class Numbering:
    """Resolve a (numId, ilvl) pair to 'ordered' or 'bullet'."""

    def __init__(self, root):
        self.levels = {}
        self.num_to_abstract = {}
        if root is None:
            return
        for abstract in root.findall(q("w", "abstractNum")):
            aid = abstract.get(q("w", "abstractNumId"))
            for lvl in abstract.findall(q("w", "lvl")):
                fmt = lvl.find(q("w", "numFmt"))
                self.levels[(aid, lvl.get(q("w", "ilvl")))] = (
                    fmt.get(W_VAL) if fmt is not None else None
                )
        for num in root.findall(q("w", "num")):
            abstract = num.find(q("w", "abstractNumId"))
            if abstract is not None:
                self.num_to_abstract[num.get(q("w", "numId"))] = abstract.get(W_VAL)

    def kind(self, num_id: str, ilvl: str) -> str:
        aid = self.num_to_abstract.get(num_id)
        fmt = self.levels.get((aid, ilvl))
        return "ordered" if fmt in ORDERED_FORMATS else "bullet"


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


def numbering_of(para):
    """Return (numId, ilvl) for a numbered paragraph, else None."""
    pPr = para.find(q("w", "pPr"))
    if pPr is None:
        return None
    numPr = pPr.find(q("w", "numPr"))
    if numPr is None:
        return None
    num_id = numPr.find(q("w", "numId"))
    ilvl = numPr.find(q("w", "ilvl"))
    if num_id is None or num_id.get(W_VAL) == "0":
        return None
    return (num_id.get(W_VAL), ilvl.get(W_VAL) if ilvl is not None else "0")


def is_list_para(para) -> bool:
    return numbering_of(para) is not None or para_style(para) in LIST_STYLES


def is_hidden(run) -> bool:
    """True for a run Word hides.

    The specification is written with the editors' notes to each other in
    hidden text, so that the working copy carries them and the published
    document does not.
    """
    rPr = run.find(q("w", "rPr"))
    if rPr is None:
        return False
    vanish = rPr.find(q("w", "vanish"))
    return vanish is not None and vanish.get(W_VAL) not in ("0", "false")


def run_text(run) -> str:
    """The literal text of one run, ignoring deletions."""
    parts = []
    for node in run:
        if node.tag == q("w", "t"):
            parts.append(node.text or "")
        elif node.tag == q("w", "tab"):
            parts.append(" ")
    return "".join(parts)


def raw_text(elem) -> str:
    """All visible literal text under an element, ignoring deletions."""
    parts = [run_text(run) for run in elem.iter(q("w", "r")) if not is_hidden(run)]
    return normalise_text("".join(parts))


def hidden_comment(para) -> list:
    """The hidden text of a paragraph, as AsciiDoc comment lines."""
    parts = [run_text(run) for run in para.iter(q("w", "r")) if is_hidden(run)]
    text = " ".join(normalise_text("".join(parts)).split())
    return [f"// {text}"] if text else []


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
        self.numbering = Numbering(package.numbering)
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

    # -- inline rendering -------------------------------------------------

    def render_inline(self, container) -> str:
        """Render the runs of a paragraph (or hyperlink) to AsciiDoc text."""
        segments = []
        self._render_children(container, segments, [])
        return self._join_segments(segments)

    def _join_segments(self, segments, plain=False) -> str:
        """Merge adjacent same-format segments, then apply inline markup."""
        if plain:
            segments = [
                (kind if kind == "raw" else "plain", text) for kind, text in segments
            ]
        merged = []
        for kind, text in segments:
            if merged and merged[-1][0] == kind and kind != "raw":
                merged[-1] = (kind, merged[-1][1] + text)
            else:
                merged.append((kind, text))
        out = []
        for position, (kind, text) in enumerate(merged):
            after = merged[position + 1][1] if position + 1 < len(merged) else ""
            # Constrained formatting only fires at word boundaries, so a span
            # that starts or ends mid-word needs the doubled delimiter.
            tight = (text[:1].strip() != "" and "".join(out)[-1:].isalnum()) or (
                text[-1:].strip() != "" and after[:1].isalnum()
            )
            if kind == "raw" or not text.strip():
                out.append(text)
                continue
            # Padding has to sit outside the delimiters, or the constrained
            # form does not fire at all.
            body = text.strip()
            lead = text[: len(text) - len(text.lstrip())]
            trail = text[len(text.rstrip()) :]
            if kind == "code":
                marked = code_span(body)
            elif kind == "italic":
                marked = emphasis(body, "_", tight)
            elif kind == "bold":
                marked = emphasis(body, "*", tight)
            elif kind == "superscript":
                marked = f"^{unspaced(body)}^"
            elif kind == "subscript":
                marked = f"~{unspaced(body)}~"
            else:
                out.append(escape_inline(text))
                continue
            out.append(lead + marked + trail)
        return "".join(out)

    def _render_children(self, container, segments, field_stack):
        for child in container:
            tag = child.tag
            if tag == q("w", "r"):
                self._render_run(child, segments, field_stack)
            elif tag == q("w", "hyperlink"):
                self._render_hyperlink(child, segments, field_stack)
            elif tag == q("w", "ins"):
                self._render_children(child, segments, field_stack)
            elif tag == q("w", "del"):
                continue
            elif tag in (q("w", "smartTag"), q("w", "sdtContent"), q("w", "sdt")):
                self._render_children(child, segments, field_stack)

    def _emit(self, segments, field_stack, text, raw=False, kind="plain"):
        if not text:
            return
        segments.append(("raw" if raw else kind, text))

    def _render_run(self, run, segments, field_stack):
        if is_hidden(run):
            return
        rPr = run.find(q("w", "rPr"))
        kind = "plain"
        if rPr is not None:
            rstyle = rPr.find(q("w", "rStyle"))
            style = rstyle.get(W_VAL) if rstyle is not None else ""
            vert = rPr.find(q("w", "vertAlign"))
            vert = vert.get(W_VAL) if vert is not None else None
            if vert in ("superscript", "subscript"):
                kind = vert
            elif style in CODE_RUN_STYLES:
                kind = "code"
            elif style in BOLD_RUN_STYLES or rPr.find(q("w", "b")) is not None:
                kind = "bold"
            elif style in ITALIC_RUN_STYLES or rPr.find(q("w", "i")) is not None:
                kind = "italic"
        for node in run:
            tag = node.tag
            if tag == q("w", "t"):
                self._emit(
                    segments, field_stack, normalise_text(node.text or ""), kind=kind
                )
            elif tag == q("w", "tab") or tag in (q("w", "br"), q("w", "cr")):
                self._emit(segments, field_stack, " ", kind=kind)
            elif tag == q("w", "noBreakHyphen"):
                self._emit(segments, field_stack, "-", kind=kind)

    def _render_hyperlink(self, node, segments, field_stack):
        # Link targets are resolved once anchors exist.
        self._render_children(node, segments, field_stack)

    # -- block rendering --------------------------------------------------

    def render_blocks(self, blocks, depth=0) -> list:
        """Render a sequence of body-level elements to AsciiDoc lines.

        ``depth`` is the table nesting level, which decides the cell separator.
        """
        lines = []
        index = 0
        while index < len(blocks):
            block = blocks[index]
            if block.tag != W_P:
                index += 1
                continue
            style = para_style(block)
            if style in DROP_STYLES:
                index += 1
                continue
            if style in CODE_STYLES:
                block_lines, index = self._render_code_run(blocks, index)
                lines.extend(block_lines)
                continue
            if is_list_para(block):
                block_lines, index = self._render_list(blocks, index)
                lines.extend(block_lines)
                continue
            if is_xml_example(block):
                block_lines, index = self._render_xml_example(blocks, index)
                lines.extend(block_lines)
                continue
            level = heading_level(style)
            if level is not None:
                lines.extend(self._render_heading(block, level))
                index += 1
                continue
            lines.extend(self._render_paragraph(block, style))
            index += 1
        return trim_blank_lines(lines)

    def _render_heading(self, para, level) -> list:
        lines = [""]
        annex = level == 1 and self.current is not None and self.current.is_annex
        if annex:
            lines.append("[appendix]")
        title = self.render_inline(para).strip()
        if annex:
            # Metanorma numbers annexes itself, so drop Word's "Appendix C:".
            title = APPENDIX_RE.sub("", title).strip(" :-") or title
        marker = "=" * (level + 1)
        lines.append(f"{marker} {title or 'Untitled'}")
        lines.append("")
        return lines

    def _render_paragraph(self, para, style) -> list:
        text = self.render_inline(para).strip()
        if not text:
            return []
        return [protect_line_start(text), ""]

    def _render_xml_example(self, blocks, index) -> tuple:
        """Emit schema fragments that Word styled as body text as a listing.

        Left as prose their quotes are curled, their comment delimiters turn
        into dashes and their tags are read as AsciiDoc markup.
        """
        text_lines = []
        while index < len(blocks):
            block = blocks[index]
            if block.tag != W_P:
                break
            if is_xml_example(block):
                text_lines.append(raw_text(block).strip())
            elif raw_text(block).strip():
                break
            index += 1
        lines = ["", listing_attributes("xml", False), "----"]
        lines.extend(text_lines)
        lines.extend(["----", ""])
        return lines, index

    def _render_code_run(self, blocks, index) -> tuple:
        style = para_style(blocks[index])
        language = CODE_STYLES[style]
        text_lines = []
        while index < len(blocks):
            block = blocks[index]
            if block.tag != W_P or para_style(block) != style:
                break
            text_lines.extend(code_lines(block))
            index += 1
        while text_lines and not text_lines[0].strip():
            text_lines.pop(0)
        while text_lines and not text_lines[-1].strip():
            text_lines.pop()
        if not text_lines:
            return [], index
        fence = "-" * 4
        while any(line.strip() == fence for line in text_lines):
            fence += "-"
        lines = [""]
        attrs = listing_attributes(language, False)
        if attrs:
            lines.append(attrs)
        lines.append(fence)
        lines.extend(text_lines)
        lines.append(fence)
        lines.append("")
        return lines, index

    def _render_list(self, blocks, index) -> tuple:
        items = []
        while index < len(blocks):
            block = blocks[index]
            if block.tag != W_P or not is_list_para(block):
                break
            if para_style(block) in DROP_STYLES:
                break
            numbering = numbering_of(block)
            if numbering is None:
                num_id, ilvl, kind = None, "0", "bullet"
            else:
                num_id, ilvl = numbering
                kind = self.numbering.kind(num_id, ilvl)
            text = self.render_inline(block).strip()
            if text:
                items.append((int(ilvl), kind, text))
            index += 1
        if not items:
            return [], index
        # Word indents lists with arbitrary ilvl values; compress them so that
        # the shallowest level in this run becomes depth 1.
        depths = {
            value: rank + 1 for rank, value in enumerate(sorted({i[0] for i in items}))
        }
        lines = [""]
        for ilvl, kind, text, comments in items:
            lines.extend(comments)
            if not text:
                continue
            marker = ("." if kind == "ordered" else "*") * depths[ilvl]
            lines.append(f"{marker} {text}")
        lines.append("")
        return lines, index

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


#: A paragraph that is really a fragment of XML rather than a sentence.
XML_EXAMPLE_RE = re.compile(r"^<[?!/]?[A-Za-z][^<>]*>")


def is_xml_example(para) -> bool:
    """True for a body-text paragraph whose whole content is XML markup."""
    if para.tag != W_P or para_style(para) in CODE_STYLES:
        return False
    if para_style(para) in CAPTION_STYLES or heading_level(para_style(para)):
        return False
    text = raw_text(para).strip()
    return bool(text) and text.endswith(">") and XML_EXAMPLE_RE.match(text) is not None


def listing_attributes(language: str, titled: bool) -> str:
    """The attribute line for a listing block.

    Metanorma draws listings from the same numbering sequence as figures, so
    an example the source neither numbers nor names would take a figure
    number away from the diagrams that do carry one.
    """
    style = f"source,{language}" if language else ""
    if titled:
        return f"[{style}]" if style else ""
    style = style.replace("source,", "source%unnumbered,") if style else "%unnumbered"
    return f"[{style}]"


def code_lines(para) -> list:
    """Verbatim lines for one listing paragraph, honouring line breaks."""
    parts = []
    for run in para.iter(q("w", "r")):
        for node in run:
            if node.tag == q("w", "t"):
                parts.append(node.text or "")
            elif node.tag == q("w", "tab"):
                parts.append("    ")
            elif node.tag in (q("w", "br"), q("w", "cr")):
                parts.append("\n")
            elif node.tag == q("w", "noBreakHyphen"):
                parts.append("-")
    text = normalise_text("".join(parts)).replace("\t", "    ")
    return text.split("\n")


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

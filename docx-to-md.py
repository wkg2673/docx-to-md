#!/usr/bin/env python3
"""
DOCX -> Markdown converter.

Interactive mode (asks for both paths):
    python docx_to_md.py

Command-line mode:
    python docx_to_md.py INPUT.docx OUTPUT.md
    python docx_to_md.py INPUT.docx OUTPUT.md --force     # overwrite without asking
    python docx_to_md.py INPUT.docx                       # asks only for the output path

Requires:  pip install python-docx

Supported: headings, bold / italic / strikethrough / inline code, hyperlinks,
bullet and numbered lists (nested), block quotes, tables, footnotes,
fenced code blocks, and images (extracted into a folder next to the output file).
"""

import argparse
import os
import re
import sys

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from lxml import etree

CODE_FONTS = {
    "courier new", "courier", "consolas", "monaco", "menlo", "lucida console",
    "source code pro", "fira code", "fira mono", "dejavu sans mono", "liberation mono",
    "cascadia code", "cascadia mono", "andale mono", "monospace",
}
# Style names that mean "code" (paragraph styles such as "Code", "Source Code",
# "Preformatted Text", "HTML Preformatted", "Verbatim"; character styles such as "HTML Code")
CODE_STYLE_RE = re.compile(r"\b(code|preformatted|verbatim|macro text)\b")
HEADING_RE = re.compile(r"heading (\d)")
BACKTICK_RUN_RE = re.compile(r"`+")
FENCE_RUN_RE = re.compile(r"`{3,}")
STRIP_BOLD_RE = re.compile(r"^\*+|\*+$")


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def clean_path(raw):
    """Strip whitespace and surrounding quotes (e.g. from 'Copy as path')."""
    raw = raw.strip()
    if len(raw) >= 2 and raw[0] == raw[-1] and raw[0] in ("'", '"'):
        raw = raw[1:-1]
    return os.path.expanduser(raw.strip())


def escape_md(text):
    """Escape characters that would otherwise be read as Markdown syntax."""
    return re.sub(r"([\\`*_\[\]<>|])", r"\\\1", text)


def split_edges(text):
    """Split text into (leading whitespace, core, trailing whitespace)."""
    core = text.strip()
    if not core:
        return text, "", ""
    lead = text[: len(text) - len(text.lstrip())]
    trail = text[len(text.rstrip()):]
    return lead, core, trail


def wrap(text, marker):
    """Wrap text in a Markdown marker, keeping edge whitespace outside it."""
    lead, core, trail = split_edges(text)
    if not core:
        return text
    return f"{lead}{marker}{core}{marker}{trail}"


def code_span(text):
    """Inline code; uses a longer backtick fence if the text contains backticks."""
    lead, core, trail = split_edges(text)
    if not core:
        return text
    longest = max((len(m) for m in BACKTICK_RUN_RE.findall(core)), default=0)
    fence = "`" * (longest + 1)
    pad = " " if core.startswith("`") or core.endswith("`") else ""
    return f"{lead}{fence}{pad}{core}{pad}{fence}{trail}"


def code_block(code):
    """Fenced code block; the fence is made longer if the code contains ```."""
    longest = max((len(m) for m in FENCE_RUN_RE.findall(code)), default=2)
    fence = "`" * max(3, longest + 1)
    return f"{fence}\n{code}\n{fence}"


# --------------------------------------------------------------------------- #
# Converter
# --------------------------------------------------------------------------- #
class DocxToMarkdown:
    def __init__(self, docx_path, md_path):
        self.doc = Document(docx_path)
        self.md_path = md_path
        self.image_dir_name = os.path.splitext(os.path.basename(md_path))[0] + "_images"
        self.image_dir = os.path.join(os.path.dirname(os.path.abspath(md_path)), self.image_dir_name)
        self.image_count = 0
        self.image_cache = {}          # rId -> relative markdown path
        self.counters = {}             # (numId, ilvl) -> current number
        self.num_formats = self._load_numbering()
        self.mono_styles = self._load_mono_styles()
        self.rels = self.doc.part.rels  # relationships used to resolve hyperlinks
        self.footnotes, self.footnote_rels = self._load_footnotes()
        self.footnote_order = []       # original footnote ids, in order of first reference
        self.footnote_numbers = {}     # original id -> sequential number in the Markdown

    # ---- numbering (list type detection) ---------------------------------- #
    def _load_numbering(self):
        """Map (numId, ilvl) -> numFmt (e.g. 'bullet', 'decimal')."""
        formats = {}
        try:
            numbering = self.doc.part.numbering_part.element
        except Exception:
            return formats

        abstract = {}
        for an in numbering.findall(qn("w:abstractNum")):
            aid = an.get(qn("w:abstractNumId"))
            for lvl in an.findall(qn("w:lvl")):
                fmt = lvl.find(qn("w:numFmt"))
                abstract[(aid, lvl.get(qn("w:ilvl")))] = fmt.get(qn("w:val")) if fmt is not None else "decimal"

        for num in numbering.findall(qn("w:num")):
            nid = num.get(qn("w:numId"))
            ref = num.find(qn("w:abstractNumId"))
            if ref is None:
                continue
            aid = ref.get(qn("w:val"))
            for (a, lvl), fmt in abstract.items():
                if a == aid:
                    formats[(nid, lvl)] = fmt
        return formats

    # ---- styles (code detection) ------------------------------------------ #
    @staticmethod
    def _style_font_name(style):
        """Font name of a style, following the chain of base styles."""
        seen = 0
        while style is not None and seen < 20:
            font = getattr(style, "font", None)
            if font is not None and font.name:
                return font.name
            style = style.base_style
            seen += 1
        return None

    def _load_mono_styles(self):
        """Ids of styles that represent code (by name or by monospace font)."""
        mono = set()
        for style in self.doc.styles:
            try:
                name = (style.name or "").lower()
                if CODE_STYLE_RE.search(name):
                    mono.add(style.style_id)
                    continue
                font = self._style_font_name(style)
                if font and font.lower() in CODE_FONTS:
                    mono.add(style.style_id)
            except Exception:
                continue
        return mono

    def _run_is_mono(self, run_el, default=False):
        """Is this w:r rendered in a monospace font? `default` applies if the run says nothing."""
        rpr = run_el.find(qn("w:rPr"))
        if rpr is not None:
            fonts = rpr.find(qn("w:rFonts"))
            if fonts is not None:
                name = fonts.get(qn("w:ascii")) or fonts.get(qn("w:hAnsi")) or ""
                if name:
                    return name.lower() in CODE_FONTS
            rstyle = rpr.find(qn("w:rStyle"))
            if rstyle is not None and rstyle.get(qn("w:val")) in self.mono_styles:
                return True
        return default

    def _is_code_paragraph(self, paragraph):
        """True for paragraphs with a code style or whose text is entirely monospace."""
        style = paragraph.style
        if style is not None and CODE_STYLE_RE.search((style.name or "").lower()):
            return True
        default_mono = style is not None and style.style_id in self.mono_styles
        runs = []
        for run in paragraph._p.iter(qn("w:r")):
            text = "".join(t.text or "" for t in run.findall(qn("w:t")))
            if text.strip():
                runs.append(run)
        return bool(runs) and all(self._run_is_mono(r, default_mono) for r in runs)

    # ---- footnotes ---------------------------------------------------------- #
    def _load_footnotes(self):
        """Read word/footnotes.xml (python-docx has no API for it)."""
        notes, rels = {}, None
        for rel in self.doc.part.rels.values():
            if rel.is_external or rel.reltype != RT.FOOTNOTES:
                continue
            try:
                root = etree.fromstring(rel.target_part.blob)
                rels = rel.target_part.rels
            except Exception:
                continue
            for fn in root.findall(qn("w:footnote")):
                if fn.get(qn("w:type")) in ("separator", "continuationSeparator", "continuationNotice"):
                    continue
                notes[fn.get(qn("w:id"))] = fn
        return notes, rels

    def _footnote_marker(self, fid):
        """Return '[^n]' for a footnote id, numbering footnotes in order of appearance."""
        if fid not in self.footnotes:
            return None
        if fid not in self.footnote_numbers:
            self.footnote_order.append(fid)
            self.footnote_numbers[fid] = len(self.footnote_order)
        return f"[^{self.footnote_numbers[fid]}]"

    def _footnote_definitions(self):
        """Markdown definitions for every footnote that was referenced."""
        if not self.footnote_order:
            return []
        definitions = []
        saved_rels = self.rels
        if self.footnote_rels is not None:
            self.rels = self.footnote_rels
        try:
            for fid in self.footnote_order:
                paragraphs = self.footnotes[fid].findall(qn("w:p"))
                texts = [self._paragraph_inline(p, refs=False).strip() for p in paragraphs]
                texts = [t for t in texts if t]
                if not texts:
                    texts = [""]
                number = self.footnote_numbers[fid]
                body = f"[^{number}]: {texts[0]}"
                for extra in texts[1:]:
                    body += "\n\n    " + extra
                definitions.append(body)
        finally:
            self.rels = saved_rels
        return definitions

    # ---- images ------------------------------------------------------------ #
    def _save_image(self, rid):
        if rid in self.image_cache:
            return self.image_cache[rid]
        try:
            part = self.doc.part.related_parts[rid]
        except KeyError:
            return None
        os.makedirs(self.image_dir, exist_ok=True)
        self.image_count += 1
        ext = os.path.splitext(part.partname)[1] or ".png"
        filename = f"image{self.image_count}{ext}"
        with open(os.path.join(self.image_dir, filename), "wb") as f:
            f.write(part.blob)
        rel = f"{self.image_dir_name}/{filename}"
        self.image_cache[rid] = rel
        return rel

    # ---- inline content ---------------------------------------------------- #
    def _run_format(self, run_el):
        """Return (bold, italic, strike, code) for a w:r element."""
        rpr = run_el.find(qn("w:rPr"))
        bold = italic = strike = False
        if rpr is not None:
            def on(tag):
                el = rpr.find(qn(tag))
                return el is not None and el.get(qn("w:val"), "true") not in ("0", "false", "off")
            bold, italic = on("w:b"), on("w:i")
            strike = on("w:strike") or on("w:dstrike")
        return bold, italic, strike, self._run_is_mono(run_el)

    def _run_pieces(self, run_el, refs=True):
        """Yield ('text'|'image'|'footnote', value) pieces from a w:r element."""
        for child in run_el:
            tag = child.tag
            if tag == qn("w:t"):
                yield "text", child.text or ""
            elif tag == qn("w:tab"):
                yield "text", "    "
            elif tag in (qn("w:br"), qn("w:cr")):
                if child.get(qn("w:type")) != "page":
                    yield "text", "\n"
            elif tag == qn("w:footnoteReference") and refs:
                marker = self._footnote_marker(child.get(qn("w:id")))
                if marker:
                    yield "footnote", marker
            elif tag in (qn("w:drawing"), qn("w:pict")) and refs:
                for blip in child.iter(qn("a:blip")):
                    rid = blip.get(qn("r:embed"))
                    if rid:
                        path = self._save_image(rid)
                        if path:
                            yield "image", path

    def _runs_to_md(self, runs, refs=True):
        """
        Convert a list of w:r elements to Markdown, merging adjacent runs
        that share formatting so we don't produce '**a****b**'.
        """
        segments = []  # [fmt, text]; fmt is a tuple for text or "raw" for images / footnotes
        for run_el in runs:
            fmt = self._run_format(run_el)
            for kind, value in self._run_pieces(run_el, refs=refs):
                if kind == "image":
                    segments.append(["raw", f"![image]({value})"])
                elif kind == "footnote":
                    segments.append(["raw", value])
                elif segments and segments[-1][0] == fmt:
                    segments[-1][1] += value
                else:
                    segments.append([fmt, value])

        out = []
        for fmt, text in segments:
            if fmt == "raw":
                out.append(text)
                continue
            bold, italic, strike, code = fmt
            if code:
                text = code_span(text)
            else:
                text = escape_md(text)
                if bold and italic:
                    text = wrap(text, "***")
                elif bold:
                    text = wrap(text, "**")
                elif italic:
                    text = wrap(text, "*")
                if strike:
                    text = wrap(text, "~~")
            out.append(text)
        return "".join(out)

    def _paragraph_inline(self, p_el, refs=True):
        """Convert all inline content of a w:p (runs + hyperlinks) to Markdown."""
        parts = []
        pending_runs = []

        def flush():
            if pending_runs:
                parts.append(self._runs_to_md(pending_runs, refs))
                pending_runs.clear()

        for child in p_el:
            if child.tag == qn("w:r"):
                pending_runs.append(child)
            elif child.tag == qn("w:hyperlink"):
                flush()
                label = self._runs_to_md(child.findall(qn("w:r")), refs).strip()
                rid = child.get(qn("r:id"))
                anchor = child.get(qn("w:anchor"))
                url = None
                if rid and rid in self.rels:
                    url = self.rels[rid].target_ref
                elif anchor:
                    url = f"#{anchor}"
                parts.append(f"[{label}]({url})" if url and label else label)
            elif child.tag in (qn("w:ins"), qn("w:smartTag")):
                # Tracked insertions / smart tags wrap normal runs
                pending_runs.extend(child.findall(qn("w:r")))
        flush()
        return "".join(parts)

    def _raw_text(self, p_el):
        """Plain text of a paragraph with no Markdown escaping (for code blocks)."""
        out = []
        for run in p_el.iter(qn("w:r")):
            for kind, value in self._run_pieces(run, refs=False):
                if kind == "text":
                    out.append(value)
        return "".join(out)

    # ---- block content ----------------------------------------------------- #
    @staticmethod
    def _is_heading(paragraph):
        style = (paragraph.style.name if paragraph.style is not None else "").lower()
        return style in ("title", "subtitle") or bool(HEADING_RE.match(style))

    def _list_info(self, paragraph):
        """Return (ordered: bool, level: int, numId) or None if not a list item."""
        ppr = paragraph._p.pPr
        num_pr = ppr.numPr if ppr is not None else None
        style = (paragraph.style.name if paragraph.style is not None else "").lower()

        if num_pr is not None and num_pr.numId is not None and num_pr.numId.val != 0:
            nid = str(num_pr.numId.val)
            lvl = str(num_pr.ilvl.val) if num_pr.ilvl is not None else "0"
            fmt = self.num_formats.get((nid, lvl), "decimal")
            return fmt != "bullet", int(lvl), nid

        if "list bullet" in style:
            m = re.search(r"(\d+)$", style)
            return False, (int(m.group(1)) - 1 if m else 0), "style-bullet"
        if "list number" in style:
            m = re.search(r"(\d+)$", style)
            return True, (int(m.group(1)) - 1 if m else 0), "style-number"
        return None

    def _paragraph_to_md(self, paragraph):
        text = self._paragraph_inline(paragraph._p)
        style = (paragraph.style.name if paragraph.style is not None else "").lower()

        # Headings
        if style == "title":
            self.counters.clear()
            return f"# {text.strip()}" if text.strip() else ""
        if style == "subtitle":
            return f"*{text.strip()}*" if text.strip() else ""
        m = HEADING_RE.match(style)
        if m and text.strip():
            self.counters.clear()
            level = min(int(m.group(1)), 6)
            heading = STRIP_BOLD_RE.sub("", text.strip())
            return f"{'#' * level} {heading}"

        # Lists
        info = self._list_info(paragraph)
        if info and text.strip():
            ordered, level, nid = info
            if ordered:
                self.counters[(nid, level)] = self.counters.get((nid, level), 0) + 1
                for key in [k for k in self.counters if k[0] == nid and k[1] > level]:
                    del self.counters[key]
                marker = f"{self.counters[(nid, level)]}."
            else:
                marker = "-"
            return f"{'    ' * level}{marker} {text.strip()}"
        self.counters.clear()

        # Quotes
        if style in ("quote", "intense quote", "block text"):
            return "> " + text.strip().replace("\n", "\n> ") if text.strip() else ""

        # Normal text: a manual line break becomes a Markdown hard break
        return text.strip().replace("\n", "  \n")

    def _table_to_md(self, table):
        rows = []
        for row in table.rows:
            cells = []
            seen = set()
            for cell in row.cells:
                if id(cell._tc) in seen:  # merged cells repeat
                    continue
                seen.add(id(cell._tc))
                content = "<br>".join(
                    self._paragraph_inline(p._p).strip() for p in cell.paragraphs if p.text.strip()
                )
                cells.append(re.sub(r"(?<!\\)\|", r"\\|", content))
            rows.append(cells)
        if not rows:
            return ""

        width = max(len(r) for r in rows)
        rows = [r + [""] * (width - len(r)) for r in rows]
        lines = ["| " + " | ".join(rows[0]) + " |", "| " + " | ".join(["---"] * width) + " |"]
        lines += ["| " + " | ".join(r) + " |" for r in rows[1:]]
        return "\n".join(lines)

    # ---- main -------------------------------------------------------------- #
    def convert(self):
        blocks = []
        previous_list_id = None  # numId of the list the previous block belonged to
        code_lines = None        # list of lines while inside a code block

        def flush_code():
            nonlocal code_lines
            if code_lines is not None:
                while code_lines and not code_lines[-1].strip():
                    code_lines.pop()
                if code_lines:
                    blocks.append(code_block("\n".join(code_lines)))
                code_lines = None

        for child in self.doc.element.body.iterchildren():
            if child.tag == qn("w:p"):
                paragraph = Paragraph(child, self.doc)

                # Code blocks: consecutive code paragraphs form one fenced block
                if (
                    not self._is_heading(paragraph)
                    and self._list_info(paragraph) is None
                    and self._is_code_paragraph(paragraph)
                ):
                    text = self._raw_text(child)
                    if code_lines is None:
                        if not text.strip():
                            continue
                        code_lines = []
                    code_lines.extend(text.split("\n"))
                    previous_list_id = None
                    continue

                flush_code()
                md = self._paragraph_to_md(paragraph)
                if not md.strip():
                    previous_list_id = None
                    continue
                info = self._list_info(paragraph)
                list_id = info[2] if info else None
                # Consecutive items of the same list are joined by a single newline
                if list_id is not None and list_id == previous_list_id and blocks:
                    blocks[-1] += "\n" + md
                else:
                    blocks.append(md)
                previous_list_id = list_id
            elif child.tag == qn("w:tbl"):
                flush_code()
                md = self._table_to_md(Table(child, self.doc))
                if md:
                    blocks.append(md)
                previous_list_id = None

        flush_code()
        blocks.extend(self._footnote_definitions())
        return "\n\n".join(blocks) + "\n"


# --------------------------------------------------------------------------- #
# Command-line / interactive handling
# --------------------------------------------------------------------------- #
def fail(message, code=1):
    print(f"Error: {message}", file=sys.stderr)
    sys.exit(code)


def ask_yes_no(question):
    try:
        return input(question).strip().lower() in ("y", "yes")
    except EOFError:
        return False


def check_input_path(path):
    """Return an error message, or None if the input path is fine."""
    if not path.lower().endswith(".docx"):
        return "The input file must have the .docx extension."
    if not os.path.isfile(path):
        return f"File not found: {path}"
    return None


def get_input_path(arg):
    if arg:
        path = clean_path(arg)
        error = check_input_path(path)
        if error:
            fail(error)
        return path
    while True:
        path = clean_path(input("Full path of the .docx file (including extension): "))
        error = check_input_path(path)
        if error:
            print(f"  -> {error}\n")
        else:
            return path


def get_output_path(arg, force):
    from_args = bool(arg)
    while True:
        raw = arg if from_args else input("Full output path (including extension, e.g. C:\\out\\result.md): ")
        path = clean_path(raw)

        if not path.lower().endswith(".md"):
            if from_args:
                fail("The output file must have the .md extension.")
            print("  -> The output file must have the .md extension.\n")
            continue

        out_dir = os.path.dirname(os.path.abspath(path))
        if not os.path.isdir(out_dir):
            if from_args or ask_yes_no(f"  Folder '{out_dir}' does not exist. Create it? (y/n): "):
                os.makedirs(out_dir, exist_ok=True)
            else:
                continue

        if os.path.exists(path) and not force:
            if not ask_yes_no("  Output file already exists. Overwrite? (y/n): "):
                if from_args:
                    fail("Output file already exists (use --force to overwrite).")
                continue
        return path


def parse_args():
    parser = argparse.ArgumentParser(
        description="Convert a Word document (.docx) to Markdown (.md). "
                    "Run without arguments for interactive prompts.",
    )
    parser.add_argument("input", nargs="?", help="path of the .docx file to convert")
    parser.add_argument("output", nargs="?", help="path of the .md file to create")
    parser.add_argument("-f", "--force", action="store_true",
                        help="overwrite the output file without asking")
    return parser.parse_args()


def main():
    args = parse_args()

    if args.input is None and args.output is None:
        print("=== DOCX to Markdown Converter ===\n")

    try:
        in_path = get_input_path(args.input)
        out_path = get_output_path(args.output, args.force)
    except (KeyboardInterrupt, EOFError):
        print("\nCancelled.")
        sys.exit(1)

    try:
        converter = DocxToMarkdown(in_path, out_path)
        markdown = converter.convert()
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(markdown)
    except Exception as exc:
        fail(f"Conversion failed: {exc}")

    print(f"\nDone! Markdown saved to: {out_path}")
    if converter.footnote_order:
        print(f"{len(converter.footnote_order)} footnote(s) converted.")
    if converter.image_count:
        print(f"{converter.image_count} image(s) saved to: {converter.image_dir}")


if __name__ == "__main__":
    main()

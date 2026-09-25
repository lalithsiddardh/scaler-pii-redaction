"""OPC (.docx) package I/O with exact per-character XML mapping.

Design constraints that drive this module:

* A .docx is a ZIP of XML parts.  We must not rebuild those parts from a DOM,
  because re-serialising a DOM silently drops namespace declarations that are
  referenced only from attribute *values* (``mc:Ignorable="w14 wp14"``) and
  makes Word refuse to open the file.
* Therefore we never re-serialise.  We locate text in the original bytes,
  remember the byte offset of every individual character, and splice new
  characters into exactly those positions.
* Word splits a sentence across many ``<w:r>`` runs and, worse, drops the space
  at a run seam ("Kushal" + "Subbayya Hegde" -> "KushalSubbayya Hegde").
  Detection therefore runs on the *concatenated* text of a paragraph, and a
  replacement is distributed back over the same ``<w:t>`` nodes it came from.
  No run, no ``<w:rPr>``, and no table cell is ever created or destroyed, so
  formatting and fixed table layouts survive untouched.

Python 3.9, standard library only.
"""

import re
import zipfile
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

# Parts whose text can contain PII.
TEXT_PART_RE = re.compile(
    r"^word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$"
)
METADATA_PART = "docProps/core.xml"

#: A table cell, including nested ones.  Word nests ``<w:tc>`` for merged
#: layouts, so the outer span is what identifies the cell a paragraph is in.
CELL_RE = re.compile(rb"<w:tc(?:\s[^>]*)?>.*?</w:tc>|<w:tc\s*/>", re.S)

# Elements that carry literal text we may need to rewrite.
TEXT_ELEMENTS = (b"w:t", b"w:instrText", b"w:delText")

_ENTITIES = {
    b"amp": "&",
    b"lt": "<",
    b"gt": ">",
    b"quot": '"',
    b"apos": "'",
}


def _element_re(name: bytes):
    """Open/close tag matcher for one element name.

    ``\\b`` after the name stops ``w:t`` from matching ``w:tab``/``w:tbl`` and
    stops ``w:p`` from matching ``w:pPr``.  The attribute group tolerates a
    literal ``>`` inside a quoted attribute value.
    """
    return re.compile(
        rb"<" + re.escape(name) + rb"\b((?:[^>\"]|\"[^\"]*\")*?)(/?)>"
        rb"|</" + re.escape(name) + rb">"
    )


def _paragraph_re():
    return re.compile(rb"<w:p\b((?:[^>\"]|\"[^\"]*\")*?)(/?)>|</w:p>")


def decode_with_offsets(raw: bytes) -> Tuple[str, List[int]]:
    """XML-unescape ``raw``, returning the text and a per-character byte offset.

    ``offsets[i]`` is the byte offset *within* ``raw`` at which character ``i``
    of the decoded string begins.  It is non-decreasing, so a slice of the
    decoded text always maps to a contiguous slice of the raw bytes even when
    the two differ in length (``&amp;`` = 1 char / 5 bytes).
    """
    chars: List[str] = []
    offsets: List[int] = []
    i, n = 0, len(raw)
    while i < n:
        if raw[i] == 0x26:  # '&'
            j = raw.find(b";", i, i + 12)
            if j == -1:
                chunk = raw[i:i + 1].decode("utf-8", "replace")
                for c in chunk:
                    chars.append(c)
                    offsets.append(i)
                i += 1
                continue
            name = raw[i + 1:j]
            if name.startswith(b"#"):
                try:
                    code = int(name[2:], 16) if name[1:2] in (b"x", b"X") else int(name[1:], 10)
                    chunk = chr(code)
                except (ValueError, IndexError):
                    chunk = "&" + name.decode("utf-8", "replace") + ";"
            else:
                chunk = _ENTITIES.get(name)
                if chunk is None:
                    chunk = "&" + name.decode("utf-8", "replace") + ";"
            for c in chunk:
                chars.append(c)
                offsets.append(i)
            i = j + 1
            continue
        byte = raw[i]
        if byte >= 0xF0:
            width = 4
        elif byte >= 0xE0:
            width = 3
        elif byte >= 0xC0:
            width = 2
        else:
            width = 1
        chunk = raw[i:i + width].decode("utf-8", "replace")
        for c in chunk:
            chars.append(c)
            offsets.append(i)
        i += width
    return "".join(chars), offsets


def xml_escape(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


class TextNode:
    """One ``<w:t>`` / ``<w:instrText>`` element inside a part's raw bytes."""

    __slots__ = (
        "elem", "open_start", "open_end", "inner_start",
        "inner_end", "elem_end", "raw_inner", "text", "offsets", "is_field",
    )

    def __init__(self, elem: bytes, open_start: int, open_end: int,
                 inner_start: int, inner_end: int, elem_end: int,
                 raw_inner: bytes, text: str, offsets: List[int], is_field: bool):
        self.elem = elem
        self.open_start = open_start
        self.open_end = open_end            # first byte after the open tag's '>'
        self.inner_start = inner_start
        self.inner_end = inner_end
        self.elem_end = elem_end
        self.raw_inner = raw_inner
        self.text = text
        self.offsets = offsets
        self.is_field = is_field

    def raw_open_tag(self, data: bytes) -> bytes:
        return data[self.open_start:self.open_end]


class Paragraph:
    """A ``<w:p>`` and the concatenated text of the text nodes inside it."""

    __slots__ = ("index", "start", "end", "nodes", "text", "char_map")

    def __init__(self, index: int, start: int, end: int, nodes: List[TextNode]):
        self.index = index
        self.start = start
        self.end = end
        self.nodes = nodes
        parts: List[str] = []
        char_map: List[Tuple[int, int]] = []
        for node_index, node in enumerate(nodes):
            parts.append(node.text)
            for k in range(len(node.text)):
                char_map.append((node_index, k))
        self.text = "".join(parts)
        self.char_map = char_map

    def __len__(self) -> int:
        return len(self.text)

    def rebuild(self) -> None:
        """Recompute the text and the node map after nodes were dropped."""
        parts: List[str] = []
        char_map: List[Tuple[int, int]] = []
        for node_index, node in enumerate(self.nodes):
            parts.append(node.text)
            for k in range(len(node.text)):
                char_map.append((node_index, k))
        self.text = "".join(parts)
        self.char_map = char_map

    def node_span(self, start: int, end: int) -> List[Tuple[int, int, int]]:
        """Split ``text[start:end]`` over the nodes it came from.

        Returns ``[(node_index, char_start, char_end), ...]``.
        """
        spans: List[Tuple[int, int, int]] = []
        for pos in range(start, end):
            node_index, offset = self.char_map[pos]
            if spans and spans[-1][0] == node_index and spans[-1][2] == offset:
                spans[-1] = (node_index, spans[-1][1], offset + 1)
            else:
                spans.append((node_index, offset, offset + 1))
        return spans


class PartDocument:
    """A single XML part plus the paragraph index used for detection."""

    def __init__(self, name: str, data: bytes):
        self.name = name
        self.data = data
        self.paragraphs: List[Paragraph] = _index_paragraphs(data)

    def paragraphs_with_text(self) -> List[Paragraph]:
        return [p for p in self.paragraphs if p.text.strip()]


def _index_paragraphs(data: bytes) -> List[Paragraph]:
    """Locate every ``w:p`` (depth aware) and index the text nodes inside it."""
    boundaries: List[Tuple[int, Optional[int]]] = []
    stack: List[int] = []
    for match in _paragraph_re().finditer(data):
        tag = match.group(0)
        if tag.startswith(b"</"):
            if stack:
                start = stack.pop()
                boundaries.append((start, match.end()))
        elif match.group(2):  # <w:p .../>
            continue
        else:
            stack.append(match.start())
    boundaries.sort()

    element_res = [(name, _element_re(name)) for name in TEXT_ELEMENTS]
    paragraphs: List[Paragraph] = []
    for index, (start, end) in enumerate(boundaries):
        chunk = data[start:end]
        nodes: List[TextNode] = []
        for name, pattern in element_res:
            open_spans: List[Tuple[int, int, int, int]] = []  # open_start, open_end, attr_end
            for match in pattern.finditer(chunk):
                if match.group(0).startswith(b"</"):
                    if not open_spans:
                        continue
                    o_start, o_end, attr_end = open_spans.pop()
                    inner_start = o_end
                    inner_end = match.start()
                    raw_inner = chunk[inner_start:inner_end]
                    text, offsets = decode_with_offsets(raw_inner)
                    # Offsets are stored part-absolute: the rewriter slices
                    # the whole part, not the paragraph's chunk.
                    nodes.append(
                        TextNode(
                            elem=name,
                            open_start=start + o_start,
                            open_end=start + o_end,
                            inner_start=start + inner_start,
                            inner_end=start + inner_end,
                            elem_end=start + match.end(),
                            raw_inner=raw_inner,
                            text=text,
                            offsets=offsets,
                            is_field=(name == b"w:instrText"),
                        )
                    )
                elif match.group(2):  # self closing, no text
                    continue
                else:
                    open_spans.append((match.start(), match.end(), match.end()))
        nodes.sort(key=lambda n: n.open_start)
        paragraphs.append(Paragraph(index, start, end, nodes))
    return _claim_innermost_nodes(paragraphs)


def _claim_innermost_nodes(paragraphs: List[Paragraph]) -> List[Paragraph]:
    """Drop text nodes that a nested paragraph already owns.

    A text box puts a ``w:p`` inside a ``w:p``; the outer paragraph's byte
    range contains the inner one, so without this step the same ``w:t``
    element would be indexed twice.  The rewriter edits nodes by byte range,
    and a node edited through two paragraphs would be edited twice with
    stale offsets, which corrupts the part.  The innermost paragraph - the
    one that really owns the run - keeps it.
    """
    claimed: Set[Tuple[int, int]] = set()
    for paragraph in sorted(paragraphs, key=lambda p: p.end - p.start):
        kept = []
        for node in paragraph.nodes:
            key = (node.inner_start, node.inner_end)
            if key in claimed:
                continue
            claimed.add(key)
            kept.append(node)
        paragraph.nodes = kept
        if len(kept) != len(paragraph.char_map):
            paragraph.rebuild()
    return paragraphs


def _rewrite_node_inner(node: TextNode, edits: Sequence[Tuple[int, int, str]]) -> bytes:
    """Rebuild one node's inner bytes, applying ``(char_start, char_end, text)``."""
    out = bytearray()
    cursor = 0
    text = node.text
    total = len(text)
    for char_start, char_end, replacement in edits:
        if char_start > cursor:
            out += node.raw_inner[node.offsets[cursor]:node.offsets[char_start]]
        out += xml_escape(replacement).encode("utf-8")
        cursor = char_end
    if cursor < total:
        out += node.raw_inner[node.offsets[cursor]:]
    return bytes(out)


class Package:
    """An in-memory .docx: every part's bytes, plus the parts we index."""

    def __init__(self, path: str):
        self.path = path
        self.entries: List[zipfile.ZipInfo] = []
        self.parts: Dict[str, bytes] = {}
        with zipfile.ZipFile(path) as archive:
            for info in archive.infolist():
                self.entries.append(info)
                self.parts[info.filename] = archive.read(info.filename)
        self.documents: Dict[str, PartDocument] = {}
        for name, data in self.parts.items():
            if TEXT_PART_RE.match(name):
                self.documents[name] = PartDocument(name, data)

    def text_parts(self) -> List[PartDocument]:
        return [self.documents[name] for name in sorted(self.documents)]

    def cell_of(self) -> Dict[str, List[int]]:
        """Map every part's paragraph indices to the table cell holding them.

        A name broken by a narrow column ("KSH" | "Distriparks" | "Private
        Limited") is only a split name when the pieces share a cell; the same
        text side by side in two cells is two separate entries.  Body text
        gets ``-1``.
        """
        result: Dict[str, List[int]] = {}
        for name, document in self.documents.items():
            data = self.parts[name]
            spans = [(m.start(), m.end()) for m in CELL_RE.finditer(data)]
            cells = [-1] * len(document.paragraphs)
            for paragraph in document.paragraphs:
                low, high = 0, len(spans) - 1
                found = -1
                while low <= high:
                    middle = (low + high) // 2
                    start, end = spans[middle]
                    if paragraph.start < start:
                        high = middle - 1
                    elif paragraph.start >= end:
                        low = middle + 1
                    else:
                        found = middle
                        break
                cells[paragraph.index] = found
            result[name] = cells
        return result

    def scrub_metadata(self) -> Dict[str, str]:
        """Blank the author-ish properties in docProps/core.xml.

        Word stores the last editor's name there, which is PII even when it is
        absent from the visible text.  Returns the properties that were changed.
        """
        if METADATA_PART not in self.parts:
            return {}
        data = self.parts[METADATA_PART]
        changed: Dict[str, str] = {}
        for tag, new_value in (
            (b"dc:creator", b"Redaction Pipeline"),
            (b"cp:lastModifiedBy", b"Redaction Pipeline"),
            (b"dc:title", b"Redacted Document"),
        ):
            pattern = re.compile(
                rb"(<" + tag + rb"\b[^>]*>)(.*?)(</" + tag + rb">)", re.S
            )
            match = pattern.search(data)
            if match and match.group(2) != new_value:
                changed[tag.decode()] = match.group(2).decode("utf-8", "replace")
                data = data[: match.start(2)] + new_value + data[match.end(2):]
        self.parts[METADATA_PART] = data
        return changed

    def save(self, path: str) -> None:
        with zipfile.ZipFile(path, "w") as archive:
            for info in self.entries:
                clone = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                clone.compress_type = info.compress_type
                clone.external_attr = info.external_attr
                clone.internal_attr = info.internal_attr
                clone.create_system = info.create_system
                archive.writestr(clone, self.parts[info.filename])

    def apply(self, edits: List[dict]) -> int:
        """Apply ``{'part','paragraph','start','end','text'}`` edits.

        Edits are distributed over the original ``<w:t>`` nodes proportionally,
        so a match that straddles a run seam leaves the run structure intact.
        Re-edits of the same node are merged before the node is rewritten once.
        """
        by_part: Dict[str, Dict[Tuple[int, int], List[Tuple[int, int, str]]]] = {}
        for edit in edits:
            doc = self.documents[edit["part"]]
            paragraph = doc.paragraphs[edit["paragraph"]]
            spans = paragraph.node_span(edit["start"], edit["end"])
            if not spans:
                continue
            original = edit["text"]
            lengths = [span[2] - span[1] for span in spans]
            source_length = sum(lengths)
            # Proportional split of the replacement across the same nodes.
            budget = len(original)
            slices: List[str] = []
            for position, length in enumerate(lengths):
                if position == len(lengths) - 1:
                    take = budget
                else:
                    take = (length * budget + source_length // 2) // source_length
                    take = min(take, budget)
                slices.append(original[:take])
                original = original[take:]
                budget -= take
            bucket = by_part.setdefault(edit["part"], {})
            for span, replacement in zip(spans, slices):
                bucket.setdefault((edit["paragraph"], span[0]), []).append(
                    (span[1], span[2], replacement)
                )

        touched = 0
        for part_name, node_map in by_part.items():
            doc = self.documents[part_name]
            data = doc.data
            # Rewrite each touched element from last to first so the byte
            # offsets of the elements still to come stay valid.
            ordered = sorted(
                node_map.items(),
                key=lambda item: doc.paragraphs[item[0][0]].nodes[item[0][1]].open_start,
                reverse=True,
            )
            for (paragraph_index, node_index), node_edits in ordered:
                node = doc.paragraphs[paragraph_index].nodes[node_index]
                node_edits = sorted(set(node_edits))
                for previous, current in zip(node_edits, node_edits[1:]):
                    if current[0] < previous[1]:
                        raise ValueError(
                            "overlapping edits in %s paragraph %d node %d: %r / %r"
                            % (part_name, paragraph_index, node_index, previous, current)
                        )
                new_inner = _rewrite_node_inner(node, node_edits)
                if new_inner == node.raw_inner:
                    continue
                open_tag = node.raw_open_tag(data)
                if b"xml:space" not in open_tag and (
                    new_inner[:1] == b" " or new_inner[-1:] == b" "
                ):
                    # Word trims leading/trailing whitespace without this hint.
                    open_tag = open_tag[:-1].rstrip() + b' xml:space="preserve">'
                closing = b"</" + node.elem + b">"
                new_element = open_tag + new_inner + closing
                data = data[: node.open_start] + new_element + data[node.elem_end:]
                touched += 1
            doc.data = data
            self.parts[part_name] = data
        return touched


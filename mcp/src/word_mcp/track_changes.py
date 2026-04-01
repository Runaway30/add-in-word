"""
Utilities for writing OOXML Track Changes markup into .docx files.

Track changes in OOXML:
  - Deleted text: <w:del w:id="N" w:author="..." w:date="..."><w:r><w:delText>...</w:delText></w:r></w:del>
  - Inserted text: <w:ins w:id="N" w:author="..." w:date="..."><w:r><w:t>...</w:t></w:r></w:ins>
"""

from __future__ import annotations

import copy
from datetime import datetime, timezone
from typing import NamedTuple

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"

NSMAP = {"w": W_NS}


def _w(tag: str) -> str:
    return f"{W}{tag}"


def _make_del_run(text: str, rpr_elem=None) -> etree._Element:
    """Build a <w:r> containing <w:delText> for use inside <w:del>."""
    r = etree.Element(_w("r"))
    if rpr_elem is not None:
        r.append(copy.deepcopy(rpr_elem))
    del_text = etree.SubElement(r, _w("delText"))
    del_text.text = text
    if text and (text[0] == " " or text[-1] == " "):
        del_text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return r


def _make_ins_run(text: str, rpr_elem=None) -> etree._Element:
    """Build a <w:r> containing <w:t> for use inside <w:ins>."""
    r = etree.Element(_w("r"))
    if rpr_elem is not None:
        r.append(copy.deepcopy(rpr_elem))
    t = etree.SubElement(r, _w("t"))
    t.text = text
    if text and (text[0] == " " or text[-1] == " "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return r


def _make_normal_run(text: str, rpr_elem=None) -> etree._Element:
    """Build a plain <w:r><w:t>...</w:t></w:r>."""
    r = etree.Element(_w("r"))
    if rpr_elem is not None:
        r.append(copy.deepcopy(rpr_elem))
    t = etree.SubElement(r, _w("t"))
    t.text = text
    if text and (text[0] == " " or text[-1] == " "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return r


def _wrap_in_del(run_elem: etree._Element, change_id: int, author: str, date: str) -> etree._Element:
    del_elem = etree.Element(_w("del"))
    del_elem.set(_w("id"), str(change_id))
    del_elem.set(_w("author"), author)
    del_elem.set(_w("date"), date)
    del_elem.append(run_elem)
    return del_elem


def _wrap_in_ins(run_elem: etree._Element, change_id: int, author: str, date: str) -> etree._Element:
    ins_elem = etree.Element(_w("ins"))
    ins_elem.set(_w("id"), str(change_id))
    ins_elem.set(_w("author"), author)
    ins_elem.set(_w("date"), date)
    ins_elem.append(run_elem)
    return ins_elem


class Replacement(NamedTuple):
    old_text: str
    new_text: str


def _collect_paragraph_text_and_runs(p_elem: etree._Element) -> tuple[str, list[tuple[etree._Element, str]]]:
    """
    Returns (full_text, [(run_element, run_text), ...]).
    Only considers direct <w:r> children (not those inside w:ins/w:del already).
    """
    runs: list[tuple[etree._Element, str]] = []
    full_text = ""
    for child in p_elem:
        tag = etree.QName(child.tag).localname if child.tag else ""
        if tag == "r":
            t_elem = child.find(_w("t"))
            run_text = (t_elem.text or "") if t_elem is not None else ""
            runs.append((child, run_text))
            full_text += run_text
        elif tag in ("ins", "del"):
            # Skip already-tracked changes when collecting text
            pass
        elif tag == "hyperlink":
            for r in child.findall(_w("r")):
                t_elem = r.find(_w("t"))
                run_text = (t_elem.text or "") if t_elem is not None else ""
                runs.append((r, run_text))
                full_text += run_text
    return full_text, runs


def apply_replacement_to_paragraph(
    p_elem: etree._Element,
    old_text: str,
    new_text: str,
    author: str,
    date: str,
    id_counter: list[int],  # mutable so we can increment across calls
) -> bool:
    """
    Find `old_text` in the paragraph and replace it with track-change markup.
    Returns True if a replacement was made.
    """
    full_text, runs = _collect_paragraph_text_and_runs(p_elem)

    if old_text not in full_text:
        return False

    start_pos = full_text.index(old_text)
    end_pos = start_pos + len(old_text)

    # Build a character-offset map: for each char position, which run it belongs to
    # char_map[i] = (run_index, offset_within_run)
    char_map: list[tuple[int, int]] = []
    for run_idx, (_, run_text) in enumerate(runs):
        for offset in range(len(run_text)):
            char_map.append((run_idx, offset))

    # Get the rPr (run properties) from the first run of the old text, if any
    first_run_idx = char_map[start_pos][0] if char_map else 0
    first_run_elem = runs[first_run_idx][0] if runs else None
    rpr_elem = None
    if first_run_elem is not None:
        rpr_elem = first_run_elem.find(_w("rPr"))

    # We will rebuild the paragraph content.
    # Strategy: collect segments: (text, kind) where kind is "normal", "del", "ins"
    # Then build new XML elements.

    # Text before the change
    before = full_text[:start_pos]
    deleted = full_text[start_pos:end_pos]
    after = full_text[end_pos:]

    # Remove all existing plain runs from the paragraph
    # (preserve pPr, bookmarks, etc.)
    elements_to_remove = []
    for child in p_elem:
        tag = etree.QName(child.tag).localname if child.tag else ""
        if tag == "r":
            elements_to_remove.append(child)

    # Find insertion position (after pPr if present, else at start)
    insert_after = None
    ppr = p_elem.find(_w("pPr"))
    if ppr is not None:
        insert_after = ppr

    for elem in elements_to_remove:
        p_elem.remove(elem)

    # Build new elements
    new_elements: list[etree._Element] = []

    if before:
        new_elements.append(_make_normal_run(before, rpr_elem))

    if deleted:
        del_id = id_counter[0]
        id_counter[0] += 1
        ins_id = id_counter[0]
        id_counter[0] += 1

        del_run = _make_del_run(deleted, rpr_elem)
        new_elements.append(_wrap_in_del(del_run, del_id, author, date))

        if new_text:
            ins_run = _make_ins_run(new_text, rpr_elem)
            new_elements.append(_wrap_in_ins(ins_run, ins_id, author, date))

    if after:
        new_elements.append(_make_normal_run(after, rpr_elem))

    # Insert new elements into the paragraph
    if insert_after is not None:
        idx = list(p_elem).index(insert_after) + 1
        for i, elem in enumerate(new_elements):
            p_elem.insert(idx + i, elem)
    else:
        for elem in new_elements:
            p_elem.append(elem)

    return True

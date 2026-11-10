"""Put back lines the layout engine dropped.

pdf2docx occasionally loses a line (a table header drawn in white on a dark
bar, text sitting on a filled shape). After conversion we walk the PDF's
lines in order, find the ones that never reached the Word file, and insert
each as a paragraph next to its nearest neighbour that did.
"""
from __future__ import annotations

import copy
from collections import Counter

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

import accuracy

W_P = qn("w:p")
MC_FALLBACK = accuracy.MC_FALLBACK


def _body_paragraphs(body) -> list:
    """Top-level-visible w:p elements in document order (tables included),
    skipping paragraphs nested inside another paragraph (text boxes)."""
    out = []

    def walk(node):
        for child in node:
            if child.tag == MC_FALLBACK:
                continue
            if child.tag == W_P:
                out.append(child)
            else:
                walk(child)
    walk(body)
    return out


def find_missing(lines: list[dict], docx_text: str) -> list[tuple[int, dict]]:
    """Content lines that did not reach the docx, as (index, line)."""
    docx_tokens = accuracy.normalize(docx_text)
    remaining = Counter(docx_tokens)
    tokens = [None if ln["diagonal"] else accuracy.normalize(ln["text"]) for ln in lines]
    missing: set[int] = set()

    # Pass 1: lines whose words are mostly absent.
    for i, words in enumerate(tokens):
        if not words:
            continue
        have = Counter()
        for w in words:
            if remaining[w] - have[w] > 0:
                have[w] += 1
        if sum(have.values()) * 2 < len(words):
            missing.add(i)
        else:
            remaining -= have

    # Pass 2: short lines whose words also occur elsewhere ("Invoice", "Date").
    # A phrase is missing when it occurs more often in the PDF than in the docx;
    # counting on space-free text keeps mere spacing differences from triggering.
    present = [i for i, w in enumerate(tokens) if w and i not in missing]
    deficit = Counter(w for i in present for w in tokens[i]) - Counter(docx_tokens)
    if deficit:
        pdf_squash = "\x00".join("".join(tokens[i]) for i in present)
        docx_squash = "".join(docx_tokens)
        groups: dict[str, list[int]] = {}
        for i in present:
            if not Counter(tokens[i]) - deficit:        # every word is in short supply
                groups.setdefault("".join(tokens[i]), []).append(i)
        for key, idxs in sorted(groups.items(), key=lambda kv: -len(kv[0])):
            if len(key) < 3 or Counter(tokens[idxs[0]]) - deficit:
                continue
            n = pdf_squash.count(key) - docx_squash.count(key)
            if n <= 0:
                continue
            # dropped lines cluster: prefer instances next to an already-missing line
            idxs.sort(key=lambda i: (not ({i - 1, i + 1} & missing), i))
            for i in idxs[:n]:
                if Counter(tokens[i]) - deficit:
                    break
                missing.add(i)
                deficit -= Counter(tokens[i])

    return [(i, lines[i]) for i in sorted(missing)]


def repair(docx_path: str, lines: list[dict]) -> int:
    """Insert missing lines. Returns how many were inserted."""
    missing = find_missing(lines, accuracy.docx_text(docx_path))
    if not missing:
        return 0
    missing_idx = {i for i, _ in missing}

    doc = Document(docx_path)
    body = doc.element.body
    paras = _body_paragraphs(body)
    keys = [" ".join(accuracy.normalize(accuracy.paragraph_text(p))) for p in paras]

    def locate(text: str):
        key = " ".join(accuracy.normalize(text))
        if len(key) < 4:
            return None
        hits = [p for p, k in zip(paras, keys) if key in k]
        return hits[0] if len(hits) == 1 else None      # only trust a unique match

    def new_para(ln):
        p = OxmlElement("w:p")
        r = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        for tag in ("w:sz", "w:szCs"):
            sz = OxmlElement(tag)
            sz.set(qn("w:val"), str(int(round(Pt(ln["size"]).pt * 2))))
            rpr.append(sz)
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = ln["text"]
        r.append(rpr)
        r.append(t)
        p.append(r)
        return p

    inserted = 0
    last_inserted: dict[int, object] = {}     # line index -> element we added
    for i, ln in missing:
        anchor, before = None, False
        # nearest earlier line on the same page that exists in the docx
        j = i - 1
        while j >= 0 and lines[j]["page"] == ln["page"] and anchor is None:
            anchor = last_inserted.get(j)
            if anchor is None and j not in missing_idx and not lines[j]["diagonal"]:
                anchor = locate(lines[j]["text"])
            j -= 1
        if anchor is None:                    # else the nearest later line: insert before it
            j = i + 1
            while j < len(lines) and lines[j]["page"] == ln["page"] and anchor is None:
                if j not in missing_idx and not lines[j]["diagonal"]:
                    anchor = locate(lines[j]["text"])
                j += 1
            before = anchor is not None
        p = new_para(ln)
        if anchor is None:
            sect = body.find(qn("w:sectPr"))
            (sect.addprevious(p) if sect is not None else body.append(p))
        elif before:
            anchor.addprevious(p)
        else:
            anchor.addnext(p)
        last_inserted[i] = p
        inserted += 1

    doc.save(docx_path)
    return inserted

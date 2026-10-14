# -*- coding: utf-8 -*-
"""Inherit {{p var }} paragraph style / alignment onto Subdoc body paragraphs."""
from __future__ import annotations

import re
from collections import namedtuple
from typing import Any, Dict, Optional

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

PlaceholderStyle = namedtuple("PlaceholderStyle", ["p_style", "alignment"])

_P_VAR_RE = re.compile(
    r"\{\{\s*p\s+([A-Za-z_][A-Za-z0-9_]*)\s*\}\}",
    re.DOTALL,
)

_ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

_SKIP_STYLE_PREFIXES = ("Heading", "List", "Title", "Subtitle")


def scan_p_placeholder_styles(tpl) -> Dict[str, PlaceholderStyle]:
    """Map variable name -> style from paragraphs containing {{p var }}.

    Must run on template XML before patch_xml strips those paragraphs.
    First occurrence wins.
    """
    tpl.init_docx()
    styles: Dict[str, PlaceholderStyle] = {}
    body = tpl.docx._element.body
    for p in body.iter(qn("w:p")):
        texts = [t.text or "" for t in p.iter(qn("w:t"))]
        joined = "".join(texts)
        compact = re.sub(r"\s+", " ", joined).strip()
        m = _P_VAR_RE.search(compact)
        if not m:
            continue
        without = _P_VAR_RE.sub("", compact).strip()
        if without:
            continue
        name = m.group(1)
        if name in styles:
            continue
        p_style = None
        alignment = None
        pPr = p.find(qn("w:pPr"))
        if pPr is not None:
            pStyle_el = pPr.find(qn("w:pStyle"))
            if pStyle_el is not None:
                p_style = pStyle_el.get(qn("w:val"))
            jc = pPr.find(qn("w:jc"))
            if jc is not None:
                alignment = jc.get(qn("w:val"))
        styles[name] = PlaceholderStyle(p_style=p_style, alignment=alignment)
    return styles


def _apply_placeholder_style_to_subdoc(subdoc, style_info: Optional[PlaceholderStyle]):
    if not style_info:
        return
    if not style_info.p_style and not style_info.alignment:
        return
    body = subdoc.element.body
    for child in list(body):
        if child.tag != qn("w:p"):
            continue
        para = None
        for p in subdoc.paragraphs:
            if p._element is child:
                para = p
                break
        if para is None:
            continue
        style_name = para.style.name if para.style is not None else "Normal"
        if any(style_name.startswith(prefix) for prefix in _SKIP_STYLE_PREFIXES):
            continue
        if style_info.p_style:
            try:
                para.style = style_info.p_style
            except (KeyError, ValueError):
                pass
        if style_info.alignment and style_info.alignment in _ALIGN_MAP:
            para.alignment = _ALIGN_MAP[style_info.alignment]


def _is_subdoc(value: Any) -> bool:
    try:
        from .subdoc import Subdoc
    except ImportError:
        return False
    return isinstance(value, Subdoc)


def _apply_styles_to_value(
    value: Any,
    placeholder_styles: Dict[str, PlaceholderStyle],
    var_name: Optional[str] = None,
) -> None:
    if _is_subdoc(value):
        if var_name and var_name in placeholder_styles:
            _apply_placeholder_style_to_subdoc(value, placeholder_styles[var_name])
        return
    if isinstance(value, list):
        for item in value:
            _apply_styles_to_value(item, placeholder_styles, var_name=None)
        return
    if isinstance(value, dict):
        for k, v in value.items():
            _apply_styles_to_value(
                v, placeholder_styles, var_name=k if isinstance(k, str) else None
            )


def apply_p_placeholder_styles(tpl, context: Any) -> None:
    """Apply {{p var }} placeholder styles onto Subdoc values in context (in place)."""
    if not isinstance(context, dict):
        return
    placeholder_styles = scan_p_placeholder_styles(tpl)
    if not placeholder_styles:
        return
    _apply_styles_to_value(context, placeholder_styles)

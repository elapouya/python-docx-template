# -*- coding: utf-8 -*-
"""JSON-serializable envelopes for Subdoc and RichText in CLI / embedders."""
from __future__ import annotations

import base64
import io
from typing import Any, Optional

from .richtext import RichText

DOCXTPL_KEY = "__docxtpl__"


def is_envelope(value: Any) -> bool:
    return isinstance(value, dict) and DOCXTPL_KEY in value


def _expand_subdoc(tpl, envelope: dict):
    try:
        from .subdoc import Subdoc  # noqa: F401
    except ImportError as e:
        raise RuntimeError(
            'Subdoc envelopes require docxcompose. Install with: pip install "docxtpl[subdoc]"'
        ) from e

    has_b64 = "docx_base64" in envelope and envelope["docx_base64"] is not None
    has_path = "path" in envelope and envelope["path"]
    if has_b64 and has_path:
        raise RuntimeError("Subdoc envelope must not set both docx_base64 and path")
    if not has_b64 and not has_path:
        raise RuntimeError("Subdoc envelope requires docx_base64 or path")

    if has_path:
        return tpl.new_subdoc(envelope["path"])
    try:
        raw = base64.b64decode(envelope["docx_base64"])
    except Exception as e:
        raise RuntimeError("Invalid docx_base64 in subdoc envelope") from e
    return tpl.new_subdoc(io.BytesIO(raw))


def _expand_richtext(tpl, envelope: dict) -> RichText:
    runs = envelope.get("runs")
    if runs is None:
        raise RuntimeError("RichText envelope requires a runs array")
    if not isinstance(runs, list):
        raise RuntimeError("RichText envelope runs must be a list")

    rt = RichText()
    for i, run in enumerate(runs):
        if not isinstance(run, dict):
            raise RuntimeError("RichText run #%s must be an object" % i)
        if "text" not in run:
            raise RuntimeError("RichText run #%s requires text" % i)
        kwargs = {}
        for key in (
            "bold",
            "italic",
            "underline",
            "strike",
            "subscript",
            "superscript",
            "color",
            "size",
            "font",
            "style",
        ):
            if key in run and run[key] is not None:
                kwargs[key] = run[key]
        url = run.get("url")
        if url:
            kwargs["url_id"] = tpl.build_url_id(url)
        rt.add(run["text"], **kwargs)
    return rt


def expand_value(tpl, value: Any) -> Any:
    if is_envelope(value):
        kind = value.get(DOCXTPL_KEY)
        if kind == "subdoc":
            return _expand_subdoc(tpl, value)
        if kind == "richtext":
            return _expand_richtext(tpl, value)
        raise RuntimeError("Unknown __docxtpl__ envelope type: %r" % kind)

    if isinstance(value, list):
        return [expand_value(tpl, item) for item in value]
    if isinstance(value, dict):
        return {k: expand_value(tpl, v) for k, v in value.items()}
    return value


def expand_context(tpl, context: Any) -> Any:
    """Expand __docxtpl__ envelopes in a JSON-loaded context before render()."""
    return expand_value(tpl, context)

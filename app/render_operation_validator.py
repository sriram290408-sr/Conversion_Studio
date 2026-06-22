"""Validate constrained HF/metadata render operations before COM execution."""

from __future__ import annotations
from typing import Any, Dict, Iterable, List, Optional, Set

ALLOWED_OPERATIONS: Set[str] = {
    "create_card",
    "create_kpi",
    "create_gauge",
    "create_slicer",
    "create_column_chart",
    "create_bar_chart",
    "create_line_chart",
    "create_area_chart",
    "create_pie_chart",
    "create_donut_chart",
    "create_table",
    "create_matrix",
    "create_treemap",
    "create_map",
    "create_textbox",
    "create_image",
    "create_shape",
    "create_placeholder",
}


def validate_render_operations(
    generated: Dict[str, Any],
    source_plan: Dict[str, Any],
    allowed_pages: Optional[Iterable[str]] = None,
) -> Dict[str, Any]:
    source_by_id = {
        str(item.get("visual_id")): item
        for item in source_plan.get("operations", [])
        if item.get("visual_id")
    }
    page_set = {str(x) for x in (allowed_pages or [])}
    valid: List[Dict[str, Any]] = []
    rejected: List[Dict[str, Any]] = []

    for item in generated.get("operations", []) if isinstance(generated, dict) else []:
        reasons: List[str] = []
        if not isinstance(item, dict):
            rejected.append(
                {"operation": item, "reasons": ["Operation is not an object."]}
            )
            continue
        visual_id = str(item.get("visual_id") or "")
        source = source_by_id.get(visual_id)
        if source is None:
            reasons.append("Unknown visual_id.")
        op = str(item.get("op") or "")
        if op not in ALLOWED_OPERATIONS:
            reasons.append("Unsupported operation.")
        page_name = str(item.get("page_name") or (source or {}).get("page_name") or "")
        if page_set and page_name not in page_set:
            reasons.append("Unknown page_name.")
        if any(
            key in item
            for key in ("code", "python", "shell", "command", "imports", "url", "path")
        ):
            reasons.append("Executable or external-access fields are forbidden.")

        layout = dict(item.get("layout") or (source or {}).get("layout") or {})
        for key in ("x_percent", "y_percent", "width_percent", "height_percent"):
            if key in layout:
                try:
                    value = float(layout[key])
                except (TypeError, ValueError):
                    reasons.append(f"Invalid {key}.")
                    continue
                if key in {"width_percent", "height_percent"} and not (
                    0.5 <= value <= 100
                ):
                    reasons.append(f"{key} is out of range.")
                elif key in {"x_percent", "y_percent"} and not (0 <= value <= 100):
                    reasons.append(f"{key} is out of range.")

        if reasons:
            rejected.append({"operation": item, "reasons": reasons})
            continue

        merged = dict(source or {})
        merged.update(
            {k: v for k, v in item.items() if k not in {"binding", "measure", "field"}}
        )
        # Never accept semantic bindings invented by HF.
        merged["visual_id"] = visual_id
        merged["page_name"] = page_name
        valid.append(merged)

    # Deterministic fallback: preserve source operations not returned by HF.
    seen = {str(item.get("visual_id")) for item in valid}
    for visual_id, source in source_by_id.items():
        if visual_id not in seen:
            valid.append(dict(source))

    return {
        "version": "1.0",
        "operations": valid,
        "rejected_operations": rejected,
        "warnings": [f"{len(rejected)} HF operation(s) rejected."] if rejected else [],
    }


__all__ = ["ALLOWED_OPERATIONS", "validate_render_operations"]

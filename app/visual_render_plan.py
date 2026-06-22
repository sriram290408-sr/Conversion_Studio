"""Build a deterministic Excel render plan from metadata, descriptions and HF vision."""

from __future__ import annotations

import math
import re
from difflib import SequenceMatcher
from typing import Any, Dict, Iterable, List, Optional, Tuple

_TYPE_ALIASES = {
    "card": "card",
    "kpi": "kpi",
    "gauge": "gauge",
    "slicer": "slicer",
    "filter": "slicer",
    "column_chart": "column_chart",
    "bar_chart": "bar_chart",
    "line_chart": "line_chart",
    "area_chart": "area_chart",
    "pie_chart": "pie_chart",
    "donut_chart": "donut_chart",
    "doughnut_chart": "donut_chart",
    "table": "table",
    "tableex": "table",
    "matrix": "matrix",
    "treemap": "treemap",
    "map": "map",
    "filledmap": "map",
    "textbox": "textbox",
    "text": "textbox",
    "image": "image",
    "shape": "shape",
}

_OPERATION_BY_TYPE = {
    "card": "create_card",
    "kpi": "create_kpi",
    "gauge": "create_gauge",
    "slicer": "create_slicer",
    "column_chart": "create_column_chart",
    "bar_chart": "create_bar_chart",
    "line_chart": "create_line_chart",
    "area_chart": "create_area_chart",
    "pie_chart": "create_pie_chart",
    "donut_chart": "create_donut_chart",
    "table": "create_table",
    "matrix": "create_matrix",
    "treemap": "create_treemap",
    "map": "create_map",
    "textbox": "create_textbox",
    "image": "create_image",
    "shape": "create_shape",
}


def _get(obj: Any, key: str, default: Any = None) -> Any:
    if isinstance(obj, dict):
        return obj.get(key, default)
    return getattr(obj, key, default)


def _norm_text(value: Any) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip().casefold()
    return re.sub(r"[^a-z0-9 ]+", "", text)


def _norm_type(value: Any) -> str:
    raw = str(value or "").strip().casefold().replace("-", "_").replace(" ", "_")
    compact = raw.replace("_", "")
    aliases = {
        "clusteredcolumnchart": "column_chart",
        "columnchart": "column_chart",
        "stackedcolumnchart": "column_chart",
        "clusteredbarchart": "bar_chart",
        "barchart": "bar_chart",
        "stackedbarchart": "bar_chart",
        "linechart": "line_chart",
        "areachart": "area_chart",
        "piechart": "pie_chart",
        "donutchart": "donut_chart",
        "doughnutchart": "donut_chart",
        "gaugevisual": "gauge",
        "cardvisual": "card",
        "multirowcard": "card",
        "kpivisual": "kpi",
        "listslicer": "slicer",
        "dropdownslicer": "slicer",
        "tableex": "table",
    }
    if compact in aliases:
        return aliases[compact]
    if raw in _TYPE_ALIASES:
        return _TYPE_ALIASES[raw]
    if "gauge" in raw:
        return "gauge"
    if "slicer" in raw:
        return "slicer"
    if "card" in raw or "kpi" in raw:
        return "kpi"
    if "column" in raw:
        return "column_chart"
    if "bar" in raw:
        return "bar_chart"
    if "line" in raw:
        return "line_chart"
    if "area" in raw:
        return "area_chart"
    if "donut" in raw or "doughnut" in raw:
        return "donut_chart"
    if "pie" in raw:
        return "pie_chart"
    if "matrix" in raw:
        return "matrix"
    if "table" in raw:
        return "table"
    if "treemap" in raw:
        return "treemap"
    if "map" in raw:
        return "map"
    return "placeholder"


def _pct_layout_from_metadata(
    layout: Dict[str, Any], canvas: Dict[str, float]
) -> Dict[str, float]:
    if all(
        k in layout
        for k in ("x_percent", "y_percent", "width_percent", "height_percent")
    ):
        return {
            k: float(layout.get(k) or 0)
            for k in ("x_percent", "y_percent", "width_percent", "height_percent")
        }
    cw = max(float(canvas.get("width") or 1280), 1.0)
    ch = max(float(canvas.get("height") or 720), 1.0)
    return {
        "x_percent": float(layout.get("x") or 0) / cw * 100,
        "y_percent": float(layout.get("y") or 0) / ch * 100,
        "width_percent": float(layout.get("width") or 0) / cw * 100,
        "height_percent": float(layout.get("height") or 0) / ch * 100,
    }


def _center(box: Dict[str, float]) -> Tuple[float, float]:
    return (
        box["x_percent"] + box["width_percent"] / 2,
        box["y_percent"] + box["height_percent"] / 2,
    )


def _position_score(a: Dict[str, float], b: Dict[str, float]) -> float:
    ax, ay = _center(a)
    bx, by = _center(b)
    distance = math.hypot(ax - bx, ay - by)
    return max(0.0, 1.0 - distance / 75.0)


def _size_score(a: Dict[str, float], b: Dict[str, float]) -> float:
    dw = abs(a["width_percent"] - b["width_percent"])
    dh = abs(a["height_percent"] - b["height_percent"])
    return max(0.0, 1.0 - (dw + dh) / 100.0)


def _title_score(a: Any, b: Any) -> float:
    na, nb = _norm_text(a), _norm_text(b)
    if not na or not nb:
        return 0.0
    return SequenceMatcher(None, na, nb).ratio()


def _match_visuals(
    metadata_visuals: List[Dict[str, Any]],
    vision_blocks: List[Dict[str, Any]],
    canvas: Dict[str, float],
) -> Dict[str, Dict[str, Any]]:
    available = set(range(len(vision_blocks)))
    matches: Dict[str, Dict[str, Any]] = {}
    for visual in metadata_visuals:
        vid = str(
            visual.get("chunk_id")
            or visual.get("visual_id")
            or visual.get("name")
            or ""
        )
        m_type = _norm_type(visual.get("visual_type") or visual.get("type"))
        m_layout = _pct_layout_from_metadata(dict(visual.get("layout") or {}), canvas)
        best_index, best_score = None, -1.0
        for index in available:
            block = vision_blocks[index]
            b_type = _norm_type(block.get("block_type") or block.get("visual_type"))
            b_layout = {
                "x_percent": float(block.get("x_percent") or 0),
                "y_percent": float(block.get("y_percent") or 0),
                "width_percent": float(block.get("width_percent") or 0),
                "height_percent": float(block.get("height_percent") or 0),
            }
            type_score = (
                1.0 if m_type == b_type else (0.35 if b_type == "placeholder" else 0.0)
            )
            score = (
                0.42 * type_score
                + 0.28 * _position_score(m_layout, b_layout)
                + 0.20
                * _title_score(
                    visual.get("visual_title") or visual.get("title"),
                    block.get("title"),
                )
                + 0.10 * _size_score(m_layout, b_layout)
            )
            if score > best_score:
                best_score, best_index = score, index
        if best_index is not None and best_score >= 0.25:
            matches[vid] = {
                **vision_blocks[best_index],
                "_match_score": round(best_score, 4),
            }
            available.remove(best_index)
    return matches


def build_visual_render_plan(
    final_chunks: Dict[str, Any],
    screenshot_analysis: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    visuals = list(
        final_chunks.get("visual_chunks") or final_chunks.get("visuals") or []
    )
    screenshot_analysis = screenshot_analysis or {}
    blocks = list(screenshot_analysis.get("visual_blocks") or [])
    canvas = final_chunks.get("canvas") or {"width": 1280, "height": 720}
    matches = _match_visuals(visuals, blocks, canvas)
    theme = dict(screenshot_analysis.get("theme") or final_chunks.get("theme") or {})
    operations: List[Dict[str, Any]] = []

    for visual in visuals:
        visual_id = str(
            visual.get("chunk_id")
            or visual.get("visual_id")
            or visual.get("name")
            or "visual_unknown"
        )
        metadata_type = _norm_type(visual.get("visual_type") or visual.get("type"))
        match = matches.get(visual_id) or {}
        # Metadata is authoritative. Vision can refine placement/style, not change a confirmed type.
        final_type = (
            metadata_type
            if metadata_type != "placeholder"
            else _norm_type(match.get("block_type"))
        )
        source_layout = dict(visual.get("layout") or {})
        pct_layout = _pct_layout_from_metadata(source_layout, canvas)
        if match:
            pct_layout = {
                "x_percent": float(match.get("x_percent") or pct_layout["x_percent"]),
                "y_percent": float(match.get("y_percent") or pct_layout["y_percent"]),
                "width_percent": float(
                    match.get("width_percent") or pct_layout["width_percent"]
                ),
                "height_percent": float(
                    match.get("height_percent") or pct_layout["height_percent"]
                ),
            }
        style = dict(visual.get("_visual_theme") or visual.get("style") or {})
        style.update(
            {
                k: v
                for k, v in dict(match.get("style") or {}).items()
                if v not in (None, "")
            }
        )
        title = str(
            visual.get("visual_title")
            or visual.get("title")
            or match.get("title")
            or ""
        ).strip()

        operation = {
            "op": _OPERATION_BY_TYPE.get(final_type, "create_placeholder"),
            "visual_id": visual_id,
            "page_name": str(visual.get("page_name") or "Dashboard"),
            "visual_type": final_type,
            "title": title,
            "description": str(
                visual.get("visual_description") or visual.get("description") or ""
            ),
            "layout": {**source_layout, **pct_layout},
            "style": style,
            "settings": dict(visual.get("settings") or {}),
            # Semantic data is copied into the deterministic source plan so it
            # survives validation and can be reattached after HF presentation
            # refinement. HF is never allowed to invent or replace these keys.
            "semantic": {
                "rows": list(
                    visual.get("rows")
                    or visual.get("axis")
                    or visual.get("uses_columns")
                    or []
                ),
                "columns": list(visual.get("columns") or []),
                "measures": list(
                    visual.get("measures")
                    or visual.get("values")
                    or visual.get("uses_measures")
                    or []
                ),
                "legend": list(visual.get("legend") or []),
                "filters": list(
                    visual.get("filters") or visual.get("visual_filters") or []
                ),
                "slicer_field": visual.get("slicer_field"),
                "minimum": visual.get("minimum") or visual.get("min_value"),
                "maximum": visual.get("maximum") or visual.get("max_value"),
            },
            "source": {
                "metadata_type": metadata_type,
                "vision_type": _norm_type(match.get("block_type")) if match else None,
                "vision_match_score": match.get("_match_score"),
            },
        }
        operations.append(operation)

    # Recover high-confidence screenshot-only cards when PBIX extraction is layout-only.
    matched_block_ids = {id(matches[key]) for key in matches}
    unmatched_blocks = [block for block in blocks if id(block) not in matched_block_ids]
    existing_measure_candidates: List[Any] = []
    for visual in visuals:
        existing_measure_candidates.extend(list(visual.get("measures") or visual.get("values") or visual.get("uses_measures") or []))
    for index, block in enumerate(unmatched_blocks, start=1):
        block_type = _norm_type(block.get("block_type") or block.get("visual_type"))
        confidence = float(block.get("confidence") or block.get("score") or 0.0)
        if block_type not in {"card", "kpi"} or confidence < 0.80:
            continue
        title = str(block.get("title") or "").strip()
        normalized_title = _norm_text(title)
        best_measure = None
        best_score = 0.0
        for measure in existing_measure_candidates:
            if isinstance(measure, dict):
                label = measure.get("display_name") or measure.get("measure_name") or measure.get("column_name") or measure.get("reference")
            else:
                label = str(measure)
            score = _title_score(normalized_title, label)
            if score > best_score:
                best_measure, best_score = measure, score
        if best_measure is None or best_score < 0.45:
            continue
        recovered_id = f"visual_recovered_card_{index:03d}"
        operations.append({
            "op": "create_card",
            "visual_id": recovered_id,
            "page_name": str(block.get("page_name") or (visuals[0].get("page_name") if visuals else "Dashboard")),
            "visual_type": "card",
            "title": title,
            "description": "Screenshot-recovered live card",
            "layout": {
                "x_percent": float(block.get("x_percent") or 0),
                "y_percent": float(block.get("y_percent") or 0),
                "width_percent": float(block.get("width_percent") or 0),
                "height_percent": float(block.get("height_percent") or 0),
                "x": float(block.get("x") or float(block.get("x_percent") or 0) / 100 * float(canvas.get("width") or 1280)),
                "y": float(block.get("y") or float(block.get("y_percent") or 0) / 100 * float(canvas.get("height") or 720)),
                "width": float(block.get("width") or float(block.get("width_percent") or 0) / 100 * float(canvas.get("width") or 1280)),
                "height": float(block.get("height") or float(block.get("height_percent") or 0) / 100 * float(canvas.get("height") or 720)),
                "layout_source": "screenshot_recovered",
            },
            "style": dict(block.get("style") or {}),
            "settings": {"number_format": "#,##0.00"},
            "semantic": {"rows": [], "columns": [], "measures": [best_measure], "legend": [], "filters": []},
            "source": {"metadata_type": None, "vision_type": "card", "vision_match_score": confidence, "recovered": True},
        })

    return {
        "version": "1.0",
        "theme": theme,
        "canvas": canvas,
        "operations": operations,
        "warnings": [],
    }


def merge_render_plan_into_chunks(
    final_chunks: Dict[str, Any], render_plan: Dict[str, Any]
) -> Dict[str, Any]:
    by_id = {str(op.get("visual_id")): op for op in render_plan.get("operations", [])}
    for visual in final_chunks.get("visual_chunks", []) or []:
        visual_id = str(
            visual.get("chunk_id")
            or visual.get("visual_id")
            or visual.get("name")
            or ""
        )
        operation = by_id.get(visual_id)
        if not operation:
            continue
        visual["render_operation"] = operation
        visual["render_style"] = operation.get("style") or {}
        visual["render_layout"] = operation.get("layout") or {}
        # Use screenshot/PBIX placement for presentation only.
        visual["layout"] = dict(operation.get("layout") or visual.get("layout") or {})

        semantic = dict(operation.get("semantic") or {})
        for key in ("rows", "columns", "measures", "legend", "filters"):
            if semantic.get(key) and not visual.get(key):
                visual[key] = semantic[key]
        if semantic.get("slicer_field") and not visual.get("slicer_field"):
            visual["slicer_field"] = semantic["slicer_field"]
        for key in ("minimum", "maximum"):
            if semantic.get(key) is not None and visual.get(key) is None:
                visual[key] = semantic[key]
    existing_ids = {str(v.get("chunk_id") or v.get("visual_id") or v.get("name") or "") for v in final_chunks.get("visual_chunks", []) or []}
    for operation in render_plan.get("operations", []) or []:
        visual_id = str(operation.get("visual_id") or "")
        if not visual_id or visual_id in existing_ids or not (operation.get("source") or {}).get("recovered"):
            continue
        semantic = dict(operation.get("semantic") or {})
        final_chunks.setdefault("visual_chunks", []).append({
            "chunk_id": visual_id, "visual_id": visual_id,
            "page_name": operation.get("page_name"),
            "visual_type": operation.get("visual_type"),
            "visual_title": operation.get("title"),
            "layout": dict(operation.get("layout") or {}),
            "rows": list(semantic.get("rows") or []),
            "columns": list(semantic.get("columns") or []),
            "measures": list(semantic.get("measures") or []),
            "legend": list(semantic.get("legend") or []),
            "filters": list(semantic.get("filters") or []),
            "render_operation": operation,
            "render_style": dict(operation.get("style") or {}),
            "recovered_from_screenshot": True,
        })
    final_chunks["visual_render_plan"] = render_plan
    return final_chunks


__all__ = ["build_visual_render_plan", "merge_render_plan_into_chunks"]

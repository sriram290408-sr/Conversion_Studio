"""Create normalized live Excel visual bindings from converter chunks."""

from __future__ import annotations

import re
from typing import Any, Dict, List

try:
    from field_normalizer import normalize_field
    from filter_engine import normalize_filters
    from live_binding_models import VisualBinding
except ImportError:
    from field_normalizer import normalize_field
    from filter_engine import normalize_filters
    from live_binding_models import VisualBinding


def _reference_leaf(value: Any) -> str:
    """Return the final semantic token from a field reference."""
    if isinstance(value, dict):
        return str(
            value.get("measure_name")
            or value.get("column_name")
            or value.get("display_name")
            or value.get("canonical_reference")
            or value.get("raw_reference")
            or ""
        ).strip()

    text = str(value or "").strip()
    bracket_parts = re.findall(r"\[([^\]]+)\]", text)
    if bracket_parts:
        return bracket_parts[-1].strip()

    if "." in text:
        return text.rsplit(".", 1)[-1].strip()

    return text.strip("[]' ")


def _implicit_measure_from_caption(value: Any) -> Dict[str, Any] | None:
    """Convert captions such as ``Sum of passengers`` into implicit measures.

    Table information is preserved when the caption contains a semantic
    reference such as ``Sum of flights[passengers]``. Bare numeric columns are
    also retained as implicit measures instead of being promoted to named
    measures.
    """
    if isinstance(value, dict):
        raw = str(
            value.get("display_name")
            or value.get("raw_reference")
            or value.get("canonical_reference")
            or value.get("measure_name")
            or ""
        ).strip()
        existing_table = str(value.get("table_name") or "").strip()
        existing_column = str(value.get("column_name") or "").strip()
    else:
        raw = str(value or "").strip()
        existing_table = ""
        existing_column = ""

    match = re.fullmatch(
        r"(?i)\s*(sum|average|avg|min|max|count|counta|distinctcount|distinct count)"
        r"\s+of\s+(.+?)\s*",
        raw,
    )
    if not match:
        return None

    aggregation = match.group(1).upper().replace(" ", "")
    if aggregation == "AVG":
        aggregation = "AVERAGE"

    reference = match.group(2).strip()
    normalized = normalize_field(reference, role="values")
    table_name = str(normalized.get("table_name") or existing_table).strip()
    column_name = str(
        normalized.get("column_name") or existing_column or _reference_leaf(reference)
    ).strip()
    canonical = (
        f"{table_name}[{column_name}]" if table_name and column_name else column_name
    )
    raw_reference = (
        f"{aggregation}({canonical})" if canonical else f"{aggregation}({reference})"
    )

    return {
        "raw_reference": raw_reference,
        "field_type": "implicit_measure",
        "role": "values",
        "table_name": table_name,
        "column_name": column_name,
        "measure_name": "",
        "aggregation": aggregation,
        "canonical_reference": canonical or reference,
        "cube_measure_path": "",
        "display_name": f"{aggregation.title()} of {column_name or reference}",
    }


def _named_measure_descriptor(name: str) -> Dict[str, Any]:
    clean_name = str(name or "").strip()
    return {
        "raw_reference": clean_name,
        "field_type": "named_measure",
        "role": "values",
        "table_name": "",
        "column_name": "",
        "measure_name": clean_name,
        "aggregation": "",
        "canonical_reference": clean_name,
        "cube_measure_path": f"[Measures].[{clean_name}]",
        "display_name": clean_name,
    }


def _promote_known_measures(
    dimensions: List[str],
    measures: List[Dict[str, Any]],
    known_measure_names: set[str],
) -> tuple[List[str], List[Dict[str, Any]]]:
    """Move Table[Measure] references from dimensions into measures."""
    retained: List[str] = []

    for reference in dimensions:
        leaf = _reference_leaf(reference)
        if leaf.casefold() in known_measure_names:
            _append_unique(
                measures,
                _named_measure_descriptor(leaf),
                "canonical_reference",
            )
        else:
            retained.append(reference)

    return retained, measures


def _repair_measure_descriptors(
    measures: List[Dict[str, Any]],
    known_measure_names: set[str],
) -> List[Dict[str, Any]]:
    """Preserve named measures and repair value-role columns as implicit measures."""
    repaired: List[Dict[str, Any]] = []

    for measure in measures:
        descriptor = (
            dict(measure)
            if isinstance(measure, dict)
            else normalize_field(measure, role="values")
        )
        leaf = _reference_leaf(descriptor)
        field_type = str(descriptor.get("field_type") or "").casefold()
        aggregation = str(descriptor.get("aggregation") or "").upper()

        if leaf.casefold() in known_measure_names or field_type == "named_measure":
            candidate = _named_measure_descriptor(leaf)
        else:
            candidate = _implicit_measure_from_caption(descriptor)
            if candidate is None and field_type == "implicit_measure":
                candidate = descriptor
            if candidate is None:
                table_name = str(descriptor.get("table_name") or "").strip()
                column_name = str(descriptor.get("column_name") or leaf).strip()
                # A field used in the values role is an implicit aggregation,
                # not a named measure. SUM is Excel/Power BI's safe default.
                aggregation = aggregation or "SUM"
                canonical = (
                    f"{table_name}[{column_name}]"
                    if table_name and column_name
                    else column_name
                )
                candidate = {
                    **descriptor,
                    "raw_reference": f"{aggregation}({canonical})",
                    "field_type": "implicit_measure",
                    "role": "values",
                    "table_name": table_name,
                    "column_name": column_name,
                    "measure_name": "",
                    "aggregation": aggregation,
                    "canonical_reference": canonical,
                    "cube_measure_path": "",
                    "display_name": f"{aggregation.title()} of {column_name}",
                }

        _append_unique(repaired, candidate, "raw_reference")

    return repaired


def _measure_identity(measure: Any) -> str:
    """Return a stable semantic identity used to remove duplicate measures."""
    if not isinstance(measure, dict):
        return str(measure or "").strip().casefold()
    field_type = str(measure.get("field_type") or "").casefold()
    if field_type == "named_measure":
        return f"named:{str(measure.get('measure_name') or _reference_leaf(measure)).casefold()}"
    table = str(measure.get("table_name") or "").strip().casefold()
    column = (
        str(measure.get("column_name") or _reference_leaf(measure)).strip().casefold()
    )
    aggregation = str(measure.get("aggregation") or "SUM").strip().upper()
    return f"implicit:{table}:{column}:{aggregation}"


def _dedupe_measures(measures: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    result: List[Dict[str, Any]] = []
    seen: set[str] = set()
    for measure in measures or []:
        key = _measure_identity(measure)
        if key and key not in seen:
            seen.add(key)
            result.append(measure)
    return result


def _select_measure_for_visual(
    measures: List[Dict[str, Any]], title: str, visual_type: str
) -> List[Dict[str, Any]]:
    """Choose the one semantic metric that belongs to a card or gauge.

    The visual title is treated as a strong semantic signal. For example:
    - ``Total Passengers`` prefers SUM(passengers).
    - ``Average Passengers`` prefers an AVERAGE aggregation or named average measure.
    This prevents a named average measure from replacing an implicit total merely
    because both contain the token ``passengers``.
    """
    unique = _dedupe_measures(measures)
    if not unique:
        return []
    if "multirow" in str(visual_type or "").casefold():
        return unique

    title_text = str(title or "").casefold()
    title_tokens = set(re.findall(r"[a-z0-9]+", title_text))

    wants_sum = any(
        token in title_tokens
        for token in ("total", "sum", "volume", "sales", "revenue", "amount")
    )
    wants_average = any(token in title_tokens for token in ("average", "avg", "mean"))
    wants_count = any(
        token in title_tokens for token in ("count", "number", "customers", "orders")
    )
    wants_min = any(token in title_tokens for token in ("minimum", "min", "lowest"))
    wants_max = any(token in title_tokens for token in ("maximum", "max", "highest"))

    scored = []
    for index, measure in enumerate(unique):
        text = " ".join(
            str(measure.get(key) or "")
            for key in (
                "display_name",
                "measure_name",
                "column_name",
                "raw_reference",
                "canonical_reference",
            )
        ).casefold()
        tokens = set(re.findall(r"[a-z0-9]+", text))
        aggregation = str(measure.get("aggregation") or "").upper()
        field_type = str(measure.get("field_type") or "").casefold()

        score = len(title_tokens & tokens) * 4

        # Strong aggregation/title agreement.
        if wants_sum:
            if aggregation == "SUM" or "sum of" in text:
                score += 12
            if "average" in text or aggregation in {"AVERAGE", "AVG"}:
                score -= 12
        if wants_average:
            if (
                aggregation in {"AVERAGE", "AVG"}
                or "average" in text
                or "avg" in tokens
            ):
                score += 12
            if aggregation == "SUM" or "sum of" in text:
                score -= 8
        if wants_count:
            if aggregation in {"COUNT", "COUNTA", "DISTINCTCOUNT"} or "count" in tokens:
                score += 10
        if wants_min and (
            aggregation == "MIN" or "minimum" in tokens or "lowest" in tokens
        ):
            score += 10
        if wants_max and (
            aggregation == "MAX" or "maximum" in tokens or "highest" in tokens
        ):
            score += 10

        # Named measures are preferred only when the title does not explicitly
        # contradict their aggregation.
        if field_type == "named_measure":
            score += 2

        scored.append((score, -index, measure))

    scored.sort(reverse=True, key=lambda item: (item[0], item[1]))
    return [scored[0][2]]


def _is_generic_title(title: str) -> bool:
    text = str(title or "").strip().casefold()
    if text in {
        "",
        "none",
        "null",
        "unknown",
        "visual",
        "visual block",
        "chart",
        "map",
        "treemap",
        "slicer",
        "filter",
        "placeholder",
    }:
        return True
    return bool(re.match(r"^visual(?:\s+block)?[_\s-]*[a-z0-9]*$", text))


def _display_name(field: Any) -> str:
    if isinstance(field, dict):
        value = (
            field.get("display_name")
            or field.get("measure_name")
            or field.get("column_name")
            or field.get("canonical_reference")
            or field.get("raw_reference")
            or "Value"
        )
        return str(value)
    normalized = normalize_field(field)
    return str(normalized.get("display_name") or field or "Value")


def _title(
    vc: Dict[str, Any], page_name: str, rows: List[str], measures: List[Any]
) -> str:
    existing = str(vc.get("visual_title") or vc.get("title") or "").strip()
    if existing and not _is_generic_title(existing):
        return existing

    visual_type = str(vc.get("visual_type") or "").casefold()
    metric = _display_name(measures[0]) if measures else "Value"
    dimension = _display_name(rows[0]) if rows else "Category"

    if "slicer" in visual_type or "filter" in visual_type:
        return f"Select {dimension}"
    if any(token in visual_type for token in ("card", "kpi", "gauge")):
        return metric
    if any(
        token in visual_type
        for token in (
            "chart",
            "bar",
            "column",
            "line",
            "area",
            "pie",
            "donut",
            "map",
            "treemap",
        )
    ):
        return f"{metric} by {dimension}" if rows and measures else metric
    if "table" in visual_type or "matrix" in visual_type:
        return f"{dimension} Details"
    return page_name or "Dashboard"


def _layout_to_excel(layout: Dict[str, Any]) -> Dict[str, Any]:
    """Preserve PBIX geometry and keep cell coordinates as fallback only.

    PBIX ``x/y/width/height`` values are the authoritative placement source.
    Approximate row/column values are retained only for cell-based KPI content
    and failure placeholders; they must not override valid PBIX geometry.
    """
    raw = dict(layout or {})
    for key in ("bbox", "position", "geometry", "bounds"):
        nested = raw.get(key)
        if isinstance(nested, dict):
            raw = {**nested, **raw}

    has_pbix_geometry = all(
        raw.get(key) is not None for key in ("x", "y", "width", "height")
    )

    if has_pbix_geometry:
        x = max(0.0, float(raw.get("x") or 0.0))
        y = max(0.0, float(raw.get("y") or 0.0))
        width = max(1.0, float(raw.get("width") or 1.0))
        height = max(1.0, float(raw.get("height") or 1.0))
        canvas_width = max(
            1.0,
            float(
                raw.get("canvas_width")
                or raw.get("page_width")
                or raw.get("report_width")
                or 1280.0
            ),
        )
        canvas_height = max(
            1.0,
            float(
                raw.get("canvas_height")
                or raw.get("page_height")
                or raw.get("report_height")
                or 720.0
            ),
        )

        # Fallback cells are intentionally approximate and are never the
        # authoritative object-placement source.
        row = max(3, int(round(3 + (y / canvas_height) * 40)))
        col = max(2, int(round(2 + (x / canvas_width) * 24)))
        row_span = max(2, int(round((height / canvas_height) * 40)))
        col_span = max(2, int(round((width / canvas_width) * 24)))

        return {
            **raw,
            "layout_source": "pbix",
            "x": x,
            "y": y,
            "width": width,
            "height": height,
            "canvas_width": canvas_width,
            "canvas_height": canvas_height,
            "row": row,
            "col": col,
            "row_span": row_span,
            "col_span": col_span,
        }

    return {
        **raw,
        "layout_source": "excel_cells",
        "row": max(3, int(raw.get("row") or 5)),
        "col": max(2, int(raw.get("col") or 2)),
        "row_span": max(3, int(raw.get("row_span") or 10)),
        "col_span": max(2, int(raw.get("col_span") or 5)),
    }


def _as_list(value: Any) -> List[Any]:
    if value is None:
        return []
    if isinstance(value, list):
        return value
    return [value]


def _collect(hint: Dict[str, Any], vc: Dict[str, Any], *keys: str) -> List[Any]:
    values: List[Any] = []
    for key in keys:
        source = hint.get(key)
        if source in (None, "", []):
            source = vc.get(key)
        values.extend(item for item in _as_list(source) if item not in (None, ""))
    return values


def _append_unique(target: List[Any], value: Any, key: str) -> None:
    seen = {
        str(item.get(key) if isinstance(item, dict) else item).casefold()
        for item in target
    }
    current = str(value.get(key) if isinstance(value, dict) else value).casefold()
    if current and current not in seen:
        target.append(value)


def _normalize_role_fields(
    raw_values: List[Any], role: str
) -> tuple[List[str], List[Dict[str, Any]]]:
    dimensions: List[str] = []
    measures: List[Dict[str, Any]] = []
    for raw in raw_values:
        normalized = normalize_field(raw, role=role)
        field_type = normalized.get("field_type")
        if field_type in {"implicit_measure", "named_measure"}:
            _append_unique(measures, normalized, "canonical_reference")
        else:
            reference = str(
                normalized.get("canonical_reference")
                or normalized.get("raw_reference")
                or raw
            ).strip()
            if reference and reference.casefold() not in {
                v.casefold() for v in dimensions
            }:
                dimensions.append(reference)
    return dimensions, measures


def _looks_like_slicer(
    visual_type: str, title: str, rows: List[str], measures: List[Any]
) -> bool:
    text = f"{visual_type} {title}".casefold()
    return (
        "slicer" in text
        or "filter" in text
        or "dropdown" in text
        or (
            bool(rows)
            and not measures
            and any(
                token in text
                for token in (
                    "year",
                    "month",
                    "date",
                    "region",
                    "zone",
                    "state",
                    "category",
                )
            )
        )
    )


def create_visual_bindings(
    visual_chunks: List[Dict[str, Any]],
    formula_chunks: List[Dict[str, Any]],
    page_name: str,
) -> List[VisualBinding]:
    """Create deterministic bindings while preserving Power BI semantic roles."""
    formulas_by_name = {
        str(item.get("measure_name") or "").strip().casefold(): item
        for item in formula_chunks or []
        if item.get("measure_name")
    }
    bindings: List[VisualBinding] = []

    for vc in visual_chunks or []:
        if str(vc.get("page_name") or "") != str(page_name):
            continue

        hint = vc.get("excel_conversion_hint") or {}
        visual_type = str(vc.get("visual_type") or "unknown").casefold()

        row_dims, row_measures = _normalize_role_fields(
            _collect(hint, vc, "axis", "rows", "uses_columns"), "rows"
        )
        column_dims, column_measures = _normalize_role_fields(
            _collect(hint, vc, "columns"), "columns"
        )
        legend_dims, legend_measures = _normalize_role_fields(
            _collect(hint, vc, "legend"), "legend"
        )
        value_dims, value_measures = _normalize_role_fields(
            _collect(hint, vc, "values", "uses_measures"), "values"
        )

        known_measure_names = set(formulas_by_name.keys())

        measures: List[Any] = []
        for collection in (
            row_measures,
            column_measures,
            legend_measures,
            value_measures,
        ):
            for measure in collection:
                _append_unique(measures, measure, "canonical_reference")

        row_dims, measures = _promote_known_measures(
            row_dims, measures, known_measure_names
        )
        column_dims, measures = _promote_known_measures(
            column_dims, measures, known_measure_names
        )
        legend_dims, measures = _promote_known_measures(
            legend_dims, measures, known_measure_names
        )
        value_dims, measures = _promote_known_measures(
            value_dims, measures, known_measure_names
        )
        measures = _dedupe_measures(
            _repair_measure_descriptors(measures, known_measure_names)
        )

        rows = row_dims + [item for item in value_dims if item not in row_dims]
        columns = column_dims
        legend = legend_dims

        # Enforce Power BI visual-role boundaries. A slicer never owns a measure;
        # a card/gauge owns one effective metric; charts keep only explicit value
        # measures and do not inherit measures accidentally discovered in row roles.
        confirmed_visual_type = str(
            (vc.get("render_operation") or {}).get("visual_type")
            or vc.get("visual_type")
            or ""
        ).casefold()
        if "slicer" in confirmed_visual_type or "filter" in confirmed_visual_type:
            measures = []
            columns = []
            legend = []
            rows = rows[:1]
        elif any(token in confirmed_visual_type for token in ("card", "kpi", "gauge")):
            measures = _select_measure_for_visual(
                measures,
                str(vc.get("visual_title") or vc.get("title") or ""),
                confirmed_visual_type,
            )
            rows = []
            columns = []
            legend = []
        elif any(
            token in confirmed_visual_type
            for token in ("chart", "bar", "column", "line", "area", "pie", "donut")
        ):
            explicit_value_measures = _dedupe_measures(value_measures)
            measures = explicit_value_measures or measures
            # A normal chart with one category should not receive unrelated metrics.
            if "combo" not in confirmed_visual_type and len(measures) > 1:
                measures = _select_measure_for_visual(
                    measures,
                    str(vc.get("visual_title") or vc.get("title") or ""),
                    confirmed_visual_type,
                )

        raw_filters: List[Any] = []
        for key, scope in (
            ("report_filters", "report"),
            ("page_filters", "page"),
            ("visual_filters", "visual"),
            ("filters", "visual"),
        ):
            for item in _as_list(vc.get(key)):
                if isinstance(item, dict):
                    raw_filters.append({**item, "scope": item.get("scope") or scope})
                elif item:
                    raw_filters.append(item)
        raw_filters.extend(_as_list(hint.get("filters")))
        filters = normalize_filters(raw_filters)

        title = _title(vc, page_name, rows, measures)
        # A title such as "Sum of passengers by month" is not a slicer.
        # Only confirmed metadata/render-plan slicer types may create slicers.
        render_operation = dict(vc.get("render_operation") or {})
        confirmed_type = str(
            render_operation.get("visual_type") or vc.get("visual_type") or ""
        ).casefold()
        if "slicer" in confirmed_type:
            binding_type = "slicer"
            slicer_field = (
                rows[0] if rows else (filters[0].get("field") if filters else None)
            )
        elif any(token in visual_type for token in ("card", "kpi", "gauge")):
            binding_type = "cube_formula"
            slicer_field = None
        elif any(
            token in visual_type
            for token in (
                "chart",
                "bar",
                "column",
                "line",
                "area",
                "pie",
                "donut",
                "treemap",
                "map",
                "table",
                "matrix",
                "decompositiontree",
                "keyinfluencers",
                "qna",
                "smartnarrative",
            )
        ):
            binding_type = "connected_pivot"
            slicer_field = None
        else:
            binding_type = "placeholder"
            slicer_field = None

        warnings: List[str] = []
        for measure in measures:
            if measure.get("field_type") != "named_measure":
                continue
            name = str(measure.get("measure_name") or "").strip()
            if name and name.casefold() not in formulas_by_name:
                warnings.append(
                    f"Named measure '{name}' is referenced by the visual but was not found in formula metadata. It will still be validated against CubeFields."
                )

        bindings.append(
            VisualBinding(
                visual_id=str(vc.get("chunk_id") or "visual_unknown"),
                page_name=page_name,
                visual_type=visual_type,
                binding_type=binding_type,
                title=title,
                layout=_layout_to_excel(vc.get("layout") or {}),
                rows=rows,
                columns=columns,
                measures=measures,
                legend=legend,
                filters=filters,
                slicer_field=slicer_field,
                source_status=(
                    "live_connected" if binding_type != "placeholder" else "fallback"
                ),
                field_mapping_status="pending",
                connection_status="pending",
                refresh_status="not_started",
                warnings=warnings,
                render_operation=render_operation,
                render_style=dict(
                    render_operation.get("style")
                    or vc.get("render_style")
                    or vc.get("_visual_theme")
                    or {}
                ),
                settings=dict(
                    vc.get("settings") or render_operation.get("settings") or {}
                ),
                description=str(
                    vc.get("visual_description") or vc.get("description") or ""
                )
                or None,
            )
        )

    return bindings


def create_all_visual_bindings(final_chunks: Dict[str, Any]) -> List[VisualBinding]:
    """Create all visual bindings for all pages in final_chunks."""
    visual_chunks = (
        final_chunks.get("visual_chunks") or final_chunks.get("visuals") or []
    )
    formula_chunks = (
        final_chunks.get("formula_chunks") or final_chunks.get("formulas") or []
    )

    pages = (
        final_chunks.get("pages")
        or final_chunks.get("page_chunks")
        or final_chunks.get("report_pages")
        or []
    )

    page_names = []
    if pages:
        for p in pages:
            name = (
                p.get("display_name")
                or p.get("page_name")
                or p.get("name")
                or "Dashboard"
            )
            if name not in page_names:
                page_names.append(name)
    else:
        seen = set()
        for vc in visual_chunks:
            name = str(vc.get("page_name") or "Dashboard")
            if name.casefold() not in seen:
                seen.add(name.casefold())
                page_names.append(name)

    if not page_names:
        page_names = ["Dashboard"]

    all_bindings = []
    for page_name in page_names:
        page_bindings = create_visual_bindings(visual_chunks, formula_chunks, page_name)
        all_bindings.extend(page_bindings or [])

    return all_bindings


__all__ = ["create_visual_bindings", "create_all_visual_bindings"]

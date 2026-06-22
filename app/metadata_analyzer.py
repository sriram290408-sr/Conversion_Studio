"""
metadata_analyzer.py

Power BI metadata intelligence layer.

This module analyzes processed PBIX chunks and builds a safe metadata summary.
It works even when DataModelSchema, table chunks, relationship chunks, or DAX
formula chunks are unavailable by inferring useful context from visual chunks.
"""

from __future__ import annotations

import logging
import re
from collections import Counter, defaultdict
from typing import Any

logger = logging.getLogger("metadata_analyzer")


FIELD_KEYS = (
    "uses_fields",
    "uses_columns",
    "uses_measures",
    "fields",
    "axis",
    "values",
    "rows",
    "columns",
    "legend",
    "filters",
    "dimension_fields",
    "measure_fields",
)

PROJECTION_KEYS = (
    "axis",
    "values",
    "rows",
    "columns",
    "legend",
    "filters",
    "category",
    "y",
)

BUSINESS_CATEGORIES = (
    "Date / Time",
    "Geography / Region",
    "Product / Brand",
    "Customer / Outlet",
    "Sales / Metrics",
    "Other",
)


def _safe_str(value: Any) -> str:
    """Return a stripped string for any value."""
    return str(value or "").strip()


def _normalize_text(value: Any) -> str:
    """Normalize text for keyword matching."""
    text = _safe_str(value).lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _unique_keep_order(values: list[Any]) -> list[str]:
    """Return unique non-empty string values while preserving original order."""
    seen: set[str] = set()
    output: list[str] = []

    for value in values:
        text = _safe_str(value)
        if not text:
            continue

        key = text.lower()
        if key in seen:
            continue

        seen.add(key)
        output.append(text)

    return output


def clean_field_label(field_ref: Any) -> str:
    """
    Convert a PBIX field reference into a display label.

    Examples:
        'Sales Table'[Amount] -> Amount
        Sales[Amount] -> Amount
        Sales.Amount -> Amount
    """
    field_ref = _safe_str(field_ref).replace('"', "").strip()

    if not field_ref:
        return ""

    quoted_match = re.match(r"^'([^']+)'\[([^\]]+)\]$", field_ref)
    if quoted_match:
        return quoted_match.group(2).strip()

    bracket_match = re.match(r"^(.+?)\[([^\]]+)\]$", field_ref)
    if bracket_match:
        return bracket_match.group(2).strip()

    if "." in field_ref and not field_ref.startswith("http"):
        parts = field_ref.split(".")
        if len(parts) >= 2:
            return ".".join(parts[1:]).strip()

    return field_ref


def parse_table_column_from_field(field_ref: Any) -> dict[str, str]:
    """
    Parse a PBIX field reference into table and column parts.

    Supported formats:
        'Table Name'[Column Name]
        Table Name[Column Name]
        Table.Column
        ColumnOnly
    """
    field_ref = _safe_str(field_ref)

    if not field_ref:
        return {"table": "", "column": "", "raw": ""}

    quoted_match = re.match(r"^'([^']+)'\[([^\]]+)\]$", field_ref)
    if quoted_match:
        return {
            "table": quoted_match.group(1).strip(),
            "column": quoted_match.group(2).strip(),
            "raw": field_ref,
        }

    bracket_match = re.match(r"^(.+?)\[([^\]]+)\]$", field_ref)
    if bracket_match:
        return {
            "table": bracket_match.group(1).strip(),
            "column": bracket_match.group(2).strip(),
            "raw": field_ref,
        }

    if "." in field_ref and not field_ref.startswith("http"):
        parts = field_ref.split(".")
        if len(parts) >= 2:
            return {
                "table": parts[0].strip(),
                "column": ".".join(parts[1:]).strip(),
                "raw": field_ref,
            }

    return {"table": "", "column": field_ref, "raw": field_ref}


def detect_table_role(table_name: str) -> str:
    """Classify a table by likely business role using naming patterns."""
    name = _normalize_text(table_name)

    if any(
        keyword in name for keyword in ("temp", "tmp", "temporary", "staging", "stage")
    ):
        return "Possible Temporary Table"

    if any(
        keyword in name
        for keyword in (
            "join",
            "joined",
            "merge",
            "merged",
            "secondary",
            "combined",
            "union",
            "append",
        )
    ):
        return "Possible Intermediate / Joined Table"

    if any(keyword in name for keyword in ("date", "calendar", "time")):
        return "Date Dimension Table"

    if any(
        keyword in name
        for keyword in ("product", "sku", "brand", "flavour", "flavor", "item", "pack")
    ):
        return "Product Dimension Table"

    if any(
        keyword in name
        for keyword in ("customer", "outlet", "dealer", "retailer", "channel")
    ):
        return "Customer / Outlet Dimension Table"

    if any(
        keyword in name
        for keyword in (
            "geo",
            "geography",
            "state",
            "region",
            "location",
            "city",
            "zone",
            "country",
        )
    ):
        return "Geography Dimension Table"

    if any(
        keyword in name
        for keyword in (
            "sales",
            "transaction",
            "fact",
            "volume",
            "value",
            "plan",
            "target",
            "actual",
        )
    ):
        return "Fact / Transaction Table"

    return "Business Model Table"


def categorize_field(field_name: str) -> str:
    """Map a field or measure name to a business category."""
    name = _normalize_text(field_name)

    if any(
        keyword in name
        for keyword in (
            "date",
            "year",
            "month",
            "quarter",
            "day",
            "week",
            "time",
            "period",
            "mtd",
            "qtd",
            "ytd",
            "lysm",
            "lysq",
        )
    ):
        return "Date / Time"

    if any(
        keyword in name
        for keyword in (
            "state",
            "city",
            "zone",
            "region",
            "country",
            "latitude",
            "longitude",
            "geo",
            "geography",
            "market type",
            "location",
            "territory",
        )
    ):
        return "Geography / Region"

    if any(
        keyword in name
        for keyword in (
            "brand",
            "brandfamily",
            "product",
            "sku",
            "flavour",
            "flavor",
            "pack",
            "category",
            "item",
            "segment",
        )
    ):
        return "Product / Brand"

    if any(
        keyword in name
        for keyword in (
            "customer",
            "outlet",
            "dealer",
            "retailer",
            "channel",
            "market",
            "consumer",
        )
    ):
        return "Customer / Outlet"

    if any(
        keyword in name
        for keyword in (
            "sales",
            "volume",
            "amount",
            "value",
            "target",
            "plan",
            "revenue",
            "profit",
            "growth",
            "achievement",
            "ach",
            "actual",
            "mtd",
            "qtd",
            "ytd",
            "lys",
            "vs",
        )
    ):
        return "Sales / Metrics"

    return "Other"


def detect_visual_category(visual_type: str) -> str:
    """Group a Power BI visual type into a dashboard category."""
    visual_type = _normalize_text(visual_type)

    if any(keyword in visual_type for keyword in ("slicer", "filter", "dropdown")):
        return "Slicers & Filters"

    if any(keyword in visual_type for keyword in ("card", "kpi", "gauge", "indicator")):
        return "Cards & KPIs"

    if any(keyword in visual_type for keyword in ("table", "matrix")):
        return "Tables & Matrices"

    if any(keyword in visual_type for keyword in ("map", "azure", "arcgis", "geo")):
        return "Maps"

    if any(
        keyword in visual_type
        for keyword in (
            "chart",
            "bar",
            "column",
            "line",
            "area",
            "pie",
            "donut",
            "treemap",
        )
    ):
        return "Charts"

    if any(keyword in visual_type for keyword in ("image", "logo")):
        return "Images"

    if any(keyword in visual_type for keyword in ("navigator", "button")):
        return "Navigation / Buttons"

    return "Other"


def _extract_field_strings_from_value(value: Any) -> list[str]:
    """Recursively collect possible field names from PBIX metadata values."""
    fields: list[str] = []

    if value is None:
        return fields

    if isinstance(value, str):
        text = value.strip()
        if text:
            fields.append(text)
        return fields

    if isinstance(value, list):
        for item in value:
            fields.extend(_extract_field_strings_from_value(item))
        return fields

    if isinstance(value, dict):
        candidate_keys = (
            "field",
            "name",
            "queryRef",
            "QueryRef",
            "displayName",
            "DisplayName",
            "measure",
            "column",
            "expr",
        )

        nested_keys = (
            "source",
            "expression",
            "select",
            "data",
            "metadata",
        )

        for key in candidate_keys:
            if value.get(key):
                fields.extend(_extract_field_strings_from_value(value[key]))

        for key in nested_keys:
            if key in value:
                fields.extend(_extract_field_strings_from_value(value[key]))

    return fields


def extract_visual_fields(visual: dict[str, Any]) -> list[str]:
    """Extract field and measure references from a visual chunk."""
    visual = visual or {}
    raw_fields: list[str] = []

    for key in FIELD_KEYS:
        raw_fields.extend(_extract_field_strings_from_value(visual.get(key)))

    hint = visual.get("excel_conversion_hint") or {}
    if isinstance(hint, dict):
        for key in PROJECTION_KEYS:
            raw_fields.extend(_extract_field_strings_from_value(hint.get(key)))

    config = visual.get("config") or {}
    if isinstance(config, dict):
        projection = (
            config.get("projection")
            or config.get("projections")
            or config.get("Projections")
            or {}
        )
        if isinstance(projection, dict):
            for key in PROJECTION_KEYS:
                raw_fields.extend(
                    _extract_field_strings_from_value(projection.get(key))
                )

        single_visual = config.get("singleVisual") or config.get("SingleVisual") or {}
        if isinstance(single_visual, dict):
            projections = (
                single_visual.get("projections")
                or single_visual.get("Projections")
                or {}
            )
            if isinstance(projections, dict):
                for value in projections.values():
                    raw_fields.extend(_extract_field_strings_from_value(value))

    ai_analysis = visual.get("ai_analysis") or visual.get("deep_analysis") or {}
    if isinstance(ai_analysis, dict):
        for key in (
            "fields",
            "measures",
            "dimensions",
            "dimension_fields",
            "measure_fields",
        ):
            raw_fields.extend(_extract_field_strings_from_value(ai_analysis.get(key)))

    return _unique_keep_order(raw_fields)


def split_dimensions_and_measures(fields: list[str]) -> dict[str, list[str]]:
    """Split fields into dimensions and measures using business keywords."""
    dimensions: list[str] = []
    measures: list[str] = []

    for field in fields:
        label = clean_field_label(field)
        if not label:
            continue

        category = categorize_field(label)
        normalized = _normalize_text(label)
        is_metric = category == "Sales / Metrics" or any(
            keyword in normalized
            for keyword in (
                "sum",
                "count",
                "average",
                "avg",
                "value",
                "volume",
                "sales",
                "target",
                "plan",
                "amount",
                "revenue",
                "profit",
                "mtd",
                "qtd",
                "ytd",
            )
        )

        if is_metric:
            measures.append(label)
        else:
            dimensions.append(label)

    return {
        "dimensions": _unique_keep_order(dimensions),
        "measures": _unique_keep_order(measures),
    }


def _infer_page_purpose(
    page_name: str, field_categories: Counter, visual_types: set[str]
) -> str:
    """Create a short business description for a report page."""
    page_name_norm = _normalize_text(page_name)
    parts: list[str] = []

    if field_categories.get("Sales / Metrics", 0):
        parts.append("sales and performance metrics")

    if field_categories.get("Geography / Region", 0):
        parts.append("regional breakdown")

    if field_categories.get("Product / Brand", 0):
        parts.append("product and brand analysis")

    if field_categories.get("Customer / Outlet", 0):
        parts.append("outlet and customer segmentation")

    if field_categories.get("Date / Time", 0):
        parts.append("time-period comparisons")

    if not parts:
        if "sales" in page_name_norm:
            parts.append("sales performance")
        elif any(
            keyword in page_name_norm
            for keyword in ("mtd", "qtd", "ytd", "lysm", "lysq")
        ):
            parts.append("time-period performance comparison")
        elif any(
            keyword in page_name_norm for keyword in ("brand", "flavour", "region")
        ):
            parts.append("brand, flavour, and regional analysis")

    normalized_visual_types = {_normalize_text(item) for item in visual_types}
    visual_descriptions: list[str] = []

    if any(
        keyword in visual_type
        for visual_type in normalized_visual_types
        for keyword in ("chart", "bar", "column", "line", "area", "pie", "treemap")
    ):
        visual_descriptions.append("charts")

    if any(
        "table" in visual_type or "matrix" in visual_type
        for visual_type in normalized_visual_types
    ):
        visual_descriptions.append("tables")

    if any(
        "map" in visual_type or "geo" in visual_type
        for visual_type in normalized_visual_types
    ):
        visual_descriptions.append("maps")

    if any(
        "slicer" in visual_type or "filter" in visual_type
        for visual_type in normalized_visual_types
    ):
        visual_descriptions.append("filters")

    if parts and visual_descriptions:
        return f"This page shows {', '.join(parts)} using {', '.join(visual_descriptions)}."

    if parts:
        return f"This page covers {', '.join(parts)}."

    if visual_descriptions:
        return f"This page contains {', '.join(visual_descriptions)}."

    return "Dashboard page with mixed business visuals."


def _collect_ai_insights(final_chunks: dict[str, Any]) -> dict[str, Any]:
    """Collect Hugging Face or deep-analysis chunk summary if available."""
    ai_chunks = (
        final_chunks.get("ai_insight_chunks")
        or final_chunks.get("deep_analysis_chunks")
        or []
    )
    page_chunks = final_chunks.get("page_chunks") or []

    business_titles: list[str] = []
    roles: Counter = Counter()
    render_types: Counter = Counter()

    for item in ai_chunks:
        if not isinstance(item, dict):
            continue

        title = item.get("business_title") or item.get("ai_title") or item.get("title")
        if title:
            business_titles.append(title)

        role = item.get("business_role") or item.get("role")
        if role:
            roles[role] += 1

        render_type = item.get("excel_render_type") or item.get("render_type")
        if render_type:
            render_types[render_type] += 1

    return {
        "ai_insight_count": len(ai_chunks),
        "page_chunk_count": len(page_chunks),
        "business_titles": _unique_keep_order(business_titles)[:25],
        "business_role_count": dict(roles),
        "excel_render_type_count": dict(render_types),
        "deep_analysis_available": bool(ai_chunks or page_chunks),
    }


def _build_table_role_analysis(
    all_tables: list[str],
    model_tables: list[str],
    table_usage_count: Counter,
) -> list[dict[str, Any]]:
    """Build table role records for the final report."""
    model_table_set = set(model_tables)

    return [
        {
            "table_name": table_name,
            "role": detect_table_role(table_name),
            "usage_count": table_usage_count.get(table_name, 0),
            "source": (
                "model_schema" if table_name in model_table_set else "visual_inferred"
            ),
            "confidence": "name_based_inference",
        }
        for table_name in all_tables
    ]


def build_metadata_analysis(final_chunks: dict[str, Any]) -> dict[str, Any]:
    """
    Build a complete metadata intelligence report from processed PBIX chunks.
    """
    final_chunks = final_chunks or {}

    visual_chunks = final_chunks.get("visual_chunks", []) or []
    table_chunks = final_chunks.get("table_chunks", []) or []
    relationship_chunks = final_chunks.get("relationship_chunks", []) or []
    formula_chunks = final_chunks.get("formula_chunks", []) or []

    unique_pages: set[str] = set()
    unique_tables_used: set[str] = set()
    unique_fields_used: set[str] = set()
    unique_measures_used: set[str] = set()

    visual_type_count: Counter = Counter()
    visual_category_count: Counter = Counter()
    table_usage_count: Counter = Counter()
    field_usage_count: Counter = Counter()
    measure_usage_count: Counter = Counter()
    field_category_count: Counter = Counter()

    business_category_fields: dict[str, list[str]] = {
        category: [] for category in BUSINESS_CATEGORIES
    }

    page_details = defaultdict(
        lambda: {
            "visual_count": 0,
            "visual_types": set(),
            "visual_categories": Counter(),
            "tables": set(),
            "fields": [],
            "measures": [],
            "field_cats": Counter(),
        }
    )

    for visual in visual_chunks:
        if not isinstance(visual, dict):
            continue

        page_name = (
            visual.get("page_name")
            or visual.get("page")
            or visual.get("report_page")
            or visual.get("target_sheet")
            or "Unknown Page"
        )

        visual_type = (
            visual.get("visual_type")
            or visual.get("type")
            or visual.get("visualType")
            or "unknown"
        )

        unique_pages.add(page_name)
        visual_type_count[visual_type] += 1

        visual_category = detect_visual_category(visual_type)
        visual_category_count[visual_category] += 1

        page_details[page_name]["visual_count"] += 1
        page_details[page_name]["visual_types"].add(visual_type)
        page_details[page_name]["visual_categories"][visual_category] += 1

        raw_fields = extract_visual_fields(visual)
        split_fields = split_dimensions_and_measures(raw_fields)

        explicit_columns = _extract_field_strings_from_value(visual.get("uses_columns"))
        explicit_measures = _extract_field_strings_from_value(
            visual.get("uses_measures")
        )
        raw_field_pool = _unique_keep_order(
            raw_fields + explicit_columns + explicit_measures
        )

        for raw_field in raw_field_pool:
            parsed = parse_table_column_from_field(raw_field)
            table_name = parsed.get("table", "")
            column_name = parsed.get("column", "")

            if not column_name:
                continue

            display_key = f"{table_name}[{column_name}]" if table_name else column_name
            unique_fields_used.add(display_key)
            field_usage_count[display_key] += 1

            if table_name:
                unique_tables_used.add(table_name)
                table_usage_count[table_name] += 1
                page_details[page_name]["tables"].add(table_name)

            category = categorize_field(column_name)
            field_category_count[category] += 1
            page_details[page_name]["field_cats"][category] += 1
            page_details[page_name]["fields"].append(column_name)

            if column_name not in business_category_fields[category]:
                business_category_fields[category].append(column_name)

        for measure in split_fields["measures"] + explicit_measures:
            measure_name = clean_field_label(measure)
            if not measure_name:
                continue

            unique_measures_used.add(measure_name)
            measure_usage_count[measure_name] += 1
            page_details[page_name]["measures"].append(measure_name)

    model_tables: list[str] = []

    for table in table_chunks:
        if not isinstance(table, dict):
            continue

        table_name = (
            table.get("table_name") or table.get("name") or table.get("displayName")
        )
        if table_name:
            model_tables.append(table_name)

        for column in table.get("columns") or []:
            column_name = _safe_str(column)
            if not column_name:
                continue

            category = categorize_field(column_name)
            if column_name not in business_category_fields[category]:
                business_category_fields[category].append(column_name)

    model_tables = _unique_keep_order(model_tables)

    for formula in formula_chunks:
        if not isinstance(formula, dict):
            continue

        measure_name = (
            formula.get("measure_name") or formula.get("name") or formula.get("title")
        )
        if measure_name:
            unique_measures_used.add(measure_name)
            measure_usage_count[measure_name] += 1

        for table_name in formula.get("used_tables", []) or []:
            table_name = _safe_str(table_name)
            if not table_name:
                continue

            unique_tables_used.add(table_name)
            table_usage_count[table_name] += 1

        for column in formula.get("used_columns", []) or []:
            parsed = parse_table_column_from_field(column)
            table_name = parsed.get("table", "")
            column_name = parsed.get("column", "")

            if not column_name:
                continue

            display_key = f"{table_name}[{column_name}]" if table_name else column_name
            unique_fields_used.add(display_key)
            field_usage_count[display_key] += 1

            category = categorize_field(column_name)
            field_category_count[category] += 1
            if column_name not in business_category_fields[category]:
                business_category_fields[category].append(column_name)

    all_detected_tables = sorted(set(model_tables) | unique_tables_used)
    table_role_analysis = _build_table_role_analysis(
        all_detected_tables, model_tables, table_usage_count
    )

    possible_intermediate_tables = [
        item["table_name"]
        for item in table_role_analysis
        if item["role"] == "Possible Intermediate / Joined Table"
    ]

    possible_temporary_tables = [
        item["table_name"]
        for item in table_role_analysis
        if item["role"] == "Possible Temporary Table"
    ]

    page_analysis: list[dict[str, Any]] = []
    for page_name, details in page_details.items():
        page_analysis.append(
            {
                "page_name": page_name,
                "visual_count": details["visual_count"],
                "visual_types": sorted(details["visual_types"]),
                "visual_categories": dict(details["visual_categories"]),
                "tables_used": sorted(details["tables"]),
                "fields_used": _unique_keep_order(details["fields"])[:20],
                "measures_used": _unique_keep_order(details["measures"])[:20],
                "field_category_count": dict(details["field_cats"]),
                "purpose": _infer_page_purpose(
                    page_name, details["field_cats"], details["visual_types"]
                ),
            }
        )

    ai_summary = _collect_ai_insights(final_chunks)

    # V16 Formulate Complexity & Binding Analysis
    classified_measures = []
    try:
        try:
            from .formula_classifier import classify_dax_measure
        except ImportError:
            from formula_classifier import classify_dax_measure
        for f in formula_chunks:
            m_name = f.get("measure_name") or f.get("name") or ""
            dax_expr = f.get("dax_formula") or f.get("expression") or ""
            if m_name:
                classified_measures.append(classify_dax_measure(m_name, dax_expr))
    except Exception as class_err:
        logger.warning(f"Failed to run measure classification in analysis: {class_err}")

    return {
        "overall_counts": {
            "total_pages": len(unique_pages),
            "total_visuals": len(visual_chunks),
            "total_model_tables": len(model_tables),
            "visual_inferred_tables": len(unique_tables_used - set(model_tables)),
            "tables_used_in_visuals": len(unique_tables_used),
            "unique_fields_used": len(unique_fields_used),
            "unique_measures_used": len(unique_measures_used),
            "total_relationships": len(relationship_chunks),
            "total_formulas": len(formula_chunks),
            "ai_insight_chunks": ai_summary.get("ai_insight_count", 0),
            "page_chunks": ai_summary.get("page_chunk_count", 0),
        },
        "model_tables": model_tables,
        "visual_inferred_tables": sorted(unique_tables_used - set(model_tables)),
        "all_detected_tables": all_detected_tables,
        "tables_used_in_visuals": dict(table_usage_count),
        "field_usage": dict(field_usage_count),
        "measure_usage": dict(measure_usage_count),
        "field_category_count": dict(field_category_count),
        "business_category_fields": {
            key: values[:25] for key, values in business_category_fields.items()
        },
        "table_role_analysis": table_role_analysis,
        "possible_intermediate_tables": possible_intermediate_tables,
        "possible_temporary_tables": possible_temporary_tables,
        "visual_type_count": dict(visual_type_count),
        "visual_category_count": dict(visual_category_count),
        "top_used_fields": [
            {"field": key, "count": value}
            for key, value in field_usage_count.most_common(25)
        ],
        "top_used_measures": [
            {"measure": key, "count": value}
            for key, value in measure_usage_count.most_common(25)
        ],
        "page_analysis": sorted(page_analysis, key=lambda item: item["page_name"]),
        "ai_deep_analysis": ai_summary,
        "classified_measures": classified_measures,
    }


# =============================================================================
# V15 optional helper: deep extraction summary for page/visual/filter/calculated
# columns. Convertor v15 includes its own integration layer, so this helper is
# optional for preview/intelligence use.
# =============================================================================


def build_deep_metadata_summary(final_chunks: dict[str, Any]) -> dict[str, Any]:
    final_chunks = final_chunks or {}
    visuals = final_chunks.get("visual_chunks", []) or []
    formulas = final_chunks.get("formula_chunks", []) or []
    relationships = final_chunks.get("relationship_chunks", []) or []
    tables = final_chunks.get("table_chunks", []) or []

    page_summary: dict[str, dict[str, Any]] = {}
    visual_filter_records = []

    for idx, visual in enumerate(visuals, 1):
        if not isinstance(visual, dict):
            continue
        page = visual.get("page_name") or visual.get("page") or "Unknown Page"
        title = visual.get("visual_title") or visual.get("title") or f"Visual {idx}"
        fields = extract_visual_fields(visual)
        filters = []
        hint = visual.get("excel_conversion_hint") or {}
        if isinstance(hint, dict):
            filters.extend(_extract_field_strings_from_value(hint.get("filters")))
        filters.extend(_extract_field_strings_from_value(visual.get("filters")))

        page_summary.setdefault(
            page,
            {
                "page_name": page,
                "visual_count": 0,
                "visual_titles": [],
                "fields": [],
                "filters": [],
            },
        )
        page_summary[page]["visual_count"] += 1
        page_summary[page]["visual_titles"].append(title)
        page_summary[page]["fields"].extend(fields)
        page_summary[page]["filters"].extend(filters)

        visual_filter_records.append(
            {
                "page_name": page,
                "visual_title": title,
                "visual_type": visual.get("visual_type") or visual.get("type") or "",
                "fields": _unique_keep_order(fields),
                "filters": _unique_keep_order(filters),
            }
        )

    for rec in page_summary.values():
        rec["visual_titles"] = _unique_keep_order(rec["visual_titles"])
        rec["fields"] = _unique_keep_order(rec["fields"])
        rec["filters"] = _unique_keep_order(rec["filters"])

    formula_dependency_records = []
    for f in formulas:
        if not isinstance(f, dict):
            continue
        formula_dependency_records.append(
            {
                "measure_name": f.get("measure_name") or f.get("name") or "",
                "dax_formula": f.get("dax_formula") or f.get("expression") or "",
                "used_tables": f.get("used_tables") or [],
                "used_columns": f.get("used_columns") or [],
                "mapped_table_chunks": f.get("mapped_table_chunks") or [],
            }
        )

    return {
        "page_summary": list(page_summary.values()),
        "visual_filter_records": visual_filter_records,
        "formula_dependency_records": formula_dependency_records,
        "relationship_count": len(relationships),
        "table_count": len(tables),
        "visual_count": len(visuals),
        "formula_count": len(formulas),
    }

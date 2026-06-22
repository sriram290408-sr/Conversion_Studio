"""Normalize PBIX/TMDL field references into deterministic semantic descriptors.

``parse_field_reference`` is the canonical entry-point for any code that needs to
decompose a raw PBIX, TMDL, or OLAP field string into its constituent parts.
``normalize_field`` is the higher-level wrapper that also classifies the field
by its semantic role (measure vs dimension) and builds the canonical reference
strings expected by the downstream OLAP mapper.

Supported input formats
-----------------------
* ``Table[Column]``                      – PBIX power query notation
* ``'Table Name'[Column]``               – quoted table name
* ``Table.Column``                       – dot-separated
* ``[Measures].[Measure Name]``          – full OLAP measure path
* ``[Measure Name]``                     – bare bracketed measure name
* ``SUM(Table.Column)``                  – implicit aggregation
* ``SUM('Table Name'[Column])``          – implicit aggregation with quoted table
* ``[Table].[Column]``                   – two-part OLAP column path
* ``[Table].[Column].[Column]``          – three-part OLAP attribute path
* ``[Table].[Hierarchy].[Level]``        – OLAP hierarchy level path
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional

_AGGREGATIONS = {
    "SUM": "SUM",
    "AVERAGE": "AVERAGE",
    "AVG": "AVERAGE",
    "MIN": "MIN",
    "MAX": "MAX",
    "COUNT": "COUNT",
    "COUNTA": "COUNTA",
    "COUNTNONNULL": "COUNTNONNULL",
    "DISTINCTCOUNT": "DISTINCTCOUNT",
}


def _clean(value: Any) -> str:
    text = str(value or "").strip()
    text = re.sub(r"[\s\)\}\]]+$", "", text)
    return text.strip()


def _quote_table(table: str) -> str:
    table = table.strip().strip("'")
    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", table):
        return table
    return "'" + table.replace("'", "''") + "'"


def _canonical_column(table: str, column: str) -> str:
    return f"{_quote_table(table)}[{column.strip()}]"


# ---------------------------------------------------------------------------
# Canonical field-reference parser
# ---------------------------------------------------------------------------


def parse_field_reference(raw: Any) -> Dict[str, Any]:
    """Decompose any supported field-reference string into its parts.

    Returns a dict with keys:
    - ``input_format``  – one of: ``olap_measure``, ``olap_column_2part``,
      ``olap_column_3part``, ``pbix_table_column``, ``pbix_dot``,
      ``aggregation``, ``bracketed_measure``, ``bare``
    - ``table_name``    – table or dimension name (may be empty)
    - ``column_name``   – column or level name (may be empty)
    - ``measure_name``  – named measure (may be empty)
    - ``aggregation``   – normalized aggregation token (may be empty)
    - ``hierarchy``     – intermediate path parts for hierarchy formats
    """
    text = str(raw or "").strip()
    result: Dict[str, Any] = {
        "input_format": "bare",
        "table_name": "",
        "column_name": "",
        "measure_name": "",
        "aggregation": "",
        "hierarchy": [],
    }

    if not text:
        return result

    # ------------------------------------------------------------------
    # 1. Implicit aggregation: SUM(Table[Column]) / SUM(Table.Column)
    # ------------------------------------------------------------------
    agg_match = re.match(
        r"(?is)^\s*(sum|average|avg|min|max|count|counta|countnonnull|distinctcount)\s*\((.+)\)\s*$",
        text,
    )
    if agg_match:
        agg = _AGGREGATIONS[agg_match.group(1).upper()]
        inner = agg_match.group(2).strip()
        result["aggregation"] = agg
        result["input_format"] = "aggregation"
        # Recurse on the inner part to extract table/column.
        inner_parsed = parse_field_reference(inner)
        result["table_name"] = inner_parsed["table_name"]
        result["column_name"] = inner_parsed["column_name"]
        return result

    # ------------------------------------------------------------------
    # 2. Full OLAP measure path: [Measures].[Name]
    # ------------------------------------------------------------------
    olap_measure = re.match(r"(?i)^\[measures\]\.\[([^\]]+)\]$", text)
    if olap_measure:
        result["input_format"] = "olap_measure"
        result["measure_name"] = olap_measure.group(1).strip()
        return result

    # ------------------------------------------------------------------
    # 3. OLAP paths – extract all bracket tokens
    # ------------------------------------------------------------------
    bracket_parts = re.findall(r"\[([^\]]+)\]", text)

    if bracket_parts:
        # 3a. Single bracket: [Name] – bare named measure
        if len(bracket_parts) == 1 and text == f"[{bracket_parts[0]}]":
            result["input_format"] = "bracketed_measure"
            result["measure_name"] = bracket_parts[0].strip()
            return result

        # 3b. Two-part OLAP column: [Table].[Column]
        if len(bracket_parts) == 2 and re.fullmatch(r"\[[^\]]+\]\.\[[^\]]+\]", text):
            result["input_format"] = "olap_column_2part"
            result["table_name"] = bracket_parts[0].strip()
            result["column_name"] = bracket_parts[1].strip()
            return result

        # 3c. Three-part OLAP path: [Table].[Hierarchy/Column].[Level/Column]
        if len(bracket_parts) >= 3:
            result["input_format"] = "olap_column_3part"
            result["table_name"] = bracket_parts[0].strip()
            result["column_name"] = bracket_parts[-1].strip()
            result["hierarchy"] = [p.strip() for p in bracket_parts[1:-1]]
            return result

        # 3d. PBIX Table[Column] (single bracket not wrapping the whole string)
        pbix_bracket = re.match(r"^'?([^'\[]+)'?\s*\[([^\]]+)\]\s*$", text)
        if pbix_bracket:
            result["input_format"] = "pbix_table_column"
            result["table_name"] = pbix_bracket.group(1).strip().strip("'")
            result["column_name"] = pbix_bracket.group(2).strip()
            return result

    # ------------------------------------------------------------------
    # 4. PBIX Table[Column] (no spaces before bracket)
    # ------------------------------------------------------------------
    pbix_bracket = re.match(r"^'?([^'\[]+)'?\s*\[([^\]]+)\]\s*$", text)
    if pbix_bracket:
        result["input_format"] = "pbix_table_column"
        result["table_name"] = pbix_bracket.group(1).strip().strip("'")
        result["column_name"] = pbix_bracket.group(2).strip()
        return result

    # ------------------------------------------------------------------
    # 5. Dot-separated: Table.Column (exactly two parts)
    # ------------------------------------------------------------------
    if "." in text and not text.startswith("["):
        parts = [p.strip() for p in text.rsplit(".", 1)]
        if len(parts) == 2 and parts[0] and parts[1]:
            result["input_format"] = "pbix_dot"
            result["table_name"] = parts[0].strip("'")
            result["column_name"] = _clean(parts[1])
            return result

    # ------------------------------------------------------------------
    # 6. Bare token – could be a measure name or unknown dimension
    # ------------------------------------------------------------------
    result["input_format"] = "bare"
    return result


# ---------------------------------------------------------------------------
# Higher-level normalization (public API used by downstream modules)
# ---------------------------------------------------------------------------


def _parse_aggregation(text: str) -> Optional[Dict[str, str]]:
    match = re.match(
        r"(?is)^\s*(sum|average|avg|min|max|count|counta|countnonnull|distinctcount)\s*\((.*)?\)\s*$",
        text,
    )
    if not match:
        return None

    aggregation = _AGGREGATIONS[match.group(1).upper()]
    inner = match.group(2).strip()

    bracket = re.match(r"^\s*'?([^'\[]+)'?\s*\[([^\]]+)\]\s*$", inner)
    if bracket:
        return {
            "aggregation": aggregation,
            "table_name": bracket.group(1).strip(),
            "column_name": bracket.group(2).strip(),
        }

    if "." in inner:
        table, column = inner.rsplit(".", 1)
        return {
            "aggregation": aggregation,
            "table_name": table.strip().strip("'"),
            "column_name": _clean(column),
        }

    return None


def normalize_field(value: Any, role: Optional[str] = None) -> Dict[str, Any]:
    """Normalize one field reference without inventing semantic-model measures."""
    if isinstance(value, dict):
        declared_type = str(value.get("field_type") or "").casefold()
        measure_name = str(
            value.get("measure_name")
            or value.get("name")
            or value.get("display_name")
            or ""
        ).strip()
        table_name = str(value.get("table_name") or value.get("table") or "").strip()
        column_name = str(
            value.get("column_name")
            or value.get("column")
            or value.get("field_name")
            or ""
        ).strip()
        aggregation = str(value.get("aggregation") or "").strip().upper()
        raw_value = (
            value.get("raw_reference")
            or value.get("canonical_reference")
            or value.get("field")
            or ""
        )

        if declared_type in {"named_measure", "measure"} and measure_name:
            raw = f"[Measures].[{measure_name}]"
        elif aggregation and (table_name or column_name):
            inner = (
                _canonical_column(table_name, column_name)
                if table_name and column_name
                else column_name
            )
            raw = f"{aggregation}({inner})"
        elif table_name and column_name:
            raw = _canonical_column(table_name, column_name)
        elif measure_name and str(role or "").casefold() in {
            "values",
            "value",
            "measure",
            "measures",
            "y",
        }:
            # Do not invent a named measure from a display caption. A real
            # named measure must be explicitly declared through field_type or
            # an OLAP [Measures].[Name] reference. Otherwise keep the value as
            # a column/implicit-measure candidate for the binding engine.
            raw = str(raw_value or measure_name).strip()
        else:
            raw = str(raw_value or measure_name or column_name or "").strip()
    else:
        raw = str(value or "").strip()
    result: Dict[str, Any] = {
        "raw_reference": raw,
        "field_type": "unknown",
        "role": role,
        "table_name": "",
        "column_name": "",
        "measure_name": "",
        "aggregation": "",
        "canonical_reference": raw,
        "cube_measure_path": "",
        "display_name": raw,
    }
    if not raw:
        return result

    # Human-readable Power BI aggregation captions, for example
    # "Sum of passengers" or "Average of flights[passengers]".
    caption_match = re.fullmatch(
        r"(?i)\s*(sum|average|avg|min|max|count|counta|distinctcount|distinct count)"
        r"\s+of\s+(.+?)\s*",
        raw,
    )
    if caption_match:
        aggregation = _AGGREGATIONS.get(
            caption_match.group(1).upper().replace(" ", ""),
            caption_match.group(1).upper().replace(" ", ""),
        )
        inner = parse_field_reference(caption_match.group(2).strip())
        table = inner.get("table_name", "")
        column = inner.get("column_name", "") or caption_match.group(2).strip()
        canonical = _canonical_column(table, column) if table else column
        result.update(
            {
                "raw_reference": f"{aggregation}({canonical})",
                "field_type": "implicit_measure",
                "aggregation": aggregation,
                "table_name": table,
                "column_name": column,
                "measure_name": "",
                "canonical_reference": canonical,
                "cube_measure_path": "",
                "display_name": f"{aggregation.title()} of {column}",
            }
        )
        return result

    # ---- Use canonical parser ----
    parsed = parse_field_reference(raw)
    fmt = parsed["input_format"]
    table = parsed["table_name"]
    column = parsed["column_name"]
    measure_name = parsed["measure_name"]
    aggregation = parsed["aggregation"]

    # Implicit aggregation (SUM/AVG/etc.)
    if fmt == "aggregation" and (table or column):
        result.update(
            {
                "field_type": "implicit_measure",
                "aggregation": aggregation,
                "table_name": table,
                "column_name": column,
                "canonical_reference": (
                    _canonical_column(table, column) if table else column
                ),
                "display_name": f"{aggregation.title()} of {column}",
            }
        )
        return result

    # Full OLAP measure path [Measures].[Name]
    if fmt == "olap_measure":
        result.update(
            {
                "field_type": "named_measure",
                "measure_name": measure_name,
                "canonical_reference": measure_name,
                "cube_measure_path": f"[Measures].[{measure_name}]",
                "display_name": measure_name,
            }
        )
        return result

    # Bracketed measure name [Name]
    if fmt == "bracketed_measure":
        result.update(
            {
                "field_type": "named_measure",
                "measure_name": measure_name,
                "canonical_reference": measure_name,
                "cube_measure_path": f"[Measures].[{measure_name}]",
                "display_name": measure_name,
            }
        )
        return result

    # OLAP 2-part column [Table].[Column]
    if fmt == "olap_column_2part":
        result.update(
            {
                "field_type": "dimension",
                "table_name": table,
                "column_name": column,
                "canonical_reference": _canonical_column(table, column),
                "display_name": column,
            }
        )
        return result

    # OLAP 3-part (hierarchy/level) [Table].[Hierarchy].[Level]
    if fmt == "olap_column_3part":
        hierarchy = parsed.get("hierarchy", [])
        result.update(
            {
                "field_type": "hierarchy" if hierarchy else "dimension",
                "table_name": table,
                "column_name": column,
                "hierarchy_path": hierarchy,
                "canonical_reference": _canonical_column(table, column),
                "display_name": column,
            }
        )
        return result

    # PBIX Table[Column] or Table.Column
    if fmt in ("pbix_table_column", "pbix_dot"):
        result.update(
            {
                "field_type": "dimension",
                "table_name": table,
                "column_name": column,
                "canonical_reference": _canonical_column(table, column),
                "display_name": column,
            }
        )
        return result

    # ---- Bare token – use role as authoritative hint ----
    role_text = str(role or "").casefold()
    if role_text in {"values", "value", "measure", "measures", "y"}:
        name = raw.strip("[]'")
        result.update(
            {
                "field_type": "named_measure",
                "measure_name": name,
                "canonical_reference": name,
                "cube_measure_path": f"[Measures].[{name}]",
                "display_name": name,
            }
        )
    elif role_text in {
        "axis",
        "category",
        "rows",
        "columns",
        "legend",
        "filters",
        "slicer",
        "tooltips",
    }:
        result.update(
            {
                "field_type": "dimension",
                "column_name": _clean(raw),
                "canonical_reference": _clean(raw),
                "display_name": _clean(raw),
            }
        )
    else:
        result["display_name"] = _clean(raw)
        result["canonical_reference"] = _clean(raw)

    return result


def semantic_comparison_key(value: Any, role: Optional[str] = None) -> str:
    """Return a stable comparison key for PBIX, TMDL, and OLAP references."""
    normalized = normalize_field(value, role=role)
    field_type = str(normalized.get("field_type") or "").casefold()

    if field_type == "named_measure":
        name = normalized.get("measure_name") or normalized.get("display_name")
        return re.sub(r"[^a-z0-9]", "", str(name or "").casefold())

    table = re.sub(
        r"[^a-z0-9]",
        "",
        str(normalized.get("table_name") or "").casefold(),
    )
    column = re.sub(
        r"[^a-z0-9]",
        "",
        str(
            normalized.get("column_name") or normalized.get("display_name") or ""
        ).casefold(),
    )
    aggregation = re.sub(
        r"[^a-z0-9]",
        "",
        str(normalized.get("aggregation") or "").casefold(),
    )
    return ":".join(part for part in (field_type, table, column, aggregation) if part)


__all__ = ["normalize_field", "parse_field_reference", "semantic_comparison_key"]

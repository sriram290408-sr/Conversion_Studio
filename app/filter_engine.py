"""Power BI filter to Excel OLAP/Pivot filter helpers.

This module normalizes Power BI filter hints, applies supported selections to
Excel OLAP PivotTables, and creates native Excel slicers through COM.

The implementation is defensive because Excel exposes different COM members for
normal PivotTables, OLAP PivotTables, and different Excel versions.
"""

from __future__ import annotations

import logging
import re
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

logger = logging.getLogger("filter_engine")

XL_PAGE_FIELD = 3

_FILTER_OP_ALIASES = {
    "=": "equals",
    "==": "equals",
    "eq": "equals",
    "is": "equals",
    "equals": "equals",
    "in": "in",
    "not in": "not_in",
    "!=": "not_equals",
    "<>": "not_equals",
    "between": "between",
    ">": "greater_than",
    ">=": "greater_than_or_equal",
    "<": "less_than",
    "<=": "less_than_or_equal",
    "all": "all",
}


class FilterEngineError(RuntimeError):
    """Raised when a filter or slicer operation cannot be completed."""


def _as_list(value: Any) -> List[Any]:
    if value is None:
        return []

    if isinstance(value, (list, tuple, set)):
        return [item for item in value if item is not None and str(item).strip()]

    return [value] if str(value).strip() else []


def _clean_field(value: Any) -> str:
    return str(value or "").strip()


def _normalize_operator(value: Any) -> str:
    operator = str(value or "equals").strip().lower()
    return _FILTER_OP_ALIASES.get(
        operator,
        operator.replace(" ", "_"),
    )


def normalize_filter(
    filter_value: Any,
    default_scope: str = "visual",
) -> Optional[Dict[str, Any]]:
    """Normalize a Power BI filter hint into a consistent dictionary.

    Supported examples:
    - ``{"field": "Date[Year]", "operator": "=", "value": 2026}``
    - ``Date[Year] = 2026``
    - ``Region in South, West``
    - ``Amount between 100 and 500``
    - ``Date[Year]`` for a slicer/filter declaration with no selection
    """
    if not filter_value:
        return None

    scope = str(default_scope or "visual").strip().lower()

    if isinstance(filter_value, dict):
        field = _clean_field(
            filter_value.get("field")
            or filter_value.get("column")
            or filter_value.get("queryRef")
            or filter_value.get("name")
        )

        if not field:
            return None

        values = _as_list(
            filter_value.get("values")
            if "values" in filter_value
            else filter_value.get("value")
        )

        return {
            "field": field,
            "operator": _normalize_operator(
                filter_value.get("operator")
                or filter_value.get("condition")
                or filter_value.get("type")
                or "equals"
            ),
            "values": values,
            "scope": str(filter_value.get("scope") or scope).strip().lower(),
        }

    text = str(filter_value).strip()
    if not text:
        return None

    between_match = re.match(
        r"^(.+?)\s+between\s+(.+?)\s+and\s+(.+)$",
        text,
        re.IGNORECASE,
    )
    if between_match:
        return {
            "field": between_match.group(1).strip(),
            "operator": "between",
            "values": [
                between_match.group(2).strip().strip("'\""),
                between_match.group(3).strip().strip("'\""),
            ],
            "scope": scope,
        }

    comparison_match = re.match(
        r"^(.+?)\s+(not\s+in|in|>=|<=|<>|!=|=|>|<)\s*(.+)$",
        text,
        re.IGNORECASE,
    )
    if comparison_match:
        field, operator, raw_values = comparison_match.groups()
        cleaned = raw_values.strip().strip("()[]{}")
        values = [
            item.strip().strip("'\"") for item in cleaned.split(",") if item.strip()
        ]

        return {
            "field": field.strip(),
            "operator": _normalize_operator(operator),
            "values": values,
            "scope": scope,
        }

    return {
        "field": text,
        "operator": "all",
        "values": [],
        "scope": scope,
    }


def normalize_filters(
    filters: Iterable[Any],
    default_scope: str = "visual",
) -> List[Dict[str, Any]]:
    """Normalize and deduplicate filter records."""
    normalized: List[Dict[str, Any]] = []
    seen: set[Tuple[Any, ...]] = set()

    for item in filters or []:
        record = normalize_filter(
            item,
            default_scope=default_scope,
        )
        if not record:
            continue

        key = (
            record["field"].lower(),
            record["operator"],
            tuple(str(value) for value in record["values"]),
            record["scope"],
        )

        if key in seen:
            continue

        seen.add(key)
        normalized.append(record)

    return normalized


def _escape_olap_member_value(value: Any) -> str:
    return str(value).replace("]", "]]")


def _member_unique_name(
    olap_field: str,
    value: Any,
) -> str:
    """Build a common OLAP member unique-name fallback.

    Example:
    ``[Date].[Year].[Year]`` becomes
    ``[Date].[Year].&[2026]``.
    """
    field = str(olap_field or "").strip()
    if not field:
        raise FilterEngineError("OLAP field path is required.")

    hierarchy = field.rsplit(".", 1)[0] if "." in field else field
    return f"{hierarchy}.&[{_escape_olap_member_value(value)}]"


def _get_cube_field(
    pivot_table: Any,
    olap_field: str,
) -> Any:
    try:
        return pivot_table.CubeFields.Item(olap_field)
    except Exception:
        return pivot_table.CubeFields(olap_field)


def _get_pivot_field(
    pivot_table: Any,
    olap_field: str,
) -> Any:
    try:
        return pivot_table.PivotFields().Item(olap_field)
    except Exception:
        return pivot_table.PivotFields(olap_field)


def _clear_pivot_field_filters(pivot_field: Any) -> None:
    try:
        pivot_field.ClearAllFilters()
    except Exception:
        pass

    try:
        pivot_field.VisibleItemsList = []
    except Exception:
        pass


def _apply_single_member(
    pivot_field: Any,
    member_unique_name: str,
    fallback_caption: str,
) -> None:
    try:
        pivot_field.CurrentPageName = member_unique_name
        return
    except Exception:
        pass

    try:
        pivot_field.CurrentPage = fallback_caption
        return
    except Exception as exc:
        raise FilterEngineError(
            f"Excel rejected the filter member '{member_unique_name}': {exc}"
        ) from exc


def _apply_multiple_members(
    pivot_field: Any,
    members: Sequence[str],
) -> None:
    try:
        pivot_field.EnableMultiplePageItems = True
    except Exception:
        pass

    try:
        pivot_field.VisibleItemsList = list(members)
    except Exception as exc:
        raise FilterEngineError(
            f"Excel rejected one or more OLAP members: {exc}"
        ) from exc


class FilterEngine:
    def __init__(self) -> None:
        self.active_filters: List[Dict[str, Any]] = []
        self.slicer_counter = 0

    def register_filter(
        self,
        field: str,
        operator: str = "equals",
        values: Optional[List[Any]] = None,
        scope: str = "visual",
    ) -> None:
        """Register a normalized filter for later use."""
        record = normalize_filter(
            {
                "field": field,
                "operator": operator,
                "values": values or [],
                "scope": scope,
            },
            default_scope=scope,
        )

        if record:
            self.active_filters.append(record)

    def apply_filters_to_pivot(
        self,
        excel_pivot_table: Any,
        filters: Iterable[Any],
        field_mapper: Any,
    ) -> List[str]:
        """Apply supported equality and IN selections to an OLAP PivotTable.

        Unsupported operators are returned as warnings rather than being
        silently ignored.
        """
        warnings: List[str] = []

        for filter_record in normalize_filters(filters):
            field = filter_record["field"]
            operator = filter_record["operator"]
            values = filter_record["values"]

            mapping = field_mapper.map_field(
                field,
                "dimension",
            )

            if mapping.get("status") == "unmapped" or not mapping.get(
                "excel_olap_field"
            ):
                warnings.append(
                    f"Could not map filter field '{field}' to an OLAP CubeField."
                )
                continue

            olap_field = mapping["excel_olap_field"]

            try:
                cube_field = _get_cube_field(
                    excel_pivot_table,
                    olap_field,
                )
                cube_field.Orientation = XL_PAGE_FIELD

                pivot_field = _get_pivot_field(
                    excel_pivot_table,
                    olap_field,
                )
                _clear_pivot_field_filters(pivot_field)

                if operator == "all" or not values:
                    logger.info(
                        "Filter field %s was added without a selected value.",
                        field,
                    )
                    continue

                if operator not in {"equals", "in"}:
                    warnings.append(
                        f"Filter '{field}' operator '{operator}' is not "
                        "supported for OLAP PivotTable page filtering."
                    )
                    continue

                members = [
                    _member_unique_name(
                        olap_field,
                        value,
                    )
                    for value in values
                ]

                if len(members) == 1:
                    _apply_single_member(
                        pivot_field,
                        members[0],
                        str(values[0]),
                    )
                else:
                    _apply_multiple_members(
                        pivot_field,
                        members,
                    )

                logger.info(
                    "Applied filter %s %s %s to PivotTable.",
                    field,
                    operator,
                    values,
                )

            except Exception as exc:
                message = f"Failed to apply filter '{field}' to PivotTable: {exc}"
                logger.warning(message)
                warnings.append(message)

        return warnings

    def create_slicer(
        self,
        workbook: Any,
        dashboard_sheet: Any,
        source_pivot: Any,
        field: str,
        field_mapper: Any,
        target_pivots: Optional[List[Any]] = None,
        title: Optional[str] = None,
        left: float = 20,
        top: float = 20,
        width: float = 140,
        height: float = 110,
    ) -> Dict[str, Any]:
        """Create a native slicer from a connected OLAP PivotTable.

        Excel versions differ in which SourceField representation Add/Add2
        accepts. The implementation makes the hierarchy visible in the source
        PivotTable, refreshes it, and then tries the supported object/string
        variants without passing fragile optional arguments positionally.
        """
        mapping = field_mapper.map_field(field, "dimension")
        if mapping.get("status") == "unmapped" or not mapping.get("excel_olap_field"):
            return {
                "status": "failed",
                "field": field,
                "error": f"Slicer field '{field}' could not be mapped to an Excel OLAP CubeField.",
            }

        olap_field = str(mapping["excel_olap_field"])
        self.slicer_counter += 1
        safe_base = re.sub(r"[^A-Za-z0-9_]", "_", str(field)).strip("_") or "Filter"
        cache_name = f"Slicer_{safe_base}_{self.slicer_counter}"[:240]
        slicer_name = f"SL_{safe_base}_{self.slicer_counter}"[:240]
        creation_errors: List[str] = []

        try:
            cube_field = _get_cube_field(source_pivot, olap_field)
            try:
                cube_field.Orientation = 1  # xlRowField: materialise hierarchy
                cube_field.Position = 1
            except Exception as exc:
                creation_errors.append(f"CubeField orientation: {exc}")
            try:
                source_pivot.RefreshTable()
            except Exception:
                pass

            candidates: List[Any] = [cube_field, olap_field]
            for attr in ("Name", "Caption"):
                try:
                    value = getattr(cube_field, attr)
                    if value and value not in candidates:
                        candidates.append(value)
                except Exception:
                    pass
            try:
                pivot_field = _get_pivot_field(source_pivot, olap_field)
                candidates.insert(0, pivot_field)
            except Exception as exc:
                creation_errors.append(f"PivotFields lookup: {exc}")

            slicer_cache = None
            for candidate in candidates:
                for method_name in ("Add2", "Add"):
                    try:
                        method = getattr(workbook.SlicerCaches, method_name)
                        slicer_cache = method(source_pivot, candidate)
                        if slicer_cache is not None:
                            break
                    except Exception as exc:
                        creation_errors.append(
                            f"{method_name}({type(candidate).__name__}): {exc}"
                        )
                if slicer_cache is not None:
                    break

            if slicer_cache is None:
                raise FilterEngineError(
                    f"Excel rejected the OLAP slicer cache for '{olap_field}'. "
                    + " | ".join(creation_errors[-8:])
                )

            try:
                slicer_cache.Name = cache_name
            except Exception:
                try:
                    cache_name = str(slicer_cache.Name)
                except Exception:
                    pass

            connected_pivots = 0
            connection_warnings: List[str] = []
            for pivot in target_pivots or []:
                if pivot is source_pivot:
                    continue
                try:
                    slicer_cache.PivotTables.AddPivotTable(pivot)
                    connected_pivots += 1
                except Exception as exc:
                    connection_warnings.append(
                        f"Could not connect slicer to PivotTable '{getattr(pivot, 'Name', 'Unknown')}': {exc}"
                    )

            try:
                slicer = slicer_cache.Slicers.Add(
                    dashboard_sheet,
                    None,
                    slicer_name,
                    title or field,
                    float(top),
                    float(left),
                    float(width),
                    float(height),
                )
            except Exception:
                slicer = slicer_cache.Slicers.Add(
                    SlicerDestination=dashboard_sheet,
                    Name=slicer_name,
                    Caption=title or field,
                    Top=float(top),
                    Left=float(left),
                    Width=float(width),
                    Height=float(height),
                )

            try:
                actual_name = str(slicer.Name)
            except Exception:
                actual_name = slicer_name
            try:
                slicer.NumberOfColumns = 1
            except Exception:
                pass
            try:
                slicer.Style = "SlicerStyleLight2"
            except Exception:
                pass
            try:
                slicer.DisplayHeader = True
            except Exception:
                pass
            try:
                slicer.Caption = title or field
            except Exception:
                pass
            logger.info("Created native OLAP slicer %s for %s.", actual_name, field)
            return {
                "status": "success",
                "field": field,
                "olap_field": olap_field,
                "slicer_cache_name": cache_name,
                "slicer_name": actual_name,
                "formula_reference": cache_name,
                "connected_pivots": connected_pivots,
                "warnings": connection_warnings,
            }
        except Exception as exc:
            logger.warning("Slicer creation failed for %s: %s", olap_field, exc)
            return {
                "status": "failed",
                "field": field,
                "olap_field": olap_field,
                "error": str(exc),
                "warnings": creation_errors,
            }


__all__ = [
    "FilterEngine",
    "FilterEngineError",
    "normalize_filter",
    "normalize_filters",
]

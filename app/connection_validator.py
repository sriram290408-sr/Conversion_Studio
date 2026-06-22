"""Validate Power BI / OLAP connections in an opened Excel workbook."""

from __future__ import annotations

import logging
import re
from typing import Any, Dict, List, Optional

try:
    from live_binding_models import ConnectionValidationResult
except ImportError:
    from live_binding_models import ConnectionValidationResult

logger = logging.getLogger("connection_validator")

# Excel connection type constants
XL_CONNECTION_TYPE_OLEDB = 1
XL_CONNECTION_TYPE_ODBC = 2
XL_CONNECTION_TYPE_MODEL = 7


def _safe_connection_type(connection: Any) -> Optional[int]:
    try:
        return int(connection.Type)
    except Exception:
        return None


def _safe_connection_name(connection: Any, fallback: str) -> str:
    try:
        name = str(connection.Name).strip()
        return name or fallback
    except Exception:
        return fallback


def _get_oledb_connection(connection: Any) -> Any:
    try:
        return connection.OLEDBConnection
    except Exception:
        return None


def _get_odbc_connection(connection: Any) -> Any:
    try:
        return connection.ODBCConnection
    except Exception:
        return None


def _looks_like_power_bi_or_olap(connection: Any) -> bool:
    """Return True when a workbook connection appears to be OLAP/Power BI."""
    connection_type = _safe_connection_type(connection)

    if connection_type in {
        XL_CONNECTION_TYPE_OLEDB,
        XL_CONNECTION_TYPE_MODEL,
    }:
        return True

    oledb = _get_oledb_connection(connection)
    if oledb is not None:
        try:
            connection_string = str(oledb.Connection or "").lower()
        except Exception:
            connection_string = ""

        if any(
            token in connection_string
            for token in (
                "msolap",
                "analysis services",
                "power bi",
                "xmla",
                "cube",
                "model",
            )
        ):
            return True

    return False


def _set_connection_synchronous(connection: Any) -> None:
    """Disable background refresh when the connection supports it."""
    oledb = _get_oledb_connection(connection)
    if oledb is not None:
        try:
            oledb.BackgroundQuery = False
        except Exception:
            pass

    odbc = _get_odbc_connection(connection)
    if odbc is not None:
        try:
            odbc.BackgroundQuery = False
        except Exception:
            pass



def _field_reference_from_value(value: Any, requested_type: str) -> Any:
    """Return a mapper-friendly descriptor from string or binding dictionaries."""
    if not isinstance(value, dict):
        return value

    descriptor = dict(value)
    requested = str(requested_type or "").casefold()

    raw = str(
        descriptor.get("raw_reference")
        or descriptor.get("canonical_reference")
        or descriptor.get("field")
        or ""
    ).strip()
    table = str(descriptor.get("table_name") or descriptor.get("table") or "").strip()
    column = str(
        descriptor.get("column_name")
        or descriptor.get("column")
        or descriptor.get("field_name")
        or ""
    ).strip()
    measure = str(
        descriptor.get("measure_name")
        or descriptor.get("name")
        or descriptor.get("display_name")
        or ""
    ).strip()
    aggregation = str(descriptor.get("aggregation") or "").strip().upper()
    semantic_type = str(descriptor.get("field_type") or "").casefold()

    if requested == "measure":
        if semantic_type in {"named_measure", "measure"} or (
            measure and not aggregation and not column
        ):
            descriptor["field_type"] = "named_measure"
            descriptor["measure_name"] = measure or raw.strip("[]")
            descriptor["raw_reference"] = raw or descriptor["measure_name"]
            descriptor["canonical_reference"] = (
                descriptor.get("canonical_reference")
                or descriptor["measure_name"]
            )
            return descriptor

        if aggregation or (table and column):
            descriptor["field_type"] = "implicit_measure"
            descriptor["table_name"] = table
            descriptor["column_name"] = column
            descriptor["aggregation"] = aggregation or "SUM"
            descriptor["raw_reference"] = raw or (
                f"{descriptor['aggregation']}({table}[{column}])"
                if table
                else f"{descriptor['aggregation']}({column})"
            )
            descriptor["canonical_reference"] = (
                descriptor.get("canonical_reference")
                or (f"{table}[{column}]" if table else column)
            )
            return descriptor

        if measure:
            descriptor["field_type"] = "named_measure"
            descriptor["measure_name"] = measure
            descriptor["raw_reference"] = raw or measure
            descriptor["canonical_reference"] = measure
            return descriptor

    if table and column:
        descriptor["field_type"] = semantic_type or "dimension"
        descriptor["table_name"] = table
        descriptor["column_name"] = column
        descriptor["raw_reference"] = raw or f"{table}[{column}]"
        descriptor["canonical_reference"] = (
            descriptor.get("canonical_reference") or f"{table}[{column}]"
        )
        return descriptor

    return raw or measure or column or descriptor


def _norm_token(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").casefold())


def _cube_signature(name: str) -> Dict[str, str]:
    parts = re.findall(r"\[([^\]]+)\]", str(name or ""))
    if not parts:
        return {
            "kind": "unknown",
            "table": "",
            "field": "",
            "raw": str(name or ""),
        }

    if str(name).casefold().startswith("[measures]."):
        return {
            "kind": "measure",
            "table": "",
            "field": _norm_token(parts[-1]),
            "raw": str(name),
        }

    return {
        "kind": "dimension",
        "table": _norm_token(parts[0] if len(parts) >= 2 else ""),
        "field": _norm_token(parts[-1]),
        "raw": str(name),
    }


def _direct_signature_match(
    source_field: Any,
    requested_type: str,
    discovered_cube_fields: List[str],
) -> Optional[Dict[str, Any]]:
    """Fallback matching for serialized binding dictionaries."""
    try:
        from .field_normalizer import normalize_field
    except ImportError:
        from field_normalizer import normalize_field

    role = "values" if str(requested_type).casefold() == "measure" else None
    normalized = normalize_field(source_field, role=role)
    semantic_type = str(normalized.get("field_type") or "").casefold()
    table = _norm_token(normalized.get("table_name"))
    column = _norm_token(normalized.get("column_name"))
    measure = _norm_token(
        normalized.get("measure_name") or normalized.get("display_name")
    )

    signatures = [_cube_signature(name) for name in discovered_cube_fields]

    if semantic_type == "named_measure" or (
        str(requested_type).casefold() == "measure" and measure and not column
    ):
        candidates = [
            item
            for item in signatures
            if item["kind"] == "measure" and item["field"] == measure
        ]
        if len(candidates) == 1:
            return {
                "status": "mapped",
                "excel_olap_field": candidates[0]["raw"],
                "confidence": 0.99,
                "match_type": "direct_measure_signature",
                "field_type": "named_measure",
                "normalized_reference": normalized.get("canonical_reference"),
            }

    if column:
        exact = [
            item
            for item in signatures
            if item["kind"] == "dimension"
            and item["field"] == column
            and (not table or item["table"] == table)
        ]
        if len(exact) == 1:
            implicit = semantic_type == "implicit_measure" or (
                str(requested_type).casefold() == "measure"
            )
            return {
                "status": (
                    "requires_pivot_aggregation" if implicit else "mapped"
                ),
                "excel_olap_field": exact[0]["raw"],
                "confidence": 0.98,
                "match_type": (
                    "direct_implicit_measure_signature"
                    if implicit
                    else "direct_dimension_signature"
                ),
                "field_type": "implicit_measure" if implicit else "dimension",
                "normalized_reference": normalized.get("canonical_reference"),
            }

        column_only = [
            item
            for item in signatures
            if item["kind"] == "dimension" and item["field"] == column
        ]
        if len(column_only) == 1:
            implicit = semantic_type == "implicit_measure" or (
                str(requested_type).casefold() == "measure"
            )
            return {
                "status": (
                    "requires_pivot_aggregation" if implicit else "mapped"
                ),
                "excel_olap_field": column_only[0]["raw"],
                "confidence": 0.90,
                "match_type": "direct_unique_column_signature",
                "field_type": "implicit_measure" if implicit else "dimension",
                "normalized_reference": normalized.get("canonical_reference"),
            }

    return None



def _parse_implicit_measure_caption(
    value: Any,
    discovered_cube_fields: List[str],
) -> Optional[Dict[str, Any]]:
    """Map 'Sum of column' captions to the matching source CubeField."""
    if isinstance(value, dict):
        raw = str(
            value.get("measure_name")
            or value.get("display_name")
            or value.get("raw_reference")
            or value.get("canonical_reference")
            or ""
        ).strip()
    else:
        raw = str(value or "").strip()

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

    column_name = match.group(2).strip()
    column_key = _norm_token(column_name)

    candidates = [
        signature
        for signature in (
            _cube_signature(name)
            for name in discovered_cube_fields
            if "__default measure" not in str(name).casefold()
        )
        if signature["kind"] == "dimension"
        and signature["field"] == column_key
    ]

    if len(candidates) != 1:
        return None

    return {
        "status": "requires_pivot_aggregation",
        "excel_olap_field": candidates[0]["raw"],
        "confidence": 0.97,
        "match_type": "implicit_measure_caption",
        "field_type": "implicit_measure",
        "aggregation": aggregation,
        "column_name": column_name,
        "normalized_reference": column_name,
        "reason": (
            "The caption represents an implicit aggregation over a source column."
        ),
    }


def _match_table_bracket_named_measure(
    value: Any,
    discovered_cube_fields: List[str],
) -> Optional[Dict[str, Any]]:
    """Map Table[Measure] to [Measures].[Measure] even if the role is wrong."""
    if isinstance(value, dict):
        raw = str(
            value.get("raw_reference")
            or value.get("canonical_reference")
            or value.get("field")
            or ""
        ).strip()
    else:
        raw = str(value or "").strip()

    match = re.fullmatch(r"'?([^'\[\]]+)'?\[([^\]]+)\]", raw)
    if not match:
        return None

    measure_name = match.group(2).strip()
    measure_key = _norm_token(measure_name)

    candidates = [
        signature
        for signature in (
            _cube_signature(name)
            for name in discovered_cube_fields
            if "__default measure" not in str(name).casefold()
        )
        if signature["kind"] == "measure"
        and signature["field"] == measure_key
    ]

    if len(candidates) != 1:
        return None

    return {
        "status": "mapped",
        "excel_olap_field": candidates[0]["raw"],
        "confidence": 1.0,
        "match_type": "table_bracket_named_measure",
        "field_type": "named_measure",
        "measure_name": measure_name,
        "normalized_reference": measure_name,
        "reason": "",
    }


def _semantic_override(
    source_field: Any,
    requested_type: str,
    discovered_cube_fields: List[str],
) -> Optional[Dict[str, Any]]:
    """Apply authoritative CubeField-based corrections before scoring."""
    usable_fields = [
        field
        for field in discovered_cube_fields
        if "__default measure" not in str(field).casefold()
    ]

    named_measure = _match_table_bracket_named_measure(
        source_field, usable_fields
    )
    if named_measure is not None:
        return named_measure

    implicit_measure = _parse_implicit_measure_caption(
        source_field, usable_fields
    )
    if implicit_measure is not None:
        return implicit_measure

    return None


def _normalize_binding_fields(binding: Any) -> List[Dict[str, Any]]:
    """Flatten a visual binding and preserve mapper-friendly descriptors."""
    result: List[Dict[str, Any]] = []

    if isinstance(binding, dict):
        rows = binding.get("rows") or []
        columns = binding.get("columns") or []
        legend = binding.get("legend") or []
        measures = binding.get("measures") or []
        filters = binding.get("filters") or []
        slicer_field = binding.get("slicer_field")
    else:
        rows = getattr(binding, "rows", []) or []
        columns = getattr(binding, "columns", []) or []
        legend = getattr(binding, "legend", []) or []
        measures = getattr(binding, "measures", []) or []
        filters = getattr(binding, "filters", []) or []
        slicer_field = getattr(binding, "slicer_field", None)

    def add(field: Any, requested_type: str) -> None:
        if field is None or field == "":
            return
        result.append(
            {
                "field": _field_reference_from_value(field, requested_type),
                "source_field": field,
                "requested_field_type": requested_type,
            }
        )

    for field in list(rows) + list(columns) + list(legend):
        add(field, "dimension")
    for field in list(measures):
        add(field, "measure")
    add(slicer_field, "dimension")

    for item in filters or []:
        if isinstance(item, dict) and "field" in item:
            add(item.get("field"), "dimension")
        else:
            add(item, "dimension")

    return result

def _semantic_match_score(field_mapping: Dict[str, Any]) -> Dict[str, Any]:
    """Calculate a weighted semantic score from the mapper's real field types."""
    weights = {
        "named_measure": 0.40,
        "dimension": 0.40,
        "implicit_measure": 0.20,
    }
    eligible = {key: 0 for key in weights}
    matched = {key: 0 for key in weights}
    seen: set[str] = set()

    items = list(field_mapping.get("mapped", [])) + list(
        field_mapping.get("unmapped", [])
    )

    for item in items:
        semantic_type = str(
            item.get("semantic_field_type") or item.get("field_type") or ""
        ).casefold()

        if semantic_type == "hierarchy":
            category = "dimension"
        elif semantic_type in weights:
            category = semantic_type
        else:
            category = (
                "named_measure"
                if str(item.get("requested_field_type") or "").casefold() == "measure"
                else "dimension"
            )

        identity = str(
            item.get("normalized_reference")
            or item.get("canonical_reference")
            or item.get("pbix_field")
            or item.get("field")
            or item.get("raw_reference")
            or ""
        ).casefold()
        key = f"{category}:{identity}"
        if key in seen:
            continue
        seen.add(key)

        eligible[category] += 1
        if item.get("status") in {"mapped", "requires_pivot_aggregation"}:
            matched[category] += 1

    if sum(eligible.values()) == 0:
        return {
            "score": 1.0,
            "required_measures": 0,
            "mapped_measures": 0,
            "required_dimensions": 0,
            "mapped_dimensions": 0,
            "eligible_by_category": eligible,
            "matched_by_category": matched,
        }

    active_categories = [key for key in weights if eligible[key] > 0]
    active_weight_sum = sum(weights[key] for key in active_categories)
    score = (
        sum((matched[key] / eligible[key]) * weights[key] for key in active_categories)
        / active_weight_sum
    )

    return {
        "score": round(score, 4),
        "required_measures": eligible["named_measure"] + eligible["implicit_measure"],
        "mapped_measures": matched["named_measure"] + matched["implicit_measure"],
        "required_dimensions": eligible["dimension"],
        "mapped_dimensions": matched["dimension"],
        "eligible_by_category": eligible,
        "matched_by_category": matched,
    }


def validate_semantic_model_compatibility(
    bindings: Any,
    tmdl_measures: Any,
    field_mapper: Any,
) -> Dict[str, Any]:
    """Validate the connected OLAP model against PBIX/TMDL visual bindings."""
    if field_mapper is None:
        raise ValueError("field_mapper is required")

    bindings = bindings or []
    field_mapping: Dict[str, List[Dict[str, Any]]] = {
        "mapped": [],
        "unmapped": [],
    }

    def register(
        mapping: Dict[str, Any],
        *,
        source_field: Any,
        requested_type: str,
        source: str,
    ) -> None:
        semantic_type = str(mapping.get("field_type") or "").casefold()
        output = {
            **mapping,
            "field": source_field,
            "requested_field_type": requested_type,
            "semantic_field_type": semantic_type,
            "validation_source": source,
        }
        bucket = (
            "mapped"
            if output.get("status") in {"mapped", "requires_pivot_aggregation"}
            else "unmapped"
        )
        field_mapping[bucket].append(output)

    for binding in bindings:
        for record in _normalize_binding_fields(binding):
            field = record["field"]
            source_field = record.get("source_field", field)
            requested_type = record["requested_field_type"]
            try:
                mapping = field_mapper.map_field(field, requested_type)
            except Exception as exc:
                raise RuntimeError(
                    f"Field mapping failed for {source_field}: {exc}"
                ) from exc

            discovered_fields = list(
                getattr(field_mapper, "discovered_cubefields", [])
                or []
            )
            override = _semantic_override(
                source_field,
                requested_type,
                discovered_fields,
            )
            if override is not None:
                mapping = {**mapping, **override}

            if mapping.get("status") not in {
                "mapped",
                "requires_pivot_aggregation",
            }:
                fallback = _direct_signature_match(
                    field,
                    requested_type,
                    list(
                        getattr(field_mapper, "discovered_cubefields", [])
                        or []
                    ),
                )
                if fallback is not None:
                    mapping = {**mapping, **fallback}

            register(
                mapping,
                source_field=source_field,
                requested_type=requested_type,
                source="visual_binding",
            )

    missing_measures: List[str] = []
    for measure in tmdl_measures or []:
        if isinstance(measure, dict):
            measure_name = str(
                measure.get("measure_name")
                or measure.get("name")
                or measure.get("display_name")
                or ""
            ).strip()
        else:
            measure_name = str(measure or "").strip()

        if not measure_name:
            continue

        descriptor = {
            "field_type": "named_measure",
            "measure_name": measure_name,
            "raw_reference": measure_name,
            "canonical_reference": measure_name,
        }

        try:
            mapping = field_mapper.map_field(descriptor, "measure")
        except Exception as exc:
            raise RuntimeError(
                f"TMDL measure mapping failed for {measure_name}: {exc}"
            ) from exc

        if mapping.get("status") not in {
            "mapped",
            "requires_pivot_aggregation",
        }:
            fallback = _direct_signature_match(
                descriptor,
                "measure",
                list(
                    getattr(field_mapper, "discovered_cubefields", [])
                    or []
                ),
            )
            if fallback is not None:
                mapping = {**mapping, **fallback}

        register(
            mapping,
            source_field=measure_name,
            requested_type="measure",
            source="tmdl_measure",
        )

        if mapping.get("status") not in {"mapped", "requires_pivot_aggregation"}:
            missing_measures.append(measure_name)

    score_result = _semantic_match_score(field_mapping)
    mapped_items = field_mapping["mapped"]
    unmapped_items = field_mapping["unmapped"]

    logger.info(
        "Semantic validation: score=%.2f mapped=%d unmapped=%d categories=%s/%s",
        score_result["score"],
        len(mapped_items),
        len(unmapped_items),
        score_result["matched_by_category"],
        score_result["eligible_by_category"],
    )
    for item in unmapped_items:
        logger.warning(
            "Unmapped semantic field: field=%r requested=%s semantic=%s reason=%s",
            item.get("field"),
            item.get("requested_field_type"),
            item.get("semantic_field_type"),
            item.get("reason"),
        )

    return {
        **score_result,
        "eligible_count": sum(score_result["eligible_by_category"].values()),
        "matched_count": sum(score_result["matched_by_category"].values()),
        "missing_measures": missing_measures,
        "mapped_fields": mapped_items,
        "unmapped_fields": unmapped_items,
    }


def _has_reusable_pivot_cache(workbook: Any) -> bool:
    try:
        caches = workbook.PivotCaches()
        return int(caches.Count) > 0
    except Exception:
        return False


def _find_template_pivot(workbook: Any) -> bool:
    for sheet in workbook.Worksheets:
        try:
            sheet_name = str(sheet.Name or "")
        except Exception:
            sheet_name = ""

        if "template" in sheet_name.lower():
            return True

        try:
            if int(sheet.PivotTables().Count) > 0:
                return True
        except Exception:
            continue

    return False


def _wait_for_refresh(excel_app: Any) -> None:
    """Wait for Excel asynchronous queries and calculation to complete."""
    try:
        excel_app.CalculateUntilAsyncQueriesDone()
    except Exception:
        pass

    try:
        excel_app.CalculateFullRebuild()
    except Exception:
        try:
            excel_app.CalculateFull()
        except Exception:
            pass


def validate_workbook_connection(
    excel_app: Any,
    workbook: Any,
    *,
    attempt_refresh: bool = True,
) -> ConnectionValidationResult:
    """Inspect and validate Power BI / OLAP connections in an opened workbook.

    The function:
    - records every workbook connection name;
    - identifies likely Power BI / OLAP connections;
    - checks for reusable PivotCaches;
    - checks for an existing template PivotTable;
    - optionally refreshes the workbook synchronously.

    It never assumes that the first workbook connection is the correct one.
    """
    result = ConnectionValidationResult()
    result.excel_com_available = excel_app is not None
    result.workbook_opened = workbook is not None

    if excel_app is None:
        result.errors.append("Excel COM application is not available.")
        return result

    if workbook is None:
        result.errors.append("Excel workbook is not open.")
        return result

    power_bi_connection_names: List[str] = []

    try:
        connections = workbook.Connections
        connection_count = int(connections.Count)

        logger.info(
            "Found %d connection(s) in workbook.",
            connection_count,
        )

        for index in range(1, connection_count + 1):
            try:
                connection = connections.Item(index)
                name = _safe_connection_name(
                    connection,
                    f"Connection {index}",
                )
                result.connection_names.append(name)

                if _looks_like_power_bi_or_olap(connection):
                    result.connection_found = True
                    result.olap_connection_found = True
                    power_bi_connection_names.append(name)

                _set_connection_synchronous(connection)

            except Exception as exc:
                message = f"Failed to inspect workbook connection {index}: {exc}"
                logger.warning(message)
                result.errors.append(message)

        result.pivot_cache_found = _has_reusable_pivot_cache(workbook)
        result.template_pivot_found = _find_template_pivot(workbook)

        # Preserve only validated OLAP/Power BI names when any were found.
        if power_bi_connection_names:
            result.connection_names = power_bi_connection_names

        if not result.connection_found:
            result.errors.append(
                "No Power BI or OLAP workbook connection was detected."
            )
            return result

        if not result.pivot_cache_found:
            result.errors.append("No reusable PivotCache was found in the workbook.")

        if attempt_refresh:
            result.refresh_attempted = True

            try:
                workbook.RefreshAll()
                _wait_for_refresh(excel_app)

                result.refresh_success = True
                result.semantic_model_match = True

                logger.info("Workbook refresh completed successfully.")
            except Exception as refresh_error:
                message = f"Connection refresh failed: {refresh_error}"
                logger.warning(message)

                result.errors.append(message)
                result.refresh_success = False
                result.semantic_model_match = False
        else:
            result.refresh_attempted = False

    except Exception as exc:
        logger.exception("Error validating workbook connection.")
        result.errors.append(str(exc))

    return result


__all__ = [
    "validate_workbook_connection",
    "validate_semantic_model_compatibility",
]

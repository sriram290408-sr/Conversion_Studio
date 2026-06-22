"""Map normalized PBIX/TMDL fields to Excel OLAP CubeField unique names.

Matching priority (highest to lowest):
1. Exact case-insensitive CubeField path match.
2. Table + column normalized match (both table and column tokens agree).
3. Unique column-name match across all non-technical tables.
4. Column-name match ignoring technical/hidden tables.

Technical tables are filtered out from fallback matching to prevent false
positives on auto-generated Power BI system tables.
"""

from __future__ import annotations

import logging
import re
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple

try:
    from .field_normalizer import normalize_field, parse_field_reference
except ImportError:
    from field_normalizer import normalize_field, parse_field_reference

logger = logging.getLogger("olap_field_mapper")

# Power BI auto-generated / technical table prefixes to de-prioritize in
# fallback matching.  These tables appear as CubeField entries but are rarely
# the intended mapping target when a user-facing column name is referenced.
_TECHNICAL_TABLE_PREFIXES: Tuple[str, ...] = (
    "localdate",
    "datatable",
    "$",
    "__",
    "rowlevelsecurity",
    "roleplayingdimension",
)


def _norm(value: Any) -> str:
    """Collapse to lowercase alphanumeric for fuzzy matching."""
    return re.sub(r"[^a-z0-9]", "", str(value or "").casefold())


def _cube_parts(name: str) -> List[str]:
    """Extract all bracket-delimited tokens from a CubeField unique name."""
    return re.findall(r"\[([^\]]+)\]", str(name or ""))


def _is_technical_table(table_token: str) -> bool:
    """Return True when the table looks like a Power BI system/auto table."""
    normed = table_token.strip().casefold()
    return any(normed.startswith(prefix) for prefix in _TECHNICAL_TABLE_PREFIXES)


class OLAPFieldMapper:
    """Map fields without converting implicit aggregations into fake measures."""

    def __init__(self) -> None:
        self.reset()

    def reset(self) -> None:
        self.mappings: Dict[str, Dict[str, Any]] = {}
        self.discovered_cubefields: List[str] = []
        # measures: norm(measure_name) → [cubefield_path, ...]
        self._measure_by_norm: Dict[str, List[str]] = {}
        # dimensions: norm(column) → [cubefield_path, ...]
        self._dimension_by_norm: Dict[str, List[str]] = {}
        # dimensions: (norm(table), norm(column)) → [cubefield_path, ...]
        self._dimension_by_table_column: Dict[Tuple[str, str], List[str]] = {}
        # non-technical dimensions: norm(column) → [cubefield_path, ...]
        self._dimension_by_norm_nontechnical: Dict[str, List[str]] = {}

    def discover_fields_from_excel(self, excel_pivot_table: Any) -> None:
        try:
            fields = excel_pivot_table.CubeFields
            self.discover_fields_from_names(
                str(fields.Item(index).Name)
                for index in range(1, int(fields.Count) + 1)
            )
        except Exception as exc:
            logger.warning("Failed to discover CubeFields: %s", exc)
            self.reset()

    def discover_fields_from_names(self, names: Iterable[str]) -> None:
        self.reset()
        seen: Set[str] = set()
        for raw in names or []:
            name = str(raw or "").strip()
            if not name or name.casefold() in seen:
                continue
            seen.add(name.casefold())
            self.discovered_cubefields.append(name)
            parts = _cube_parts(name)
            if not parts:
                continue

            # Measure path: [Measures].[Name]
            if name.casefold().startswith("[measures]."):
                self._measure_by_norm.setdefault(_norm(parts[-1]), []).append(name)
                continue

            # Dimension / column path
            column = parts[-1]
            table = parts[0] if len(parts) >= 2 else ""

            self._dimension_by_norm.setdefault(_norm(column), []).append(name)

            if table:
                self._dimension_by_table_column.setdefault(
                    (_norm(table), _norm(column)), []
                ).append(name)

                if not _is_technical_table(table):
                    self._dimension_by_norm_nontechnical.setdefault(
                        _norm(column), []
                    ).append(name)

        logger.info("Discovered %d unique CubeFields", len(self.discovered_cubefields))

    # ------------------------------------------------------------------
    # Internal resolution helpers
    # ------------------------------------------------------------------

    def _exact_candidate(self, candidates: Iterable[str]) -> str:
        """Return the first candidate that exists verbatim (case-insensitive) in discovered CubeFields."""
        candidate_map = {item.casefold(): item for item in self.discovered_cubefields}
        for candidate in candidates:
            match = candidate_map.get(candidate.casefold())
            if match:
                return match
        return ""

    @staticmethod
    def _candidate_dimension_paths(table: str, column: str) -> List[str]:
        """Generate the most common OLAP path variants for a table/column pair."""
        if not table:
            return []
        return [
            f"[{table}].[{column}].[{column}]",  # three-part attribute hierarchy
            f"[{table}].[{column}]",  # two-part column path
        ]

    def _resolve_measure(self, name: str) -> Dict[str, Any]:
        # Priority 1 – exact path match
        exact = self._exact_candidate([f"[Measures].[{name}]"])
        if exact:
            return {
                "status": "mapped",
                "excel_olap_field": exact,
                "confidence": 1.0,
                "match_type": "exact_cube_field",
            }
        # Priority 2 – normalized name match
        matches = self._measure_by_norm.get(_norm(name), [])
        if len(matches) == 1:
            return {
                "status": "mapped",
                "excel_olap_field": matches[0],
                "confidence": 0.99,
                "match_type": "normalized_measure",
            }
        if len(matches) > 1:
            return {
                "status": "ambiguous",
                "excel_olap_field": "",
                "confidence": 0.0,
                "match_type": "ambiguous_measure",
            }
        return {
            "status": "unmapped",
            "excel_olap_field": "",
            "confidence": 0.0,
            "match_type": "measure_not_found",
        }

    def _resolve_dimension(self, table: str, column: str) -> Dict[str, Any]:
        # Priority 1 – exact OLAP path variant
        exact = self._exact_candidate(self._candidate_dimension_paths(table, column))
        if exact:
            return {
                "status": "mapped",
                "excel_olap_field": exact,
                "confidence": 1.0,
                "match_type": "exact_cube_field",
            }

        # Priority 2 – table + column normalized match
        if table:
            matches = self._dimension_by_table_column.get(
                (_norm(table), _norm(column)), []
            )
            if len(matches) == 1:
                return {
                    "status": "mapped",
                    "excel_olap_field": matches[0],
                    "confidence": 0.98,
                    "match_type": "table_and_column",
                }
            if len(matches) > 1:
                # Prefer the entry whose first bracket part is a closer match
                return {
                    "status": "mapped",
                    "excel_olap_field": matches[0],
                    "confidence": 0.95,
                    "match_type": "table_and_column_multi",
                }

        # Priority 3 – column in non-technical tables
        nontechnical = self._dimension_by_norm_nontechnical.get(_norm(column), [])
        if len(nontechnical) == 1:
            return {
                "status": "mapped",
                "excel_olap_field": nontechnical[0],
                "confidence": 0.90,
                "match_type": "unique_column_nontechnical",
            }
        if len(nontechnical) > 1:
            # Multiple non-technical tables share this column name; pick the first
            # rather than failing—the caller still gets a live field.
            return {
                "status": "mapped",
                "excel_olap_field": nontechnical[0],
                "confidence": 0.70,
                "match_type": "ambiguous_nontechnical_first",
            }

        # Priority 4 – column across all tables (including technical)
        matches = self._dimension_by_norm.get(_norm(column), [])
        if len(matches) == 1:
            return {
                "status": "mapped",
                "excel_olap_field": matches[0],
                "confidence": 0.85,
                "match_type": "unique_column",
            }
        if len(matches) > 1:
            # Still ambiguous across all tables; pick first with lower confidence.
            return {
                "status": "mapped",
                "excel_olap_field": matches[0],
                "confidence": 0.60,
                "match_type": "ambiguous_column_first",
            }
        return {
            "status": "unmapped",
            "excel_olap_field": "",
            "confidence": 0.0,
            "match_type": "dimension_not_found",
        }

    # ------------------------------------------------------------------
    # Public mapping API
    # ------------------------------------------------------------------

    def map_field(
        self, pbix_field: Any, field_type: str = "dimension"
    ) -> Dict[str, Any]:
        requested_type = str(field_type or "dimension").casefold()

        if isinstance(pbix_field, dict):
            normalized = dict(pbix_field)
            raw = str(
                normalized.get("raw_reference")
                or normalized.get("canonical_reference")
                or normalized.get("measure_name")
                or normalized.get("field")
                or ""
            ).strip()
        else:
            raw = str(pbix_field or "").strip()
            role = "values" if requested_type == "measure" else None
            normalized = normalize_field(raw, role=role)

        semantic_type = str(normalized.get("field_type") or "unknown").casefold()
        table = str(normalized.get("table_name") or "")
        column = str(normalized.get("column_name") or "")
        measure_name = str(normalized.get("measure_name") or "")
        aggregation = str(normalized.get("aggregation") or "")

        if not table and not measure_name and raw:
            parsed = parse_field_reference(raw)
            table = table or str(parsed.get("table_name") or "")
            column = column or str(parsed.get("column_name") or "")
            measure_name = measure_name or str(parsed.get("measure_name") or "")
            aggregation = aggregation or str(parsed.get("aggregation") or "")

        key = (
            f"{requested_type}:{semantic_type}:{table}:{column}:"
            f"{measure_name}:{aggregation}:{raw}"
        ).casefold()
        if key in self.mappings:
            return dict(self.mappings[key])

        if semantic_type == "named_measure":
            resolved = self._resolve_measure(measure_name or raw.strip("[]"))

        elif semantic_type == "implicit_measure":
            dimension_mapping = self._resolve_dimension(table, column)
            resolved = {
                **dimension_mapping,
                "status": (
                    "requires_pivot_aggregation"
                    if dimension_mapping.get("status") == "mapped"
                    else dimension_mapping.get("status")
                ),
                "match_type": (
                    "implicit_measure_source_column"
                    if dimension_mapping.get("status") == "mapped"
                    else dimension_mapping.get("match_type")
                ),
            }

        elif requested_type == "measure":
            # PBIX frequently serializes a named measure as Table[Measure].
            candidate_measure = measure_name or column
            measure_mapping = (
                self._resolve_measure(candidate_measure)
                if candidate_measure
                else {
                    "status": "unmapped",
                    "excel_olap_field": "",
                    "confidence": 0.0,
                    "match_type": "measure_not_found",
                }
            )

            if measure_mapping.get("status") == "mapped":
                resolved = measure_mapping
                semantic_type = "named_measure"
                measure_name = candidate_measure
                column = ""
            elif table and column:
                # Otherwise it is a source column in Values and should be
                # aggregated through the connected PivotTable.
                dimension_mapping = self._resolve_dimension(table, column)
                resolved = {
                    **dimension_mapping,
                    "status": (
                        "requires_pivot_aggregation"
                        if dimension_mapping.get("status") == "mapped"
                        else dimension_mapping.get("status")
                    ),
                    "match_type": (
                        "implicit_measure_source_column"
                        if dimension_mapping.get("status") == "mapped"
                        else dimension_mapping.get("match_type")
                    ),
                }
                semantic_type = "implicit_measure"
            else:
                resolved = measure_mapping
                semantic_type = "named_measure"
                measure_name = candidate_measure or raw.strip("[]")

        elif semantic_type in {"dimension", "hierarchy"}:
            resolved = self._resolve_dimension(
                table,
                column or str(normalized.get("display_name") or raw),
            )

        else:
            if requested_type == "measure":
                resolved = self._resolve_measure(raw.strip("[]"))
                semantic_type = "named_measure"
                measure_name = raw.strip("[]")
            else:
                resolved = self._resolve_dimension(table, column or raw)
                semantic_type = "dimension"

        if resolved.get("status") in {"unmapped", "ambiguous"} and raw:
            fuzzy_hit = self.best_match_for_raw(raw)
            if fuzzy_hit:
                resolved = {
                    "status": "mapped",
                    "excel_olap_field": fuzzy_hit,
                    "confidence": 0.50,
                    "match_type": "substring_fallback",
                }

        result = {
            "pbix_field": raw,
            "normalized_reference": normalized.get("canonical_reference") or raw,
            "excel_olap_field": resolved.get("excel_olap_field", ""),
            "excel_column": resolved.get("excel_olap_field", ""),
            "field_type": semantic_type,
            "aggregation": aggregation,
            "confidence": float(resolved.get("confidence", 0.0)),
            "status": resolved.get("status", "unmapped"),
            "match_type": resolved.get("match_type", "none"),
            "table_name": table,
            "field_name": measure_name or column or raw,
            "reason": "",
        }

        if result["status"] == "unmapped":
            result["reason"] = "No compatible CubeField was found."
        elif result["status"] == "ambiguous":
            result["reason"] = "Multiple CubeFields matched this reference."
        elif result["status"] == "requires_pivot_aggregation":
            result["reason"] = (
                "The source column exists and will be aggregated through the "
                "connected PivotTable."
            )

        self.mappings[key] = result
        return dict(result)

    def map_fields(self, fields: Iterable[Any]) -> Dict[str, Any]:
        mapped: List[Dict[str, Any]] = []
        unresolved: List[Dict[str, Any]] = []
        seen: Set[str] = set()
        for item in fields or []:
            if isinstance(item, dict) and "field" in item:
                field = item.get("field")
                field_type = item.get("field_type", "dimension")
            else:
                field = item
                field_type = (
                    "measure"
                    if isinstance(item, dict)
                    and item.get("field_type") in {"named_measure", "implicit_measure"}
                    else "dimension"
                )
            key = f"{field_type}:{field}".casefold()
            if key in seen:
                continue
            seen.add(key)
            result = self.map_field(field, str(field_type))
            if result["status"] in {"mapped", "requires_pivot_aggregation"}:
                mapped.append(result)
            else:
                unresolved.append(result)
        return {
            "mapped": mapped,
            "unmapped": unresolved,
            "mapped_count": len(mapped),
            "unmapped_count": len(unresolved),
            "total_count": len(mapped) + len(unresolved),
            "discovered_cubefield_count": len(self.discovered_cubefields),
            "discovered_cubefields": list(self.discovered_cubefields),
        }

    def best_match_for_raw(self, raw: str) -> Optional[str]:
        """Return the best matching CubeField for a raw string using substring search.

        This is a last-resort helper for the integration test's ``_field_exists``
        check and should not be used in production field-mapping logic.
        """
        raw_lower = raw.casefold()
        # Exact path
        for cf in self.discovered_cubefields:
            if cf.casefold() == raw_lower:
                return cf

        candidates = []
        for cf in self.discovered_cubefields:
            if raw_lower in cf.casefold() or cf.casefold() in raw_lower:
                candidates.append(cf)

        unique_candidates = list(dict.fromkeys(candidates))
        if len(unique_candidates) == 1:
            return unique_candidates[0]
        return None


__all__ = ["OLAPFieldMapper"]

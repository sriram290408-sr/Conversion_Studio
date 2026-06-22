"""Excel COM orchestration for a copied Power BI-connected workbook."""

from __future__ import annotations

import logging
import os
import re
import shutil
import threading
from collections import defaultdict
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple

import pythoncom
import win32com.client as win32

from connection_validator import validate_workbook_connection
from cube_formula_builder import build_cube_value_formula
from excel_visual_renderer import render_visual_to_dashboard
from filter_engine import FilterEngine
from olap_field_mapper import OLAPFieldMapper
from pivot_factory import PivotFactory
from refresh_manager import refresh_and_calculate_workbook

logger = logging.getLogger("excel_com_renderer")
excel_com_lock = threading.Lock()

FORMULA_ERRORS = {
    "#N/A",
    "#VALUE!",
    "#NAME?",
    "#REF!",
    "#NUM!",
    "#DIV/0!",
    "#GETTING_DATA",
}


def _safe_sheet_name(value: str, used: Set[str]) -> str:
    base = re.sub(r"[:\\/?*\[\]]", "", str(value or "Dashboard")).strip() or "Dashboard"
    base = base[:31]
    candidate = base
    suffix = 2
    while candidate.casefold() in used:
        tail = f" ({suffix})"
        candidate = f"{base[:31-len(tail)]}{tail}"
        suffix += 1
    used.add(candidate.casefold())
    return candidate


def _measure_reference(measure: Any) -> str:
    if isinstance(measure, dict):
        return str(
            measure.get("measure_name")
            or measure.get("canonical_reference")
            or measure.get("raw_reference")
            or measure.get("display_name")
            or ""
        ).strip()
    return str(measure or "").strip()


def _measure_key(measure: Any) -> str:
    return _measure_reference(measure).casefold()


def _extract_binding_fields(binding: Any) -> List[Dict[str, Any]]:
    records: List[Dict[str, Any]] = []
    for role in ("rows", "columns", "legend"):
        for value in getattr(binding, role, None) or []:
            if value:
                records.append(
                    {"field": value, "field_type": "dimension", "role": role}
                )
    for value in getattr(binding, "measures", None) or []:
        if value:
            records.append({"field": value, "field_type": "measure", "role": "values"})
    slicer_field = getattr(binding, "slicer_field", None)
    if slicer_field:
        records.append(
            {"field": slicer_field, "field_type": "dimension", "role": "slicer"}
        )
    for item in getattr(binding, "filters", None) or []:
        if isinstance(item, dict) and item.get("field"):
            records.append(
                {"field": item["field"], "field_type": "dimension", "role": "filter"}
            )
    return records


def _mapping_key(record: Dict[str, Any]) -> str:
    field = record.get("field")
    if isinstance(field, dict):
        value = (
            field.get("measure_name")
            or field.get("canonical_reference")
            or field.get("raw_reference")
            or str(field)
        )
    else:
        value = field
    return f"{record.get('field_type')}:{value}".casefold()


def _build_field_mapping(
    mapper: OLAPFieldMapper, bindings: List[Any]
) -> Dict[str, Any]:
    mapped: List[Dict[str, Any]] = []
    unmapped: List[Dict[str, Any]] = []
    seen: Set[str] = set()
    for binding in bindings:
        for record in _extract_binding_fields(binding):
            key = _mapping_key(record)
            if key in seen:
                continue
            seen.add(key)
            result = mapper.map_field(record["field"], record["field_type"])
            output = {
                **result,
                "role": record["role"],
                "visual_id": str(getattr(binding, "visual_id", "") or ""),
                "page_name": str(getattr(binding, "page_name", "") or ""),
                "visual_title": str(getattr(binding, "title", "") or ""),
            }
            if output.get("status") in {"mapped", "requires_pivot_aggregation"}:
                mapped.append(output)
            else:
                unmapped.append(output)
    return {
        "mapped": mapped,
        "unmapped": unmapped,
        "mapped_count": len(mapped),
        "unmapped_count": len(unmapped),
        "total_count": len(mapped) + len(unmapped),
        "discovered_cubefields": list(mapper.discovered_cubefields),
    }


def _semantic_match_score(field_mapping: Dict[str, Any]) -> Dict[str, Any]:
    """Compute a weighted semantic-match score from the field mapping result.

    Category weights follow validate_semantic_model_compatibility so both code
    paths agree: 40 % named_measure / 40 % dimension / 20 % implicit_measure.
    Active-weight normalisation means a model with no measures still scores
    purely on its dimensions and vice-versa.

    "unknown"-typed fields are excluded from scoring (they cannot be validated
    without a CubeField hit).  A score of 1.0 is returned when no eligible
    fields exist at all.
    """
    all_records = list(field_mapping.get("mapped", [])) + list(
        field_mapping.get("unmapped", [])
    )

    weights = {"named_measure": 0.40, "dimension": 0.40, "implicit_measure": 0.20}
    eligible: Dict[str, int] = {c: 0 for c in weights}
    matched: Dict[str, int] = {c: 0 for c in weights}

    seen_required: Set[str] = set()
    seen_matched: Set[str] = set()

    for item in all_records:
        ftype = str(item.get("field_type") or "").casefold()
        if ftype not in {"named_measure", "implicit_measure", "dimension", "hierarchy"}:
            # unknown / unclassified – cannot score reliably
            continue

        cat = (
            "named_measure" if ftype == "named_measure"
            else "implicit_measure" if ftype == "implicit_measure"
            else "dimension"   # dimension + hierarchy both use the dimension bucket
        )

        dedup_key = (
            f"{cat}:{str(item.get('normalized_reference') or item.get('pbix_field') or '').casefold()}"
        )
        if dedup_key in seen_required:
            continue
        seen_required.add(dedup_key)
        eligible[cat] += 1

        if item.get("status") in {"mapped", "requires_pivot_aggregation"}:
            if dedup_key not in seen_matched:
                seen_matched.add(dedup_key)
                matched[cat] += 1

    total_eligible = sum(eligible.values())
    if total_eligible == 0:
        # No classifiable fields to test – no evidence of a mismatch.
        return {
            "score": 1.0,
            "measure_match_ratio": 1.0,
            "dimension_match_ratio": 1.0,
            "implicit_measure_match_ratio": 1.0,
            "required_measures": 0,
            "mapped_measures": 0,
            "required_dimensions": 0,
            "mapped_dimensions": 0,
            "eligible_by_category": eligible,
            "matched_by_category": matched,
        }

    cat_scores: Dict[str, float] = {}
    for cat in weights:
        if eligible[cat] > 0:
            cat_scores[cat] = matched[cat] / eligible[cat]
        else:
            cat_scores[cat] = 1.0  # not penalised when category absent

    active_weight_sum = sum(weights[c] for c in weights if eligible[c] > 0)
    if active_weight_sum > 0:
        score = sum(
            cat_scores[c] * weights[c]
            for c in weights
            if eligible[c] > 0
        ) / active_weight_sum
    else:
        score = 1.0

    score = round(score, 4)
    return {
        "score": score,
        "measure_match_ratio": round(cat_scores["named_measure"], 4),
        "dimension_match_ratio": round(cat_scores["dimension"], 4),
        "implicit_measure_match_ratio": round(cat_scores["implicit_measure"], 4),
        "required_measures": eligible["named_measure"],
        "mapped_measures": matched["named_measure"],
        "required_dimensions": eligible["dimension"],
        "mapped_dimensions": matched["dimension"],
        "eligible_by_category": dict(eligible),
        "matched_by_category": dict(matched),
    }


def _discover_mapper(workbook: Any) -> OLAPFieldMapper:
    names: Set[str] = set()
    for sheet_index in range(1, int(workbook.Worksheets.Count) + 1):
        sheet = workbook.Worksheets(sheet_index)
        try:
            pivots = sheet.PivotTables()
            count = int(pivots.Count)
        except Exception:
            continue
        for pivot_index in range(1, count + 1):
            try:
                pivot = pivots.Item(pivot_index)
                if not bool(pivot.PivotCache().OLAP):
                    continue
                fields = pivot.CubeFields
                for field_index in range(1, int(fields.Count) + 1):
                    names.add(str(fields.Item(field_index).Name))
            except Exception:
                continue
    mapper = OLAPFieldMapper()
    mapper.discover_fields_from_names(sorted(names))
    return mapper


def _required_kpi_measure_keys(bindings: Iterable[Any]) -> Set[str]:
    result: Set[str] = set()
    for binding in bindings:
        if str(getattr(binding, "binding_type", "")) != "cube_formula":
            continue
        measures = getattr(binding, "measures", None) or []
        if measures:
            result.add(_measure_key(measures[0]))
    return result


def _materialize_formula_chunks(
    formula_chunks: Optional[List[Any]],
    required_keys: Set[str],
    mapper: OLAPFieldMapper,
    connection_name: str,
) -> Tuple[List[Dict[str, Any]], Dict[str, Dict[str, Any]], List[str]]:
    materialized: List[Dict[str, Any]] = []
    by_measure: Dict[str, Dict[str, Any]] = {}
    warnings: List[str] = []
    for original in formula_chunks or []:
        if not isinstance(original, dict):
            continue
        chunk = original
        measure_name = str(chunk.get("measure_name") or "").strip()
        if not measure_name or measure_name.casefold() not in required_keys:
            continue
        if chunk.get("conversion_status") not in {"prepared", "materialized"}:
            continue
        mapping = mapper.map_field(
            {
                "field_type": "named_measure",
                "measure_name": measure_name,
                "raw_reference": measure_name,
                "canonical_reference": measure_name,
            },
            "measure",
        )
        if mapping.get("status") != "mapped" or not mapping.get("excel_olap_field"):
            chunk.update(
                {
                    "mapping_status": mapping.get("status", "unmapped"),
                    "mapping_reason": mapping.get(
                        "reason", "Named measure mapping failed"
                    ),
                }
            )
            warnings.append(f"KPI named measure '{measure_name}' could not be mapped.")
            continue
        formula = build_cube_value_formula(
            connection_name,
            mapping["excel_olap_field"],
            [],
        )
        chunk.update(
            {
                "connection_name": connection_name,
                "cube_measure_path": mapping["excel_olap_field"],
                "cube_formula": formula,
                "excel_formula": formula,
                "conversion_status": "materialized",
                "mapping_status": "mapped",
                "mapping_confidence": mapping.get("confidence", 1.0),
                "required_tables": [],
                "required_hidden_sheets": [],
            }
        )
        materialized.append(chunk)
        by_measure[measure_name.casefold()] = chunk
    return materialized, by_measure, warnings


def _formula_value_is_valid(cell: Any) -> bool:
    try:
        display = str(cell.Text or "").strip()
    except Exception:
        display = ""
    if not display or display.upper() in FORMULA_ERRORS or display.startswith("#"):
        return False
    try:
        value = cell.Value
    except Exception:
        value = None
    return value not in (None, "")


def _verify_saved_workbook(
    output_path: str, required_formula_cells: List[Tuple[str, str]]
) -> Dict[str, Any]:
    result: Dict[str, Any] = {
        "file_exists": os.path.exists(output_path),
        "workbook_opened": False,
        "connection_count": 0,
        "pivot_cache_count": 0,
        "connected_pivot_count": 0,
        "cube_field_accessible": False,
        "cube_formula_valid": not required_formula_cells,
        "validated_formula_cells": [],
        "repair_detected": False,
        "verification_passed": False,
        "warnings": [],
        "errors": [],
    }
    if not result["file_exists"]:
        result["errors"].append("Output workbook does not exist.")
        return result

    app = None
    workbook = None
    try:
        app = win32.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.AskToUpdateLinks = False
        app.EnableEvents = False
        app.ScreenUpdating = False
        workbook = app.Workbooks.Open(output_path, ReadOnly=True, UpdateLinks=0)
        result["workbook_opened"] = True
        result["connection_count"] = int(workbook.Connections.Count)
        result["pivot_cache_count"] = int(workbook.PivotCaches().Count)

        connected = 0
        cube_accessible = False
        for sheet_index in range(1, int(workbook.Worksheets.Count) + 1):
            sheet = workbook.Worksheets(sheet_index)
            try:
                pivots = sheet.PivotTables()
                count = int(pivots.Count)
            except Exception:
                continue
            for pivot_index in range(1, count + 1):
                try:
                    pivot = pivots.Item(pivot_index)
                    if not bool(pivot.PivotCache().OLAP):
                        continue
                    fields = pivot.CubeFields
                    if int(fields.Count) > 0:
                        connected += 1
                        cube_accessible = True
                except Exception:
                    continue
        result["connected_pivot_count"] = connected
        result["cube_field_accessible"] = cube_accessible

        valid_formula_count = 0
        for sheet_name, address in required_formula_cells:
            record = {
                "sheet": sheet_name,
                "address": address,
                "valid": False,
                "display": "",
            }
            try:
                cell = workbook.Worksheets(sheet_name).Range(address)
                record["display"] = str(cell.Text or "")
                record["valid"] = _formula_value_is_valid(cell)
                if record["valid"]:
                    valid_formula_count += 1
            except Exception as exc:
                record["error"] = str(exc)
            result["validated_formula_cells"].append(record)
        if required_formula_cells:
            result["cube_formula_valid"] = valid_formula_count == len(
                required_formula_cells
            )

        result["verification_passed"] = bool(
            result["workbook_opened"]
            and result["connection_count"] > 0
            and result["pivot_cache_count"] > 0
            and result["connected_pivot_count"] > 0
            and result["cube_field_accessible"]
            and result["cube_formula_valid"]
        )
    except Exception as exc:
        result["errors"].append(str(exc))
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
    return result


class ExcelCOMRenderer:
    def __init__(self, template_path: str, output_path: str) -> None:
        self.template_path = os.path.abspath(template_path)
        self.output_path = os.path.abspath(output_path)
        self.excel_app: Any = None
        self.workbook: Any = None
        self.lock_acquired = False
        self.com_initialized = False

    def start_excel(self) -> Any:
        pythoncom.CoInitialize()
        self.com_initialized = True
        app = win32.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.AskToUpdateLinks = False
        app.EnableEvents = False
        app.ScreenUpdating = False
        self.excel_app = app
        return app

    def _close_rendering_excel(self) -> None:
        if self.workbook is not None:
            try:
                self.workbook.Close(SaveChanges=False)
            except Exception:
                pass
            self.workbook = None
        if self.excel_app is not None:
            try:
                self.excel_app.Quit()
            except Exception:
                pass
            self.excel_app = None

    def run_workflow(
        self,
        visual_bindings: List[Any],
        formula_chunks: Optional[List[Any]] = None,
        page_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        self.lock_acquired = excel_com_lock.acquire(timeout=120)
        if not self.lock_acquired:
            raise RuntimeError(
                "Excel COM resource lock timeout. Another conversion is using Excel."
            )

        validation = None
        refresh = None
        verification: Dict[str, Any] = {}
        field_mapping: Dict[str, Any] = {}
        render_results: List[Dict[str, Any]] = []
        workflow_log: List[str] = []
        page_sheet_names: Dict[str, str] = {}
        slicers_created = 0
        slicers_failed = 0
        required_formula_cells: List[Tuple[str, str]] = []
        materialized_chunks: List[Dict[str, Any]] = []
        semantic_score: Dict[str, Any] = {}

        try:
            os.makedirs(os.path.dirname(self.output_path) or ".", exist_ok=True)
            if os.path.normcase(self.template_path) != os.path.normcase(
                self.output_path
            ):
                shutil.copy2(self.template_path, self.output_path)

            self.start_excel()
            self.workbook = self.excel_app.Workbooks.Open(
                self.output_path, UpdateLinks=0
            )
            validation = validate_workbook_connection(self.excel_app, self.workbook)
            if not validation.usable_for_live_conversion:
                raise RuntimeError(
                    "LIVE_CONNECTION_INVALID: " + "; ".join(validation.errors)
                )

            connection_name = str(validation.selected_connection_name or "").strip()
            if not connection_name:
                raise RuntimeError(
                    "No validated Power BI/OLAP connection name was selected."
                )

            mapper = _discover_mapper(self.workbook)
            field_mapping = _build_field_mapping(mapper, visual_bindings)
            semantic_score = _semantic_match_score(field_mapping)
            validation.semantic_model_match_score = semantic_score["score"]
            validation.semantic_model_match = semantic_score["score"] >= 0.50
            validation.cube_field_count = len(mapper.discovered_cubefields)
            validation.cube_fields_discovered = validation.cube_field_count > 0
            validation.usable_for_live_conversion = bool(
                validation.usable_for_live_conversion
                and validation.cube_field_count > 0
                and semantic_score["score"] >= 0.30
            )
            if semantic_score["score"] < 0.30:
                raise RuntimeError(
                    f"SEMANTIC_MODEL_MISMATCH: match score {semantic_score['score']:.2f} is below 0.30 "
                    f"(measures={semantic_score.get('measure_match_ratio', 'n/a'):.0%}, "
                    f"dimensions={semantic_score.get('dimension_match_ratio', 'n/a'):.0%}). "
                    "The Analyze-in-Excel workbook appears to be from a different Power BI semantic model."
                )
            if semantic_score["score"] < 0.65:
                workflow_log.append(
                    f"Warning: semantic-model score is {semantic_score['score']:.2f} – partial mapping; "
                    f"some visuals may use fallback CubeField assignments."
                )

            required_kpis = _required_kpi_measure_keys(visual_bindings)
            materialized_chunks, materialized_by_measure, formula_warnings = (
                _materialize_formula_chunks(
                    formula_chunks,
                    required_kpis,
                    mapper,
                    connection_name,
                )
            )
            workflow_log.extend(f"Warning: {item}" for item in formula_warnings)
            field_mapping["materialized_formula_chunks"] = materialized_chunks

            factory = PivotFactory()
            filter_engine = FilterEngine()
            bindings_by_page: Dict[str, List[Any]] = defaultdict(list)
            for binding in visual_bindings:
                bindings_by_page[
                    str(getattr(binding, "page_name", None) or page_name or "Dashboard")
                ].append(binding)

            used_names = {
                str(self.workbook.Worksheets(index).Name).casefold()
                for index in range(1, int(self.workbook.Worksheets.Count) + 1)
            }

            for current_page, page_bindings in bindings_by_page.items():
                sheet_name = _safe_sheet_name(current_page, used_names)
                page_sheet_names[current_page] = sheet_name
                try:
                    dashboard = self.workbook.Worksheets(sheet_name)
                    dashboard.Cells.Clear()
                    for index in range(int(dashboard.ChartObjects().Count), 0, -1):
                        dashboard.ChartObjects(index).Delete()
                    for index in range(int(dashboard.Shapes.Count), 0, -1):
                        dashboard.Shapes.Item(index).Delete()
                except Exception:
                    dashboard = self.workbook.Worksheets.Add()
                    dashboard.Name = sheet_name
                try:
                    dashboard.Activate()
                    self.excel_app.ActiveWindow.DisplayGridlines = False
                except Exception:
                    pass

                pivot_by_visual: Dict[str, Dict[str, Any]] = {}
                page_pivots: List[Any] = []
                for binding in page_bindings:
                    if str(getattr(binding, "binding_type", "")) != "connected_pivot":
                        continue
                    info = factory.get_or_create_pivot(
                        self.excel_app,
                        self.workbook,
                        connection_name,
                        list(getattr(binding, "rows", None) or []),
                        list(getattr(binding, "columns", None) or []),
                        list(getattr(binding, "measures", None) or []),
                        list(getattr(binding, "legend", None) or []),
                        str(getattr(binding, "visual_type", "") or ""),
                        mapper,
                        filters=list(getattr(binding, "filters", None) or []),
                    )
                    pivot_by_visual[str(binding.visual_id)] = info
                    if info.get("pivot_table") is not None:
                        page_pivots.append(info["pivot_table"])
                        binding.pivot_name = info.get("pivot_name")
                        binding.source_sheet = info.get("sheet_name")
                        binding.pivot_signature = info.get("signature")
                    binding.warnings.extend(info.get("warnings") or [])
                    if info.get("error"):
                        binding.errors.append(str(info["error"]))

                slicer_refs: Dict[str, Dict[str, Any]] = {}
                for binding in page_bindings:
                    if str(getattr(binding, "binding_type", "")) != "slicer":
                        continue
                    field = getattr(binding, "slicer_field", None)
                    if not field:
                        binding.errors.append("No slicer field was detected.")
                        slicers_failed += 1
                        continue
                    source = factory.create_filter_source_pivot(
                        self.excel_app,
                        self.workbook,
                        connection_name,
                        str(field),
                        mapper,
                    )
                    if source.get("pivot_table") is None:
                        binding.errors.append(
                            source.get("error") or "Slicer source PivotTable failed."
                        )
                        slicers_failed += 1
                        continue
                    layout = getattr(binding, "layout", None) or {}
                    target = dashboard.Cells(
                        max(1, int(layout.get("row", 5))),
                        max(1, int(layout.get("col", 2))),
                    )
                    slicer = filter_engine.create_slicer(
                        workbook=self.workbook,
                        dashboard_sheet=dashboard,
                        source_pivot=source["pivot_table"],
                        field=str(field),
                        field_mapper=mapper,
                        target_pivots=page_pivots,
                        title=str(getattr(binding, "title", None) or field),
                        left=float(target.Left),
                        top=float(target.Top),
                        width=max(float(layout.get("col_span", 3)) * 75.0, 120.0),
                        height=max(float(layout.get("row_span", 8)) * 20.0, 90.0),
                    )
                    if slicer.get("status") == "success":
                        slicers_created += 1
                        slicer_refs[str(field).casefold()] = slicer
                        binding.render_status = "success"
                    else:
                        slicers_failed += 1
                        binding.errors.append(
                            slicer.get("error") or "Slicer creation failed."
                        )

                for binding in page_bindings:
                    if str(getattr(binding, "binding_type", "")) == "slicer":
                        continue
                    cube_refs: List[Dict[str, Any]] = []
                    for item in getattr(binding, "filters", None) or []:
                        if isinstance(item, dict):
                            key = str(item.get("field") or "").casefold()
                            if key in slicer_refs:
                                cube_refs.append(slicer_refs[key])
                    if not cube_refs:
                        cube_refs = list(slicer_refs.values())

                    render = render_visual_to_dashboard(
                        self.excel_app,
                        self.workbook,
                        dashboard,
                        binding,
                        pivot_by_visual.get(str(binding.visual_id)),
                        mapper,
                        theme={
                            "card_color": "FFFFFF",
                            "accent_color": "2563EB",
                            "text_color": "111827",
                            "muted_text_color": "64748B",
                            "header_color": "0F172A",
                        },
                        cube_filter_refs=cube_refs,
                        connection_name=connection_name,
                        materialized_formulas=materialized_by_measure,
                    )
                    render_results.append(render)
                    binding.render_status = str(render.get("status") or "failed")
                    if render.get("error"):
                        binding.errors.append(str(render["error"]))
                    if render.get("cube_formula") and render.get("value_cell"):
                        binding.cube_formula = str(render["cube_formula"])
                        binding.value_cell = str(render["value_cell"])
                        required_formula_cells.append(
                            (sheet_name, str(render["value_cell"]))
                        )

            refresh = refresh_and_calculate_workbook(self.excel_app, self.workbook)
            if refresh.timeout:
                raise RuntimeError(
                    "REFRESH_TIMEOUT: Excel refresh exceeded the configured timeout."
                )

            for index in range(1, int(self.workbook.Worksheets.Count) + 1):
                sheet = self.workbook.Worksheets(index)
                try:
                    if str(sheet.Name).startswith("_PBI_") or "Template" in str(
                        sheet.Name
                    ):
                        sheet.Visible = 0
                except Exception:
                    pass

            self.workbook.Save()
            self.workbook.Close(SaveChanges=True)
            self.workbook = None
            # Quit the rendering process before opening a fresh verification process.
            self.excel_app.Quit()
            self.excel_app = None

            verification = _verify_saved_workbook(
                self.output_path, required_formula_cells
            )
            if not verification.get("verification_passed"):
                raise RuntimeError(
                    "OUTPUT_VERIFICATION_FAILED: "
                    + "; ".join(
                        verification.get("errors") or ["post-save checks failed"]
                    )
                )

            return {
                "status": (
                    "completed_with_warnings"
                    if any(item.get("status") != "success" for item in render_results)
                    else "completed"
                ),
                "conversion_mode": "live_semantic_model",
                "output_path": self.output_path,
                "connection_preserved": True,
                "validation": validation.model_dump() if validation is not None else {},
                "semantic_match_score": semantic_score.get("score", 0.0),
                "semantic_match_details": semantic_score,
                "refresh": refresh.model_dump() if refresh is not None else {},
                "verification_result": verification,
                "field_mapping": field_mapping,
                "discovered_cubefields": field_mapping.get("discovered_cubefields", []),
                "materialized_formula_chunks": materialized_chunks,
                "render_results": render_results,
                "visual_bindings": [
                    item.model_dump() if hasattr(item, "model_dump") else item
                    for item in visual_bindings
                ],
                "rendered_visuals_count": sum(
                    1 for item in render_results if item.get("status") == "success"
                ),
                "failed_visuals_count": sum(
                    1 for item in render_results if item.get("status") == "failed"
                ),
                "slicers_created": slicers_created,
                "slicers_failed": slicers_failed,
                "dashboard_pages": list(page_sheet_names.values()),
                "logs": workflow_log,
            }

        except Exception:
            logger.exception("Excel COM workflow failed")
            self._close_rendering_excel()
            raise
        finally:
            self._close_rendering_excel()
            if self.com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
                self.com_initialized = False
            if self.lock_acquired:
                excel_com_lock.release()
                self.lock_acquired = False


__all__ = ["ExcelCOMRenderer"]

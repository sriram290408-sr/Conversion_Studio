"""Interactive Excel COM connection service for Power BI live-conversion.

All public functions in this module MUST be called from the session's
dedicated COM thread (managed by COMSessionManager).  They are never called
directly from FastAPI request handlers.

Workflow
--------
1. ``launch_excel_for_connection``   — open Excel visibly, create blank workbook
2. ``detect_and_validate_connection``— inspect workbook, discover CubeFields,
                                       validate semantic model against PBIX/TMDL
3. ``build_live_dashboard``          — create PivotTables, charts, slicers, KPIs,
                                       refresh, save, verify
4. ``cancel_session_com``            — close session workbook/Excel cleanly
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
import re
import tempfile
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger("interactive_excel_connection")

DEFAULT_THEME = {
    "accent_color": "118DFF",
    "background_color": "FFFFFF",
    "text_color": "111827",
    "secondary_text_color": "6B7280",
    "muted_text_color": "6B7280",
    "card_background": "FFFFFF",
    "border_color": "E5E7EB",
    "remaining_color": "E6E6E6",
    "gridline_color": "D9D9D9",
}

# ---------------------------------------------------------------------------
# Excel COM constants
# ---------------------------------------------------------------------------
XL_HIDDEN = 0
XL_VISIBLE = -1  # xlSheetVisible
XL_ROW_FIELD = 1
XL_COLUMN_FIELD = 2
XL_PAGE_FIELD = 3
XL_DATA_FIELD = 4
XL_CHART_TYPE_COLUMN_CLUSTERED = 51
XL_CHART_TYPE_BAR_CLUSTERED = 57
XL_CHART_TYPE_LINE = 4
XL_CHART_TYPE_AREA = 1
XL_CHART_TYPE_PIE = 5
XL_CHART_TYPE_COLUMN_STACKED = 52
XL_CHART_TYPE_DOUGHNUT = -4120
XL_SHAPE_RECTANGLE = 1
XL_SHAPE_TEXTBOX = 17
XL_LOCATION_AS_OBJECT = 2  # xlLocationAsObject

# Chart-type map from Power BI visual type to Excel constant
_CHART_TYPE_MAP: Dict[str, int] = {
    "clusteredcolumnchart": XL_CHART_TYPE_COLUMN_CLUSTERED,
    "columnchart": XL_CHART_TYPE_COLUMN_CLUSTERED,
    "barchart": XL_CHART_TYPE_BAR_CLUSTERED,
    "clusteredbarchart": XL_CHART_TYPE_BAR_CLUSTERED,
    "linechart": XL_CHART_TYPE_LINE,
    "areachart": XL_CHART_TYPE_AREA,
    "piechart": XL_CHART_TYPE_PIE,
    "donutchart": XL_CHART_TYPE_PIE,
    "stackedcolumnchart": XL_CHART_TYPE_COLUMN_STACKED,
    "stackedbarchart": 58,
    "combochart": XL_CHART_TYPE_COLUMN_CLUSTERED,
}

# Visual types the renderer can handle; everything else → placeholder
_SUPPORTED_CHART_TYPES = frozenset(_CHART_TYPE_MAP.keys())
_CARD_TYPES = frozenset({"card", "multirowcard", "kpivisual", "kpi"})
_SLICER_TYPES = frozenset({"slicer"})
_GAUGE_TYPES = frozenset({"gauge", "gaugevisual"})
_TABLE_TYPES = frozenset({"tableex", "pivottable", "matrix"})
_UNSUPPORTED_TYPES = frozenset({"map", "filledmap", "arcgismap", "custom"})

# Sheet name max length in Excel
_MAX_SHEET_NAME = 31


# ---------------------------------------------------------------------------
# Import helpers (graceful local / package imports)
# ---------------------------------------------------------------------------
def _import_com_retry():
    try:
        from .com_retry import com_call
    except ImportError:
        from com_retry import com_call
    return com_call


def _import_session_cls():
    try:
        from .com_session_manager import COMSession
    except ImportError:
        from com_session_manager import COMSession
    return COMSession


def _import_field_tools():
    try:
        from .olap_field_mapper import OLAPFieldMapper
        from .field_normalizer import normalize_field
        from .connection_validator import validate_semantic_model_compatibility
    except ImportError:
        from olap_field_mapper import OLAPFieldMapper
        from field_normalizer import normalize_field
        from connection_validator import validate_semantic_model_compatibility
    return OLAPFieldMapper, normalize_field, validate_semantic_model_compatibility


def _import_pivot_factory():
    try:
        from .pivot_factory import PivotFactory
    except ImportError:
        from pivot_factory import PivotFactory
    return PivotFactory


def _import_filter_engine():
    try:
        from .filter_engine import FilterEngine
    except ImportError:
        from filter_engine import FilterEngine
    return FilterEngine


def _import_refresh_manager():
    try:
        from .refresh_manager import refresh_and_calculate_workbook
    except ImportError:
        from refresh_manager import refresh_and_calculate_workbook
    return refresh_and_calculate_workbook


def _import_binding_engine():
    try:
        from .binding_engine import create_all_visual_bindings
    except ImportError:
        from binding_engine import create_all_visual_bindings
    return create_all_visual_bindings


def _import_native_capture():
    try:
        from .powerbi_native_capture import PowerBINativeCaptureService
        from .visual_compatibility_registry import (
            POWERBI_NATIVE_CAPTURE,
            render_mode_for,
        )
    except ImportError:
        from powerbi_native_capture import PowerBINativeCaptureService
        from visual_compatibility_registry import (
            POWERBI_NATIVE_CAPTURE,
            render_mode_for,
        )
    return PowerBINativeCaptureService, POWERBI_NATIVE_CAPTURE, render_mode_for


def _import_visual_renderer():
    try:
        from .excel_visual_renderer import (
            configure_dashboard_canvas,
            finalize_live_visual_values,
            render_visual_to_dashboard,
        )
    except ImportError:
        from excel_visual_renderer import (
            configure_dashboard_canvas,
            finalize_live_visual_values,
            render_visual_to_dashboard,
        )
    return (
        render_visual_to_dashboard,
        finalize_live_visual_values,
        configure_dashboard_canvas,
    )


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------
def _safe_sheet_name(raw: str, existing: set, prefix: str = "") -> str:
    """Return a unique, Excel-safe worksheet name (<= 31 chars)."""
    cleaned = re.sub(r"[\\/*?:\[\]]", "_", str(raw or "Sheet"))
    candidate = (prefix + cleaned)[:_MAX_SHEET_NAME]
    if not candidate:
        candidate = "Sheet"
    base = candidate
    suffix = 2
    while candidate.casefold() in existing:
        tail = f"_{suffix}"
        candidate = base[: _MAX_SHEET_NAME - len(tail)] + tail
        suffix += 1
    return candidate


def _existing_sheet_names(workbook: Any) -> set:
    com_call = _import_com_retry()
    count = int(com_call(lambda: workbook.Worksheets.Count))
    return {
        str(com_call(lambda: workbook.Worksheets(i).Name)).casefold()
        for i in range(1, count + 1)
    }


def _clear_excel_recovery_files() -> None:
    """Remove Excel auto-recovery lock files that cause recovery dialogs."""
    import glob

    patterns = [
        os.path.join(tempfile.gettempdir(), "~$*.xl*"),
        os.path.join(
            os.path.expanduser("~"),
            "AppData",
            "Local",
            "Microsoft",
            "Excel",
            "**",
            "~$*.xl*",
        ),
    ]
    for pattern in patterns:
        for f in glob.glob(pattern, recursive=True):
            try:
                os.remove(f)
            except OSError:
                pass


# ---------------------------------------------------------------------------
# Step 1 — Launch Excel visibly and create a blank workbook
# ---------------------------------------------------------------------------
def launch_excel_for_connection(session: Any) -> None:
    """Open Excel (Visible=True) and create a blank workbook.

    Updates session state to ``waiting_for_user_connection`` on success
    or ``error`` on failure.  Must be called on the session's COM thread.
    """
    com_call = _import_com_retry()

    session.update_state("excel_launching")
    logger.info("Session %s: launching Excel.", session.session_id)

    _clear_excel_recovery_files()

    try:
        import win32com.client as win32  # type: ignore[import]
    except ImportError as exc:
        session.errors.append(f"pywin32 not available: {exc}")
        session.update_state("error")
        return

    excel = None
    workbook = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        com_call(lambda: setattr(excel, "Visible", True), label="Visible=True")
        com_call(lambda: setattr(excel, "DisplayAlerts", True), label="DisplayAlerts")
        com_call(
            lambda: setattr(excel, "AskToUpdateLinks", False), label="AskToUpdateLinks"
        )
        com_call(lambda: setattr(excel, "EnableEvents", False), label="EnableEvents")

        workbook = com_call(lambda: excel.Workbooks.Add(), label="Workbooks.Add")

        # Save the blank workbook so it has a real path.
        out_dir = str(Path(session.output_path).parent)
        wb_path = str(Path(out_dir) / f"_live_session_{session.session_id[:8]}.xlsx")
        com_call(
            lambda: workbook.SaveAs(wb_path, 51),  # 51 = xlOpenXMLWorkbook
            label="SaveAs blank",
        )

        session._com_excel = excel
        session._com_workbook = workbook
        session._com_workbook_path = wb_path

        session.update_state("waiting_for_user_connection")
        logger.info(
            "Session %s: Excel open, workbook at %s. Waiting for user.",
            session.session_id,
            wb_path,
        )

    except Exception as exc:
        logger.exception("Session %s: Excel launch failed.", session.session_id)
        session.errors.append(f"Excel launch failed: {exc}")
        # Best-effort cleanup
        _quit_excel_safely(excel, workbook, session)
        session.update_state("live_conversion_failed")


# ---------------------------------------------------------------------------
# Step 2 — Detect the Power BI connection created by the user
# ---------------------------------------------------------------------------
def _find_olap_pivot_and_connection(
    workbook: Any,
) -> Tuple[Optional[Any], Optional[Any], Optional[Any]]:
    """Return (pivot_cache, pivot_table, workbook_connection) for the first
    connected OLAP PivotTable.  If a workbook connection exists but no
    PivotTable has been inserted yet, returns (None, None, wb_conn) so callers
    can still discover CubeFields via the connection alone."""
    com_call = _import_com_retry()

    wb_conn_count = int(com_call(lambda: workbook.Connections.Count))
    pc_count = int(com_call(lambda: workbook.PivotCaches().Count))

    logger.info(
        "detect: %d workbook connection(s), %d PivotCache(s)", wb_conn_count, pc_count
    )

    if wb_conn_count == 0:
        return None, None, None

    if pc_count == 0:
        # No pivot yet — return the first connection so we can inspect CubeFields
        wb_conn = com_call(lambda: workbook.Connections.Item(1))
        return None, None, wb_conn

    ws_count = int(com_call(lambda: workbook.Worksheets.Count))
    for ws_idx in range(1, ws_count + 1):
        ws = com_call(lambda: workbook.Worksheets(ws_idx))
        try:
            pt_count = int(com_call(lambda: ws.PivotTables().Count))
        except Exception:
            continue
        for pt_idx in range(1, pt_count + 1):
            try:
                pt = com_call(lambda: ws.PivotTables(pt_idx))
                cache = com_call(lambda: pt.PivotCache())
                if not bool(com_call(lambda: cache.OLAP)):
                    continue

                # Find the matching workbook connection
                wb_conn = None
                for ci in range(1, wb_conn_count + 1):
                    try:
                        conn = com_call(lambda: workbook.Connections.Item(ci))
                        conn_name = str(com_call(lambda: conn.Name) or "")
                        cache_conn = str(com_call(lambda: cache.Connection) or "")
                        if conn_name and (
                            conn_name.casefold() in cache_conn.casefold()
                            or cache_conn.casefold() in conn_name.casefold()
                        ):
                            wb_conn = conn
                            break
                    except Exception:
                        continue

                if wb_conn is None and wb_conn_count > 0:
                    # Fall back to the first workbook connection
                    wb_conn = com_call(lambda: workbook.Connections.Item(1))

                return cache, pt, wb_conn
            except Exception:
                continue

    return None, None, None


def _discover_cube_fields(pivot_table: Any) -> List[Dict[str, Any]]:
    """Return all CubeField unique names and metadata from a PivotTable."""
    com_call = _import_com_retry()
    fields: List[Dict[str, Any]] = []
    seen: set = set()

    try:
        count = int(com_call(lambda: pivot_table.CubeFields.Count))
    except Exception:
        return fields

    for i in range(1, count + 1):
        try:
            cf = com_call(lambda: pivot_table.CubeFields(i))
            name = str(com_call(lambda: cf.Name) or "").strip()
            if not name or name.casefold() in seen:
                continue
            seen.add(name.casefold())

            try:
                caption = str(com_call(lambda: cf.Caption) or "")
            except Exception:
                caption = ""
            try:
                is_measure = bool(com_call(lambda: cf.IsMeasure))
            except Exception:
                is_measure = name.casefold().startswith("[measures]")
            try:
                cf_type = int(com_call(lambda: cf.CubeFieldType))
            except Exception:
                cf_type = 0

            fields.append(
                {
                    "name": name,
                    "caption": caption,
                    "is_measure": is_measure,
                    "cube_field_type": cf_type,
                }
            )
        except Exception:
            continue

    logger.info("Discovered %d CubeField(s).", len(fields))
    return fields


# ---------------------------------------------------------------------------
# Step 2 public entry-point — called via session.dispatch()
# ---------------------------------------------------------------------------
def detect_and_validate_connection(session: Any) -> Dict[str, Any]:
    """Inspect the workbook for an OLAP connection and validate it.

    Returns a result dict.  Updates session state to:
    - ``connection_not_detected``  — no OLAP connection or PivotTable found
    - ``semantic_model_mismatch``  — score < 0.50
    - ``connection_detected``      — pass (caller should then call build)
    Must be called on the session's COM thread.
    """
    com_call = _import_com_retry()
    OLAPFieldMapper, normalize_field, validate_semantic_compat = _import_field_tools()
    create_all_visual_bindings = _import_binding_engine()

    session.update_state("detecting_connection")
    excel = session._com_excel
    workbook = session._com_workbook

    if excel is None or workbook is None:
        session.errors.append("Excel or workbook not initialised.")
        session.update_state("live_conversion_failed")
        return {"state": "live_conversion_failed", "message": "Excel not initialised."}

    # --- Detect connection --------------------------------------------------
    pivot_cache, pivot_table, wb_conn = _find_olap_pivot_and_connection(workbook)

    if wb_conn is None or pivot_cache is None or pivot_table is None:
        session.update_state("connection_not_detected")
        return {
            "state": "connection_not_detected",
            "message": (
                "No connected Power BI PivotTable was found in the session workbook. "
                "In Excel use Insert > PivotTable > From Power BI, insert one empty "
                "PivotTable, then click 'Connection Completed' again."
            ),
        }

    conn_name = ""
    try:
        conn_name = str(com_call(lambda: wb_conn.Name) or "")
    except Exception:
        pass

    session.selected_connection_name = conn_name
    logger.info("Selected connection: %r", conn_name)
    session.update_state("connection_detected")

    # --- Discover CubeFields ------------------------------------------------
    cube_fields: List[Dict[str, Any]] = []
    if pivot_table is not None:
        cube_fields = _discover_cube_fields(pivot_table)
    cube_field_names = [cf["name"] for cf in cube_fields]
    session.cube_field_count = len(cube_field_names)

    if not cube_field_names:
        message = (
            "No CubeFields discovered from the connected PivotTable. "
            "Wait for Excel to finish loading the semantic model, then try again."
        )
        logger.warning(message)
        session.errors.append(message)
        session.update_state("connection_not_detected")
        return {
            "state": "connection_not_detected",
            "message": message,
            "cube_field_count": 0,
            "selected_connection_name": conn_name,
        }

    # --- Build OLAPFieldMapper from discovered CubeFields -------------------
    mapper = OLAPFieldMapper()
    mapper.discover_fields_from_names(cube_field_names)
    measures_count = sum(1 for cf in cube_fields if cf["is_measure"])
    dims_count = len(cube_field_names) - measures_count
    logger.info(
        "OLAPFieldMapper: %d measures, %d dimensions.", measures_count, dims_count
    )
    logger.info(
        "Discovered CubeFields: %s",
        [
            {
                "name": item.get("name"),
                "caption": item.get("caption"),
                "is_measure": item.get("is_measure"),
            }
            for item in cube_fields
        ],
    )

    # --- Create visual bindings from PBIX metadata --------------------------
    final_chunks = session.metadata or {}
    try:
        visual_bindings = create_all_visual_bindings(final_chunks)
    except Exception as exc:
        logger.warning(
            "create_all_visual_bindings failed: %s — using raw bindings.", exc
        )
        visual_bindings = session.visual_bindings or []

    session.visual_bindings = visual_bindings

    # --- Validate semantic model --------------------------------------------
    session.update_state("validating_semantic_model")
    try:
        tmdl_measures = final_chunks.get("measures", [])
        validation = validate_semantic_compat(
            bindings=visual_bindings,
            tmdl_measures=tmdl_measures,
            field_mapper=mapper,
        )
    except Exception as exc:
        logger.exception("Semantic-model validation failed")
        message = f"Semantic-model validation failed: {exc}"
        session.errors.append(message)
        session.semantic_match_score = 0.0
        session.update_state("live_conversion_failed")
        return {
            "state": "live_conversion_failed",
            "semantic_match_score": 0.0,
            "message": message,
        }

    score = validation.get("score", 1.0)
    session.semantic_match_score = round(float(score), 4)
    session.result["semantic_validation"] = {
        "score": session.semantic_match_score,
        "eligible_count": validation.get("eligible_count", 0),
        "matched_count": validation.get("matched_count", 0),
        "eligible_by_category": validation.get("eligible_by_category", {}),
        "matched_by_category": validation.get("matched_by_category", {}),
        "missing_measures": validation.get("missing_measures", []),
        "unmapped_fields": validation.get("unmapped_fields", []),
    }

    VALIDATION_THRESHOLD = 0.50
    if score < VALIDATION_THRESHOLD:
        unmapped_fields = validation.get("unmapped_fields", [])
        unmapped_names = [
            str(item.get("field") or item.get("pbix_field") or "")
            for item in unmapped_fields
            if item.get("field") or item.get("pbix_field")
        ]

        logger.warning(
            "Session %s: semantic score %.2f < %.2f — mismatch. "
            "matched=%s eligible=%s unmapped=%s",
            session.session_id,
            score,
            VALIDATION_THRESHOLD,
            validation.get("matched_count", 0),
            validation.get("eligible_count", 0),
            unmapped_names,
        )
        session.update_state("semantic_model_mismatch")
        return {
            "state": "semantic_model_mismatch",
            "semantic_match_score": session.semantic_match_score,
            "cube_field_count": session.cube_field_count,
            "selected_connection_name": conn_name,
            "message": (
                f"The selected semantic model could not be matched confidently "
                f"(score {score:.0%}). Review the unmapped fields before changing "
                f"the connected model."
            ),
            "eligible_count": validation.get("eligible_count", 0),
            "matched_count": validation.get("matched_count", 0),
            "eligible_by_category": validation.get("eligible_by_category", {}),
            "matched_by_category": validation.get("matched_by_category", {}),
            "missing_measures": validation.get("missing_measures", []),
            "unmapped_fields": unmapped_fields,
        }

    if score < 0.70:
        session.warnings.append(
            f"Semantic match score is {score:.0%} — some visuals may use "
            "fallback CubeField assignments."
        )

    logger.info(
        "Session %s: validation passed (score=%.2f, cubefields=%d).",
        session.session_id,
        score,
        session.cube_field_count,
    )

    # Store for build phase
    if not hasattr(session, "runtime_objects"):
        session.runtime_objects = {}
    session.runtime_objects.update(
        {
            "olap_mapper": mapper,
            "olap_pivot_cache": pivot_cache,
            "template_pivot_table": pivot_table,
            "workbook_connection": wb_conn,
        }
    )
    session.result.update(
        {
            "conn_name": conn_name,
            "semantic_match_score": session.semantic_match_score,
            "olap_mapper_summary": {
                "cube_field_count": len(cube_fields),
                "measure_count": measures_count,
                "dimension_count": dims_count,
            },
        }
    )
    session.update_state("connection_detected")
    return {
        "state": "connection_detected",
        "semantic_match_score": session.semantic_match_score,
        "cube_field_count": session.cube_field_count,
        "selected_connection_name": conn_name,
        "eligible_count": validation.get("eligible_count", 0),
        "matched_count": validation.get("matched_count", 0),
    }


# ---------------------------------------------------------------------------
# Step 3 — Build the live dashboard
# ---------------------------------------------------------------------------
def _get_or_create_sheet(workbook: Any, name: str, visible: bool = False) -> Any:
    com_call = _import_com_retry()
    try:
        return com_call(lambda: workbook.Worksheets(name))
    except Exception:
        pass
    sheet = com_call(lambda: workbook.Worksheets.Add())
    com_call(lambda: setattr(sheet, "Name", name), label="sheet.Name")
    com_call(lambda: setattr(sheet, "Visible", XL_VISIBLE if visible else XL_HIDDEN))
    return sheet


def _delete_sheet_if_exists(workbook: Any, name: str) -> None:
    com_call = _import_com_retry()
    try:
        ws = com_call(lambda: workbook.Worksheets(name))
        com_call(lambda: ws.Delete(), label="ws.Delete")
    except Exception:
        pass


def _number(value: Any, default: float) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return float(default)


def _visual_layout_points(
    layout: Optional[Dict[str, Any]],
    *,
    fallback_left: float,
    fallback_top: float,
    fallback_width: float,
    fallback_height: float,
) -> Tuple[float, float, float, float]:
    """Convert PBIX/screenshot layout coordinates into Excel points.

    The binding engine may expose x/y/width/height directly or inside
    position/geometry/bounds.  Power BI canvas coordinates are normally close
    to pixels, so a 0.75 conversion produces stable Excel point placement.
    """
    raw = dict(layout or {})
    for key in ("position", "geometry", "bounds"):
        nested = raw.get(key)
        if isinstance(nested, dict):
            raw = {**nested, **raw}

    x = raw.get("x", raw.get("left"))
    y = raw.get("y", raw.get("top"))
    width = raw.get("width", raw.get("w"))
    height = raw.get("height", raw.get("h"))

    # Power BI layout is generally pixel based.  Excel object placement is in
    # points.  1 px ~= 0.75 points at 96 DPI.
    scale = _number(raw.get("excel_scale"), 0.75)

    left = _number(x, fallback_left / scale) * scale
    top = _number(y, fallback_top / scale) * scale
    out_width = max(70.0, _number(width, fallback_width / scale) * scale)
    out_height = max(45.0, _number(height, fallback_height / scale) * scale)

    return left, top, out_width, out_height


def _safe_shape_name(raw: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_]", "_", str(raw or "Visual"))
    return cleaned[:120] or "Visual"


def _format_kpi_value(value: Any, format_hint: str = "") -> str:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value or "")

    hint = str(format_hint or "").casefold()
    if "percent" in hint or "%" in hint:
        return f"{number:.1%}"
    if abs(number) >= 1_000_000:
        return f"{number / 1_000_000:.1f}M"
    if abs(number) >= 1_000:
        rounded = number / 1_000
        if abs(rounded - round(rounded)) < 0.05:
            return f"{round(rounded):.0f}K"
        return f"{rounded:.1f}K"
    if abs(number - round(number)) < 0.000001:
        return f"{number:,.0f}"
    return f"{number:,.2f}"


def _write_live_reference_cell(
    dashboard_sheet: Any,
    source_cell: Any,
    helper_row: int,
    helper_col: int,
) -> Any:
    """Write a normal Excel formula linked to a hidden semantic PivotTable."""
    com_call = _import_com_retry()

    source_sheet = com_call(
        lambda: source_cell.Parent,
        label="visual source worksheet",
    )
    source_sheet_name = str(
        com_call(lambda: source_sheet.Name, label="visual source worksheet name") or ""
    )
    source_address = str(
        com_call(lambda: source_cell.Address, label="visual source address") or ""
    )

    helper_cell = com_call(
        lambda: dashboard_sheet.Cells.Item(helper_row, helper_col),
        label="visual helper cell",
    )
    formula = _excel_sheet_formula_reference(source_sheet_name, source_address)
    try:
        com_call(
            lambda: setattr(helper_cell, "Formula", formula),
            label="visual helper formula",
        )
    except Exception:
        com_call(
            lambda: setattr(helper_cell, "Formula2", formula),
            label="visual helper formula2",
        )

    return helper_cell


def _create_kpi_card(
    dashboard_sheet: Any,
    pivot_table: Any,
    title: str,
    layout: Dict[str, Any],
    visual_id: str,
    helper_row: int,
    helper_col: int,
) -> bool:
    """Create a screenshot-positioned KPI card linked by an Excel formula."""
    com_call = _import_com_retry()

    try:
        data_range = com_call(lambda: pivot_table.DataBodyRange)
        source_cell = com_call(
            lambda: data_range.Cells.Item(1, 1),
            label="KPI PivotTable value",
        )
    except Exception:
        return False

    helper_cell = _write_live_reference_cell(
        dashboard_sheet,
        source_cell,
        helper_row,
        helper_col,
    )

    left, top, width, height = _visual_layout_points(
        layout,
        fallback_left=760.0,
        fallback_top=50.0,
        fallback_width=165.0,
        fallback_height=105.0,
    )

    try:
        shape = com_call(
            lambda: dashboard_sheet.Shapes.AddShape(
                XL_SHAPE_RECTANGLE,
                left,
                top,
                width,
                height,
            ),
            label="KPI card shape",
        )
        try:
            shape.Name = _safe_shape_name(f"KPI_{visual_id}")
        except Exception:
            pass
        try:
            shape.Fill.ForeColor.RGB = 16777215
            shape.Line.ForeColor.RGB = 14474460
            shape.Line.Weight = 0.75
        except Exception:
            pass

        value_box = com_call(
            lambda: dashboard_sheet.Shapes.AddTextbox(
                1,
                left + 10,
                top + 12,
                max(50.0, width - 20),
                max(28.0, height * 0.48),
            ),
            label="KPI value textbox",
        )
        try:
            value_box.Name = _safe_shape_name(f"KPI_VALUE_{visual_id}")
        except Exception:
            pass

        # Link the textbox directly to the formula cell so it updates after refresh.
        helper_address = str(com_call(lambda: helper_cell.Address) or "")
        sheet_name = str(com_call(lambda: dashboard_sheet.Name) or "")
        shape_formula = _excel_sheet_formula_reference(sheet_name, helper_address)
        try:
            value_box.Formula = shape_formula
        except Exception:
            # Fallback for Excel builds that reject Shape.Formula.
            value = com_call(lambda: helper_cell.Value)
            value_box.TextFrame2.TextRange.Text = _format_kpi_value(value)

        try:
            value_box.Line.Visible = 0
            value_box.Fill.Visible = 0
            value_box.TextFrame2.TextRange.Font.Size = 22
            value_box.TextFrame2.TextRange.Font.Bold = -1
        except Exception:
            pass

        label_box = com_call(
            lambda: dashboard_sheet.Shapes.AddTextbox(
                1,
                left + 10,
                top + height * 0.60,
                max(50.0, width - 20),
                max(20.0, height * 0.25),
            ),
            label="KPI title textbox",
        )
        try:
            label_box.Name = _safe_shape_name(f"KPI_TITLE_{visual_id}")
            label_box.TextFrame2.TextRange.Text = title or "KPI"
            label_box.TextFrame2.TextRange.Font.Size = 10
            label_box.Line.Visible = 0
            label_box.Fill.Visible = 0
        except Exception:
            pass

        return True
    except Exception as exc:
        logger.warning("KPI card creation failed for %s: %s", title, exc)
        return False


def _create_gauge_visual(
    dashboard_sheet: Any,
    pivot_table: Any,
    title: str,
    layout: Dict[str, Any],
    visual_id: str,
    helper_row: int,
    helper_col: int,
    maximum: Optional[float] = None,
) -> bool:
    """Create a semi-doughnut gauge linked to a semantic-model PivotTable."""
    com_call = _import_com_retry()

    try:
        data_range = com_call(lambda: pivot_table.DataBodyRange)
        source_cell = com_call(
            lambda: data_range.Cells.Item(1, 1),
            label="Gauge PivotTable value",
        )
    except Exception:
        return False

    value_cell = _write_live_reference_cell(
        dashboard_sheet,
        source_cell,
        helper_row,
        helper_col,
    )

    # Gauge helper values: current, remaining, hidden lower half.
    max_cell = com_call(lambda: dashboard_sheet.Cells.Item(helper_row, helper_col + 1))
    remaining_cell = com_call(
        lambda: dashboard_sheet.Cells.Item(helper_row, helper_col + 2)
    )
    hidden_cell = com_call(
        lambda: dashboard_sheet.Cells.Item(helper_row, helper_col + 3)
    )

    current_value = _number(com_call(lambda: value_cell.Value), 0.0)
    max_value = _number(maximum, max(current_value * 2.0, 1.0))
    if max_value <= current_value:
        max_value = max(current_value * 2.0, 1.0)

    com_call(lambda: setattr(max_cell, "Value", max_value))
    com_call(
        lambda: setattr(
            remaining_cell,
            "Formula",
            f"=MAX(0,{max_cell.Address}-{value_cell.Address})",
        )
    )
    com_call(lambda: setattr(hidden_cell, "Formula", f"={max_cell.Address}"))

    left, top, width, height = _visual_layout_points(
        layout,
        fallback_left=385.0,
        fallback_top=40.0,
        fallback_width=315.0,
        fallback_height=145.0,
    )

    try:
        chart_obj = com_call(
            lambda: dashboard_sheet.ChartObjects().Add(
                left,
                top,
                width,
                height,
            ),
            label="Gauge chart object",
        )
        chart = com_call(lambda: chart_obj.Chart)
        chart.ChartType = XL_CHART_TYPE_DOUGHNUT

        source_range = com_call(
            lambda: dashboard_sheet.Range(
                value_cell,
                hidden_cell,
            )
        )
        chart.SetSourceData(source_range)
        chart.HasLegend = False
        chart.HasTitle = True
        chart.ChartTitle.Text = title or "Gauge"

        try:
            series = chart.SeriesCollection(1)
            series.DoughnutHoleSize = 62
            series.FirstSliceAngle = 270
            # Hide the lower-half slice.
            hidden_point = series.Points(3)
            hidden_point.Format.Fill.Visible = 0
            hidden_point.Format.Line.Visible = 0
        except Exception:
            pass

        # Center value textbox linked to the live helper cell.
        value_box = com_call(
            lambda: dashboard_sheet.Shapes.AddTextbox(
                1,
                left + width * 0.34,
                top + height * 0.52,
                width * 0.32,
                height * 0.24,
            )
        )
        helper_address = str(com_call(lambda: value_cell.Address) or "")
        sheet_name = str(com_call(lambda: dashboard_sheet.Name) or "")
        try:
            value_box.Formula = _excel_sheet_formula_reference(
                sheet_name,
                helper_address,
            )
        except Exception:
            value_box.TextFrame2.TextRange.Text = _format_kpi_value(
                com_call(lambda: value_cell.Value)
            )
        try:
            value_box.Line.Visible = 0
            value_box.Fill.Visible = 0
            value_box.TextFrame2.TextRange.Font.Size = 20
        except Exception:
            pass

        return True
    except Exception as exc:
        logger.warning("Gauge creation failed for %s: %s", title, exc)
        return False


def _apply_pivot_value_sort(pivot_table: Any) -> None:
    """Sort the first row field descending by the first data field."""
    try:
        row_field = pivot_table.RowFields(1)
        data_field = pivot_table.DataFields(1)
        row_field.AutoSort(2, data_field.Name)  # xlDescending = 2
    except Exception:
        pass


def _create_kpi_formula(
    sheet: Any,
    row: int,
    col: int,
    conn_name: str,
    measure_olap_field: str,
    title: str,
) -> bool:
    """Write a CUBEVALUE formula into sheet[row, col]. Returns True on success."""
    com_call = _import_com_retry()
    formula = f'=CUBEVALUE("{conn_name}","{measure_olap_field}")'
    cell = com_call(lambda: sheet.Cells(row, col))
    try:
        com_call(lambda: setattr(cell, "Formula", formula), label="Formula")
    except Exception:
        try:
            com_call(lambda: setattr(cell, "Formula2", formula), label="Formula2")
        except Exception as exc:
            logger.warning("CUBEVALUE write failed for %s: %s", measure_olap_field, exc)
            return False

    # Label cell above
    label_cell = com_call(lambda: sheet.Cells(row - 1, col))
    try:
        com_call(lambda: setattr(label_cell, "Value", title or measure_olap_field))
    except Exception:
        pass
    return True


def _excel_sheet_formula_reference(sheet_name: str, address: str) -> str:
    safe_name = str(sheet_name or "").replace("'", "''")
    return f"='{safe_name}'!{address}"


def _create_visible_kpi_from_pivot(
    pivot_table: Any,
    dashboard_sheet: Any,
    title: str,
    layout: Dict[str, Any],
    measure_count: int = 1,
) -> int:
    """Create visible KPI cells linked to a hidden semantic-model PivotTable."""
    com_call = _import_com_retry()

    row = max(2, int(layout.get("row") or 2))
    col = max(2, int(layout.get("col") or 2))
    created = 0

    try:
        data_range = com_call(lambda: pivot_table.DataBodyRange)
    except Exception:
        data_range = None

    if data_range is None:
        try:
            table_range = com_call(lambda: pivot_table.TableRange1)
            data_range = com_call(lambda: table_range.Offset(1, 0))
        except Exception:
            return 0

    for index in range(max(1, int(measure_count or 1))):
        try:
            source_cell = com_call(
                lambda i=index: data_range.Cells.Item(1, i + 1),
                label="KPI source cell",
            )

            source_sheet_obj = com_call(
                lambda: source_cell.Parent,
                label="KPI source worksheet",
            )
            source_sheet = str(
                com_call(
                    lambda: source_sheet_obj.Name,
                    label="KPI source worksheet name",
                )
                or ""
            )

            source_address = str(
                com_call(
                    lambda: source_cell.Address,
                    label="KPI source address",
                )
                or ""
            )

            label_cell = com_call(
                lambda i=index: dashboard_sheet.Cells.Item(
                    row,
                    col + i * 3,
                ),
                label="KPI label cell",
            )
            value_cell = com_call(
                lambda i=index: dashboard_sheet.Cells.Item(
                    row + 1,
                    col + i * 3,
                ),
                label="KPI value cell",
            )

            label_text = title if index == 0 else f"{title} {index + 1}"
            com_call(
                lambda: setattr(label_cell, "Value", label_text),
                label="KPI label value",
            )
            com_call(
                lambda: setattr(label_cell.Font, "Bold", True),
                label="KPI label bold",
            )
            com_call(
                lambda: setattr(label_cell.Font, "Size", 11),
                label="KPI label font size",
            )

            formula = _excel_sheet_formula_reference(
                source_sheet,
                source_address,
            )
            try:
                com_call(
                    lambda: setattr(value_cell, "Formula", formula),
                    label="KPI visible formula",
                )
            except Exception:
                com_call(
                    lambda: setattr(value_cell, "Formula2", formula),
                    label="KPI visible formula2",
                )

            com_call(
                lambda: setattr(value_cell.Font, "Bold", True),
                label="KPI value bold",
            )
            com_call(
                lambda: setattr(value_cell.Font, "Size", 20),
                label="KPI value font size",
            )
            try:
                com_call(
                    lambda: setattr(value_cell, "NumberFormat", "#,##0.00"),
                    label="KPI number format",
                )
            except Exception:
                pass

            try:
                column_obj = com_call(
                    lambda i=index: dashboard_sheet.Columns.Item(col + i * 3),
                    label="KPI column",
                )
                com_call(
                    lambda: setattr(column_obj, "ColumnWidth", 18),
                    label="KPI column width",
                )
            except Exception:
                pass

            created += 1
        except Exception as exc:
            logger.warning(
                "Visible KPI creation failed for %s measure %d: %s",
                title,
                index + 1,
                exc,
            )

    return created


def _create_pivot_chart(
    workbook: Any,
    pivot_table: Any,
    pivot_sheet: Any,
    dashboard_sheet: Any,
    chart_type: int,
    left: float,
    top: float,
    width: float,
    height: float,
    title: str = "",
) -> bool:
    """Embed a PivotChart on dashboard_sheet linked to pivot_table."""
    com_call = _import_com_retry()
    try:
        chart_obj = com_call(
            lambda: dashboard_sheet.ChartObjects().Add(left, top, width, height)
        )
        chart = com_call(lambda: chart_obj.Chart)
        # TableRange1 keeps the PivotChart connected to the semantic-model
        # PivotTable.  Disable grand totals so they are not plotted.
        try:
            pivot_table.RowGrand = False
            pivot_table.ColumnGrand = False
        except Exception:
            pass
        _apply_pivot_value_sort(pivot_table)

        com_call(lambda: chart.SetSourceData(pivot_table.TableRange1))
        com_call(lambda: setattr(chart, "ChartType", chart_type))
        if title:
            try:
                chart.HasTitle = True
                chart.ChartTitle.Text = title
            except Exception:
                pass
        try:
            chart.HasLegend = False
        except Exception:
            pass
        return True
    except Exception as exc:
        logger.warning("PivotChart creation failed: %s", exc)
        return False


def _create_slicer(
    workbook: Any,
    pivot_table: Any,
    field_olap_name: str,
    slicer_caption: str,
    dashboard_sheet: Any,
    left: float,
    top: float,
    width: float = 120.0,
    height: float = 200.0,
) -> bool:
    """Create a native OLAP slicer connected to pivot_table."""
    com_call = _import_com_retry()
    try:
        slicers = com_call(lambda: workbook.SlicerCaches)
        sc = com_call(lambda: slicers.Add(pivot_table, field_olap_name, slicer_caption))
        sl = com_call(
            lambda: sc.Slicers.Add(
                dashboard_sheet,
                Type=1,  # xlSlicer
                Top=top,
                Left=left,
                Width=width,
                Height=height,
            )
        )
        return True
    except Exception as exc:
        logger.warning("Slicer creation failed for %s: %s", field_olap_name, exc)
        return False


def _query_shape_hash(rows, columns, measures, legend, filters) -> str:
    payload = {
        "rows": sorted(str(v).casefold() for v in rows if v),
        "columns": sorted(str(v).casefold() for v in columns if v),
        "measures": sorted(str(v).casefold() for v in measures if v),
        "legend": sorted(str(v).casefold() for v in legend if v),
        "filters": sorted(str(f).casefold() for f in (filters or []) if f),
    }
    return hashlib.sha256(
        json.dumps(payload, sort_keys=True, default=str).encode()
    ).hexdigest()[:16]


def _resolve_visual_type(
    vb_dict: Dict[str, Any],
    rows: List[str],
    columns: List[str],
    measures: List[Any],
    slicer_field: Any,
    title: str,
) -> str:
    """Resolve incorrect or generic visual types from binding structure.

    Layout-only PBIX extraction can lose the original Power BI visual type.
    This fallback prevents cards, gauges and slicers from becoming charts.
    """
    raw_type = (
        str(
            vb_dict.get("visual_type")
            or vb_dict.get("type")
            or vb_dict.get("visualType")
            or vb_dict.get("visual_kind")
            or ""
        )
        .casefold()
        .strip()
    )

    aliases = {
        "cardvisual": "card",
        "card": "card",
        "multirowcard": "multirowcard",
        "multi row card": "multirowcard",
        "kpivisual": "kpivisual",
        "kpi": "kpi",
        "gauge": "gauge",
        "gaugevisual": "gauge",
        "slicer": "slicer",
        "table": "tableex",
        "matrix": "matrix",
    }
    if raw_type in aliases:
        return aliases[raw_type]

    if "slicer" in raw_type or slicer_field:
        return "slicer"
    if "gauge" in raw_type:
        return "gauge"
    if "card" in raw_type:
        return "multirowcard" if len(measures) > 1 else "card"
    if "kpi" in raw_type:
        return "kpivisual"

    no_categories = not rows and not columns
    if no_categories and measures:
        settings = vb_dict.get("settings") or {}
        has_gauge_metadata = any(
            vb_dict.get(key) is not None
            for key in ("maximum", "max_value", "minimum", "target")
        )
        if isinstance(settings, dict):
            has_gauge_metadata = has_gauge_metadata or any(
                settings.get(key) is not None
                for key in ("maximum", "max_value", "minimum", "target")
            )

        title_key = str(title or "").casefold()
        if has_gauge_metadata or "gauge" in title_key:
            return "gauge"

        # A single measure without a category is a KPI/card, not a chart.
        if len(measures) == 1:
            return "card"

        # Multiple measures without categories represent a multi-row card.
        return "multirowcard"

    return raw_type or "clusteredcolumnchart"


def save_workbook_safely(excel: Any, workbook: Any, output_path_str: str, session: Any) -> Path:
    import pywintypes
    output_path = Path(output_path_str).expanduser().resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    ext = output_path.suffix.lower()
    file_format = 51
    if ext == ".xlsm":
        file_format = 52
    elif ext == ".xlsb":
        file_format = 50
        
    logger.info("Starting workbook save: %s", output_path)
    
    try:
        if hasattr(excel, "CalculateUntilAsyncQueriesDone"):
            excel.CalculateUntilAsyncQueriesDone()
    except Exception:
        pass
        
    try:
        for _ in range(20):
            if int(getattr(excel, "CalculationState", 0)) == 0:
                break
            time.sleep(0.5)
    except Exception:
        pass

    try:
        com_call = _import_com_retry()
        com_call(lambda: setattr(excel, "DisplayAlerts", False))
    except Exception:
        pass

    last_exc = None
    for attempt in range(1, 4):
        try:
            current_path = Path(com_call(lambda: workbook.FullName)).resolve()
            if current_path == output_path:
                com_call(lambda: workbook.Save(), label="workbook.Save")
            else:
                com_call(lambda: workbook.SaveAs(
                    Filename=str(output_path),
                    FileFormat=file_format,
                    AddToMru=False,
                    Local=True,
                ), label="workbook.SaveAs")
            break
        except Exception as exc:
            last_exc = exc
            logger.warning(
                "Excel save attempt %s failed; retrying: %s",
                attempt,
                exc,
            )
            time.sleep(2)
    else:
        raise RuntimeError(f"Failed to save workbook after 3 attempts: {last_exc}")

    for _ in range(20):
        if output_path.exists() and output_path.stat().st_size > 0:
            break
        time.sleep(0.5)
    else:
        raise RuntimeError(
            f"Excel reported save completion but output file was not created: {output_path}"
        )

    logger.info("Workbook save completed: %s size=%d", output_path, output_path.stat().st_size)
    return output_path


def build_live_dashboard(session: Any) -> Dict[str, Any]:
    """Build the full live dashboard: PivotTables, charts, slicers, KPIs.

    Must be called on the session's COM thread AFTER a successful
    detect_and_validate_connection call.
    """
    com_call = _import_com_retry()
    PivotFactory = _import_pivot_factory()
    FilterEngine = _import_filter_engine()
    refresh_and_calculate = _import_refresh_manager()
    (
        render_visual_to_dashboard,
        finalize_live_visual_values,
        configure_dashboard_canvas,
    ) = _import_visual_renderer()
    filter_engine = FilterEngine()
    PowerBINativeCaptureService, POWERBI_NATIVE_CAPTURE, render_mode_for = (
        _import_native_capture()
    )
    native_capture_service = PowerBINativeCaptureService()

    session.update_state("building")

    excel = session._com_excel
    workbook = session._com_workbook
    mapper = session.runtime_objects.get("olap_mapper")
    pivot_cache = session.runtime_objects.get("olap_pivot_cache")
    pivot_table_template = session.runtime_objects.get("template_pivot_table")
    workbook_connection = session.runtime_objects.get("workbook_connection")
    conn_name = session.result.get("conn_name", "")

    if not all([excel, workbook, mapper, pivot_cache, pivot_table_template]):
        message = "Build prerequisites missing — a connected OLAP PivotTable and cache are required."
        session.errors.append(message)
        session.update_state("live_conversion_failed")
        return {"state": "live_conversion_failed", "errors": [message]}

    warnings: List[str] = []
    errors: List[str] = []
    rendered_visual_ids: set = set()
    failed_visual_ids: set = set()
    native_capture_count = 0
    capture_dir = Path(session.output_path).parent / f"_pbi_native_{session.session_id}"

    # ── Disable user interaction and alerts during build ──────────────────
    com_call(lambda: setattr(excel, "DisplayAlerts", False))
    com_call(lambda: setattr(excel, "EnableEvents", False))
    com_call(lambda: setattr(excel, "ScreenUpdating", False))

    # ── Collect Power BI pages and visuals ────────────────────────────────
    final_chunks = session.metadata or {}
    visual_bindings = session.visual_bindings or []

    pages = final_chunks.get("page_chunks") or final_chunks.get("pages") or []
    if not pages:
        # Group visuals by page_name from bindings
        page_names: List[str] = []
        seen_pages: set = set()
        for vb in visual_bindings:
            pn = str(
                (vb.page_name if hasattr(vb, "page_name") else vb.get("page_name", ""))
                or "Dashboard"
            )
            if pn.casefold() not in seen_pages:
                seen_pages.add(pn.casefold())
                page_names.append(pn)
        pages = [{"name": pn} for pn in page_names] or [{"name": "Dashboard"}]

    # ── Create one dashboard sheet per page ───────────────────────────────
    existing = _existing_sheet_names(workbook)
    dashboard_sheets: Dict[str, Any] = {}

    # Remove the default blank sheets created with the workbook
    ws_count = int(com_call(lambda: workbook.Worksheets.Count))
    default_names = []
    template_sheet_name = ""
    try:
        template_sheet_name = str(
            com_call(lambda: pivot_table_template.Parent.Name) or ""
        )
    except Exception:
        template_sheet_name = ""

    for i in range(1, ws_count + 1):
        n = str(com_call(lambda: workbook.Worksheets(i).Name))
        if (
            re.match(r"^Sheet\d*$", n, re.IGNORECASE)
            and n.casefold() != template_sheet_name.casefold()
        ):
            default_names.append(n)

    for page in pages:
        page_name = str(page.get("name") or page.get("page_name") or "Dashboard")
        safe = _safe_sheet_name(page_name, existing)
        existing.add(safe.casefold())
        sheet = _get_or_create_sheet(workbook, safe, visible=True)
        configure_dashboard_canvas(excel, sheet)
        dashboard_sheets[page_name] = sheet

    # Remove default blank sheets (keep at least one)
    for dname in default_names:
        ws_count_now = int(com_call(lambda: workbook.Worksheets.Count))
        if ws_count_now > 1:
            _delete_sheet_if_exists(workbook, dname)

    # ── Hidden pivot sheets ───────────────────────────────────────────────
    pivot_sheet_name = _safe_sheet_name("_Pivots", existing)
    existing.add(pivot_sheet_name.casefold())
    pivot_sheet = _get_or_create_sheet(workbook, pivot_sheet_name, visible=False)

    report_sheet_name = _safe_sheet_name("_Live_Conversion_Report", existing)
    existing.add(report_sheet_name.casefold())
    report_sheet = _get_or_create_sheet(workbook, report_sheet_name, visible=False)

    # ── PivotFactory for query-shape deduplication ────────────────────────
    factory = PivotFactory(
        pivot_cache=pivot_cache,
        template_pivot_table=pivot_table_template,
        workbook_connection=workbook_connection,
    )
    pivot_map: Dict[str, Any] = {}  # shape_hash -> pivot_info

    def _ensure_pivot(
        rows,
        cols,
        measures,
        legend,
        filters,
        vtype,
        visual_key,
    ):
        # One hidden connected PivotTable per Power BI visual.
        h = f"{visual_key}:{_query_shape_hash(rows, cols, measures, legend, filters)}"
        if h in pivot_map:
            session.pivot_tables_reused += 1
            return pivot_map[h]

        info = factory.get_or_create_pivot(
            excel_app=excel,
            workbook=workbook,
            connection_name=conn_name,
            rows=rows,
            columns=cols,
            measures=measures,
            legend=legend,
            visual_type=vtype,
            field_mapper=mapper,
            filters=filters,
            unique_key=visual_key,
        )
        if "error" not in info:
            session.pivot_tables_created += 1
            pivot_map[h] = info

        warnings.extend(info.get("warnings") or [])
        return info

    # ── KPI sheet for CUBEVALUE formulas ─────────────────────────────────
    kpi_sheet_name = _safe_sheet_name("_KPI_Values", existing)
    existing.add(kpi_sheet_name.casefold())
    kpi_sheet = _get_or_create_sheet(workbook, kpi_sheet_name, visible=False)
    kpi_row = 2

    # ── Iterate visuals ───────────────────────────────────────────────────
    slicer_fields_done: set = set()
    required_pivot_count = 0
    chart_left = 20.0
    chart_top = 220.0
    chart_w = 850.0
    chart_h = 360.0
    chart_gap = 20.0
    slicer_left = 20.0
    slicer_top = 45.0
    slicer_w = 330.0
    slicer_h = 130.0
    dashboard_helper_row = 2
    dashboard_helper_col = 200

    for vb in visual_bindings:
        # Normalise: VisualBinding object or dict
        if hasattr(vb, "model_dump"):
            vb_dict = vb.model_dump()
        elif hasattr(vb, "__dict__"):
            vb_dict = vars(vb)
        else:
            vb_dict = dict(vb) if isinstance(vb, dict) else {}

        render_operation = dict(vb_dict.get("render_operation") or {})
        if render_operation:
            # PBIX geometry remains authoritative. Screenshot/HF output may
            # improve title, visual classification and styling only.
            if render_operation.get("title"):
                vb_dict["title"] = str(render_operation["title"])
            if render_operation.get("visual_type"):
                vb_dict["visual_type"] = str(render_operation["visual_type"])
            vb_dict["render_style"] = dict(render_operation.get("style") or {})
            vb_dict["screenshot_style"] = dict(render_operation.get("style") or {})

        page_name = str(vb_dict.get("page_name") or list(dashboard_sheets.keys())[0])
        dash_sheet = (
            dashboard_sheets.get(page_name) or list(dashboard_sheets.values())[0]
        )

        rows_f = [str(f) for f in (vb_dict.get("rows") or []) if f]
        cols_f = [str(f) for f in (vb_dict.get("columns") or []) if f]
        measures_f = vb_dict.get("measures") or []
        legend_f = [str(f) for f in (vb_dict.get("legend") or []) if f]
        filters_f = vb_dict.get("filters") or []
        slicer_field = vb_dict.get("slicer_field")
        title = str(vb_dict.get("title") or "")
        vtype = _resolve_visual_type(
            vb_dict=vb_dict,
            rows=rows_f,
            columns=cols_f,
            measures=measures_f,
            slicer_field=slicer_field,
            title=title,
        )

        # Defensive runtime role cleanup. This guarantees that metadata leakage
        # cannot make a slicer carry measures or a single-value card create three
        # cards, even when an older cached binding is supplied.
        if vtype in _SLICER_TYPES:
            measures_f = []
            cols_f = []
            legend_f = []
            rows_f = rows_f[:1]
            vb_dict["measures"] = []
            vb_dict["columns"] = []
            vb_dict["legend"] = []
            vb_dict["rows"] = rows_f
        elif vtype in _GAUGE_TYPES or (
            vtype in _CARD_TYPES and "multirow" not in vtype
        ):
            measures_f = measures_f[:1]
            vb_dict["measures"] = measures_f

        visual_layout = dict(vb_dict.get("layout") or {})
        logger.info(
            "Visual %s placement_source=%s pbix=(%s,%s,%s,%s)",
            vb_dict.get("visual_id") or vb_dict.get("chunk_id") or "",
            visual_layout.get("layout_source")
            or (
                "pbix"
                if all(
                    visual_layout.get(k) is not None
                    for k in ("x", "y", "width", "height")
                )
                else "excel_cells"
            ),
            visual_layout.get("x"),
            visual_layout.get("y"),
            visual_layout.get("width"),
            visual_layout.get("height"),
        )

        logger.info(
            "Visual %s resolved type=%s raw_type=%s rows=%d cols=%d measures=%d",
            vb_dict.get("visual_id") or vb_dict.get("chunk_id") or "",
            vtype,
            vb_dict.get("visual_type") or vb_dict.get("type") or "",
            len(rows_f),
            len(cols_f),
            len(measures_f),
        )
        visual_id = str(
            vb_dict.get("visual_id")
            or vb_dict.get("chunk_id")
            or f"visual_{len(pivot_map) + 1:03d}"
        )

        # ── Slicers ───────────────────────────────────────────────────────
        if vtype in _SLICER_TYPES or (slicer_field and not measures_f):
            field = slicer_field or (rows_f[0] if rows_f else None)
            if field and field.casefold() not in slicer_fields_done:
                mapping = mapper.map_field(field, "dimension")
                if mapping.get("status") == "mapped":
                    # Need a backing pivot for the slicer
                    required_pivot_count += 1
                    pivot_info = _ensure_pivot(
                        [field],
                        [],
                        [],
                        [],
                        [],
                        "slicer",
                        visual_id,
                    )
                    pt = pivot_info.get("pivot_table")
                    if pt is not None:
                        visual_left, visual_top, visual_width, visual_height = (
                            _visual_layout_points(
                                vb_dict.get("layout") or {},
                                fallback_left=slicer_left,
                                fallback_top=slicer_top,
                                fallback_width=slicer_w,
                                fallback_height=slicer_h,
                            )
                        )
                        slicer_result = filter_engine.create_slicer(
                            workbook=workbook,
                            dashboard_sheet=dash_sheet,
                            source_pivot=pt,
                            field=field,
                            field_mapper=mapper,
                            target_pivots=[
                                item.get("pivot_table")
                                for item in pivot_map.values()
                                if item.get("pivot_table") is not None
                            ],
                            title=title or field,
                            left=visual_left,
                            top=visual_top,
                            width=visual_width,
                            height=visual_height,
                        )
                        if slicer_result.get("status") == "success":
                            session.slicers_created += 1
                            rendered_visual_ids.add(visual_id)
                            slicer_fields_done.add(field.casefold())
                            warnings.extend(slicer_result.get("warnings") or [])
                            slicer_top += slicer_h + chart_gap
                        else:
                            warnings.append(
                                f"Slicer creation failed for {field}: "
                                f"{slicer_result.get('error') or 'Excel rejected the slicer'}"
                            )
                else:
                    warnings.append(f"Slicer field unmapped: {field}")
            continue

        # ── Gauge — connected PivotTable + semi-doughnut visual ─────────
        if vtype in _GAUGE_TYPES and measures_f:
            required_pivot_count += 1
            pivot_info = _ensure_pivot(
                [],
                [],
                measures_f[:1],
                [],
                filters_f,
                "gauge",
                visual_id,
            )

            if "error" in pivot_info:
                errors.append(
                    f"Gauge PivotTable failed for visual '{title}': "
                    f"{pivot_info['error']}"
                )
                continue

            gauge_pivot = pivot_info.get("pivot_table")
            if gauge_pivot is None:
                errors.append(
                    f"Gauge PivotTable was not returned for visual '{title}'."
                )
                continue

            maximum = (
                vb_dict.get("maximum")
                or vb_dict.get("max_value")
                or (vb_dict.get("settings") or {}).get("maximum")
                if isinstance(vb_dict.get("settings"), dict)
                else None
            )

            if maximum is not None:
                vb_dict["maximum"] = maximum
            gauge_render = render_visual_to_dashboard(
                excel_app=excel,
                workbook=workbook,
                dashboard_sheet=dash_sheet,
                binding=vb_dict,
                pivot_info=pivot_info,
                field_mapper=mapper,
                theme=dict(DEFAULT_THEME),
                cube_filter_refs=[],
                connection_name=conn_name,
                materialized_formulas={},
            )
            if gauge_render.get("status") not in {"success", "live_approximation"}:
                warnings.append(
                    f"Gauge could not be created for '{title}': "
                    f"{gauge_render.get('error') or 'renderer failed'}"
                )
            else:
                session.result["gauges_created"] = (
                    int(session.result.get("gauges_created", 0)) + 1
                )
                rendered_visual_ids.add(visual_id)
            continue

        # ── KPI / card — connected PivotTable + visible KPI ──────────────
        if vtype in _CARD_TYPES and measures_f:
            required_pivot_count += 1
            pivot_info = _ensure_pivot(
                [],
                [],
                measures_f,
                [],
                filters_f,
                "kpi",
                visual_id,
            )

            if "error" in pivot_info:
                errors.append(
                    f"KPI PivotTable failed for visual '{title}': "
                    f"{pivot_info['error']}"
                )
                continue

            kpi_pivot = pivot_info.get("pivot_table")
            if kpi_pivot is None:
                errors.append(f"KPI PivotTable was not returned for visual '{title}'.")
                continue

            # Keep CUBEVALUE for named measures when possible.
            for measure in measures_f:
                if not isinstance(measure, dict):
                    continue

                if str(measure.get("field_type") or "") != "named_measure":
                    continue

                mapping = mapper.map_field(measure, "measure")
                if mapping.get("status") != "mapped":
                    continue

                if _create_kpi_formula(
                    kpi_sheet,
                    kpi_row,
                    2,
                    conn_name,
                    mapping["excel_olap_field"],
                    title or str(measure.get("display_name") or "KPI"),
                ):
                    session.cube_formulas_created += 1
                    kpi_row += 3

            card_render = render_visual_to_dashboard(
                excel_app=excel,
                workbook=workbook,
                dashboard_sheet=dash_sheet,
                binding=vb_dict,
                pivot_info=pivot_info,
                field_mapper=mapper,
                theme=dict(DEFAULT_THEME),
                cube_filter_refs=[],
                connection_name=conn_name,
                materialized_formulas={},
            )
            if card_render.get("status") not in {"success", "live_approximation"}:
                warnings.append(
                    f"KPI PivotTable was created but the visible KPI card "
                    f"could not be placed for '{title}': "
                    f"{card_render.get('error') or 'renderer failed'}"
                )
            else:
                visible_count = int(card_render.get("visible_kpi_count") or 1)
                session.result["kpi_cards_created"] = (
                    int(session.result.get("kpi_cards_created", 0)) + visible_count
                )
                rendered_visual_ids.add(visual_id)
            continue

        # ── Power BI-native/custom visuals: render through Power BI Service ──
        if render_mode_for(vtype) == POWERBI_NATIVE_CAPTURE:
            try:
                image_path = native_capture_service.render_visual(
                    vb_dict,
                    dict(final_chunks.get("canvas") or {"width": 1280, "height": 720}),
                    capture_dir,
                )
                vb_dict["image_path"] = str(image_path)
                vb_dict["render_mode"] = "powerbi_image"
                native_render = render_visual_to_dashboard(
                    excel_app=excel,
                    workbook=workbook,
                    dashboard_sheet=dash_sheet,
                    binding=vb_dict,
                    pivot_info=None,
                    field_mapper=mapper,
                    theme={**DEFAULT_THEME, **dict(final_chunks.get("theme") or {})},
                    cube_filter_refs=[],
                    connection_name=conn_name,
                    materialized_formulas={},
                )
                if native_render.get("status") != "powerbi_native_image":
                    raise RuntimeError(
                        native_render.get("error")
                        or "Power BI-native image renderer failed"
                    )
                native_capture_count += 1
                rendered_visual_ids.add(visual_id)
            except Exception as exc:
                failed_visual_ids.add(visual_id)
                errors.append(
                    f"Power BI-native visual '{title or visual_id}' failed: {exc}"
                )
            continue

        # ── Chart / table / matrix — PivotTable + optional chart ──────────
        if not (rows_f or cols_f or measures_f):
            continue

        required_pivot_count += 1
        pivot_info = _ensure_pivot(
            rows_f,
            cols_f,
            measures_f,
            legend_f,
            filters_f,
            vtype,
            visual_id,
        )
        if "error" in pivot_info:
            errors.append(
                f"PivotTable failed for visual '{title}': {pivot_info['error']}"
            )
            continue

        pt = pivot_info.get("pivot_table")
        pt_sheet_obj = None
        try:
            pt_sheet_obj = com_call(
                lambda: workbook.Worksheets(pivot_info["sheet_name"])
            )
        except Exception:
            pass

        # Centralized visible rendering: apply validated screenshot layout/style
        # while keeping the hidden PivotTable as the live semantic data source.
        if pt is not None:
            visible_render = render_visual_to_dashboard(
                excel_app=excel,
                workbook=workbook,
                dashboard_sheet=dash_sheet,
                binding=vb_dict,
                pivot_info=pivot_info,
                field_mapper=mapper,
                theme=dict(DEFAULT_THEME),
                cube_filter_refs=[],
                connection_name=conn_name,
                materialized_formulas={},
            )
            if visible_render.get("status") in {"success", "live_approximation"}:
                if vtype not in _TABLE_TYPES:
                    session.pivot_charts_created += 1
                rendered_visual_ids.add(visual_id)
            else:
                warnings.append(
                    f"Visual renderer failed for '{title}': "
                    f"{visible_render.get('error') or 'unknown renderer error'}"
                )

    # Hide formula/helper cells used by KPI cards and gauges.
    for dashboard_sheet in dashboard_sheets.values():
        try:
            dashboard_sheet.Columns.Item(dashboard_helper_col).Hidden = True
            dashboard_sheet.Columns.Item(dashboard_helper_col + 1).Hidden = True
            dashboard_sheet.Columns.Item(dashboard_helper_col + 2).Hidden = True
            dashboard_sheet.Columns.Item(dashboard_helper_col + 3).Hidden = True
        except Exception:
            pass

    # Preserve the user-created connected PivotTable and hide its helper sheet.
    if template_sheet_name:
        try:
            template_sheet = com_call(lambda: workbook.Worksheets(template_sheet_name))
            com_call(lambda: setattr(template_sheet, "Visible", XL_HIDDEN))
        except Exception as exc:
            warnings.append(f"Could not hide template PivotTable sheet: {exc}")

    # ── Validate build results before refresh/save ────────────────────────
    expected_pivots = required_pivot_count
    session.expected_counts = {
        "pivot_tables": expected_pivots,
        "pivot_charts": session.pivot_charts_created,
        "slicers": session.slicers_created,
        "cube_formulas": session.cube_formulas_created,
        "dashboard_sheets": len(dashboard_sheets),
    }
    session.actual_counts = {
        "pivot_tables": factory.created_count,
        "pivot_charts": session.pivot_charts_created,
        "slicers": session.slicers_created,
        "cube_formulas": session.cube_formulas_created,
        "dashboard_sheets": len(dashboard_sheets),
    }
    if factory.errors:
        errors.extend(
            f"{item.get('pivot_name')}: {item.get('error')}" for item in factory.errors
        )
    if factory.created_count < expected_pivots:
        message = (
            "Not all required semantic-model PivotTables were created: "
            f"expected={expected_pivots}, "
            f"created={factory.created_count}, "
            f"failed={factory.failed_count}."
        )
        errors.append(message)
        session.errors.extend(errors)
        session.update_state("live_conversion_failed")
        return {
            "state": "live_conversion_failed",
            "expected_pivots": expected_pivots,
            "created_pivots": factory.created_count,
            "failed_pivots": factory.failed_count,
            "errors": errors,
        }

    # ── Add page titles on dashboard sheets ───────────────────────────────
    for page_name, ds in dashboard_sheets.items():
        try:
            title_cell = com_call(lambda: ds.Cells(1, 1))
            com_call(lambda: setattr(title_cell, "Value", page_name))
            com_call(lambda: setattr(title_cell.Font, "Size", 16))
            com_call(lambda: setattr(title_cell.Font, "Bold", True))
        except Exception:
            pass

    # ── Refresh ───────────────────────────────────────────────────────────
    session.update_state("refreshing")
    com_call(lambda: setattr(excel, "ScreenUpdating", True))

    # Keep alerts disabled while the automated refresh/save pipeline runs.
    # A save is attempted only after refresh_manager confirms that Excel,
    # all connections, and all PivotCaches are fully idle.
    com_call(lambda: setattr(excel, "DisplayAlerts", False))

    refresh_result = refresh_and_calculate(
        excel_app=excel,
        workbook=workbook,
        timeout_seconds=int(os.getenv("LIVE_REFRESH_TIMEOUT", "240")),
        poll_interval=float(os.getenv("LIVE_REFRESH_POLL_INTERVAL", "0.5")),
    )
    warnings.extend(refresh_result.warnings)
    errors.extend(refresh_result.errors)

    if (
        not bool(getattr(refresh_result, "refresh_completed", False))
        or bool(getattr(refresh_result, "timeout", False))
        or bool(getattr(refresh_result, "errors", []))
    ):
        message = (
            "The Power BI/OLAP refresh did not complete cleanly. "
            "The workbook was not saved because doing so could cancel a pending "
            "refresh or preserve stale values."
        )
        errors.append(message)
        session.errors.extend(errors)
        session.update_state("live_conversion_failed")
        try:
            com_call(lambda: setattr(excel, "DisplayAlerts", True))
        except Exception:
            pass
        return {
            "state": "live_conversion_failed",
            "refresh_completed": False,
            "refresh_timeout": bool(getattr(refresh_result, "timeout", False)),
            "warnings": warnings,
            "errors": errors,
        }

    logger.info(
        "Refresh confirmed idle before save: connections=%s pivots=%s duration=%ss",
        getattr(refresh_result, "connections_refreshed", 0),
        getattr(refresh_result, "pivot_caches_refreshed", 0),
        getattr(refresh_result, "duration_seconds", 0),
    )

    # Excel can reject linked text-box formulas while OLAP objects are being
    # created. Reapply those links now that refresh is complete and Excel is idle.
    try:
        live_value_result = finalize_live_visual_values(workbook)
        session.result["live_value_links"] = live_value_result
        com_call(lambda: excel.CalculateFull())
    except Exception as exc:
        warnings.append(f"Could not finalize live KPI/gauge labels: {exc}")

    # ── Save ──────────────────────────────────────────────────────────────
    detected_visual_count = len(visual_bindings)
    rendered_visual_count = len(rendered_visual_ids)
    session.result["native_captures_created"] = native_capture_count
    session.result["detected_visuals"] = detected_visual_count
    session.result["rendered_visuals"] = rendered_visual_count
    session.result["failed_visuals"] = len(failed_visual_ids)
    session.result["skipped_visuals"] = max(
        0, detected_visual_count - rendered_visual_count - len(failed_visual_ids)
    )
    if errors or rendered_visual_count + len(failed_visual_ids) < detected_visual_count:
        message = (
            f"Render-all validation failed: detected={detected_visual_count} "
            f"rendered={rendered_visual_count} failed={len(failed_visual_ids)}. "
            + (
                " | ".join(errors[:5])
                if errors
                else "One or more visuals were not rendered."
            )
        )
        session.errors.append(message)
        session.update_state("live_conversion_failed")
        return {
            "state": "live_conversion_failed",
            "errors": session.errors,
            "warnings": warnings,
        }

    session.update_state("saving")
    output_path = session.output_path
    try:
        save_workbook_safely(excel, workbook, output_path)
        logger.info("Session %s: workbook saved to %s", session.session_id, output_path)
        try:
            com_call(lambda: setattr(excel, "DisplayAlerts", True))
        except Exception:
            pass
    except Exception as exc:
        errors.append(f"SaveAs failed: {exc}")
        session.errors.extend(errors)
        session.update_state("live_conversion_failed")
        return {"state": "live_conversion_failed", "errors": errors}

    # ── Close session workbook / Excel ────────────────────────────────────
    _quit_excel_safely(excel, workbook, session, save=False)

    # ── Fresh-reopen verification ─────────────────────────────────────────
    session.update_state("verifying")
    verification = _verify_saved_workbook(
        output_path, conn_name, session.expected_counts
    )
    warnings.extend(verification.get("warnings", []))
    errors.extend(verification.get("errors", []))

    # ── Build result dict ─────────────────────────────────────────────────
    verification_passed = bool(verification.get("passed", False))
    session.post_save_verification_passed = verification_passed
    session.live_conversion_succeeded = verification_passed
    final_state = "completed_live" if verification_passed else "live_conversion_failed"
    result = {
        "state": final_state,
        "conversion_mode": "interactive_live_semantic_model",
        "connection_preserved": verification.get("connection_preserved", False),
        "live_conversion_succeeded": verification_passed,
        "selected_connection_name": conn_name,
        "cube_field_count": session.cube_field_count,
        "semantic_match_score": session.semantic_match_score,
        "pivot_tables_created": session.pivot_tables_created,
        "pivot_tables_reused": session.pivot_tables_reused,
        "pivot_charts_created": session.pivot_charts_created,
        "slicers_created": session.slicers_created,
        "cube_formulas_created": session.cube_formulas_created,
        "kpi_cards_created": int(session.result.get("kpi_cards_created", 0)),
        "gauges_created": int(session.result.get("gauges_created", 0)),
        "post_save_verification_passed": verification_passed,
        "expected_counts": dict(session.expected_counts),
        "actual_counts": verification.get("actual_counts", dict(session.actual_counts)),
        "download_url": (
            f"/live-connect/{session.session_id}/download"
            if verification_passed
            else None
        ),
        "report_url": f"/live-connect/{session.session_id}/report",
        "warnings": warnings,
        "errors": errors,
    }

    session.result.update(result)
    session.warnings.extend(warnings)
    session.errors.extend(errors)
    session.update_state(final_state)

    logger.info(
        "Session %s: build complete — pivots=%d charts=%d slicers=%d cube_formulas=%d, cards=%d, gauges=%d.",
        session.session_id,
        session.pivot_tables_created,
        session.pivot_charts_created,
        session.slicers_created,
        session.cube_formulas_created,
        int(session.result.get("kpi_cards_created", 0)),
        int(session.result.get("gauges_created", 0)),
    )
    return result


# ---------------------------------------------------------------------------
# Reopen verification
# ---------------------------------------------------------------------------
def _verify_saved_workbook(
    output_path: str,
    expected_conn_name: str,
    expected_counts: Optional[Dict[str, int]] = None,
) -> Dict[str, Any]:
    """Open the saved workbook in a new COM instance and verify it rigorously."""
    com_call = _import_com_retry()
    result: Dict[str, Any] = {
        "passed": False,
        "connection_preserved": False,
        "olap_cache_found": False,
        "olap_pivot_found": False,
        "has_charts_or_slicers_or_formulas": False,
        "warnings": [],
        "errors": [],
    }

    _clear_excel_recovery_files()

    import pythoncom  # type: ignore[import]

    pythoncom.CoInitialize()
    excel2 = None
    wb2 = None

    try:
        import win32com.client as win32  # type: ignore[import]

        excel2 = win32.DispatchEx("Excel.Application")
        com_call(lambda: setattr(excel2, "Visible", False))
        com_call(lambda: setattr(excel2, "DisplayAlerts", False))
        com_call(lambda: setattr(excel2, "AskToUpdateLinks", False))

        wb2 = com_call(
            lambda: excel2.Workbooks.Open(
                output_path,
                UpdateLinks=False,
                ReadOnly=True,
            ),
            label="Workbooks.Open (verify)",
        )

        # 1. Workbook connections
        conn_count = int(com_call(lambda: wb2.Connections.Count))
        result["connection_preserved"] = conn_count > 0
        if conn_count == 0:
            result["errors"].append("No workbook connections found after save.")

        # 2. OLAP PivotCaches
        pc_count = int(com_call(lambda: wb2.PivotCaches().Count))
        olap_cache_found = False
        for i in range(1, pc_count + 1):
            try:
                pc = com_call(lambda: wb2.PivotCaches().Item(i))
                if bool(com_call(lambda: pc.OLAP)):
                    olap_cache_found = True
                    break
            except Exception:
                continue
        result["olap_cache_found"] = olap_cache_found
        if not olap_cache_found:
            result["warnings"].append(
                "No OLAP PivotCache found; connection may be missing."
            )

        # 3. Connected PivotTables
        ws_count = int(com_call(lambda: wb2.Worksheets.Count))
        result["sheet_count"] = ws_count
        olap_pivot_found = False
        chart_count = 0
        slicer_count = 0
        formula_count = 0

        for ws_idx in range(1, ws_count + 1):
            try:
                ws = com_call(lambda: wb2.Worksheets(ws_idx))
                pt_count = int(com_call(lambda: ws.PivotTables().Count))
                for pt_idx in range(1, pt_count + 1):
                    try:
                        pt = com_call(lambda: ws.PivotTables(pt_idx))
                        cache = com_call(lambda: pt.PivotCache())
                        if bool(com_call(lambda: cache.OLAP)):
                            olap_pivot_found = True
                    except Exception:
                        continue

                # Charts
                try:
                    chart_count += int(com_call(lambda: ws.ChartObjects().Count))
                except Exception:
                    pass

                # CUBE formulas in used range
                try:
                    used = com_call(lambda: ws.UsedRange)
                    cell_count = int(com_call(lambda: used.Count))
                    for ci in range(1, min(cell_count + 1, 500)):
                        try:
                            cell = com_call(lambda: used.Item(ci))
                            val = str(com_call(lambda: cell.Formula) or "")
                            if val.upper().startswith("=CUBE"):
                                formula_count += 1
                        except Exception:
                            pass
                except Exception:
                    pass
            except Exception:
                continue

        # Slicers
        try:
            sc_count = int(com_call(lambda: wb2.SlicerCaches.Count))
            slicer_count = sc_count
        except Exception:
            pass

        result["olap_pivot_found"] = olap_pivot_found
        result["chart_count"] = chart_count
        result["slicer_count"] = slicer_count
        result["cube_formula_count"] = formula_count
        result["has_charts_or_slicers_or_formulas"] = (
            chart_count > 0 or slicer_count > 0 or formula_count > 0
        )

        actual_counts = {
            "pivot_tables": 1 if olap_pivot_found else 0,
            "pivot_charts": chart_count,
            "slicers": slicer_count,
            "cube_formulas": formula_count,
            "dashboard_sheets": ws_count,
        }
        # Count all connected PivotTables, not just a boolean.
        connected_pivot_count = 0
        for ws_idx in range(1, ws_count + 1):
            try:
                ws = com_call(lambda: wb2.Worksheets(ws_idx))
                pt_count = int(com_call(lambda: ws.PivotTables().Count))
                for pt_idx in range(1, pt_count + 1):
                    pt = com_call(lambda: ws.PivotTables(pt_idx))
                    cache = com_call(lambda: pt.PivotCache())
                    if bool(com_call(lambda: cache.OLAP)):
                        connected_pivot_count += 1
            except Exception:
                continue
        actual_counts["pivot_tables"] = connected_pivot_count
        result["actual_counts"] = actual_counts

        mandatory_ok = (
            result["connection_preserved"]
            and result["olap_cache_found"]
            and connected_pivot_count > 0
        )
        expected = expected_counts or {}
        for key in ("pivot_tables", "pivot_charts", "slicers", "cube_formulas"):
            required = int(expected.get(key, 0) or 0)
            actual = int(actual_counts.get(key, 0) or 0)

            # The workbook contains the one user-created template PivotTable
            # plus the PivotTables generated for visuals/KPIs.
            if key == "pivot_tables" and required > 0:
                required += 1
            if required > 0 and actual < required:
                result["errors"].append(
                    f"Expected at least {required} {key}, found {actual}."
                )
        result["passed"] = mandatory_ok and not result["errors"]

        if not result["has_charts_or_slicers_or_formulas"]:
            result["warnings"].append(
                "No charts, slicers, or CUBE formulas detected; dashboard may be empty."
            )

    except Exception as exc:
        result["errors"].append(f"Reopen verification failed: {exc}")
    finally:
        try:
            if wb2:
                com_call(lambda: wb2.Close(False))
        except Exception:
            pass
        try:
            if excel2:
                com_call(lambda: excel2.Quit())
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    return result


# ---------------------------------------------------------------------------
# Cancellation helper
# ---------------------------------------------------------------------------
def run_continue_workflow_com(session: Any) -> None:
    """Compound workflow dispatched to the COM thread when user clicks Continue.

    Runs detect_and_validate_connection, then (if it passes) build_live_dashboard.
    If detection fails with connection_not_detected or semantic_model_mismatch the
    session stays in that retriable state so the user can fix and click again.
    Must be called via session.dispatch(..., wait=False).
    """
    try:
        detect_result = detect_and_validate_connection(session)
        detect_state = detect_result.get("state", session.state)

        if detect_state in ("connection_not_detected", "semantic_model_mismatch"):
            return

        if detect_state == "live_conversion_failed":
            return

        if detect_state != "connection_detected":
            logger.error(
                "Session %s: unexpected detect state %r — aborting.",
                session.session_id,
                detect_state,
            )
            if not session.is_terminal():
                session.update_state("live_conversion_failed")
            return

        build_live_dashboard(session)
    finally:
        session.continue_job_enqueued = False


def cancel_session_com(session: Any) -> None:
    """Close only the session's workbook and Excel instance.

    Called on the COM thread by COMSessionManager.cancel_session.
    """
    _quit_excel_safely(session._com_excel, session._com_workbook, session, save=False)


def _quit_excel_safely(
    excel: Any,
    workbook: Any,
    session: Any,
    save: bool = False,
) -> None:
    """Best-effort close workbook and quit Excel, ignoring errors."""
    com_call = _import_com_retry()
    try:
        if workbook is not None:
            com_call(lambda: workbook.Close(save), label="Close workbook")
    except Exception as exc:
        logger.debug("Close workbook error: %s", exc)

    try:
        if excel is not None:
            com_call(lambda: excel.Quit(), label="excel.Quit")
    except Exception as exc:
        logger.debug("Excel quit error: %s", exc)

    session._com_excel = None
    session._com_workbook = None


__all__ = [
    "launch_excel_for_connection",
    "detect_and_validate_connection",
    "build_live_dashboard",
    "run_continue_workflow_com",
    "cancel_session_com",
]

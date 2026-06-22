"""Render Power BI visual bindings into an Excel workbook through COM.

The renderer keeps semantic-model data logic separate from screenshot layout:
- PBIX/TMDL/semantic bindings decide fields, measures, filters and aggregation.
- Screenshot/Hugging Face analysis may improve visual type, position and styling.
- Excel COM creates live PivotTables, charts, slicers, gauges and KPI cards.
"""

from __future__ import annotations

import json
import logging
import re
import time
from typing import Any, Dict, List, Optional, Tuple

from cube_formula_builder import build_cube_value_formula
from powerbi_native_live_renderer import (
    LIVE_APPROXIMATION_VISUALS,
    render_powerbi_native_live_approximation,
)

try:
    from .visual_compatibility_registry import (
        EXCEL_NATIVE,
        LIVE_APPROXIMATION,
        POWERBI_NATIVE_CAPTURE,
        render_mode_for,
    )
except ImportError:
    from visual_compatibility_registry import (
        EXCEL_NATIVE,
        LIVE_APPROXIMATION,
        POWERBI_NATIVE_CAPTURE,
        render_mode_for,
    )

logger = logging.getLogger("excel_visual_renderer")

XL_COLUMN_CLUSTERED = 51
XL_BAR_CLUSTERED = 57
XL_LINE = 4
XL_AREA = 1
XL_PIE = 5
XL_DOUGHNUT = -4120
XL_COLUMN_STACKED = 52
XL_BAR_STACKED = 58
XL_LINE_MARKERS = 65
XL_SCATTER = -4169
XL_RADAR = -4151
XL_TREEMAP = 117
XL_SUNBURST = 120
XL_WATERFALL = 119
XL_FUNNEL = 123
XL_RECTANGLE = 1
XL_TEXTBOX_HORIZONTAL = 1
XL_DESCENDING = 2

MSO_FALSE = 0
MSO_TRUE = -1
MSO_SAVE_WITH_DOCUMENT = -1


def excel_rgb(hex_color: str, default: str = "118DFF") -> int:
    """Convert #RRGGBB or RRGGBB to an Excel OLE/BGR integer."""
    value = str(hex_color or default).strip().lstrip("#")
    if len(value) == 3:
        value = "".join(character * 2 for character in value)
    if len(value) != 6:
        value = default
    try:
        red = int(value[0:2], 16)
        green = int(value[2:4], 16)
        blue = int(value[4:6], 16)
    except ValueError:
        red, green, blue = 17, 141, 255
    return red | (green << 8) | (blue << 16)


def _compact_number_format(measure: Any) -> str:
    descriptor = _as_dict(measure)
    text = " ".join(
        str(descriptor.get(key) or "")
        for key in ("aggregation", "display_name", "measure_name", "column_name")
    ).casefold()
    if any(token in text for token in ("average", "avg", "mean", "rate", "ratio")):
        return "#,##0.00"
    return '[>=1000000]0.0,,"M";[>=1000]0,"K";0'


def _measure_display_label(measure: Any) -> str:
    descriptor = _as_dict(measure)
    field_type = str(descriptor.get("field_type") or "").casefold()
    display_name = str(
        descriptor.get("display_name") or descriptor.get("measure_name") or ""
    ).strip()
    if field_type == "named_measure":
        return display_name or _clean_ref_label(_measure_reference(measure))
    aggregation = str(descriptor.get("aggregation") or "SUM").strip().title()
    column_name = str(
        descriptor.get("column_name") or _clean_ref_label(_measure_reference(measure))
    ).strip()
    return f"{aggregation} of {column_name.lower()}"


_RETRYABLE_COM_CODES = (
    "-2147418111",
    "80010001",
    "-2147417846",
    "8001010a",
    "-2146777998",
    "800ac472",
)


def _pump_messages() -> None:
    try:
        import pythoncom  # type: ignore[import]

        pythoncom.PumpWaitingMessages()
    except Exception:
        pass


def _com_retry(func, *, attempts: int = 40, delay: float = 0.20):
    """Retry Excel calls rejected while OLAP/Pivot operations are busy."""
    last_error = None
    for attempt in range(max(1, attempts)):
        try:
            return func()
        except Exception as exc:
            last_error = exc
            text = str(exc).casefold()
            retryable = any(code.casefold() in text for code in _RETRYABLE_COM_CODES)
            if not retryable or attempt >= attempts - 1:
                raise
            _pump_messages()
            time.sleep(delay)
    if last_error is not None:
        raise last_error


def _set_shape_link_metadata(shape: Any, sheet_name: str, address: str) -> None:
    payload = json.dumps(
        {"kind": "live_value", "sheet": str(sheet_name), "address": str(address)}
    )
    try:
        shape.AlternativeText = payload
    except Exception:
        pass


def _link_or_fill_shape(shape: Any, sheet: Any, cell: Any) -> bool:
    """Link a shape to a cell; fall back to its calculated text."""
    sheet_name = str(_com_retry(lambda: sheet.Name)).replace("'", "''")
    address = str(_com_retry(lambda: cell.Address))
    _set_shape_link_metadata(shape, sheet_name, address)

    try:
        _com_retry(lambda: setattr(shape, "Formula", f"='{sheet_name}'!{address}"))
        return True
    except Exception:
        try:
            text = str(
                _com_retry(lambda: cell.Text) or _com_retry(lambda: cell.Value) or "—"
            )
            _com_retry(lambda: setattr(shape.TextFrame2.TextRange, "Text", text))
        except Exception:
            pass
        return False


def _render_cell_kpi_fallback(
    sheet: Any,
    row: int,
    col: int,
    row_span: int,
    col_span: int,
    title: str,
    formula: str,
    number_format: str,
    theme: Dict[str, Any],
) -> str:
    """Create a live KPI with worksheet cells when Excel rejects drawing calls."""
    end_row = max(row + max(row_span, 4) - 1, row + 3)
    end_col = max(col + max(col_span, 3) - 1, col + 2)
    full = _com_retry(
        lambda: sheet.Range(sheet.Cells(row, col), sheet.Cells(end_row, end_col))
    )
    try:
        _com_retry(lambda: full.UnMerge())
    except Exception:
        pass

    value_end_row = max(row + 1, end_row - 1)
    value_range = _com_retry(
        lambda: sheet.Range(sheet.Cells(row, col), sheet.Cells(value_end_row, end_col))
    )
    title_range = _com_retry(
        lambda: sheet.Range(
            sheet.Cells(value_end_row + 1, col), sheet.Cells(end_row, end_col)
        )
    )
    _com_retry(lambda: value_range.Merge())
    _com_retry(lambda: title_range.Merge())

    value_cell = _com_retry(lambda: sheet.Cells(row, col))
    title_cell = _com_retry(lambda: sheet.Cells(value_end_row + 1, col))
    _set_formula(value_cell, formula)
    value_cell.NumberFormat = number_format
    title_cell.Value = title

    for rng in (value_range, title_range):
        try:
            rng.HorizontalAlignment = -4108
            rng.VerticalAlignment = -4108
        except Exception:
            pass

    try:
        value_cell.Font.Bold = False
        value_cell.Font.Size = 22
        title_cell.Font.Bold = False
        title_cell.Font.Size = 9
        title_cell.Font.Color = excel_rgb(theme.get("secondary_text_color"), "9CA3AF")
        full.Interior.Color = excel_rgb(theme.get("card_background"), "FFFFFF")
        full.Borders.Color = excel_rgb(theme.get("border_color"), "E5E7EB")
    except Exception:
        pass

    return f"{sheet.Name}!{value_range.Address}"


# Visuals shared by Power BI and Excel. These render as native/live Excel objects.
EXCEL_NATIVE_VISUALS = {
    "clusteredcolumnchart",
    "clusteredbarchart",
    "stackedcolumnchart",
    "stackedbarchart",
    "linechart",
    "areachart",
    "piechart",
    "donutchart",
    "scatterchart",
    "waterfallchart",
    "funnelchart",
    "treemap",
    "sunburst",
    "radarchart",
    "combochart",
    "ribbonchart",
    "tableex",
    "matrix",
    "card",
    "multirowcard",
    "kpi",
    "gauge",
    "slicer",
    "textbox",
    "text",
    "shape",
    "image",
}

# Power BI-native visuals that have no fully editable Excel equivalent.
# These are rendered from Power BI export or screenshot crops.
POWERBI_NATIVE_LIVE_APPROX_VISUALS = set(LIVE_APPROXIMATION_VISUALS)

POWERBI_NATIVE_IMAGE_VISUALS = {
    "paginatedreport",
    "powerapps",
    "powerautomate",
    "pythonvisual",
    "rvisual",
    "custom",
}


def _value(obj: Any, key: str, default: Any = None) -> Any:
    if isinstance(obj, dict):
        return obj.get(key, default)
    return getattr(obj, key, default)


def _as_dict(obj: Any) -> Dict[str, Any]:
    if isinstance(obj, dict):
        return dict(obj)
    if hasattr(obj, "model_dump"):
        return obj.model_dump()
    if hasattr(obj, "__dict__"):
        return dict(vars(obj))
    return {}


def _is_generic_title(title_str: str) -> bool:
    if not title_str:
        return True
    text = str(title_str).strip().lower()
    return text in {
        "",
        "none",
        "null",
        "unknown",
        "visual",
        "chart",
        "slicer",
        "filter",
        "placeholder",
        "value",
    } or bool(re.match(r"^visual(?:\s+block)?[_\s-]*\d*$", text))


def _clean_ref_label(ref: str) -> str:
    text = str(ref or "").strip()
    if not text:
        return "Value"

    matches = re.findall(r"\[([^\]]+)\]", text)
    if matches:
        text = matches[-1]
    elif "." in text:
        text = text.rsplit(".", 1)[-1]

    text = re.sub(r"[\]\)\}]+$", "", text).strip("[]'\" ")
    text = re.sub(
        r"(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])",
        " ",
        text,
    )
    return text.strip().title() or "Value"


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


def _chart_type(visual_type: str) -> int:
    vt = str(visual_type or "").casefold().replace("_", "").replace("-", "")

    if "waterfall" in vt:
        return XL_WATERFALL
    if "funnel" in vt:
        return XL_FUNNEL
    if "treemap" in vt:
        return XL_TREEMAP
    if "sunburst" in vt:
        return XL_SUNBURST
    if "scatter" in vt or "bubble" in vt:
        return XL_SCATTER
    if "radar" in vt:
        return XL_RADAR
    if "stackedbar" in vt:
        return XL_BAR_STACKED
    if "stackedcolumn" in vt:
        return XL_COLUMN_STACKED
    if "donut" in vt or "doughnut" in vt:
        return XL_DOUGHNUT
    if "pie" in vt:
        return XL_PIE
    if "line" in vt:
        return XL_LINE_MARKERS if "marker" in vt else XL_LINE
    if "area" in vt:
        return XL_AREA
    if "bar" in vt and "column" not in vt:
        return XL_BAR_CLUSTERED
    return XL_COLUMN_CLUSTERED


def _is_chart_visual(visual_type: str) -> bool:
    vt = str(visual_type or "").casefold()
    return any(
        token in vt
        for token in (
            "chart",
            "bar",
            "column",
            "line",
            "area",
            "pie",
            "donut",
            "doughnut",
            "scatter",
            "bubble",
            "waterfall",
            "funnel",
            "treemap",
            "sunburst",
            "radar",
            "combo",
            "ribbon",
        )
    )


def _is_table_visual(visual_type: str) -> bool:
    vt = str(visual_type or "").casefold()
    return "table" in vt or "matrix" in vt


def _is_card_visual(visual_type: str) -> bool:
    vt = str(visual_type or "").casefold()
    return any(token in vt for token in ("card", "kpi"))


def _is_gauge_visual(visual_type: str) -> bool:
    return "gauge" in str(visual_type or "").casefold()


def _is_slicer_visual(visual_type: str) -> bool:
    return "slicer" in str(visual_type or "").casefold()


def _normalize_visual_type(raw_type: str) -> str:
    raw = str(raw_type or "").casefold().strip()
    compact = raw.replace("_", "").replace("-", "").replace(" ", "")

    aliases = {
        "clusteredcolumnchart": "clusteredcolumnchart",
        "columnchart": "clusteredcolumnchart",
        "clusteredbarchart": "clusteredbarchart",
        "barchart": "clusteredbarchart",
        "stackedcolumnchart": "stackedcolumnchart",
        "stackedbarchart": "stackedbarchart",
        "linechart": "linechart",
        "areachart": "areachart",
        "piechart": "piechart",
        "donutchart": "donutchart",
        "doughnutchart": "donutchart",
        "scatterchart": "scatterchart",
        "bubblechart": "scatterchart",
        "waterfallchart": "waterfallchart",
        "funnelchart": "funnelchart",
        "treemap": "treemap",
        "sunburst": "sunburst",
        "radarchart": "radarchart",
        "combochart": "combochart",
        "lineandclusteredcolumnchart": "combochart",
        "lineandstackedcolumnchart": "combochart",
        "tableex": "tableex",
        "table": "tableex",
        "matrix": "matrix",
        "cardvisual": "card",
        "card": "card",
        "multirowcard": "multirowcard",
        "kpivisual": "kpi",
        "kpi": "kpi",
        "gaugevisual": "gauge",
        "gauge": "gauge",
        "listslicer": "slicer",
        "dropdownslicer": "slicer",
        "slicer": "slicer",
        "map": "map",
        "filledmap": "filledmap",
        "arcgismap": "arcgismap",
        "decompositiontree": "decompositiontree",
        "keyinfluencers": "keyinfluencers",
        "azuremap": "azuremap",
        "shapemap": "shapemap",
        "qna": "qna",
        "qnavisual": "qna",
        "smartnarrative": "smartnarrative",
        "paginatedreport": "paginatedreport",
        "powerapps": "powerapps",
        "powerautomate": "powerautomate",
        "pythonvisual": "pythonvisual",
        "rvisual": "rvisual",
        "customvisual": "custom",
    }
    return aliases.get(compact, raw or "unknown")


def _resolve_visual_type(binding: Any) -> str:
    """Resolve the visual type using metadata first and structure second."""
    raw = _normalize_visual_type(
        _value(binding, "visual_type", "")
        or _value(binding, "type", "")
        or _value(binding, "visualType", "")
        or ""
    )

    rows = list(_value(binding, "rows", []) or [])
    columns = list(_value(binding, "columns", []) or [])
    measures = list(_value(binding, "measures", []) or [])
    slicer_field = _value(binding, "slicer_field", None)
    title = str(_value(binding, "title", "") or "").casefold()

    if slicer_field or "slicer" in raw:
        return "slicer"
    if "gauge" in raw or "gauge" in title:
        return "gauge"
    if "multirowcard" in raw or "multi row card" in raw:
        return "multirowcard"
    if "card" in raw:
        return "card"
    if "kpi" in raw:
        return "kpi"
    if _is_table_visual(raw):
        return raw
    if _is_chart_visual(raw):
        return raw

    no_categories = not rows and not columns
    if no_categories and measures:
        return "multirowcard" if len(measures) > 1 else "card"

    return raw or "clusteredcolumnchart"


def _resolve_title(binding: Any, dashboard_sheet: Any) -> str:
    title = _value(binding, "title", None)
    visual_type = _resolve_visual_type(binding)
    measures = list(_value(binding, "measures", []) or [])
    rows = list(_value(binding, "rows", []) or [])

    if title and not _is_generic_title(str(title)):
        return str(title)

    metric = _clean_ref_label(_measure_reference(measures[0])) if measures else "Value"
    dimension = _clean_ref_label(rows[0]) if rows else "Category"

    if _is_card_visual(visual_type) or _is_gauge_visual(visual_type):
        return metric
    if _is_chart_visual(visual_type):
        return f"{metric} by {dimension}"
    if _is_table_visual(visual_type):
        return f"{dimension} Details"
    if _is_slicer_visual(visual_type):
        return dimension

    return str(
        _value(binding, "page_name", "")
        or getattr(dashboard_sheet, "Name", "")
        or "Dashboard"
    )


def _number_format_for_measure(measure_name: str) -> str:
    name = str(measure_name or "").casefold()
    if any(
        token in name
        for token in ("percent", "percentage", "margin %", "growth %", "rate", "ratio")
    ):
        return "0.00%"
    if any(
        token in name
        for token in ("count", "quantity", "volume", "units", "orders", "customers")
    ):
        return "#,##0"
    if any(
        token in name
        for token in (
            "sales",
            "revenue",
            "amount",
            "cost",
            "profit",
            "price",
            "value",
            "budget",
            "target",
        )
    ):
        return "#,##0.00"
    return "#,##0.00"


def _set_formula(cell: Any, formula: str) -> None:
    try:
        cell.Formula = formula
    except Exception:
        cell.Formula2 = formula


def configure_dashboard_canvas(
    excel_app: Any,
    dashboard_sheet: Any,
    *,
    start_row: int = 3,
    start_col: int = 2,
    end_row: int = 43,
    end_col: int = 26,
    column_width: float = 4.5,
    row_height: float = 15.0,
    zoom: int = 85,
) -> Dict[str, float]:
    """Create one predictable Excel canvas for all Power BI pages."""
    for column in range(start_col, end_col + 1):
        try:
            _com_retry(
                lambda c=column: setattr(
                    dashboard_sheet.Columns.Item(c), "ColumnWidth", column_width
                )
            )
        except Exception:
            pass

    for row in range(start_row, end_row + 1):
        try:
            _com_retry(
                lambda r=row: setattr(
                    dashboard_sheet.Rows.Item(r), "RowHeight", row_height
                )
            )
        except Exception:
            pass

    try:
        for title_row in range(1, 5):
            for title_col in range(1, 5):
                title_cell = dashboard_sheet.Cells(title_row, title_col)
                title_value = str(title_cell.Value or "").strip()
                if re.fullmatch(r"(?i)page\s*\d+", title_value):
                    title_cell.ClearContents()
        dashboard_sheet.Activate()
        excel_app.ActiveWindow.DisplayGridlines = False
        excel_app.ActiveWindow.DisplayHeadings = False
        excel_app.ActiveWindow.Zoom = zoom
    except Exception:
        pass

    left = float(dashboard_sheet.Cells(start_row, start_col).Left)
    top = float(dashboard_sheet.Cells(start_row, start_col).Top)
    right = float(dashboard_sheet.Cells(end_row + 1, end_col + 1).Left)
    bottom = float(dashboard_sheet.Cells(end_row + 1, end_col + 1).Top)

    return {
        "left": left,
        "top": top,
        "width": max(1.0, right - left),
        "height": max(1.0, bottom - top),
        "start_row": float(start_row),
        "start_col": float(start_col),
        "end_row": float(end_row),
        "end_col": float(end_col),
    }


def _flatten_layout(binding: Any) -> Dict[str, Any]:
    layout = dict(_value(binding, "layout", {}) or {})
    for key in ("bbox", "position", "geometry", "bounds"):
        nested = layout.get(key)
        if isinstance(nested, dict):
            layout = {**nested, **layout}
    return layout


def _layout_box(
    binding: Any, dashboard_sheet: Any
) -> Tuple[float, float, float, float, int, int, int, int]:
    """Map PBIX geometry directly onto the configured Excel canvas.

    Priority:
    1. PBIX ``x/y/width/height``.
    2. Excel row/column fallback when PBIX geometry is unavailable.

    Screenshot-derived layout is never used here; screenshot analysis is styling
    input only.
    """
    layout = _flatten_layout(binding)

    row = max(3, min(160, int(layout.get("row") or 5)))
    col = max(2, min(30, int(layout.get("col") or 2)))
    row_span = max(2, min(60, int(layout.get("row_span") or 10)))
    col_span = max(2, min(24, int(layout.get("col_span") or 5)))

    has_pbix_geometry = all(
        layout.get(key) is not None for key in ("x", "y", "width", "height")
    )

    if has_pbix_geometry:
        canvas_width = max(
            1.0,
            float(
                layout.get("canvas_width")
                or layout.get("page_width")
                or layout.get("report_width")
                or 1280.0
            ),
        )
        canvas_height = max(
            1.0,
            float(
                layout.get("canvas_height")
                or layout.get("page_height")
                or layout.get("report_height")
                or 720.0
            ),
        )

        canvas_left = float(dashboard_sheet.Cells(3, 2).Left)
        canvas_top = float(dashboard_sheet.Cells(3, 2).Top)
        canvas_right = float(dashboard_sheet.Cells(44, 27).Left)
        canvas_bottom = float(dashboard_sheet.Cells(44, 27).Top)
        excel_canvas_width = max(1.0, canvas_right - canvas_left)
        excel_canvas_height = max(1.0, canvas_bottom - canvas_top)

        x = max(0.0, min(canvas_width, float(layout.get("x") or 0.0)))
        y = max(0.0, min(canvas_height, float(layout.get("y") or 0.0)))
        visual_width = max(
            1.0,
            min(canvas_width - x, float(layout.get("width") or 1.0)),
        )
        visual_height = max(
            1.0,
            min(canvas_height - y, float(layout.get("height") or 1.0)),
        )

        left = canvas_left + (x / canvas_width) * excel_canvas_width
        top = canvas_top + (y / canvas_height) * excel_canvas_height
        width = max(45.0, (visual_width / canvas_width) * excel_canvas_width)
        height = max(30.0, (visual_height / canvas_height) * excel_canvas_height)

        # Approximate cells only for cell-backed KPI/gauge labels.
        row = max(3, min(43, int(round(3 + (y / canvas_height) * 40))))
        col = max(2, min(26, int(round(2 + (x / canvas_width) * 24))))
        row_span = max(2, int(round((visual_height / canvas_height) * 40)))
        col_span = max(2, int(round((visual_width / canvas_width) * 24)))

        logger.info(
            "PBIX canvas placement visual=%s x=%.1f y=%.1f w=%.1f h=%.1f "
            "-> left=%.1f top=%.1f width=%.1f height=%.1f",
            _value(binding, "visual_id", ""),
            x,
            y,
            visual_width,
            visual_height,
            left,
            top,
            width,
            height,
        )
        return left, top, width, height, row, col, row_span, col_span

    start_cell = dashboard_sheet.Cells(row, col)
    end_cell = dashboard_sheet.Cells(
        min(160, row + row_span),
        min(30, col + col_span),
    )
    left = float(start_cell.Left)
    top = float(start_cell.Top)
    width = max(45.0, float(end_cell.Left) - left)
    height = max(30.0, float(end_cell.Top) - top)

    logger.warning(
        "PBIX geometry unavailable for visual %s; using Excel-cell fallback.",
        _value(binding, "visual_id", ""),
    )
    return left, top, width, height, row, col, row_span, col_span


def _apply_title_style(
    sheet: Any, row: int, col: int, title: str, theme: Dict[str, Any]
) -> None:
    cell = sheet.Cells(row, col)
    cell.Value = title
    cell.Font.Bold = True
    cell.Font.Size = 11
    try:
        cell.Font.Color = excel_rgb(theme.get("text_color"), "111827")
    except Exception:
        pass


def _write_failure_placeholder(
    sheet: Any, row: int, col: int, title: str, message: str
) -> None:
    sheet.Cells(row, col).Value = title
    sheet.Cells(row, col).Font.Bold = True
    message_cell = sheet.Cells(row + 1, col)
    message_cell.Value = f"[Could not render: {message}]"
    message_cell.Font.Italic = True


def _prepare_pivot_for_chart(pivot_table: Any) -> None:
    try:
        pivot_table.RowGrand = False
        pivot_table.ColumnGrand = False
    except Exception:
        pass
    try:
        row_field = pivot_table.RowFields(1)
        data_field = pivot_table.DataFields(1)
        row_field.AutoSort(XL_DESCENDING, data_field.Name)
        logger.info(
            "Applied value-descending sort to PivotTable %s using %s",
            getattr(pivot_table, "Name", "?"),
            getattr(data_field, "Name", "?"),
        )
    except Exception as exc:
        logger.warning("Could not apply sort to PivotTable: %s", exc)


def _style_chart(
    chart: Any,
    title: str,
    theme: Dict[str, Any],
    visual_type: str = "",
    category_title: str = "",
    value_title: str = "",
) -> None:
    visual_key = str(visual_type or "").casefold()
    try:
        chart.HasTitle = True
        chart.ChartTitle.Text = title
        chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 11
        chart.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = -1
        chart.ChartTitle.Format.TextFrame2.TextRange.ParagraphFormat.Alignment = 1
    except Exception:
        pass
    try:
        series_count = int(chart.SeriesCollection().Count)
    except Exception:
        series_count = 0
    try:
        chart.HasLegend = series_count > 1 or any(
            token in visual_key for token in ("pie", "donut", "doughnut")
        )
    except Exception:
        pass
    try:
        chart.ChartArea.Format.Fill.ForeColor.RGB = excel_rgb(
            theme.get("background_color"), "FFFFFF"
        )
        chart.ChartArea.Format.Line.Visible = 0
        chart.PlotArea.Format.Fill.Visible = 0
    except Exception:
        pass
    try:
        if series_count:
            first = chart.SeriesCollection(1)
            accent = excel_rgb(theme.get("accent_color"), "118DFF")
            first.Format.Fill.ForeColor.RGB = accent
            first.Format.Line.ForeColor.RGB = accent
            if "column" in visual_key or "bar" in visual_key:
                first.Format.Line.Visible = 0
    except Exception:
        pass
    try:
        if "column" in visual_key or "bar" in visual_key:
            chart.ChartGroups(1).GapWidth = 35
            chart.ChartGroups(1).Overlap = 0
    except Exception:
        pass
    try:
        category_axis = chart.Axes(1)
        category_axis.HasTitle = bool(category_title)
        if category_title:
            category_axis.AxisTitle.Text = category_title
        category_axis.TickLabels.Font.Size = 9
    except Exception:
        pass
    try:
        value_axis = chart.Axes(2)
        value_axis.HasTitle = bool(value_title)
        if value_title:
            value_axis.AxisTitle.Text = value_title
        value_axis.TickLabels.Font.Size = 9
        value_axis.HasMajorGridlines = True
        value_axis.MajorGridlines.Format.Line.ForeColor.RGB = excel_rgb(
            theme.get("gridline_color"), "D9D9D9"
        )
        value_axis.MajorGridlines.Format.Line.DashStyle = 4
    except Exception:
        pass


def _render_live_chart(
    dashboard_sheet: Any,
    pivot_table: Any,
    binding: Any,
    visual_type: str,
    title: str,
    theme: Dict[str, Any],
) -> str:
    left, top, width, height, *_ = _layout_box(binding, dashboard_sheet)
    _prepare_pivot_for_chart(pivot_table)

    chart_obj = dashboard_sheet.ChartObjects().Add(left, top, width, height)
    chart = chart_obj.Chart
    chart.SetSourceData(pivot_table.TableRange1)
    chart.ChartType = _chart_type(visual_type)
    rows = list(_value(binding, "rows", []) or [])
    measures = list(_value(binding, "measures", []) or [])
    category_title = _clean_ref_label(rows[0]) if rows else ""
    value_title = _measure_display_label(measures[0]) if measures else ""
    _style_chart(
        chart,
        title,
        theme,
        visual_type=visual_type,
        category_title=category_title,
        value_title=value_title,
    )
    logger.info(
        "Created native Excel chart type=%s object=%s left=%.1f top=%.1f width=%.1f height=%.1f",
        visual_type,
        getattr(chart_obj, "Name", ""),
        left,
        top,
        width,
        height,
    )
    return str(chart_obj.Name)


def _build_cube_formula_for_measure(
    measure_descriptor: Any,
    field_mapper: Any,
    connection_name: Optional[str],
    cube_filter_refs: List[Dict[str, Any]],
    materialized_formulas: Dict[str, Dict[str, Any]],
) -> Tuple[str, str]:
    measure_ref = _measure_reference(measure_descriptor)
    if not measure_ref:
        raise RuntimeError("The KPI measure descriptor is empty.")

    prepared = materialized_formulas.get(_measure_key(measure_descriptor))
    if prepared and prepared.get("excel_formula"):
        return (
            str(prepared["excel_formula"]),
            str(prepared.get("cube_measure_path") or ""),
        )

    mapping = field_mapper.map_field(measure_descriptor, "measure")
    if mapping.get("status") != "mapped" or not mapping.get("excel_olap_field"):
        raise RuntimeError(
            f"Measure could not be mapped to an Excel CubeField: {measure_ref}"
        )
    if not connection_name:
        raise RuntimeError(
            "A validated Power BI/OLAP connection name was not supplied."
        )

    mapped_measure = str(mapping["excel_olap_field"])
    formula = build_cube_value_formula(
        str(connection_name),
        mapped_measure,
        cube_filter_refs,
    )
    return formula, mapped_measure


def _measure_field_type(measure: Any) -> str:
    if isinstance(measure, dict):
        return str(measure.get("field_type") or "").casefold()
    return ""


def _pivot_data_formula(
    dashboard_sheet: Any,
    pivot_info: Optional[Dict[str, Any]],
    measure_index: int,
) -> Tuple[str, str]:
    """Return a normal Excel reference to a live PivotTable data cell.

    This is used for implicit measures such as SUM(Table[Column]), which do not
    necessarily exist as named CubeFields and therefore cannot always be used
    directly with CUBEVALUE.
    """
    if not pivot_info or pivot_info.get("pivot_table") is None:
        raise RuntimeError(
            "A connected PivotTable is required for an implicit measure."
        )
    pivot_table = pivot_info["pivot_table"]
    try:
        data_range = pivot_table.DataBodyRange
        source_cell = data_range.Cells(
            1, min(measure_index + 1, int(data_range.Columns.Count))
        )
    except Exception as exc:
        raise RuntimeError(f"The connected PivotTable has no data cell: {exc}") from exc
    source_sheet = str(source_cell.Parent.Name).replace("'", "''")
    return f"='{source_sheet}'!{source_cell.Address}", str(source_cell.Address)


def _render_live_cell_card(
    dashboard_sheet: Any,
    *,
    row: int,
    col: int,
    row_span: int,
    col_span: int,
    title: str,
    formula: str,
    number_format: str,
    theme: Dict[str, Any],
) -> str:
    """Render a guaranteed-visible live KPI card with merged worksheet cells.

    Worksheet cells are significantly more reliable than drawing objects while
    Excel is busy creating OLAP PivotTables. The value remains live because the
    merged value cell contains the Pivot/CUBE formula directly.
    """
    row = max(2, int(row))
    col = max(2, int(col))
    row_span = max(4, int(row_span))
    col_span = max(3, int(col_span))

    value_end_row = row + max(2, row_span - 2)
    end_row = row + row_span - 1
    end_col = col + col_span - 1

    full_range = _com_retry(
        lambda: dashboard_sheet.Range(
            dashboard_sheet.Cells(row, col),
            dashboard_sheet.Cells(end_row, end_col),
        )
    )
    try:
        _com_retry(lambda: full_range.UnMerge())
    except Exception:
        pass

    value_range = _com_retry(
        lambda: dashboard_sheet.Range(
            dashboard_sheet.Cells(row, col),
            dashboard_sheet.Cells(value_end_row, end_col),
        )
    )
    title_range = _com_retry(
        lambda: dashboard_sheet.Range(
            dashboard_sheet.Cells(value_end_row + 1, col),
            dashboard_sheet.Cells(end_row, end_col),
        )
    )

    _com_retry(lambda: value_range.Merge())
    _com_retry(lambda: title_range.Merge())

    value_cell = _com_retry(lambda: dashboard_sheet.Cells(row, col))
    title_cell = _com_retry(lambda: dashboard_sheet.Cells(value_end_row + 1, col))

    _set_formula(value_cell, formula)
    value_cell.NumberFormat = number_format
    title_cell.Value = title

    try:
        value_range.HorizontalAlignment = -4108
        value_range.VerticalAlignment = -4108
        title_range.HorizontalAlignment = -4108
        title_range.VerticalAlignment = -4108

        value_cell.Font.Bold = True
        value_cell.Font.Size = 24
        title_cell.Font.Bold = True
        title_cell.Font.Size = 10

        full_range.Interior.Color = int(
            str(theme.get("card_background", "FFFFFF")).replace("#", ""),
            16,
        )
        full_range.Borders.Color = excel_rgb(theme.get("border_color"), "E5E7EB")
        full_range.Borders.Weight = 2
    except Exception:
        pass

    return str(value_range.Address)


def _render_kpi_cards(
    dashboard_sheet: Any,
    binding: Any,
    pivot_info: Optional[Dict[str, Any]],
    title: str,
    theme: Dict[str, Any],
    field_mapper: Any,
    cube_filter_refs: List[Dict[str, Any]],
    connection_name: Optional[str],
    materialized_formulas: Dict[str, Dict[str, Any]],
) -> Dict[str, Any]:
    """Render live KPI cards with worksheet cells at the PBIX position."""
    measures = list(_value(binding, "measures", []) or [])
    if not measures:
        raise RuntimeError("No measure is assigned to this KPI visual.")

    if "multirow" not in _resolve_visual_type(binding):
        measures = measures[:1]

    _, _, _, _, row, col, row_span, col_span = _layout_box(binding, dashboard_sheet)
    created: List[str] = []
    formulas: List[str] = []
    card_span = max(3, col_span // max(1, len(measures)))

    for index, measure in enumerate(measures):
        if _measure_field_type(measure) == "implicit_measure":
            formula, _ = _pivot_data_formula(dashboard_sheet, pivot_info, index)
        else:
            formula, _ = _build_cube_formula_for_measure(
                measure,
                field_mapper,
                connection_name,
                cube_filter_refs,
                materialized_formulas,
            )

        label = _measure_display_label(measure)
        if len(measures) > 1 and _is_generic_title(label):
            label = _clean_ref_label(_measure_reference(measure))
        address = _render_cell_kpi_fallback(
            dashboard_sheet,
            row,
            col + index * card_span,
            max(4, row_span),
            card_span,
            label,
            formula,
            _compact_number_format(measure),
            theme,
        )
        created.append(address)
        formulas.append(formula)

    return {
        "object_name": ",".join(created),
        "cube_formulas": formulas,
        "visible_kpi_count": len(created),
        "render_strategy": "pbix_positioned_live_cells",
    }


def _style_gauge_point(
    series: Any, index: int, color: Optional[int] = None, transparent: bool = False
) -> None:
    """Style one doughnut point with independent COM retries."""
    point = _com_retry(lambda: series.Points(index), attempts=60, delay=0.15)
    if transparent:
        _com_retry(lambda: setattr(point.Format.Fill, "Visible", MSO_TRUE), attempts=60)
        _com_retry(lambda: setattr(point.Format.Fill, "Transparency", 1.0), attempts=60)
        _com_retry(
            lambda: setattr(point.Format.Line, "Visible", MSO_FALSE), attempts=60
        )
        return
    if color is not None:
        _com_retry(
            lambda: setattr(point.Format.Fill.ForeColor, "RGB", color), attempts=60
        )
        _com_retry(lambda: setattr(point.Format.Fill, "Transparency", 0.0), attempts=60)
        _com_retry(
            lambda: setattr(point.Format.Line, "Visible", MSO_FALSE), attempts=60
        )


def _render_gauge(
    dashboard_sheet: Any,
    binding: Any,
    pivot_info: Optional[Dict[str, Any]],
    title: str,
    theme: Dict[str, Any],
    field_mapper: Any,
    cube_filter_refs: List[Dict[str, Any]],
    connection_name: Optional[str],
    materialized_formulas: Dict[str, Dict[str, Any]],
) -> Dict[str, Any]:
    """Render one live three-point semicircular doughnut series."""
    measures = list(_value(binding, "measures", []) or [])
    if not measures:
        raise RuntimeError("No measure is assigned to this gauge visual.")
    measure = measures[0]
    if _measure_field_type(measure) == "implicit_measure":
        formula, mapped_measure = _pivot_data_formula(dashboard_sheet, pivot_info, 0)
    else:
        formula, mapped_measure = _build_cube_formula_for_measure(
            measure,
            field_mapper,
            connection_name,
            cube_filter_refs,
            materialized_formulas,
        )
    left, top, width, height, row, col, row_span, col_span = _layout_box(
        binding, dashboard_sheet
    )
    helper_col = min(50, max(28, col + col_span + 2))
    helper_row = max(3, row)
    current_cell = _com_retry(lambda: dashboard_sheet.Cells(helper_row, helper_col))
    remaining_cell = _com_retry(
        lambda: dashboard_sheet.Cells(helper_row + 1, helper_col)
    )
    hidden_cell = _com_retry(lambda: dashboard_sheet.Cells(helper_row + 2, helper_col))
    max_cell = _com_retry(lambda: dashboard_sheet.Cells(helper_row + 3, helper_col))
    _set_formula(current_cell, formula)
    operation = dict(_value(binding, "render_operation", {}) or {})
    settings = {
        **dict(_value(binding, "settings", {}) or {}),
        **dict(operation.get("settings") or {}),
    }
    maximum = (
        operation.get("maximum")
        or operation.get("max_value")
        or _value(binding, "maximum", None)
        or _value(binding, "max_value", None)
        or settings.get("maximum")
        or settings.get("max_value")
    )
    minimum = (
        operation.get("minimum")
        if operation.get("minimum") is not None
        else _value(binding, "minimum", settings.get("minimum", 0))
    )
    if maximum is None:
        _set_formula(max_cell, f"=MAX(1,ROUNDUP({current_cell.Address}*2,0))")
    else:
        max_cell.Value = float(maximum)
    _set_formula(remaining_cell, f"=MAX(0,{max_cell.Address}-{current_cell.Address})")
    _set_formula(hidden_cell, f"={max_cell.Address}")
    chart_obj = _com_retry(
        lambda: dashboard_sheet.ChartObjects().Add(
            left, top, width, max(80.0, height * 0.86)
        )
    )
    chart = _com_retry(lambda: chart_obj.Chart)
    _com_retry(lambda: setattr(chart, "ChartType", XL_DOUGHNUT))
    try:
        while int(chart.SeriesCollection().Count) > 0:
            chart.SeriesCollection(1).Delete()
    except Exception:
        pass
    source_range = _com_retry(lambda: dashboard_sheet.Range(current_cell, hidden_cell))
    series = _com_retry(lambda: chart.SeriesCollection().NewSeries())
    _com_retry(lambda: setattr(series, "Values", source_range))
    try:
        series.Name = title
    except Exception:
        pass
    chart.HasLegend = False
    chart.HasTitle = True
    chart.ChartTitle.Text = title
    try:
        chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 10
        chart.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = -1
        chart.ChartTitle.Format.TextFrame2.TextRange.ParagraphFormat.Alignment = 1
        chart.ChartArea.Format.Fill.Visible = 0
        chart.ChartArea.Format.Line.Visible = 0
        chart.PlotArea.Format.Fill.Visible = 0
    except Exception:
        pass
    try:
        chart.ChartGroups(1).DoughnutHoleSize = int(settings.get("hole_size", 68) or 68)
        chart.ChartGroups(1).FirstSliceAngle = 270
    except Exception:
        pass
    # Excel may reject doughnut-point formatting immediately after series creation.
    # Force a chart refresh, pump COM messages, and style each point independently.
    try:
        _com_retry(lambda: chart.Refresh(), attempts=60, delay=0.15)
    except Exception:
        pass
    _pump_messages()
    time.sleep(0.25)
    point_errors = []
    for point_index, point_color, transparent in (
        (1, excel_rgb(theme.get("accent_color"), "118DFF"), False),
        (2, excel_rgb(theme.get("remaining_color"), "E6E6E6"), False),
        (3, None, True),
    ):
        try:
            _style_gauge_point(series, point_index, point_color, transparent)
        except Exception as exc:
            point_errors.append(f"point {point_index}: {exc}")
    if point_errors:
        logger.warning(
            "Gauge point styling partially failed: %s", " | ".join(point_errors)
        )
    # center/min/max values are ordinary cells so they remain live
    value_row = min(43, row + max(2, row_span - 3))
    value_col = min(26, col + max(1, col_span // 3))
    value_end_col = min(26, col + max(2, (col_span * 2) // 3))
    value_range = _com_retry(
        lambda: dashboard_sheet.Range(
            dashboard_sheet.Cells(value_row, value_col),
            dashboard_sheet.Cells(min(43, value_row + 1), value_end_col),
        )
    )
    try:
        value_range.UnMerge()
    except Exception:
        pass
    _com_retry(lambda: value_range.Merge())
    value_cell = _com_retry(lambda: dashboard_sheet.Cells(value_row, value_col))
    _set_formula(value_cell, f"={current_cell.Address}")
    value_cell.NumberFormat = _compact_number_format(measure)
    min_cell = _com_retry(
        lambda: dashboard_sheet.Cells(min(43, row + row_span - 1), col)
    )
    max_label_cell = _com_retry(
        lambda: dashboard_sheet.Cells(
            min(43, row + row_span - 1), min(26, col + col_span - 1)
        )
    )
    min_cell.Value = minimum if minimum is not None else 0
    _set_formula(max_label_cell, f"={max_cell.Address}")
    try:
        value_range.HorizontalAlignment = -4108
        value_range.VerticalAlignment = -4108
        value_cell.Font.Bold = False
        value_cell.Font.Size = 18
        min_cell.Font.Size = 9
        max_label_cell.Font.Size = 9
    except Exception:
        pass
    # Overlay live value/min/max text boxes. They are finalized after RefreshAll.
    try:
        value_box = _com_retry(
            lambda: dashboard_sheet.Shapes.AddTextbox(
                XL_TEXTBOX_HORIZONTAL,
                left + width * 0.25,
                top + height * 0.46,
                width * 0.50,
                max(24.0, height * 0.26),
            )
        )
        value_box.Name = f"GaugeValue_{_value(binding, 'visual_id', 'Gauge')}"
        value_box.Line.Visible = MSO_FALSE
        value_box.Fill.Visible = MSO_FALSE
        value_box.TextFrame2.TextRange.Text = "—"
        value_box.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        value_box.TextFrame2.TextRange.Font.Size = 18
        _set_shape_link_metadata(
            value_box, str(dashboard_sheet.Name), str(current_cell.Address)
        )

        min_box = _com_retry(
            lambda: dashboard_sheet.Shapes.AddTextbox(
                XL_TEXTBOX_HORIZONTAL, left + 2, top + height * 0.72, width * 0.20, 18
            )
        )
        min_box.Line.Visible = MSO_FALSE
        min_box.Fill.Visible = MSO_FALSE
        min_box.TextFrame2.TextRange.Text = str(minimum if minimum is not None else 0)
        min_box.TextFrame2.TextRange.Font.Size = 8

        max_box = _com_retry(
            lambda: dashboard_sheet.Shapes.AddTextbox(
                XL_TEXTBOX_HORIZONTAL,
                left + width * 0.78,
                top + height * 0.72,
                width * 0.20,
                18,
            )
        )
        max_box.Line.Visible = MSO_FALSE
        max_box.Fill.Visible = MSO_FALSE
        max_box.TextFrame2.TextRange.Text = "—"
        max_box.TextFrame2.TextRange.Font.Size = 8
        _set_shape_link_metadata(
            max_box, str(dashboard_sheet.Name), str(max_cell.Address)
        )
    except Exception as exc:
        logger.warning("Gauge live labels could not be created: %s", exc)

    try:
        dashboard_sheet.Columns.Item(helper_col).Hidden = True
    except Exception:
        pass
    logger.info(
        "Created semicircular gauge object=%s current=%s max=%s",
        getattr(chart_obj, "Name", ""),
        current_cell.Address,
        max_cell.Address,
    )
    return {
        "object_name": str(chart_obj.Name),
        "cube_formula": formula,
        "mapped_measure": mapped_measure,
        "visible_gauge_count": 1,
        "render_strategy": "single_series_vertical_semicircle",
    }


def _render_slicer(
    workbook: Any,
    dashboard_sheet: Any,
    binding: Any,
    pivot_info: Optional[Dict[str, Any]],
    field_mapper: Any,
    title: str,
) -> str:
    """Create a native OLAP slicer using the exact PivotTable CubeField.

    Excel is strict about OLAP slicer arguments. Passing only the OLAP unique
    name string can raise E_INVALIDARG (-2147024809), so this implementation
    resolves the CubeField object first and uses Add2/Add fallbacks.
    """
    if (
        not pivot_info
        or pivot_info.get("error")
        or pivot_info.get("pivot_table") is None
    ):
        raise RuntimeError("A connected PivotTable is required for the slicer.")

    pivot_table = pivot_info["pivot_table"]
    field = (
        _value(binding, "slicer_field", None)
        or (list(_value(binding, "rows", []) or []) or [None])[0]
    )
    if not field:
        raise RuntimeError("No slicer field is assigned.")

    mapping = field_mapper.map_field(field, "dimension")
    olap_field = str(mapping.get("excel_olap_field") or "").strip()
    if mapping.get("status") != "mapped" or not olap_field:
        raise RuntimeError(f"Slicer field could not be mapped: {field}")

    left, top, width, height, *_ = _layout_box(binding, dashboard_sheet)

    # Resolve the exact CubeField object from the connected PivotTable.
    cube_field = None
    try:
        cube_field = pivot_table.CubeFields.Item(olap_field)
    except Exception:
        try:
            cube_field = pivot_table.CubeFields(olap_field)
        except Exception:
            pass

    if cube_field is None:
        # Caption fallback for Excel builds that do not accept the unique name.
        wanted = _clean_ref_label(olap_field).casefold()
        try:
            count = int(pivot_table.CubeFields.Count)
            for index in range(1, count + 1):
                candidate = pivot_table.CubeFields.Item(index)
                candidate_name = str(getattr(candidate, "Name", "") or "")
                candidate_caption = str(getattr(candidate, "Caption", "") or "")
                if (
                    candidate_name.casefold() == olap_field.casefold()
                    or candidate_caption.casefold() == wanted
                    or _clean_ref_label(candidate_name).casefold() == wanted
                ):
                    cube_field = candidate
                    break
        except Exception:
            pass

    if cube_field is None:
        raise RuntimeError(
            f"Excel CubeField could not be resolved for slicer field: {olap_field}"
        )

    cache = None
    slicer_cache_name = re.sub(
        r"[^A-Za-z0-9_]",
        "_",
        f"SlicerCache_{title or _clean_ref_label(olap_field)}",
    )[:80]

    # Preferred for modern Excel/OLAP.
    try:
        cache = workbook.SlicerCaches.Add2(
            pivot_table,
            cube_field,
            slicer_cache_name,
        )
    except Exception:
        pass

    # Some Excel versions accept Add2 without the cache-name argument.
    if cache is None:
        try:
            cache = workbook.SlicerCaches.Add2(
                pivot_table,
                cube_field,
            )
        except Exception:
            pass

    # Legacy fallback.
    if cache is None:
        try:
            cache = workbook.SlicerCaches.Add(
                pivot_table,
                cube_field,
                slicer_cache_name,
            )
        except Exception:
            pass

    if cache is None:
        try:
            cache = workbook.SlicerCaches.Add(
                pivot_table,
                cube_field,
            )
        except Exception as exc:
            raise RuntimeError(
                f"Excel could not create an OLAP slicer cache for {olap_field}: {exc}"
            ) from exc

    slicer_name = re.sub(
        r"[^A-Za-z0-9_]",
        "_",
        f"Slicer_{title or _clean_ref_label(olap_field)}",
    )[:80]
    slicer_caption = title or _clean_ref_label(olap_field)

    slicer = None

    # Positional COM signature is more reliable than named arguments.
    try:
        slicer = cache.Slicers.Add(
            dashboard_sheet,
            "",
            slicer_name,
            slicer_caption,
            top,
            left,
            width,
            height,
        )
    except Exception:
        pass

    # Named-argument fallback for other Excel builds.
    if slicer is None:
        try:
            slicer = cache.Slicers.Add(
                SlicerDestination=dashboard_sheet,
                Name=slicer_name,
                Caption=slicer_caption,
                Top=top,
                Left=left,
                Width=width,
                Height=height,
            )
        except Exception as exc:
            raise RuntimeError(
                f"Slicer cache was created, but Excel could not place the slicer: {exc}"
            ) from exc

    try:
        slicer.Caption = slicer_caption
    except Exception:
        pass

    try:
        slicer.NumberOfColumns = max(1, int(width // 90))
    except Exception:
        pass

    try:
        slicer.Shape.Locked = True
    except Exception:
        pass

    logger.info(
        "Created OLAP slicer %s for field %s at left=%.1f top=%.1f width=%.1f height=%.1f",
        getattr(slicer, "Name", slicer_name),
        olap_field,
        left,
        top,
        width,
        height,
    )

    return str(getattr(slicer, "Name", slicer_name))


def _render_textbox_visual(
    dashboard_sheet: Any,
    binding: Any,
    title: str,
    body: str,
    theme: Dict[str, Any],
) -> str:
    left, top, width, height, *_ = _layout_box(binding, dashboard_sheet)
    shape = dashboard_sheet.Shapes.AddTextbox(
        XL_TEXTBOX_HORIZONTAL,
        left,
        top,
        width,
        height,
    )
    try:
        shape.TextFrame2.TextRange.Text = body or title
        shape.TextFrame2.TextRange.Font.Size = 11
        shape.Line.Visible = 0
        shape.Fill.Visible = 0
    except Exception:
        pass
    return str(shape.Name)


def _visual_image_path(binding: Any) -> Optional[str]:
    """Return a Power BI exported visual image or screenshot crop path."""
    candidates = (
        _value(binding, "rendered_image_path", None),
        _value(binding, "visual_image_path", None),
        _value(binding, "screenshot_crop", None),
        _value(binding, "image_path", None),
    )

    for candidate in candidates:
        if not candidate:
            continue

        if isinstance(candidate, dict):
            candidate = (
                candidate.get("path")
                or candidate.get("file_path")
                or candidate.get("image_path")
            )

        if candidate:
            return str(candidate)

    return None


def _render_powerbi_native_image(
    dashboard_sheet: Any,
    binding: Any,
    title: str,
    visual_type: str,
) -> str:
    """Embed a Power BI-native visual image at the exact dashboard position."""
    image_path = _visual_image_path(binding)

    if not image_path:
        raise RuntimeError(
            f"No rendered image was supplied for Power BI-native visual "
            f"'{visual_type}'."
        )

    left, top, width, height, *_ = _layout_box(binding, dashboard_sheet)

    visual_id = str(
        _value(binding, "visual_id", "")
        or _value(binding, "chunk_id", "")
        or visual_type
    )
    shape_name = f"PBI_NATIVE_{re.sub(r'[^A-Za-z0-9_]+', '_', visual_id)[:180]}"
    try:
        dashboard_sheet.Shapes(shape_name).Delete()
    except Exception:
        pass
    picture = dashboard_sheet.Shapes.AddPicture(
        image_path, MSO_FALSE, MSO_SAVE_WITH_DOCUMENT, left, top, width, height
    )
    try:
        picture.Name = shape_name
    except Exception:
        pass

    try:
        picture.LockAspectRatio = MSO_FALSE
        picture.Left = left
        picture.Top = top
        picture.Width = width
        picture.Height = height
        picture.Locked = True
        picture.AlternativeText = f"Power BI native visual: {title or visual_type}"
    except Exception:
        pass

    logger.info(
        "Rendered Power BI-native visual %s as embedded image %s",
        visual_type,
        image_path,
    )

    return str(picture.Name)


def _render_mode_for_visual(visual_type: str, binding: Any) -> str:
    explicit_mode = str(_value(binding, "render_mode", "") or "").casefold()
    if explicit_mode in {
        EXCEL_NATIVE,
        LIVE_APPROXIMATION,
        POWERBI_NATIVE_CAPTURE,
        "powerbi_image",
    }:
        return (
            "powerbi_image"
            if explicit_mode == POWERBI_NATIVE_CAPTURE
            else explicit_mode
        )
    policy_mode = render_mode_for(visual_type)
    return "powerbi_image" if policy_mode == POWERBI_NATIVE_CAPTURE else policy_mode


def _render_unsupported_visual_placeholder(
    dashboard_sheet: Any,
    binding: Any,
    title: str,
    visual_type: str,
    theme: Dict[str, Any],
) -> str:
    """Create a styled placeholder at the exact screenshot position.

    Unsupported Power BI-only visuals still preserve their intended position,
    title, and dimensions instead of shifting the dashboard layout.
    """
    left, top, width, height, *_ = _layout_box(binding, dashboard_sheet)
    shape = dashboard_sheet.Shapes.AddShape(
        XL_RECTANGLE,
        left,
        top,
        width,
        height,
    )
    try:
        shape.Fill.ForeColor.RGB = excel_rgb(theme.get("card_background"), "FFFFFF")
        shape.Line.ForeColor.RGB = excel_rgb(theme.get("border_color"), "D9D9D9")
        shape.TextFrame2.TextRange.Text = (
            f"{title}\n[{visual_type} is not natively supported by Excel]"
        )
        shape.TextFrame2.TextRange.Font.Size = 10
    except Exception:
        pass
    return str(shape.Name)


def _apply_dashboard_surface(excel_app: Any, dashboard_sheet: Any) -> None:
    try:
        for title_row in range(1, 5):
            for title_col in range(1, 5):
                title_cell = dashboard_sheet.Cells(title_row, title_col)
                title_value = str(title_cell.Value or "").strip()
                if re.fullmatch(r"(?i)page\s*\d+", title_value):
                    title_cell.ClearContents()
        dashboard_sheet.Activate()
        excel_app.ActiveWindow.DisplayGridlines = False
        excel_app.ActiveWindow.DisplayHeadings = False
    except Exception:
        pass
    try:
        dashboard_sheet.Cells.Interior.Color = excel_rgb("FFFFFF")
    except Exception:
        pass


def _render_live_pivot_view(
    dashboard_sheet: Any,
    pivot_table: Any,
    binding: Any,
    title: str,
    theme: Dict[str, Any],
) -> str:
    left, top, width, height, row, col, *_ = _layout_box(binding, dashboard_sheet)
    _apply_title_style(dashboard_sheet, row, col, title, theme)

    pivot_sheet_name = ""
    try:
        pivot_sheet_name = str(pivot_table.Parent.Name)
    except Exception:
        pass

    dashboard_name = str(getattr(dashboard_sheet, "Name", "") or "")
    if pivot_sheet_name and pivot_sheet_name.casefold() == dashboard_name.casefold():
        try:
            pivot_table.TableStyle2 = "PivotStyleMedium2"
            pivot_table.ShowTableStyleRowStripes = True
        except Exception:
            pass
        return str(pivot_table.Name)

    pivot_range = pivot_table.TableRange2
    pivot_range.CopyPicture(Appearance=1, Format=2)
    dashboard_sheet.Paste(Destination=dashboard_sheet.Cells(row, col))
    shape = dashboard_sheet.Shapes(dashboard_sheet.Shapes.Count)
    shape.Left = left
    shape.Top = top + 20
    shape.Width = width
    shape.Height = max(100.0, height - 20)

    logger.warning(
        "PivotTable %s was rendered as a static picture because it was created "
        "on worksheet %s instead of dashboard worksheet %s.",
        getattr(pivot_table, "Name", "Unknown"),
        pivot_sheet_name or "Unknown",
        dashboard_name or "Unknown",
    )
    return str(shape.Name)


def supported_visual_matrix() -> Dict[str, List[str]]:
    return {
        "excel_native_live": sorted(EXCEL_NATIVE_VISUALS),
        "powerbi_native_live_approximation": sorted(POWERBI_NATIVE_LIVE_APPROX_VISUALS),
        "powerbi_native_auto_regenerated_image": sorted(POWERBI_NATIVE_IMAGE_VISUALS),
        "unknown_positioned_fallback": ["unknown"],
    }


def render_visual_to_dashboard(
    excel_app: Any,
    workbook: Any,
    dashboard_sheet: Any,
    binding: Any,
    pivot_info: Optional[Dict[str, Any]],
    field_mapper: Any,
    theme: Dict[str, Any],
    show_tech: bool = False,
    cube_filter_refs: Optional[List[Dict[str, Any]]] = None,
    connection_name: Optional[str] = None,
    materialized_formulas: Optional[Dict[str, Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    result: Dict[str, Any] = {
        "visual_id": str(_value(binding, "visual_id", "") or ""),
        "object_name": "",
        "status": "success",
        "error": None,
        "warnings": [],
    }

    cube_filter_refs = cube_filter_refs or []
    materialized_formulas = materialized_formulas or {}

    operation = dict(_value(binding, "render_operation", {}) or {})
    operation_style = dict(
        operation.get("style") or _value(binding, "render_style", {}) or {}
    )
    if operation_style:
        theme = {**theme, **operation_style}

    # Metadata-confirmed render operation controls Excel renderer selection.
    operation_type = str(operation.get("visual_type") or "").strip()
    visual_type = (
        _normalize_visual_type(operation_type)
        if operation_type
        else _resolve_visual_type(binding)
    )
    render_mode = _render_mode_for_visual(visual_type, binding)
    title = str(operation.get("title") or "").strip() or _resolve_title(
        binding, dashboard_sheet
    )
    binding_type = str(_value(binding, "binding_type", "") or "")
    _, _, _, _, row, col, _, _ = _layout_box(binding, dashboard_sheet)

    _apply_dashboard_surface(excel_app, dashboard_sheet)

    logger.info(
        "Rendering visual id=%s resolved_type=%s render_mode=%s "
        "binding_type=%s title=%s",
        result["visual_id"],
        visual_type,
        render_mode,
        binding_type,
        title,
    )

    try:
        # Cards and gauges have dedicated live Excel renderers in this module.
        # Dispatch them before the generic approximation renderer; otherwise the
        # registry's excel_live_approximation mode incorrectly routes them to
        # powerbi_native_live_renderer, where they are intentionally unsupported.
        if _is_gauge_visual(visual_type):
            gauge_result = _render_gauge(
                dashboard_sheet,
                binding,
                pivot_info,
                title,
                theme,
                field_mapper,
                cube_filter_refs,
                connection_name,
                materialized_formulas,
            )
            result.update(gauge_result)
            result["status"] = "live_approximation"

        elif _is_card_visual(visual_type) or (
            binding_type == "cube_formula"
            and visual_type in {"card", "multicard", "multirowcard", "kpi"}
        ):
            card_result = _render_kpi_cards(
                dashboard_sheet,
                binding,
                pivot_info,
                title,
                theme,
                field_mapper,
                cube_filter_refs,
                connection_name,
                materialized_formulas,
            )
            result.update(card_result)
            result["status"] = "live_approximation"

        elif render_mode == "excel_live_approximation":
            approximation = render_powerbi_native_live_approximation(
                visual_type=visual_type,
                dashboard_sheet=dashboard_sheet,
                binding=binding,
                pivot_info=pivot_info,
                title=title,
                formula_cells=None,
            )
            result.update(approximation)
            result["status"] = "live_approximation"

        elif render_mode == "powerbi_image":
            result["status"] = "powerbi_native_image"
            result["object_name"] = _render_powerbi_native_image(
                dashboard_sheet,
                binding,
                title,
                visual_type,
            )

        elif _is_slicer_visual(visual_type) or binding_type == "slicer":
            result["object_name"] = _render_slicer(
                workbook,
                dashboard_sheet,
                binding,
                pivot_info,
                field_mapper,
                title,
            )

        elif binding_type == "connected_pivot":
            if (
                not pivot_info
                or pivot_info.get("error")
                or pivot_info.get("pivot_table") is None
            ):
                message = (pivot_info or {}).get(
                    "error"
                ) or "Connected PivotTable was not created"
                _write_failure_placeholder(dashboard_sheet, row, col, title, message)
                raise RuntimeError(message)

            pivot_table = pivot_info["pivot_table"]

            if _is_chart_visual(visual_type):
                result["object_name"] = _render_live_chart(
                    dashboard_sheet,
                    pivot_table,
                    binding,
                    visual_type,
                    title,
                    theme,
                )
            else:
                result["object_name"] = _render_live_pivot_view(
                    dashboard_sheet,
                    pivot_table,
                    binding,
                    title,
                    theme,
                )

        elif (
            _is_table_visual(visual_type)
            and pivot_info
            and pivot_info.get("pivot_table") is not None
        ):
            result["object_name"] = _render_live_pivot_view(
                dashboard_sheet,
                pivot_info["pivot_table"],
                binding,
                title,
                theme,
            )

        elif visual_type in {"textbox", "text", "shape", "image"}:
            result["object_name"] = _render_textbox_visual(
                dashboard_sheet,
                binding,
                title,
                str(_value(binding, "text", "") or title),
                theme,
            )

        else:
            raise RuntimeError(
                f"No renderer or Power BI-native capture was supplied for visual type '{visual_type}'."
            )

    except Exception as exc:
        logger.error(
            "Error rendering visual %s to Excel COM: %s",
            result["visual_id"],
            exc,
        )
        result["status"] = "failed"
        result["error"] = str(exc)

    layout = _flatten_layout(binding)
    left, top, width, height, *_ = _layout_box(binding, dashboard_sheet)
    logger.info(
        "visual id=%s type=%s mode=%s binding=%s pbix=(%s,%s,%s,%s) "
        "excel=(%.1f,%.1f,%.1f,%.1f) object=%s status=%s",
        result.get("visual_id"),
        visual_type,
        render_mode,
        binding_type,
        layout.get("x"),
        layout.get("y"),
        layout.get("width"),
        layout.get("height"),
        left,
        top,
        width,
        height,
        result.get("object_name"),
        result.get("status"),
    )
    return result


def finalize_live_visual_values(workbook: Any) -> Dict[str, int]:
    """Relink/update gauge and KPI value shapes after refresh completes.

    Excel may reject Shape.Formula while PivotTables are being created. This
    second pass runs after refresh, when Excel is idle, and makes the value
    displays live whenever the Excel build supports linked text boxes.
    """
    result = {"linked": 0, "static_fallback": 0, "failed": 0}

    try:
        sheet_count = int(_com_retry(lambda: workbook.Worksheets.Count))
    except Exception:
        return result

    for sheet_index in range(1, sheet_count + 1):
        try:
            sheet = _com_retry(lambda i=sheet_index: workbook.Worksheets(i))
            shape_count = int(_com_retry(lambda: sheet.Shapes.Count))
        except Exception:
            continue

        for shape_index in range(1, shape_count + 1):
            try:
                shape = _com_retry(lambda i=shape_index: sheet.Shapes.Item(i))
                alternative = str(getattr(shape, "AlternativeText", "") or "")
                if not alternative.startswith("{"):
                    continue
                payload = json.loads(alternative)
                if payload.get("kind") != "live_value":
                    continue

                source_sheet = _com_retry(
                    lambda: workbook.Worksheets(str(payload["sheet"]))
                )
                source_cell = _com_retry(
                    lambda: source_sheet.Range(str(payload["address"]))
                )
                linked = _link_or_fill_shape(shape, source_sheet, source_cell)
                if linked:
                    result["linked"] += 1
                else:
                    result["static_fallback"] += 1
            except Exception:
                result["failed"] += 1

    logger.info(
        "Finalized live visual values: linked=%d static=%d failed=%d",
        result["linked"],
        result["static_fallback"],
        result["failed"],
    )
    return result

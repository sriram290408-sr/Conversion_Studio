"""Live Excel approximations for Power BI-only visual families.

Every renderer in this module remains connected to the Power BI semantic model
through a connected PivotTable or worksheet formula. These are controlled Excel
recreations, not static screenshots.
"""

from __future__ import annotations

import logging
import re
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger("powerbi_native_live_renderer")

XL_COLUMN_CLUSTERED = 51
XL_BAR_CLUSTERED = 57
XL_LINE_MARKERS = 65
XL_AREA_STACKED = 76
XL_MAP = 140
XL_RECTANGLE = 1
XL_CENTER = -4108
XL_LEFT = -4131


def _value(obj: Any, key: str, default: Any = None) -> Any:
    if isinstance(obj, dict):
        return obj.get(key, default)
    return getattr(obj, key, default)


def _flatten_layout(binding: Any) -> Dict[str, Any]:
    layout = dict(_value(binding, "layout", {}) or {})
    for key in ("bbox", "position", "geometry", "bounds"):
        nested = layout.get(key)
        if isinstance(nested, dict):
            layout = {**nested, **layout}
    return layout


def _layout_box(
    binding: Any,
    dashboard_sheet: Any,
) -> Tuple[float, float, float, float, int, int, int, int]:
    """Use PBIX geometry on the same B3:Z43 Excel canvas as native visuals."""
    layout = _flatten_layout(binding)

    row = max(3, min(160, int(layout.get("row") or 5)))
    col = max(2, min(30, int(layout.get("col") or 2)))
    row_span = max(3, min(60, int(layout.get("row_span") or 10)))
    col_span = max(2, min(24, int(layout.get("col_span") or 5)))

    if all(layout.get(k) is not None for k in ("x", "y", "width", "height")):
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
        excel_width = max(1.0, canvas_right - canvas_left)
        excel_height = max(1.0, canvas_bottom - canvas_top)

        x = max(0.0, min(canvas_width, float(layout.get("x") or 0)))
        y = max(0.0, min(canvas_height, float(layout.get("y") or 0)))
        width = max(1.0, min(canvas_width - x, float(layout.get("width") or 1)))
        height = max(1.0, min(canvas_height - y, float(layout.get("height") or 1)))

        left = canvas_left + (x / canvas_width) * excel_width
        top = canvas_top + (y / canvas_height) * excel_height
        object_width = max(60.0, (width / canvas_width) * excel_width)
        object_height = max(45.0, (height / canvas_height) * excel_height)

        row = max(3, min(43, int(round(3 + (y / canvas_height) * 40))))
        col = max(2, min(26, int(round(2 + (x / canvas_width) * 24))))
        row_span = max(3, int(round((height / canvas_height) * 40)))
        col_span = max(2, int(round((width / canvas_width) * 24)))

        return (
            left,
            top,
            object_width,
            object_height,
            row,
            col,
            row_span,
            col_span,
        )

    start = dashboard_sheet.Cells(row, col)
    end = dashboard_sheet.Cells(row + row_span, col + col_span)
    return (
        float(start.Left),
        float(start.Top),
        max(60.0, float(end.Left) - float(start.Left)),
        max(45.0, float(end.Top) - float(start.Top)),
        row,
        col,
        row_span,
        col_span,
    )


def _pivot_or_raise(pivot_info: Optional[Dict[str, Any]]) -> Any:
    if (
        not pivot_info
        or pivot_info.get("error")
        or pivot_info.get("pivot_table") is None
    ):
        raise RuntimeError(
            "A connected OLAP PivotTable is required for this live approximation."
        )
    return pivot_info["pivot_table"]


def _hex_to_rgb_int(value: Any, default: str) -> int:
    text = str(value or default).replace("#", "").strip()
    try:
        return int(text, 16)
    except Exception:
        return int(default, 16)


def _style_chart(chart: Any, title: str, theme: Dict[str, Any]) -> None:
    try:
        chart.HasTitle = True
        chart.ChartTitle.Text = title
        chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 11
    except Exception:
        pass

    try:
        chart.ChartArea.Format.Line.Visible = 0
        chart.ChartArea.Format.Fill.ForeColor.RGB = _hex_to_rgb_int(
            theme.get("background_color"), "FFFFFF"
        )
        chart.PlotArea.Format.Fill.Visible = 0
    except Exception:
        pass

    try:
        series_count = int(chart.SeriesCollection().Count)
        chart.HasLegend = series_count > 1
        if series_count:
            chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = _hex_to_rgb_int(
                theme.get("accent_color"), "118DFF"
            )
            chart.SeriesCollection(1).Format.Line.ForeColor.RGB = _hex_to_rgb_int(
                theme.get("accent_color"), "118DFF"
            )
    except Exception:
        pass


def _source_cell_reference(source_cell: Any) -> str:
    sheet_name = str(source_cell.Parent.Name).replace("'", "''")
    return f"='{sheet_name}'!{source_cell.Address}"


def _write_linked_pivot_grid(
    dashboard_sheet: Any,
    pivot_table: Any,
    row: int,
    col: int,
    row_span: int,
    col_span: int,
    title: str,
    theme: Dict[str, Any],
) -> str:
    """Copy a PivotTable into a live formula grid on the dashboard."""
    source = pivot_table.TableRange2
    source_rows = min(int(source.Rows.Count), max(2, row_span - 1))
    source_cols = min(int(source.Columns.Count), max(1, col_span))

    title_cell = dashboard_sheet.Cells(row, col)
    title_cell.Value = title
    title_cell.Font.Bold = True
    title_cell.Font.Size = 11

    for r in range(1, source_rows + 1):
        for c in range(1, source_cols + 1):
            source_cell = source.Cells(r, c)
            target_cell = dashboard_sheet.Cells(row + r, col + c - 1)
            try:
                target_cell.Formula = _source_cell_reference(source_cell)
            except Exception:
                target_cell.Formula2 = _source_cell_reference(source_cell)

    target_range = dashboard_sheet.Range(
        dashboard_sheet.Cells(row + 1, col),
        dashboard_sheet.Cells(row + source_rows, col + source_cols - 1),
    )
    try:
        target_range.Borders.Color = _hex_to_rgb_int(
            theme.get("border_color"), "D9D9D9"
        )
        target_range.Font.Size = 9
        target_range.Columns.AutoFit()
    except Exception:
        pass

    return f"{dashboard_sheet.Name}!{target_range.Address}"


def render_map_approximation(
    dashboard_sheet: Any,
    binding: Any,
    pivot_info: Optional[Dict[str, Any]],
    title: str,
    theme: Dict[str, Any],
) -> Dict[str, Any]:
    """Use an Excel Filled Map when supported, otherwise a live ranked bar."""
    pivot_table = _pivot_or_raise(pivot_info)
    left, top, width, height, *_ = _layout_box(binding, dashboard_sheet)

    chart_obj = dashboard_sheet.ChartObjects().Add(left, top, width, height)
    chart = chart_obj.Chart
    chart.SetSourceData(pivot_table.TableRange1)

    rendered_as = "excel_filled_map"
    try:
        chart.ChartType = XL_MAP
    except Exception:
        chart.ChartType = XL_BAR_CLUSTERED
        rendered_as = "geographic_ranked_bar"

    _style_chart(chart, title, theme)
    return {
        "object_name": str(chart_obj.Name),
        "rendered_as": rendered_as,
        "live_connected": True,
        "approximation_type": "map",
    }


def render_decomposition_tree_approximation(
    dashboard_sheet: Any,
    binding: Any,
    pivot_info: Optional[Dict[str, Any]],
    title: str,
    theme: Dict[str, Any],
) -> Dict[str, Any]:
    """Represent the decomposition path as a live hierarchical Pivot grid."""
    pivot_table = _pivot_or_raise(pivot_info)
    _, _, _, _, row, col, row_span, col_span = _layout_box(binding, dashboard_sheet)

    try:
        pivot_table.RowAxisLayout(1)
        pivot_table.RepeatAllLabels(2)
        pivot_table.RowGrand = False
        pivot_table.ColumnGrand = False
        pivot_table.TableStyle2 = "PivotStyleMedium2"
    except Exception:
        pass

    object_name = _write_linked_pivot_grid(
        dashboard_sheet,
        pivot_table,
        row,
        col,
        row_span,
        col_span,
        title,
        theme,
    )
    return {
        "object_name": object_name,
        "rendered_as": "live_hierarchical_pivot_grid",
        "live_connected": True,
        "approximation_type": "decomposition_tree",
    }


def render_key_influencers_approximation(
    dashboard_sheet: Any,
    binding: Any,
    pivot_info: Optional[Dict[str, Any]],
    title: str,
    theme: Dict[str, Any],
) -> Dict[str, Any]:
    """Render ranked influencing categories as a live horizontal bar chart."""
    pivot_table = _pivot_or_raise(pivot_info)
    left, top, width, height, *_ = _layout_box(binding, dashboard_sheet)

    try:
        row_field = pivot_table.RowFields(1)
        data_field = pivot_table.DataFields(1)
        row_field.AutoSort(2, data_field.Name)
        pivot_table.RowGrand = False
        pivot_table.ColumnGrand = False
    except Exception:
        pass

    chart_obj = dashboard_sheet.ChartObjects().Add(left, top, width, height)
    chart = chart_obj.Chart
    chart.SetSourceData(pivot_table.TableRange1)
    chart.ChartType = XL_BAR_CLUSTERED
    _style_chart(chart, title, theme)

    return {
        "object_name": str(chart_obj.Name),
        "rendered_as": "live_ranked_influencer_bar",
        "live_connected": True,
        "approximation_type": "key_influencers",
    }


def render_smart_narrative_approximation(
    dashboard_sheet: Any,
    binding: Any,
    pivot_info: Optional[Dict[str, Any]],
    title: str,
    theme: Dict[str, Any],
) -> Dict[str, Any]:
    """Create a formula-driven narrative block without Shape.Formula."""
    pivot_table = _pivot_or_raise(pivot_info)
    _, _, _, _, row, col, row_span, col_span = _layout_box(binding, dashboard_sheet)

    source = pivot_table.DataBodyRange
    source_cell = source.Cells(1, 1)
    value_reference = _source_cell_reference(source_cell)
    template = str(
        _value(binding, "narrative_template", "")
        or _value(binding, "text", "")
        or f"{title}: "
    )

    end_row = row + max(3, row_span) - 1
    end_col = col + max(3, col_span) - 1
    target_range = dashboard_sheet.Range(
        dashboard_sheet.Cells(row, col),
        dashboard_sheet.Cells(end_row, end_col),
    )
    try:
        target_range.UnMerge()
    except Exception:
        pass
    target_range.Merge()

    target = dashboard_sheet.Cells(row, col)
    safe_template = template.replace('"', '""')
    try:
        target.Formula = f'="{safe_template}"&TEXT({value_reference},"#,##0.00")'
    except Exception:
        target.Formula2 = f'="{safe_template}"&TEXT({value_reference},"#,##0.00")'

    try:
        target.WrapText = True
        target.HorizontalAlignment = XL_LEFT
        target.VerticalAlignment = -4160
        target.Font.Size = 11
        target_range.Interior.Color = _hex_to_rgb_int(
            theme.get("card_background"), "FFFFFF"
        )
        target_range.Borders.Color = _hex_to_rgb_int(
            theme.get("border_color"), "D9D9D9"
        )
    except Exception:
        pass

    return {
        "object_name": f"{dashboard_sheet.Name}!{target_range.Address}",
        "rendered_as": "formula_driven_narrative",
        "live_connected": True,
        "approximation_type": "smart_narrative",
    }


def render_qna_approximation(
    dashboard_sheet: Any,
    binding: Any,
    pivot_info: Optional[Dict[str, Any]],
    title: str,
    theme: Dict[str, Any],
) -> Dict[str, Any]:
    """Render the stored Q&A query result as a live chart."""
    pivot_table = _pivot_or_raise(pivot_info)
    left, top, width, height, *_ = _layout_box(binding, dashboard_sheet)

    chart_obj = dashboard_sheet.ChartObjects().Add(left, top, width, height)
    chart = chart_obj.Chart
    chart.SetSourceData(pivot_table.TableRange1)
    chart.ChartType = XL_COLUMN_CLUSTERED
    _style_chart(chart, title, theme)

    return {
        "object_name": str(chart_obj.Name),
        "rendered_as": "live_predefined_query_chart",
        "live_connected": True,
        "approximation_type": "qna",
    }


def render_ribbon_approximation(
    dashboard_sheet: Any,
    binding: Any,
    pivot_info: Optional[Dict[str, Any]],
    title: str,
    theme: Dict[str, Any],
) -> Dict[str, Any]:
    """Approximate ribbon rank movement with a live stacked-area chart."""
    pivot_table = _pivot_or_raise(pivot_info)
    left, top, width, height, *_ = _layout_box(binding, dashboard_sheet)

    chart_obj = dashboard_sheet.ChartObjects().Add(left, top, width, height)
    chart = chart_obj.Chart
    chart.SetSourceData(pivot_table.TableRange1)
    chart.ChartType = XL_AREA_STACKED
    _style_chart(chart, title, theme)

    return {
        "object_name": str(chart_obj.Name),
        "rendered_as": "live_stacked_area_rank_view",
        "live_connected": True,
        "approximation_type": "ribbon_chart",
    }


def render_anomaly_approximation(
    dashboard_sheet: Any,
    binding: Any,
    pivot_info: Optional[Dict[str, Any]],
    title: str,
    theme: Dict[str, Any],
) -> Dict[str, Any]:
    """Render the source series as a live line chart for anomaly review."""
    pivot_table = _pivot_or_raise(pivot_info)
    left, top, width, height, *_ = _layout_box(binding, dashboard_sheet)

    chart_obj = dashboard_sheet.ChartObjects().Add(left, top, width, height)
    chart = chart_obj.Chart
    chart.SetSourceData(pivot_table.TableRange1)
    chart.ChartType = XL_LINE_MARKERS
    _style_chart(chart, title, theme)

    return {
        "object_name": str(chart_obj.Name),
        "rendered_as": "live_line_with_markers",
        "live_connected": True,
        "approximation_type": "anomaly_detection",
    }


def render_scorecard_approximation(
    dashboard_sheet: Any,
    binding: Any,
    pivot_info: Optional[Dict[str, Any]],
    title: str,
    theme: Dict[str, Any],
) -> Dict[str, Any]:
    """Render scorecard metrics as a live linked Pivot grid."""
    pivot_table = _pivot_or_raise(pivot_info)
    _, _, _, _, row, col, row_span, col_span = _layout_box(binding, dashboard_sheet)
    object_name = _write_linked_pivot_grid(
        dashboard_sheet,
        pivot_table,
        row,
        col,
        row_span,
        col_span,
        title,
        theme,
    )
    return {
        "object_name": object_name,
        "rendered_as": "live_scorecard_grid",
        "live_connected": True,
        "approximation_type": "scorecard",
    }


def render_powerbi_native_live_approximation(
    visual_type: str,
    dashboard_sheet: Any,
    binding: Any,
    pivot_info: Optional[Dict[str, Any]],
    title: str,
    formula_cells: Optional[List[Any]] = None,
    theme: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Dispatch a Power BI-only visual to a live Excel approximation."""
    del formula_cells
    visual_type = re.sub(r"[\s_-]+", "", str(visual_type or "").casefold())
    theme = dict(theme or {})

    if visual_type in {"map", "filledmap", "azuremap", "arcgismap", "shapemap"}:
        return render_map_approximation(
            dashboard_sheet, binding, pivot_info, title, theme
        )
    if visual_type == "decompositiontree":
        return render_decomposition_tree_approximation(
            dashboard_sheet, binding, pivot_info, title, theme
        )
    if visual_type == "keyinfluencers":
        return render_key_influencers_approximation(
            dashboard_sheet, binding, pivot_info, title, theme
        )
    if visual_type == "smartnarrative":
        return render_smart_narrative_approximation(
            dashboard_sheet, binding, pivot_info, title, theme
        )
    if visual_type in {"qna", "qnavisual"}:
        return render_qna_approximation(
            dashboard_sheet, binding, pivot_info, title, theme
        )
    if visual_type == "ribbonchart":
        return render_ribbon_approximation(
            dashboard_sheet, binding, pivot_info, title, theme
        )
    if visual_type in {"anomalydetection", "anomaly"}:
        return render_anomaly_approximation(
            dashboard_sheet, binding, pivot_info, title, theme
        )
    if visual_type in {"scorecard", "goals"}:
        return render_scorecard_approximation(
            dashboard_sheet, binding, pivot_info, title, theme
        )

    raise RuntimeError(
        f"No live Excel approximation is implemented for '{visual_type}'."
    )


LIVE_APPROXIMATION_VISUALS = {
    "map",
    "filledmap",
    "azuremap",
    "arcgismap",
    "shapemap",
    "decompositiontree",
    "keyinfluencers",
    "smartnarrative",
    "qna",
    "qnavisual",
    "ribbonchart",
    "anomalydetection",
    "scorecard",
    "goals",
}


__all__ = (
    "LIVE_APPROXIMATION_VISUALS",
    "render_powerbi_native_live_approximation",
)

"""Central visual rendering policy for the Power BI to Excel converter."""

from __future__ import annotations

from typing import Dict

EXCEL_NATIVE = "excel_native"
LIVE_APPROXIMATION = "excel_live_approximation"
POWERBI_NATIVE_CAPTURE = "powerbi_native_capture"

EXCEL_NATIVE_VISUALS = {
    "clusteredcolumnchart",
    "columnchart",
    "stackedcolumnchart",
    "100stackedcolumnchart",
    "clusteredbarchart",
    "barchart",
    "stackedbarchart",
    "100stackedbarchart",
    "linechart",
    "areachart",
    "stackedareachart",
    "lineandclusteredcolumnchart",
    "lineandstackedcolumnchart",
    "combochart",
    "piechart",
    "donutchart",
    "doughnutchart",
    "scatterchart",
    "bubblechart",
    "waterfallchart",
    "funnelchart",
    "treemap",
    "table",
    "tableex",
    "matrix",
    "pivottable",
    "slicer",
    "textbox",
    "text",
    "image",
    "shape",
}

LIVE_APPROXIMATION_VISUALS = {
    "card",
    "multirowcard",
    "kpi",
    "kpivisual",
    "gauge",
    "gaugevisual",
    "filledmap",
    "map",
    "dynamictext",
    "scorecard",
    "goals",
}

POWERBI_NATIVE_CAPTURE_VISUALS = {
    "ribbonchart",
    "decompositiontree",
    "keyinfluencers",
    "smartnarrative",
    "anomalydetection",
    "qna",
    "qnavisual",
    "paginatedreport",
    "powerapps",
    "powervisual",
    "rvisual",
    "pythonvisual",
    "azuremap",
    "azuremaps",
    "arcgismap",
    "arcgis",
    "buttons",
    "bookmarknavigator",
    "pagenavigator",
    "reportpagetooltip",
    "drillthrough",
    "custom",
    "customvisual",
    "appsource",
}

ALIASES: Dict[str, str] = {
    "column_chart": "clusteredcolumnchart",
    "bar_chart": "clusteredbarchart",
    "line_chart": "linechart",
    "area_chart": "areachart",
    "pie_chart": "piechart",
    "donut_chart": "donutchart",
    "doughnut_chart": "donutchart",
    "filled_map": "filledmap",
    "decomposition_tree": "decompositiontree",
    "key_influencers": "keyinfluencers",
    "smart_narrative": "smartnarrative",
}


def normalize_visual_type(value: str) -> str:
    raw = str(value or "").strip().casefold().replace("-", "_").replace(" ", "_")
    compact = raw.replace("_", "")
    return ALIASES.get(raw) or ALIASES.get(compact) or compact


def render_mode_for(value: str) -> str:
    visual_type = normalize_visual_type(value)
    if visual_type in EXCEL_NATIVE_VISUALS:
        return EXCEL_NATIVE
    if visual_type in LIVE_APPROXIMATION_VISUALS:
        return LIVE_APPROXIMATION
    return POWERBI_NATIVE_CAPTURE

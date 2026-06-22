"""Pydantic models for live Power BI semantic-model Excel rendering."""

from __future__ import annotations

from typing import Any, Dict, List, Literal, Optional

from pydantic import BaseModel, ConfigDict, Field

BindingType = Literal[
    "connected_pivot",
    "cube_formula",
    "slicer",
    "navigation",
    "image",
    "placeholder",
]

ConversionMode = Literal[
    "live_semantic_model",
    "interactive_powerbi_connection",
    "interactive_live_semantic_model",
    "standalone",
    "standalone_fallback",
]


class RenderLayout(BaseModel):
    model_config = ConfigDict(extra="allow")

    x: Optional[float] = None
    y: Optional[float] = None
    width: Optional[float] = None
    height: Optional[float] = None
    x_percent: Optional[float] = None
    y_percent: Optional[float] = None
    width_percent: Optional[float] = None
    height_percent: Optional[float] = None
    row: Optional[int] = None
    col: Optional[int] = None
    row_span: Optional[int] = None
    col_span: Optional[int] = None


class RenderOperation(BaseModel):
    model_config = ConfigDict(extra="allow")

    op: str
    visual_id: str
    page_name: str
    visual_type: str
    title: Optional[str] = None
    description: Optional[str] = None
    layout: Dict[str, Any] = Field(default_factory=dict)
    style: Dict[str, Any] = Field(default_factory=dict)
    settings: Dict[str, Any] = Field(default_factory=dict)


class VisualBinding(BaseModel):
    """Normalized rendering instructions for one Power BI visual."""

    model_config = ConfigDict(extra="allow", arbitrary_types_allowed=True)

    visual_id: str
    page_name: str
    visual_type: str
    binding_type: BindingType

    title: Optional[str] = None
    description: Optional[str] = None
    layout: Dict[str, Any] = Field(default_factory=dict)
    render_operation: Dict[str, Any] = Field(default_factory=dict)
    render_style: Dict[str, Any] = Field(default_factory=dict)
    settings: Dict[str, Any] = Field(default_factory=dict)

    rows: List[str] = Field(default_factory=list)
    columns: List[str] = Field(default_factory=list)
    measures: List[Any] = Field(default_factory=list)
    legend: List[str] = Field(default_factory=list)
    filters: List[Dict[str, Any]] = Field(default_factory=list)
    slicer_field: Optional[str] = None
    tooltips: List[str] = Field(default_factory=list)

    pivot_signature: Optional[str] = None
    connection_name: Optional[str] = None
    pivot_name: Optional[str] = None
    source_sheet: Optional[str] = None

    source_status: str = "pending"
    field_mapping_status: str = "pending"
    connection_status: str = "pending"
    refresh_status: str = "not_started"
    render_status: str = "pending"

    mapped_rows: List[str] = Field(default_factory=list)
    mapped_columns: List[str] = Field(default_factory=list)
    mapped_measures: List[str] = Field(default_factory=list)
    mapped_legend: List[str] = Field(default_factory=list)
    mapped_fields: List[Dict[str, Any]] = Field(default_factory=list)
    unmapped_fields: List[Dict[str, Any]] = Field(default_factory=list)
    mapped_formula_chunks: List[str] = Field(default_factory=list)

    cube_formula: Optional[str] = None
    value_cell: Optional[str] = None
    fallback_reason: Optional[str] = None

    warnings: List[str] = Field(default_factory=list)
    errors: List[str] = Field(default_factory=list)


class ConnectionValidationResult(BaseModel):
    model_config = ConfigDict(extra="allow")
    excel_com_available: bool = False
    workbook_opened: bool = False
    connection_found: bool = False
    connection_count: int = 0
    connection_names: List[str] = Field(default_factory=list)
    selected_connection_name: Optional[str] = None
    selected_connection_type: Optional[str] = None
    olap_connection_found: bool = False
    pivot_cache_found: bool = False
    pivot_cache_count: int = 0
    template_pivot_found: bool = False
    template_pivot_name: Optional[str] = None
    template_pivot_sheet: Optional[str] = None
    cube_fields_discovered: bool = False
    cube_field_count: int = 0
    measure_count: int = 0
    dimension_count: int = 0
    refresh_attempted: bool = False
    refresh_success: bool = False
    semantic_model_match: Optional[bool] = None
    semantic_model_match_score: Optional[float] = None
    missing_fields: List[str] = Field(default_factory=list)
    missing_measures: List[str] = Field(default_factory=list)
    usable_for_live_conversion: bool = False
    conversion_mode: ConversionMode = "standalone_fallback"
    warnings: List[str] = Field(default_factory=list)
    errors: List[str] = Field(default_factory=list)


class RefreshResult(BaseModel):
    model_config = ConfigDict(extra="allow")
    refresh_started: bool = False
    refresh_completed: bool = False
    duration_seconds: float = 0.0
    connection_count: int = 0
    connections_refreshed: int = 0
    connections_failed: List[str] = Field(default_factory=list)
    pivot_cache_count: int = 0
    pivot_caches_refreshed: int = 0
    pivot_caches_failed: List[str] = Field(default_factory=list)
    async_queries_completed: bool = False
    calculation_completed: bool = False
    timeout: bool = False
    timeout_seconds: int = 0
    warnings: List[str] = Field(default_factory=list)
    errors: List[str] = Field(default_factory=list)


class LiveConversionResult(BaseModel):
    model_config = ConfigDict(extra="allow")
    status: str = "pending"
    conversion_mode: ConversionMode = "standalone_fallback"
    output_path: Optional[str] = None
    connection_validation: Optional[ConnectionValidationResult] = None
    refresh_result: Optional[RefreshResult] = None
    visual_bindings: List[VisualBinding] = Field(default_factory=list)
    field_mapping: Dict[str, Any] = Field(
        default_factory=lambda: {"mapped": [], "unmapped": []}
    )
    discovered_cubefields: List[Any] = Field(default_factory=list)
    verification_result: Dict[str, Any] = Field(default_factory=dict)
    render_results: List[Dict[str, Any]] = Field(default_factory=list)
    rendered_visuals: int = 0
    failed_visuals: int = 0
    warning_count: int = 0
    warnings: List[str] = Field(default_factory=list)
    errors: List[str] = Field(default_factory=list)


class InteractiveSessionStatus(BaseModel):
    model_config = ConfigDict(extra="allow")
    session_id: str
    state: str = "created"
    pivot_tables_created: int = 0
    pivot_tables_reused: int = 0
    pivot_charts_created: int = 0
    slicers_created: int = 0
    cube_formulas_created: int = 0
    cube_field_count: int = 0
    semantic_match_score: float = 0.0
    selected_connection_name: str = ""
    conversion_mode: Optional[str] = None
    connection_preserved: Optional[bool] = None
    post_save_verification_passed: Optional[bool] = None
    download_url: Optional[str] = None
    report_url: Optional[str] = None
    warnings: List[str] = Field(default_factory=list)
    errors: List[str] = Field(default_factory=list)
    last_activity_age_seconds: float = 0.0


__all__ = [
    "RenderLayout",
    "RenderOperation",
    "VisualBinding",
    "ConnectionValidationResult",
    "RefreshResult",
    "LiveConversionResult",
    "InteractiveSessionStatus",
]

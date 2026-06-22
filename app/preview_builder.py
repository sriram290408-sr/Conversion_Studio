import json
from html import escape
from typing import Dict, Any, List

# Plotly is intentionally not used in the preview builder.
# Returning Plotly HTML without loading Plotly JS in the frontend caused
# a large blank area above the preview cards.
go = None

# ---------------------------------------------------------------------------
# Safe helper functions
# ---------------------------------------------------------------------------


def safe(value: Any) -> str:
    if value is None:
        return "—"
    return escape(str(value))


def list_text(value: Any) -> str:
    if value is None:
        return "—"
    if isinstance(value, list):
        return ", ".join(str(v) for v in value) if value else "—"
    return str(value)


def safe_list(value: Any) -> str:
    return safe(list_text(value))


def dataframe_records(items: List[Dict]) -> List[Dict]:
    """
    Kept the same function name for compatibility, but removed Pandas usage.
    The preview page should be card-only, so this returns plain JSON-safe records.
    """
    if not items:
        return []

    records = []
    for item in items:
        clean_item = {}
        for key, value in item.items():
            if value is None:
                clean_item[key] = "—"
            elif isinstance(value, (str, int, float, bool)):
                clean_item[key] = value
            elif isinstance(value, list):
                clean_item[key] = ", ".join(str(v) for v in value) if value else "—"
            else:
                clean_item[key] = str(value)
        records.append(clean_item)

    return records


def _chips(items):
    if not items:
        return "<span style='font-size:10px;color:#475569;'>None</span>"

    return "".join(
        [
            f"<span style='display:inline-block;background:rgba(30,41,59,0.8);border:1px solid rgba(71,85,105,0.5);padding:2px 6px;border-radius:4px;font-size:10px;font-family:monospace;color:#cbd5e1;margin:2px;'>{safe(it)}</span>"
            for it in items
        ]
    )


# ---------------------------------------------------------------------------
# Individual card builders
# ---------------------------------------------------------------------------


def _build_table_card(t: Dict) -> str:
    columns_val = t.get("columns", []) or []

    if columns_val:
        columns_chips = "".join(
            [
                f"<span class='chip' style='display:inline-block;background:rgba(30,41,59,0.8);border:1px solid rgba(71,85,105,0.6);color:#cbd5e1;font-size:10px;padding:2px 8px;border-radius:4px;margin:2px;'>{safe(c)}</span>"
                for c in columns_val
            ]
        )
    else:
        columns_chips = "<span style='color:#64748b;font-size:10px;'>No columns</span>"

    return f"""
    <div class="preview-card">
        <div style='display:flex;justify-content:space-between;align-items:flex-start;gap:8px;'>
            <div>
                <span style='display:inline-block;padding:2px 8px;background:rgba(14,144,235,0.1);color:#38a9f8;border:1px solid rgba(14,144,235,0.2);font-size:9px;font-weight:700;border-radius:4px;text-transform:uppercase;'>Table Chunk</span>
                <h4 style='margin:6px 0 2px;font-size:14px;font-weight:600;color:#fff;'>{safe(t.get('table_name'))}</h4>
                <p style='margin:0;font-size:10px;color:#64748b;font-family:monospace;'>Chunk ID: {safe(t.get('chunk_id'))}</p>
            </div>
            <div style='text-align:right;'>
                <p style='margin:0;font-size:12px;color:#94a3b8;font-family:monospace;'>Excel: {safe(t.get('excel_table_name'))}</p>
                <p style='margin:2px 0 0;font-size:10px;color:#64748b;font-family:monospace;'>Sheet: {safe(t.get('hidden_sheet'))}</p>
            </div>
        </div>
        <div style='margin-top:12px;'>
            <span style='font-size:9px;color:#64748b;text-transform:uppercase;font-weight:700;letter-spacing:0.05em;'>Columns</span>
            <div style='display:flex;flex-wrap:wrap;gap:4px;margin-top:6px;'>
                {columns_chips}
            </div>
        </div>
    </div>"""


def _build_relationship_card(r: Dict) -> str:
    return f"""
    <div class="preview-card" style='display:flex;align-items:center;gap:16px;'>
        <div style='padding:10px;background:rgba(99,102,241,0.1);border:1px solid rgba(99,102,241,0.2);border-radius:12px;flex-shrink:0;font-size:18px;'>&#x26D4;</div>
        <div style='flex-grow:1;min-width:0;'>
            <div style='display:flex;align-items:center;gap:8px;margin-bottom:4px;'>
                <span style='padding:2px 8px;background:rgba(99,102,241,0.1);color:#818cf8;border:1px solid rgba(99,102,241,0.2);font-size:9px;font-weight:700;border-radius:4px;text-transform:uppercase;'>Relationship</span>
                <span style='font-size:10px;color:#64748b;font-family:monospace;'>{safe(r.get('chunk_id'))}</span>
            </div>
            <div style='display:flex;flex-wrap:wrap;align-items:center;gap:8px;margin-top:6px;font-size:13px;font-weight:600;color:#fff;'>
                <span>{safe(r.get('from_table'))}[{safe(r.get('from_column'))}]</span>
                <span style='color:#475569;'>&rarr;</span>
                <span>{safe(r.get('to_table'))}[{safe(r.get('to_column'))}]</span>
            </div>
            <p style='margin:4px 0 0;font-size:10px;color:#64748b;'>Type: {safe(r.get('relationship_type', 'many_to_one'))}</p>
        </div>
    </div>"""


def _build_formula_card(f: Dict) -> str:
    """Build a formula preview card for normal Excel and CUBE conversions."""
    status = str(f.get("conversion_status", "unknown") or "unknown")
    cube_formula = str(f.get("cube_formula") or "").strip()
    excel_formula = str(f.get("excel_formula") or "").strip()
    displayed_formula = cube_formula or excel_formula or "—"

    chunk_type = str(f.get("chunk_type", "") or "").lower()
    conversion_type = str(f.get("conversion_type", "") or "").lower()
    source_key = str(f.get("conversion_source", "") or "")
    is_cube = bool(
        cube_formula
        or chunk_type == "cube_formula_chunk"
        or "cube" in conversion_type
        or "cube" in source_key.lower()
        or displayed_formula.upper().startswith(
            ("=CUBEVALUE(", "=CUBEMEMBER(", "=CUBESET(")
        )
    )

    if status == "converted":
        badge_html = '<span class="status-badge status-converted">converted</span>'
    elif status == "needs_review":
        badge_html = '<span class="status-badge status-review">needs_review</span>'
    else:
        badge_html = f'<span class="status-badge status-review">{safe(status)}</span>'

    src_map = {
        "huggingface": "Hugging Face API",
        "huggingface_router": "Hugging Face Router",
        "rule_based": "Rule Engine",
        "rule_based_simple_aggregation": "Rule Engine",
        "rule_based_fallback_hf_unavailable": "Rule Engine (HF Offline)",
        "rule_based_fallback_hf_disabled": "Rule Engine (HF Disabled)",
        "rule_based_fallback_relationship_filter": "Rule Engine (Relationship Fallback)",
        "rule_based_measure_reference": "Rule Engine (Measure Reference)",
        "cube_measure_reference": "Excel CUBE Measure Reference",
        "cube_member_reference": "Excel CUBE Member Reference",
        "cube_category_measure": "Excel CUBE Category Measure",
        "tmdl_metadata_rule_based": "TMDL Metadata / CUBE Rule Engine",
    }
    source_text = src_map.get(source_key, source_key or "Unknown")

    req_tables = f.get("required_tables", []) or []
    req_sheets = f.get("required_hidden_sheets", []) or []
    map_tables = f.get("mapped_table_chunks", []) or []
    map_rels = f.get("mapped_relationship_chunks", []) or []

    connection_name = (
        f.get("cube_connection")
        or f.get("connection_name")
        or f.get("cube_connection_name")
        or "ThisWorkbookDataModel"
    )
    measure_member = f.get("cube_measure_member") or f.get("measure_member") or ""
    dimension_member = f.get("cube_dimension_member") or f.get("dimension_member") or ""

    notes_section = ""
    hf_err = f.get("hf_error")
    notes_val = f.get("notes") or ""
    if hf_err or notes_val:
        content = ""
        if hf_err:
            content += (
                f"<strong style='color:#f87171;'>HF Error:</strong> {safe(hf_err)}<br/>"
            )
        if notes_val:
            content += f"<strong>Note:</strong> {safe(notes_val)}"
        notes_section = f"""
        <div style='padding:10px;background:rgba(15,23,42,0.5);border:1px solid rgba(30,41,59,0.8);border-radius:8px;font-size:10px;color:#94a3b8;display:flex;gap:8px;margin-top:12px;'>
            <span style='flex-shrink:0;'>&#x2139;&#xFE0F;</span>
            <p style='margin:0;line-height:1.5;'>{content}</p>
        </div>"""

    formula_badge = "CUBE Formula Chunk" if is_cube else "Formula Chunk"
    formula_heading = (
        "Converted Excel CUBE Formula" if is_cube else "Converted Excel Formula"
    )
    formula_color = "#67e8f9" if is_cube else "#e2e8f0"

    cube_details = ""
    if is_cube:
        cube_details = f"""
        <div style='display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px;'>
            <div style='padding:10px;background:rgba(6,182,212,0.06);border:1px solid rgba(6,182,212,0.18);border-radius:8px;'>
                <span style='font-size:9px;color:#67e8f9;text-transform:uppercase;font-weight:700;display:block;margin-bottom:4px;'>CUBE Connection</span>
                <span style='font-family:monospace;font-size:11px;color:#cffafe;'>{safe(connection_name)}</span>
            </div>
            <div style='padding:10px;background:rgba(6,182,212,0.06);border:1px solid rgba(6,182,212,0.18);border-radius:8px;'>
                <span style='font-size:9px;color:#67e8f9;text-transform:uppercase;font-weight:700;display:block;margin-bottom:4px;'>Formula Type</span>
                <span style='font-family:monospace;font-size:11px;color:#cffafe;'>{safe(f.get('conversion_type') or 'cube_measure_reference')}</span>
            </div>
            {f"<div><span style='font-size:9px;color:#64748b;text-transform:uppercase;font-weight:700;display:block;margin-bottom:4px;'>Measure Member</span><span style='font-family:monospace;font-size:11px;color:#cbd5e1;'>{safe(measure_member)}</span></div>" if measure_member else ""}
            {f"<div><span style='font-size:9px;color:#64748b;text-transform:uppercase;font-weight:700;display:block;margin-bottom:4px;'>Dimension Member</span><span style='font-family:monospace;font-size:11px;color:#cbd5e1;'>{safe(dimension_member)}</span></div>" if dimension_member else ""}
        </div>"""

    return f"""
    <div class="preview-card">
        <div style='display:flex;justify-content:space-between;align-items:flex-start;gap:8px;margin-bottom:12px;'>
            <div>
                <span style='padding:2px 8px;background:rgba(16,185,129,0.1);color:#34d399;border:1px solid rgba(16,185,129,0.2);font-size:9px;font-weight:700;border-radius:4px;text-transform:uppercase;'>{formula_badge}</span>
                <h4 style='margin:6px 0 2px;font-size:14px;font-weight:600;color:#fff;'>{safe(f.get('measure_name'))}</h4>
                <p style='margin:0;font-size:10px;color:#64748b;font-family:monospace;'>Chunk ID: {safe(f.get('chunk_id'))}</p>
            </div>
            <div style='display:flex;flex-direction:column;align-items:flex-end;gap:4px;'>
                {badge_html}
                <span style='font-size:9px;color:#64748b;'>Source: {safe(source_text)}</span>
            </div>
        </div>

        <div style='background:rgba(2,6,23,0.8);border:1px solid rgba(15,23,42,0.9);border-radius:8px;padding:12px;margin-bottom:12px;'>
            <span style='font-size:9px;color:#64748b;text-transform:uppercase;font-weight:700;letter-spacing:0.05em;display:block;margin-bottom:6px;'>Original DAX Formula</span>
            <pre class="formula-code" style='margin:0;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;user-select:all;color:#a7f3d0;'>{safe(f.get('dax_formula'))}</pre>
        </div>

        <div style='background:rgba(2,6,23,0.8);border:1px solid rgba(6,182,212,0.25);border-radius:8px;padding:12px;margin-bottom:12px;'>
            <span style='font-size:9px;color:#67e8f9;text-transform:uppercase;font-weight:700;letter-spacing:0.05em;display:block;margin-bottom:6px;'>{formula_heading}</span>
            <pre class="formula-code" style='margin:0;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;user-select:all;color:{formula_color};'>{safe(displayed_formula)}</pre>
        </div>

        {cube_details}

        <div style='display:grid;grid-template-columns:1fr 1fr;gap:12px;'>
            <div>
                <span style='font-size:9px;color:#64748b;text-transform:uppercase;font-weight:700;display:block;margin-bottom:4px;'>Required Tables</span>
                <div style='display:flex;flex-wrap:wrap;gap:4px;'>{_chips(req_tables)}</div>
            </div>
            <div>
                <span style='font-size:9px;color:#64748b;text-transform:uppercase;font-weight:700;display:block;margin-bottom:4px;'>Hidden Sheets</span>
                <div style='display:flex;flex-wrap:wrap;gap:4px;'>{_chips(req_sheets)}</div>
            </div>
            <div>
                <span style='font-size:9px;color:#64748b;text-transform:uppercase;font-weight:700;display:block;margin-bottom:4px;'>Mapped Table Chunks</span>
                <div style='display:flex;flex-wrap:wrap;gap:4px;'>{_chips(map_tables)}</div>
            </div>
            <div>
                <span style='font-size:9px;color:#64748b;text-transform:uppercase;font-weight:700;display:block;margin-bottom:4px;'>Mapped Relationships</span>
                <div style='display:flex;flex-wrap:wrap;gap:4px;'>{_chips(map_rels)}</div>
            </div>
        </div>
        {notes_section}
    </div>"""


def _build_visual_card(v: Dict) -> str:
    hint = v.get("excel_conversion_hint", {}) or {}
    map_t = v.get("mapped_table_chunks", []) or []
    map_f = v.get("mapped_formula_chunks", []) or []
    map_r = v.get("mapped_relationship_chunks", []) or []

    return f"""
    <div class="preview-card">
        <div style='display:flex;justify-content:space-between;align-items:flex-start;gap:8px;margin-bottom:12px;'>
            <div>
                <span style='padding:2px 8px;background:rgba(168,85,247,0.1);color:#c084fc;border:1px solid rgba(168,85,247,0.2);font-size:9px;font-weight:700;border-radius:4px;text-transform:uppercase;'>Visual Chunk</span>
                <h4 style='margin:6px 0 2px;font-size:14px;font-weight:600;color:#fff;'>{safe(v.get('visual_title'))}</h4>
                <p style='margin:0;font-size:10px;color:#64748b;font-family:monospace;'>Chunk ID: {safe(v.get('chunk_id'))}</p>
            </div>
            <div style='text-align:right;'>
                <span style='display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;background:rgba(30,41,59,0.8);border:1px solid rgba(71,85,105,0.5);color:#cbd5e1;'>
                    {safe(v.get('visual_type'))}
                </span>
                <p style='margin:4px 0 0;font-size:10px;color:#64748b;font-weight:600;'>Page: {safe(v.get('page_name'))}</p>
            </div>
        </div>

        <div style='background:rgba(2,6,23,0.8);border:1px solid rgba(15,23,42,0.9);border-radius:8px;padding:12px;margin-bottom:12px;'>
            <span style='font-size:9px;color:#64748b;text-transform:uppercase;font-weight:700;display:block;margin-bottom:8px;'>Excel Conversion Hints</span>
            <div style='display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:11px;font-family:monospace;color:#cbd5e1;'>
                <div><span style='color:#64748b;font-family:sans-serif;'>Type:</span> {safe(hint.get('output_type'))}</div>
                <div><span style='color:#64748b;font-family:sans-serif;'>Target:</span> {safe(hint.get('target_sheet'))}</div>
                <div><span style='color:#64748b;font-family:sans-serif;'>Axis:</span> {safe(hint.get('axis'))}</div>
                <div><span style='color:#64748b;font-family:sans-serif;'>Values:</span> {safe(hint.get('values'))}</div>
            </div>
        </div>

        <div style='display:grid;grid-template-columns:1fr;gap:8px;'>
            <div>
                <span style='font-size:9px;color:#64748b;text-transform:uppercase;font-weight:700;display:block;margin-bottom:4px;'>Mapped Table Chunks</span>
                <div style='display:flex;flex-wrap:wrap;gap:4px;'>{_chips(map_t)}</div>
            </div>
            <div>
                <span style='font-size:9px;color:#64748b;text-transform:uppercase;font-weight:700;display:block;margin-bottom:4px;'>Mapped Formula Chunks</span>
                <div style='display:flex;flex-wrap:wrap;gap:4px;'>{_chips(map_f)}</div>
            </div>
            <div>
                <span style='font-size:9px;color:#64748b;text-transform:uppercase;font-weight:700;display:block;margin-bottom:4px;'>Mapped Relationship Chunks</span>
                <div style='display:flex;flex-wrap:wrap;gap:4px;'>{_chips(map_r)}</div>
            </div>
        </div>
    </div>"""


def _build_visual_description_card(v: Dict) -> str:
    map_t = v.get("mapped_table_chunks", []) or []
    map_f = v.get("mapped_formula_chunks", []) or []
    hint = v.get("excel_conversion_hint", {}) or {}

    return f"""
    <div class="preview-card">
        <div style='display:flex;justify-content:space-between;align-items:flex-start;gap:8px;margin-bottom:12px;'>
            <div>
                <span style='padding:2px 8px;background:rgba(245,158,11,0.1);color:#fbbf24;border:1px solid rgba(245,158,11,0.2);font-size:9px;font-weight:700;border-radius:4px;text-transform:uppercase;'>Visual Description</span>
                <h4 style='margin:6px 0 2px;font-size:14px;font-weight:600;color:#fff;'>{safe(v.get('visual_title'))}</h4>
                <p style='margin:0;font-size:10px;color:#64748b;font-family:monospace;'>Chunk ID: {safe(v.get('chunk_id'))}</p>
            </div>
            <div style='text-align:right;'>
                <span style='display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;background:rgba(30,41,59,0.8);border:1px solid rgba(71,85,105,0.5);color:#cbd5e1;'>
                    {safe(v.get('visual_type'))}
                </span>
                <p style='margin:4px 0 0;font-size:10px;color:#64748b;font-weight:600;'>Page: {safe(v.get('page_name'))}</p>
            </div>
        </div>

        <div style='background:rgba(251,191,36,0.05);border:1px solid rgba(251,191,36,0.15);border-radius:8px;padding:12px;margin-bottom:12px;color:#fde68a;font-size:12px;line-height:1.6;'>
            <strong>Business Summary:</strong><br/>
            {safe(v.get('visual_description'))}
        </div>

        <div style='display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px;'>
            <div>
                <span style='font-size:9px;color:#64748b;text-transform:uppercase;font-weight:700;display:block;margin-bottom:4px;'>Axis</span>
                <span style='font-family:monospace;font-size:11px;color:#cbd5e1;'>{safe(hint.get('axis'))}</span>
            </div>
            <div>
                <span style='font-size:9px;color:#64748b;text-transform:uppercase;font-weight:700;display:block;margin-bottom:4px;'>Values</span>
                <span style='font-family:monospace;font-size:11px;color:#cbd5e1;'>{safe(hint.get('values'))}</span>
            </div>
        </div>

        <div style='display:grid;grid-template-columns:1fr 1fr;gap:12px;'>
            <div>
                <span style='font-size:9px;color:#64748b;text-transform:uppercase;font-weight:700;display:block;margin-bottom:4px;'>Mapped Tables</span>
                <div style='display:flex;flex-wrap:wrap;gap:4px;'>{_chips(map_t)}</div>
            </div>
            <div>
                <span style='font-size:9px;color:#64748b;text-transform:uppercase;font-weight:700;display:block;margin-bottom:4px;'>Mapped Formulas</span>
                <div style='display:flex;flex-wrap:wrap;gap:4px;'>{_chips(map_f)}</div>
            </div>
        </div>
    </div>"""


# ---------------------------------------------------------------------------
# Preview builders - card-only
# ---------------------------------------------------------------------------


def build_tables_preview(table_chunks: List[Dict]) -> Dict[str, Any]:
    if not table_chunks:
        return {
            "html": '<div class="empty-tab">No table chunks found</div>',
            "records": [],
            "count": 0,
        }

    cards_html = "\n".join(_build_table_card(t) for t in table_chunks)

    return {
        "html": f"<div class='space-y-4'>{cards_html}</div>",
        "records": dataframe_records(table_chunks),
        "count": len(table_chunks),
    }


def build_relationships_preview(relationship_chunks: List[Dict]) -> Dict[str, Any]:
    """
    Build relationship preview without Plotly.

    Reason:
    The frontend preview page does not load Plotly JS. When fig.to_html()
    is returned with include_plotlyjs=False, the browser shows a large blank
    Plotly container before the relationship cards. This function returns
    only card HTML, so the cards start immediately below the tab header.
    """
    if not relationship_chunks:
        return {
            "html": '<div class="empty-tab">No relationship chunks found</div>',
            "records": [],
            "count": 0,
        }

    cards_html = "\n".join(_build_relationship_card(r) for r in relationship_chunks)

    return {
        "html": f"<div class='space-y-4'>{cards_html}</div>",
        "records": dataframe_records(relationship_chunks),
        "count": len(relationship_chunks),
    }


def build_formulas_preview(formula_chunks: List[Dict]) -> Dict[str, Any]:
    if not formula_chunks:
        return {
            "html": '<div class="empty-tab">No formula chunks found</div>',
            "records": [],
            "count": 0,
        }

    cards_html = "\n".join(_build_formula_card(f) for f in formula_chunks)

    return {
        "html": f"<div class='space-y-4'>{cards_html}</div>",
        "records": dataframe_records(formula_chunks),
        "count": len(formula_chunks),
    }


def build_visuals_preview(visual_chunks: List[Dict]) -> Dict[str, Any]:
    """
    Build visual preview without Plotly.

    Reason:
    The frontend preview page does not load Plotly JS. Returning Plotly HTML
    creates a large empty area before the visual cards. This function returns
    only visual cards for a clean preview layout.
    """
    if not visual_chunks:
        return {
            "html": '<div class="empty-tab">No visual chunks found</div>',
            "records": [],
            "count": 0,
        }

    cards_html = "\n".join(_build_visual_card(v) for v in visual_chunks)

    return {
        "html": f"<div class='space-y-4'>{cards_html}</div>",
        "records": dataframe_records(visual_chunks),
        "count": len(visual_chunks),
    }


def build_visual_descriptions_preview(visual_chunks: List[Dict]) -> Dict[str, Any]:
    if not visual_chunks:
        return {
            "html": '<div class="empty-tab">No visual descriptions found</div>',
            "records": [],
            "count": 0,
        }

    cards_html = "\n".join(_build_visual_description_card(v) for v in visual_chunks)

    return {
        "html": f"<div class='space-y-4'>{cards_html}</div>",
        "records": dataframe_records(visual_chunks),
        "count": len(visual_chunks),
    }


# ---------------------------------------------------------------------------
# Metadata / Live Excel / Field Mapping preview builders
# ---------------------------------------------------------------------------


def _build_metadata_analysis_preview(analysis: Dict[str, Any]) -> Dict[str, Any]:
    """Build the metadata intelligence preview tab."""
    if not analysis:
        return {
            "html": '<div class="empty-tab">No metadata analysis found</div>',
            "records": [],
            "count": 0,
        }

    counts = analysis.get("overall_counts", {}) or {}
    visual_categories = analysis.get("visual_category_count", {}) or {}
    visual_types = analysis.get("visual_type_count", {}) or {}
    page_analysis = analysis.get("page_analysis", []) or []
    table_roles = analysis.get("table_role_analysis", []) or []
    possible_int = analysis.get("possible_intermediate_tables", []) or []
    possible_tmp = analysis.get("possible_temporary_tables", []) or []
    biz_cats = analysis.get("business_category_fields", {}) or {}
    top_fields = analysis.get("top_used_fields", []) or []
    ai_deep = analysis.get("ai_deep_analysis", {}) or {}

    has_data = any(
        [
            counts,
            visual_categories,
            visual_types,
            page_analysis,
            table_roles,
            ai_deep,
        ]
    )
    if not has_data:
        return {
            "html": '<div class="empty-tab">No analysis data found</div>',
            "records": [],
            "count": 0,
        }

    def row(label: str, value: Any) -> str:
        return (
            "<tr>"
            f"<td style='padding:4px 12px 4px 0;font-weight:700;color:#cbd5e1;white-space:nowrap;'>{safe(label)}</td>"
            f"<td style='padding:4px 0;color:#e2e8f0;'>{safe(value)}</td>"
            "</tr>"
        )

    html = "<div class='preview-card metadata-analysis' style='font-size:12px;line-height:1.6;'>"

    html += "<h4 style='margin:0 0 8px;font-size:14px;color:#fff;'>Metadata Intelligence Summary</h4>"
    html += "<table style='border-collapse:collapse;margin-bottom:12px;'>"
    count_labels = {
        "total_pages": "Total Pages",
        "total_visuals": "Total Visuals",
        "total_model_tables": "Model Tables",
        "visual_inferred_tables": "Visual-Inferred Tables",
        "tables_used_in_visuals": "Tables Used in Visuals",
        "unique_fields_used": "Unique Fields Used",
        "unique_measures_used": "Unique Measures Used",
        "total_relationships": "Relationships",
        "total_formulas": "DAX / Formula Chunks",
        "ai_insight_chunks": "AI Insight Chunks",
        "page_chunks": "Page Chunks",
    }
    for key, label in count_labels.items():
        if key in counts:
            html += row(label, counts.get(key, 0))
    html += "</table>"

    if ai_deep:
        html += "<div style='margin:10px 0;padding:10px;background:rgba(37,99,235,0.08);border:1px solid rgba(37,99,235,0.2);border-radius:8px;'>"
        html += "<strong style='color:#93c5fd;'>AI Deep Analysis</strong>"
        if isinstance(ai_deep, dict):
            html += f"<p style='margin:6px 0 0;color:#cbd5e1;'>Source: {safe(ai_deep.get('source', '—'))}</p>"
            html += f"<p style='margin:2px 0 0;color:#cbd5e1;'>Visual insights: {safe(len(ai_deep.get('visual_insights', []) or []))}</p>"
            if ai_deep.get("hf_error"):
                html += f"<p style='margin:2px 0 0;color:#fca5a5;'>HF warning: {safe(ai_deep.get('hf_error'))}</p>"
            if ai_deep.get("hf_warnings"):
                html += f"<p style='margin:2px 0 0;color:#fbbf24;'>Warnings: {safe_list(ai_deep.get('hf_warnings'))}</p>"
        html += "</div>"

    if table_roles:
        html += "<h4 style='margin:12px 0 8px;font-size:13px;color:#fff;'>Detected Tables</h4>"
        html += (
            "<table style='border-collapse:collapse;width:100%;margin-bottom:12px;'>"
        )
        html += "<tr style='background:#1e293b;color:#facc15;font-weight:700;'><td style='padding:5px 8px;'>Table</td><td style='padding:5px 8px;'>Role</td><td style='padding:5px 8px;'>Source</td><td style='padding:5px 8px;text-align:right;'>Uses</td></tr>"
        for i, rec in enumerate(table_roles):
            bg = "#111827" if i % 2 == 0 else "#1e293b"
            html += f"<tr style='background:{bg};'>"
            html += f"<td style='padding:4px 8px;color:#e5e7eb;'>{safe(rec.get('table_name'))}</td>"
            html += f"<td style='padding:4px 8px;color:#94a3b8;'>{safe(rec.get('role'))}</td>"
            html += f"<td style='padding:4px 8px;color:#38bdf8;'>{safe(rec.get('source'))}</td>"
            html += f"<td style='padding:4px 8px;text-align:right;color:#cbd5e1;'>{safe(rec.get('usage_count', 0))}</td>"
            html += "</tr>"
        html += "</table>"

    html += "<div style='display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px;'>"
    html += "<div style='padding:10px;background:rgba(15,23,42,0.7);border-radius:8px;border:1px solid rgba(51,65,85,0.8);'>"
    html += "<strong style='color:#fff;'>Intermediate / Joined Tables</strong>"
    if possible_int:
        html += (
            "<ul style='margin:6px 0 0;padding-left:18px;color:#cbd5e1;'>"
            + "".join(f"<li>{safe(t)}</li>" for t in possible_int)
            + "</ul>"
        )
    else:
        html += "<p style='color:#34d399;margin:6px 0 0;'>None detected.</p>"
    html += "</div>"
    html += "<div style='padding:10px;background:rgba(15,23,42,0.7);border-radius:8px;border:1px solid rgba(51,65,85,0.8);'>"
    html += "<strong style='color:#fff;'>Temporary / Staging Tables</strong>"
    if possible_tmp:
        html += (
            "<ul style='margin:6px 0 0;padding-left:18px;color:#cbd5e1;'>"
            + "".join(f"<li>{safe(t)}</li>" for t in possible_tmp)
            + "</ul>"
        )
    else:
        html += "<p style='color:#34d399;margin:6px 0 0;'>No clearly marked temporary tables detected.</p>"
    html += "</div></div>"

    if visual_categories:
        html += "<h4 style='margin:12px 0 8px;font-size:13px;color:#fff;'>Visual Category Count</h4>"
        html += (
            "<table style='border-collapse:collapse;width:100%;margin-bottom:12px;'>"
        )
        total = max(sum(visual_categories.values()), 1)
        for key, value in sorted(visual_categories.items(), key=lambda x: -x[1]):
            pct = round((value / total) * 100)
            html += row(key, f"{value} visuals ({pct}%)")
        html += "</table>"

    if biz_cats and any(biz_cats.values()):
        html += "<h4 style='margin:12px 0 8px;font-size:13px;color:#fff;'>Business Field Categories</h4>"
        html += (
            "<table style='border-collapse:collapse;width:100%;margin-bottom:12px;'>"
        )
        for category, fields in biz_cats.items():
            if fields:
                html += row(category, list_text(fields[:10]))
        html += "</table>"

    if top_fields:
        html += "<h4 style='margin:12px 0 8px;font-size:13px;color:#fff;'>Most Used Fields</h4>"
        html += (
            "<table style='border-collapse:collapse;width:100%;margin-bottom:12px;'>"
        )
        for field in top_fields[:10]:
            if isinstance(field, dict):
                html += row(field.get("field", "—"), f"{field.get('count', 0)} uses")
            else:
                html += row(str(field), "—")
        html += "</table>"

    if page_analysis:
        html += "<h4 style='margin:12px 0 8px;font-size:13px;color:#fff;'>Page-wise Analysis</h4>"
        for page in page_analysis:
            html += "<div style='margin-bottom:8px;padding:10px;background:#1e293b;border-left:3px solid #2563eb;border-radius:8px;'>"
            html += f"<strong style='color:#fff;'>{safe(page.get('page_name', 'Unknown'))}</strong>"
            html += f" <span style='color:#facc15;margin-left:8px;'>{safe(page.get('visual_count', 0))} visuals</span>"
            if page.get("tables_used"):
                html += f"<br><span style='color:#94a3b8;font-size:11px;'>Tables: {safe_list(page.get('tables_used'))}</span>"
            if page.get("purpose"):
                html += f"<br><em style='color:#cbd5e1;font-size:11px;'>{safe(page.get('purpose'))}</em>"
            html += "</div>"

    html += "</div>"
    return {
        "html": html,
        "records": dataframe_records([analysis]),
        "count": counts.get("total_visuals", len(page_analysis) or len(table_roles)),
    }


def _build_live_excel_preview(analysis: Dict[str, Any]) -> Dict[str, Any]:
    """Build uploaded live Excel workbook preview."""
    if not analysis:
        return {
            "html": '<div class="empty-tab">No live Excel workbook uploaded.</div>',
            "records": [],
            "count": 0,
        }

    sheets = analysis.get("sheets", []) or []
    html = "<div class='preview-card live-excel' style='font-size:12px;color:#cbd5e1;'>"
    html += f"<h4 style='margin:0 0 8px;color:#fff;'>Live Excel Workbook</h4>"
    html += f"<p><strong>Workbook Type:</strong> {safe(analysis.get('workbook_type', 'N/A'))}</p>"
    html += f"<p><strong>Sheet Count:</strong> {safe(analysis.get('sheet_count', len(sheets)))}</p>"
    html += "<hr style='border-color:#334155;margin:10px 0;'>"
    for sheet in sheets:
        html += "<div style='margin-bottom:8px;'>"
        html += f"<strong style='color:#fff;'>{safe(sheet.get('sheet_name'))}</strong>"
        html += f" — Rows: {safe(sheet.get('max_row', 0))}, Cols: {safe(sheet.get('max_column', 0))}"
        if sheet.get("headers"):
            html += f"<br><span style='font-size:11px;color:#94a3b8;'>Headers: {safe_list(sheet.get('headers')[:12])}</span>"
        if sheet.get("tables"):
            html += f"<br><span style='font-size:11px;color:#38bdf8;'>Tables: {safe_list(sheet.get('tables'))}</span>"
        html += "</div>"
    html += "</div>"
    return {"html": html, "records": dataframe_records(sheets), "count": len(sheets)}


def _build_field_mapping_preview(mapping: Dict[str, Any]) -> Dict[str, Any]:
    """Build PBIX field to Excel column mapping preview."""
    if not mapping:
        return {
            "html": '<div class="empty-tab">No field mapping available. Upload a live Excel workbook to enable mapping.</div>',
            "records": [],
            "count": 0,
        }

    mapped = mapping.get("mapped", []) or []
    unmapped = mapping.get("unmapped", []) or []

    html = "<div class='preview-card mapping-preview' style='font-size:12px;color:#cbd5e1;'>"
    html += f"<h4 style='margin:0 0 8px;color:#fff;'>Mapped Fields ({len(mapped)})</h4>"
    if mapped:
        html += (
            "<table style='border-collapse:collapse;width:100%;margin-bottom:12px;'>"
        )
        for rec in mapped:
            html += "<tr>"
            html += f"<td style='padding:4px 8px;font-weight:700;color:#fff;'>{safe(rec.get('pbix_field'))}</td>"
            html += f"<td style='padding:4px 8px;color:#34d399;'>&rarr; {safe(rec.get('excel_column'))}</td>"
            html += f"<td style='padding:4px 8px;color:#94a3b8;font-size:10px;'>{safe(rec.get('match_type'))}</td>"
            html += "</tr>"
        html += "</table>"
    else:
        html += "<p style='color:#94a3b8;'>No mapped fields.</p>"

    html += f"<h4 style='margin:12px 0 8px;color:#fff;'>Unmapped Fields ({len(unmapped)})</h4>"
    if unmapped:
        html += (
            "<ul style='margin:0;padding-left:18px;'>"
            + "".join(f"<li>{safe(item)}</li>" for item in unmapped)
            + "</ul>"
        )
    else:
        html += "<p style='color:#34d399;'>No unmapped fields.</p>"
    html += "</div>"
    return {"html": html, "records": dataframe_records(mapped), "count": len(mapped)}


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------


def build_chunks_preview(
    final_chunks: Dict[str, Any],
    metadata_analysis: Dict[str, Any] = None,
    live_excel_analysis: Dict[str, Any] = None,
    field_mapping: Dict[str, Any] = None,
) -> Dict[str, Any]:
    """Build all frontend previews.

    Backward compatible:
    - Existing code can still call build_chunks_preview(final_chunks)
    - New code can call build_chunks_preview(final_chunks, metadata_analysis, live_excel_analysis, field_mapping)
    - Also reads these values from final_chunks if they are already attached there
    """
    table_chunks = final_chunks.get("table_chunks", []) or []
    relationship_chunks = final_chunks.get("relationship_chunks", []) or []
    formula_chunks = final_chunks.get("formula_chunks", []) or []
    visual_chunks = final_chunks.get("visual_chunks", []) or []

    metadata_analysis = (
        metadata_analysis
        or final_chunks.get("metadata_analysis")
        or final_chunks.get("metadata_intelligence")
        or final_chunks.get("analysis")
        or {}
    )
    live_excel_analysis = (
        live_excel_analysis or final_chunks.get("live_excel_analysis") or {}
    )
    field_mapping = (
        field_mapping
        or final_chunks.get("field_mapping")
        or final_chunks.get("pbix_excel_mapping")
        or {}
    )

    tables_preview = build_tables_preview(table_chunks)
    relationships_preview = build_relationships_preview(relationship_chunks)
    formulas_preview = build_formulas_preview(formula_chunks)
    visuals_preview = build_visuals_preview(visual_chunks)
    descriptions_preview = build_visual_descriptions_preview(visual_chunks)
    analysis_preview = _build_metadata_analysis_preview(metadata_analysis)
    live_excel_preview = _build_live_excel_preview(live_excel_analysis)
    mapping_preview = _build_field_mapping_preview(field_mapping)

    # The frontend PDF exporter uses these exact HTML blocks, so the PDF
    # preserves the same card design shown in the browser preview.
    pdf_sections = [
        {
            "id": "analysis",
            "label": "Analysis",
            "html": analysis_preview.get("html", ""),
            "count": analysis_preview.get("count", 0),
        },
        {
            "id": "tables",
            "label": "Tables",
            "html": tables_preview.get("html", ""),
            "count": tables_preview.get("count", 0),
        },
        {
            "id": "relationships",
            "label": "Relationships",
            "html": relationships_preview.get("html", ""),
            "count": relationships_preview.get("count", 0),
        },
        {
            "id": "formulas",
            "label": "Formulas",
            "html": formulas_preview.get("html", ""),
            "count": formulas_preview.get("count", 0),
        },
        {
            "id": "visuals",
            "label": "Visuals",
            "html": visuals_preview.get("html", ""),
            "count": visuals_preview.get("count", 0),
        },
        {
            "id": "descriptions",
            "label": "Visual Descriptions",
            "html": descriptions_preview.get("html", ""),
            "count": descriptions_preview.get("count", 0),
        },
        {
            "id": "live_excel",
            "label": "Live Excel",
            "html": live_excel_preview.get("html", ""),
            "count": live_excel_preview.get("count", 0),
        },
        {
            "id": "mapping",
            "label": "Field Mapping",
            "html": mapping_preview.get("html", ""),
            "count": mapping_preview.get("count", 0),
        },
    ]

    return {
        "tables_preview": tables_preview,
        "relationships_preview": relationships_preview,
        "formulas_preview": formulas_preview,
        "visuals_preview": visuals_preview,
        "visual_descriptions_preview": descriptions_preview,
        "metadata_analysis_preview": analysis_preview,
        "live_excel_analysis_preview": live_excel_preview,
        "field_mapping_preview": mapping_preview,
        "pdf_export": {
            "version": 1,
            "layout": "one-card-per-page",
            "sections": pdf_sections,
        },
        "raw_json_preview": {
            "text": json.dumps(final_chunks, indent=2, ensure_ascii=False)
        },
    }

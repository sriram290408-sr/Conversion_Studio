from __future__ import annotations

import html
import json
import logging
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Sequence, Tuple

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

logger = logging.getLogger("pdf_preview_renderer")

PAGE_BG = colors.white
CARD_BG = colors.white
TEXT = colors.HexColor("#172033")
MUTED = colors.HexColor("#667085")
GOLD = colors.HexColor("#75621f")
BLUE = colors.HexColor("#2563eb")
BORDER = colors.HexColor("#d8dee9")
HEADER_BG = colors.HexColor("#f4f6f9")
CODE_BG = colors.HexColor("#f7f8fa")
WARNING_BG = colors.HexColor("#fff8e1")

# Keys that are useful to business readers. Everything else is omitted from compact
# sections unless the section explicitly asks for technical content.
VISUAL_KEYS = [
    "page_name",
    "visual_type",
    "visual_description",
    "business_title",
    "business_role",
    "excel_render_type",
    "dimension_fields",
    "measure_fields",
    "filter_fields",
    "mapped_table_chunks",
    "mapped_formula_chunks",
    "conversion_status",
    "layout",
]

VISUAL_DESCRIPTION_KEYS = [
    "page_name",
    "visual_type",
    "visual_description",
    "business_title",
    "business_role",
    "excel_render_type",
    "dimension_fields",
    "measure_fields",
    "filter_fields",
]

FORMULA_KEYS = [
    "measure_name",
    "table_name",
    "dax_formula",
    "excel_formula",
    "conversion_status",
    "conversion_source",
    "required_tables",
    "required_hidden_sheets",
    "notes",
]

TABLE_KEYS = [
    "table_name",
    "columns",
    "source",
    "excel_table_name",
    "hidden_sheet",
    "row_count",
    "description",
]

RELATIONSHIP_KEYS = [
    "from_table",
    "from_column",
    "to_table",
    "to_column",
    "cardinality",
    "cross_filter_direction",
    "active",
]


def _as_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (dict, list, tuple)):
        return json.dumps(value, ensure_ascii=False, default=str, indent=2)
    return str(value)


def _safe(value: Any, limit: int = 14000) -> str:
    text = _as_text(value)[:limit]
    return html.escape(text).replace("\r", "").replace("\n", "<br/>")


def _normalize_semantic_reference(value: Any) -> Any:
    """Normalize display-only semantic references without changing source payload."""
    if isinstance(value, list):
        return [_normalize_semantic_reference(item) for item in value]
    if isinstance(value, tuple):
        return tuple(_normalize_semantic_reference(item) for item in value)
    if isinstance(value, dict):
        return {k: _normalize_semantic_reference(v) for k, v in value.items()}
    if not isinstance(value, str):
        return value

    text = value.strip()
    # Common malformed extraction: Sum(flights[passengers)]
    text = re.sub(r"\b(Sum|Avg|Average|Min|Max|Count|CountNonNull)\(([^\[]+)\[([^\]\)]+)\)\]", r"\1(\2[\3])", text, flags=re.I)
    # Common malformed extraction: Sum(flights.passengers)
    text = re.sub(r"\b(Sum|Avg|Average|Min|Max|Count|CountNonNull)\(([^\.\)]+)\.([^\)]+)\)", r"\1(\2[\3])", text, flags=re.I)
    # Remove duplicated trailing parenthesis in field names.
    text = re.sub(r"\[([^\]]+?)\)\]", r"[\1]", text)
    return text


def _records(preview_section: Any) -> List[Dict[str, Any]]:
    if not isinstance(preview_section, dict):
        return []
    records = preview_section.get("records")
    if isinstance(records, list):
        return [item for item in records if isinstance(item, dict)]
    return []


def _chunk_records(chunks: Dict[str, Any], key: str) -> List[Dict[str, Any]]:
    value = chunks.get(key) or []
    if isinstance(value, list):
        return [item for item in value if isinstance(item, dict)]
    if isinstance(value, dict):
        return [value]
    return []


def _title(record: Dict[str, Any], index: int) -> str:
    return str(
        record.get("business_title")
        or record.get("ai_title")
        or record.get("title")
        or record.get("visual_title")
        or record.get("measure_name")
        or record.get("table_name")
        or record.get("page_name")
        or record.get("chunk_id")
        or f"Record {index}"
    )


def _non_empty(value: Any) -> bool:
    return value not in (None, "", [], {}, ())


def _select_items(record: Dict[str, Any], allowed_keys: Sequence[str] | None) -> List[Tuple[str, Any]]:
    if allowed_keys is None:
        items = list(record.items())
    else:
        items = [(key, record.get(key)) for key in allowed_keys if key in record]

    result: List[Tuple[str, Any]] = []
    for key, value in items:
        if key in {"title", "visual_title", "ai_title", "business_title"}:
            continue
        if not _non_empty(value):
            continue
        result.append((key, _normalize_semantic_reference(value)))
    return result


def _summary_analysis_record(record: Dict[str, Any]) -> Dict[str, Any]:
    """Reduce a giant metadata-analysis object into readable high-value rows."""
    preferred = [
        "overall_counts",
        "model_tables",
        "visual_inferred_tables",
        "tables_used_in_visuals",
        "visual_category_count",
        "top_used_fields",
        "top_used_measures",
        "page_analysis",
        "ai_deep_analysis",
        "classified_measures",
    ]
    compact: Dict[str, Any] = {}
    for key in preferred:
        value = record.get(key)
        if _non_empty(value):
            compact[key] = _normalize_semantic_reference(value)
    return compact or record


def _infer_mapped_tables(record: Dict[str, Any], chunks: Dict[str, Any]) -> List[str]:
    existing = record.get("mapped_table_chunks") or []
    if existing:
        return list(existing) if isinstance(existing, list) else [str(existing)]

    references: List[str] = []
    for key in ("uses_fields", "dimension_fields", "measure_fields", "uses_measures"):
        value = record.get(key) or []
        if isinstance(value, str):
            value = [value]
        references.extend(str(item) for item in value)

    table_chunks = _chunk_records(chunks, "table_chunks")
    inferred: List[str] = []
    for table_record in table_chunks:
        table_name = str(table_record.get("table_name") or table_record.get("name") or "").strip()
        chunk_id = str(table_record.get("chunk_id") or "").strip()
        if not table_name:
            continue
        if any(
            re.search(rf"(^|[\[\.(]){re.escape(table_name)}([\]\.)]|$)", ref, flags=re.I)
            or ref.lower().startswith(table_name.lower() + ".")
            for ref in references
        ):
            inferred.append(chunk_id or f"table_{table_name}")
    return list(dict.fromkeys(inferred))


def _compact_visual_record(record: Dict[str, Any], chunks: Dict[str, Any], description_only: bool) -> Dict[str, Any]:
    compact = dict(record)
    deep = compact.get("ai_deep_analysis")
    if isinstance(deep, dict):
        compact.setdefault("business_title", deep.get("recommended_title"))
        compact.setdefault("business_role", deep.get("business_role"))
        compact.setdefault("excel_render_type", deep.get("excel_render_type"))
        compact.setdefault("dimension_fields", deep.get("dimension_fields"))
        compact.setdefault("measure_fields", deep.get("measure_fields"))
        compact.setdefault("filter_fields", deep.get("filter_fields"))
        compact.setdefault("visual_description", deep.get("description"))

    inferred_tables = _infer_mapped_tables(compact, chunks)
    if inferred_tables:
        compact["mapped_table_chunks"] = inferred_tables

    # Correct display strategy for card-like visuals without altering conversion data.
    visual_type = str(compact.get("visual_type") or "").lower()
    role = str(compact.get("business_role") or "").lower()
    if visual_type in {"card", "cardvisual", "kpi", "multicard"} or role == "kpi":
        compact["excel_render_type"] = "kpi_card"

    if description_only:
        return {key: compact.get(key) for key in VISUAL_DESCRIPTION_KEYS if _non_empty(compact.get(key))}
    return compact


def generate_pdf_preview(data: dict, output_pdf_path: str) -> dict:
    output = Path(output_pdf_path).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    logger.info("Generating compact Chunk Visualizer PDF at: %s", output)

    # Portrait for business-readable output. Large technical dictionaries are compacted.
    document = SimpleDocTemplate(
        str(output),
        pagesize=A4,
        rightMargin=12 * mm,
        leftMargin=12 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm,
        title="Power BI to Excel - Chunk Visualizer Preview",
        author="Power BI to Excel Conversion Studio",
        allowSplitting=True,
    )

    base = getSampleStyleSheet()
    heading = ParagraphStyle(
        "ChunkSection",
        parent=base["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=17,
        leading=21,
        textColor=GOLD,
        spaceAfter=3 * mm,
        keepWithNext=True,
    )
    section_meta = ParagraphStyle(
        "SectionMeta",
        parent=base["Normal"],
        fontSize=8,
        leading=10,
        textColor=MUTED,
        spaceAfter=3 * mm,
    )
    card_title = ParagraphStyle(
        "CardTitle",
        parent=base["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10,
        leading=12.5,
        textColor=BLUE,
    )
    body = ParagraphStyle(
        "CardBody",
        parent=base["Normal"],
        fontName="Helvetica",
        fontSize=8.1,
        leading=10.5,
        textColor=TEXT,
    )
    label = ParagraphStyle(
        "CardLabel",
        parent=body,
        fontName="Helvetica-Bold",
        textColor=TEXT,
    )
    code = ParagraphStyle(
        "CardCode",
        parent=body,
        fontName="Courier",
        fontSize=6.8,
        leading=8.5,
        backColor=CODE_BG,
        borderPadding=2,
    )

    page_state = {"count": 0}

    def paint_page(canvas, _doc) -> None:
        page_state["count"] += 1
        canvas.saveState()
        canvas.setFillColor(PAGE_BG)
        canvas.rect(0, 0, A4[0], A4[1], fill=1, stroke=0)
        canvas.setStrokeColor(BORDER)
        canvas.line(12 * mm, 10 * mm, A4[0] - 12 * mm, 10 * mm)
        canvas.setFillColor(MUTED)
        canvas.setFont("Helvetica", 7)
        canvas.drawString(12 * mm, 6.5 * mm, "Power BI Conversion Preview - Chunk Visualizer")
        canvas.drawRightString(A4[0] - 12 * mm, 6.5 * mm, f"Page {page_state['count']}")
        canvas.restoreState()

    def value_style(key: str) -> ParagraphStyle:
        lowered = key.lower()
        return code if any(token in lowered for token in ("formula", "dax", "json", "query", "expression", "sql", "m_code")) else body

    def _split_text_chunks(
        value: Any,
        max_chars: int = 1200,
        max_lines: int = 28,
        max_total_chars: int = 12000,
        max_chunks: int = 10,
    ) -> List[str]:
        """Split a large value into page-safe chunks.

        ReportLab Tables cannot split a single row across pages. A very large JSON or
        analysis value therefore creates a row taller than the available frame and
        raises LayoutError. This helper bounds every table row by splitting values
        into multiple continuation rows before creating flowables.
        """
        raw = _as_text(value).replace("\r", "")
        truncated = len(raw) > max_total_chars
        if truncated:
            raw = raw[:max_total_chars]
        if not raw:
            return [""]

        # Hard-wrap unusually long tokens so Paragraph can always find break points.
        normalized_lines: List[str] = []
        for source_line in raw.split("\n") or [raw]:
            line = source_line if source_line else " "
            while len(line) > 160:
                cut = line.rfind(" ", 0, 160)
                if cut < 60:
                    cut = 160
                normalized_lines.append(line[:cut])
                line = line[cut:].lstrip()
            normalized_lines.append(line)

        chunks: List[str] = []
        current: List[str] = []
        current_chars = 0
        for line in normalized_lines:
            line_len = len(line) + 1
            if current and (
                current_chars + line_len > max_chars or len(current) >= max_lines
            ):
                chunks.append("\n".join(current))
                if len(chunks) >= max_chunks:
                    break
                current = []
                current_chars = 0
            current.append(line)
            current_chars += line_len

        if current and len(chunks) < max_chunks:
            chunks.append("\n".join(current))
        if truncated or len(chunks) >= max_chunks:
            notice = "[Content truncated in PDF preview. Full value remains available in converted_metadata.json.]"
            if chunks:
                chunks[-1] = chunks[-1] + "\n" + notice
            else:
                chunks = [notice]
        return chunks or [""]

    def _single_row_table(label_text: str, value_text: str, continuation: bool) -> Table:
        display_label = f"{label_text} (continued)" if continuation else label_text
        row = [
            Paragraph(html.escape(display_label), label),
            Paragraph(_safe(value_text, 4000), body),
        ]
        table = Table(
            [row],
            colWidths=[42 * mm, 144 * mm],
            splitByRow=1,
            hAlign="LEFT",
        )
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), CARD_BG),
                    ("BOX", (0, 0), (-1, -1), 0.45, BORDER),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        return table

    def card_for_record(
        record: Dict[str, Any],
        index: int,
        allowed_keys: Sequence[str] | None,
        section_name: str,
    ) -> List[Any]:
        """Return page-safe flowables for one record.

        Each property is emitted as an independent one-row table. Long values are
        divided into continuation rows, so no single Table row can become taller
        than the page frame.
        """
        flowables: List[Any] = []

        title_table = Table(
            [[Paragraph(_safe(_title(record, index), 700), card_title)]],
            colWidths=[186 * mm],
            hAlign="LEFT",
        )
        title_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), HEADER_BG),
                    ("BOX", (0, 0), (-1, -1), 0.6, BORDER),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ]
            )
        )
        flowables.append(title_table)

        items = _select_items(record, allowed_keys)
        if not items:
            flowables.append(
                _single_row_table("Details", "No additional details available.", False)
            )
            return flowables

        for key, value in items:
            label_text = str(key).replace("_", " ").title()
            chunks_for_value = _split_text_chunks(value)
            for chunk_index, chunk in enumerate(chunks_for_value):
                # Preserve code-like formatting while still keeping every row bounded.
                style = value_style(str(key))
                display_label = (
                    f"{label_text} (continued)" if chunk_index else label_text
                )
                row_table = Table(
                    [[
                        Paragraph(html.escape(display_label), label),
                        Paragraph(_safe(chunk, 4000), style),
                    ]],
                    colWidths=[42 * mm, 144 * mm],
                    splitByRow=1,
                    hAlign="LEFT",
                )
                row_table.setStyle(
                    TableStyle(
                        [
                            ("BACKGROUND", (0, 0), (-1, -1), CARD_BG),
                            ("BOX", (0, 0), (-1, -1), 0.45, BORDER),
                            ("VALIGN", (0, 0), (-1, -1), "TOP"),
                            ("LEFTPADDING", (0, 0), (-1, -1), 6),
                            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                            ("TOPPADDING", (0, 0), (-1, -1), 4),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                        ]
                    )
                )
                flowables.append(row_table)

        return flowables

    chunks = data.get("chunks") or {}
    previews = data.get("chunks_preview") or {}
    story: List[Any] = []

    section_specs = [
        ("Analysis", "metadata_analysis_preview", "metadata_analysis", None),
        ("Tables", "tables_preview", "table_chunks", TABLE_KEYS),
        ("Relationships", "relationships_preview", "relationship_chunks", RELATIONSHIP_KEYS),
        ("Formulas", "formulas_preview", "formula_chunks", FORMULA_KEYS),
        ("Visuals", "visuals_preview", "visual_chunks", VISUAL_KEYS),
        ("Visual Descriptions", "visual_descriptions_preview", "visual_description_chunks", VISUAL_DESCRIPTION_KEYS),
        ("Live Excel Analysis", "live_excel_analysis_preview", "live_excel_analysis", None),
        ("Field Mapping", "field_mapping_preview", "excel_field_mapping", None),
    ]

    rendered_sections = 0
    rendered_records = 0

    for section_title, preview_key, chunk_key, allowed_keys in section_specs:
        records = _records(previews.get(preview_key))
        if not records:
            records = _chunk_records(chunks, chunk_key)
        if not records:
            continue

        processed: List[Dict[str, Any]] = []
        for record in records:
            current = dict(record)
            if section_title == "Analysis":
                current = _summary_analysis_record(current)
            elif section_title == "Visuals":
                current = _compact_visual_record(current, chunks, description_only=False)
            elif section_title == "Visual Descriptions":
                current = _compact_visual_record(current, chunks, description_only=True)
            else:
                current = _normalize_semantic_reference(current)
            processed.append(current)

        if rendered_sections:
            story.append(PageBreak())
        story.append(Paragraph(section_title, heading))
        story.append(Paragraph(f"{len(processed)} detailed record{'s' if len(processed) != 1 else ''}", section_meta))

        for index, record in enumerate(processed, 1):
            card_flowables = card_for_record(record, index, allowed_keys, section_title)
            story.extend(card_flowables)
            story.append(Spacer(1, 3 * mm))
            rendered_records += 1
        rendered_sections += 1

    if not story:
        raise RuntimeError("No Chunk Visualizer records are available for PDF generation")

    document.build(story, onFirstPage=paint_page, onLaterPages=paint_page)

    size = output.stat().st_size if output.exists() else 0
    if size < 1000:
        raise RuntimeError(f"Chunk Visualizer PDF output is empty or incomplete: {size} bytes")

    logger.info(
        "Chunk Visualizer PDF complete. Pages=%d Sections=%d Records=%d Size=%d bytes",
        page_state["count"], rendered_sections, rendered_records, size,
    )
    return {
        "pdf_status": "success",
        "renderer": "reportlab_chunk_visualizer_compact",
        "sections_rendered": rendered_sections,
        "records_rendered": rendered_records,
        "pages_created": page_state["count"],
        "file_size_bytes": size,
        "download_ready": True,
    }

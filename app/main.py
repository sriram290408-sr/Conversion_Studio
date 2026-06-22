import os
import json
import uuid
import time
import shutil
import logging
import tempfile
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, UploadFile, File, HTTPException, Body
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

# Configure logging before emitting startup diagnostics.
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("main")

# Package-local imports. The relative imports are the supported production path
# when the server is started from the project root as ``app.main:app``.
try:
    from .Convertor import (
        read_powerbi_metadata,
        process_metadata_to_chunks,
        compile_chunks_to_xlsx,
        check_huggingface_connectivity,
        analyze_live_excel_workbook,
        map_pbix_fields_to_excel_columns,
        apply_tmdl_metadata_to_chunks,
        build_hf_render_artifacts,
    )
    from .preview_builder import build_chunks_preview
except ImportError:
    # Backward-compatible fallback for direct execution from inside app/.
    from Convertor import (
        read_powerbi_metadata,
        process_metadata_to_chunks,
        compile_chunks_to_xlsx,
        check_huggingface_connectivity,
        analyze_live_excel_workbook,
        map_pbix_fields_to_excel_columns,
        apply_tmdl_metadata_to_chunks,
        build_hf_render_artifacts,
    )
    from preview_builder import build_chunks_preview

import platform

# Live-connect session infrastructure (Windows + pywin32 only).
LIVE_CONNECT_AVAILABLE = False
LIVE_CONNECT_UNAVAILABLE_REASON: str | None = None
session_manager = None
launch_excel_for_connection = None
detect_and_validate_connection = None
build_live_dashboard = None
run_continue_workflow_com = None
create_all_visual_bindings = None

if platform.system() != "Windows":
    LIVE_CONNECT_UNAVAILABLE_REASON = (
        f"Unsupported platform: {platform.system()}. Windows is required."
    )
else:
    try:
        import pythoncom  # type: ignore[import]
        import pywintypes  # type: ignore[import]
        import win32com.client  # type: ignore[import]
    except Exception as exc:
        logger.exception("pywin32 initialization failed")
        LIVE_CONNECT_UNAVAILABLE_REASON = (
            f"pywin32 initialization failed: {type(exc).__name__}: {exc}"
        )
    else:
        try:
            try:
                from .com_session_manager import session_manager
                from .interactive_excel_connection import (
                    launch_excel_for_connection,
                    detect_and_validate_connection,
                    build_live_dashboard,
                    run_continue_workflow_com,
                )
                from .binding_engine import create_all_visual_bindings
            except ImportError:
                from com_session_manager import session_manager
                from interactive_excel_connection import (
                    launch_excel_for_connection,
                    detect_and_validate_connection,
                    build_live_dashboard,
                    run_continue_workflow_com,
                )
                from binding_engine import create_all_visual_bindings
        except Exception as exc:
            logger.exception("Live-connect application import failed")
            LIVE_CONNECT_UNAVAILABLE_REASON = (
                f"Live-connect application import failed: {type(exc).__name__}: {exc}"
            )
        else:
            LIVE_CONNECT_AVAILABLE = True

logger.info(
    "Live-connect startup check: platform=%s available=%s reason=%s",
    platform.system(),
    LIVE_CONNECT_AVAILABLE,
    LIVE_CONNECT_UNAVAILABLE_REASON,
)


app = FastAPI(title="Power BI PBIX to Excel Converter", version="1.0.0")

APP_DIR = Path(__file__).resolve().parent
BASE_DIR = APP_DIR.parent

ROOT_INDEX = BASE_DIR / "index.html"
TEMPLATES_INDEX = APP_DIR / "templates" / "index.html"
STATIC_DIR = BASE_DIR / "static"

TEMP_DIR = Path(tempfile.gettempdir()) / "powerbi_converter"
TEMP_DIR.mkdir(parents=True, exist_ok=True)

MAX_UPLOAD_SIZE_BYTES = int(os.getenv("MAX_UPLOAD_SIZE_BYTES", str(100 * 1024 * 1024)))
TEMP_EXPIRY_SECONDS = int(os.getenv("TEMP_EXPIRY_SECONDS", str(60 * 60)))
# Local origins are always supported. Add hosted frontend origins through the
# CORS_ORIGINS environment variable as a comma-separated list.
LOCAL_CORS_ORIGINS = [
    "http://127.0.0.1:8000",
    "http://localhost:8000",
    "http://127.0.0.1:5500",
    "http://localhost:5500",
]

DEPLOYED_CORS_ORIGINS = [
    origin.strip().rstrip("/")
    for origin in os.getenv("CORS_ORIGINS", "").split(",")
    if origin.strip()
]

ALLOWED_ORIGINS = list(
    dict.fromkeys(
        [origin.rstrip("/") for origin in LOCAL_CORS_ORIGINS] + DEPLOYED_CORS_ORIGINS
    )
)

if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def safe_delete_dir(path: Path) -> None:
    try:
        if path and path.exists() and path.is_dir():
            shutil.rmtree(path, ignore_errors=True)
            logger.info("Deleted temp folder: %s", path)
    except Exception as e:
        logger.warning("Failed to delete temp folder %s: %s", path, e)


def cleanup_old_temp_files(max_age_minutes: int = None) -> None:
    try:
        now = time.time()
        expiry_seconds = (
            (max_age_minutes * 60)
            if max_age_minutes is not None
            else TEMP_EXPIRY_SECONDS
        )
        if not TEMP_DIR.exists():
            return
        for session_dir in TEMP_DIR.iterdir():
            if not session_dir.is_dir():
                continue
            try:
                if now - session_dir.stat().st_mtime > expiry_seconds:
                    safe_delete_dir(session_dir)
            except Exception as inner_error:
                logger.warning(
                    "Skipping temp cleanup for %s: %s", session_dir, inner_error
                )
    except Exception as e:
        logger.warning("Temp cleanup failed: %s", e)


def _validate_pdf_output(pdf_path: Path) -> dict:
    result = {
        "status": "failed",
        "download_ready": False,
        "file_size_bytes": 0,
        "page_count": 0,
        "text_length": 0,
        "error": None,
    }
    try:
        if not pdf_path.exists():
            raise RuntimeError("PDF file was not created.")
        file_size = pdf_path.stat().st_size
        if file_size < 1000:
            raise RuntimeError(f"PDF file is empty or incomplete: {file_size} bytes.")
        with pdf_path.open("rb") as handle:
            if handle.read(5) != b"%PDF-":
                raise RuntimeError("Generated file does not have a valid PDF header.")

        from pypdf import PdfReader

        reader = PdfReader(str(pdf_path))
        page_count = len(reader.pages)
        if page_count < 1:
            raise RuntimeError("Generated PDF contains no pages.")
        extracted = "\n".join((page.extract_text() or "") for page in reader.pages)
        text_length = len(extracted.strip())
        if text_length < 20:
            raise RuntimeError("Generated PDF contains no meaningful extractable text.")
        required_any = ("Analysis", "Tables", "Visuals", "Formulas")
        if not any(token in extracted for token in required_any):
            raise RuntimeError(
                "Generated PDF does not contain expected Chunk Visualizer sections."
            )

        result.update(
            {
                "status": "success",
                "download_ready": True,
                "file_size_bytes": file_size,
                "page_count": page_count,
                "text_length": text_length,
            }
        )
    except Exception as exc:
        result["error"] = str(exc)
    return result


def create_preview_pdf(pdf_path: Path, payload: dict) -> dict:
    try:
        import inspect

        try:
            from .pdf_preview_renderer import generate_pdf_preview
        except ImportError:
            from pdf_preview_renderer import generate_pdf_preview

        logger.info(
            "Using PDF renderer from: %s", inspect.getfile(generate_pdf_preview)
        )
        renderer_result = generate_pdf_preview(payload, str(pdf_path))
        validation = _validate_pdf_output(pdf_path)
        if validation["status"] != "success":
            raise RuntimeError(
                validation.get("error") or "Generated PDF failed validation."
            )
        result = {**renderer_result, **validation}
        logger.info(
            "Chunk Visualizer PDF validated: pages=%s sections=%s records=%s size=%s text=%s",
            result.get("page_count"),
            result.get("sections_rendered"),
            result.get("records_rendered"),
            result.get("file_size_bytes"),
            result.get("text_length"),
        )
        return result
    except Exception as exc:
        logger.exception("Chunk Visualizer PDF generation failed")
        if pdf_path.exists():
            try:
                pdf_path.unlink()
            except OSError:
                logger.warning("Failed to remove invalid PDF file: %s", pdf_path)
        return {
            "status": "failed",
            "pdf_status": "failed",
            "renderer": "reportlab_chunk_visualizer_compact",
            "download_ready": False,
            "file_size_bytes": 0,
            "page_count": 0,
            "text_length": 0,
            "sections_rendered": 0,
            "records_rendered": 0,
            "error": str(exc),
        }


def get_index_html_path() -> Path:
    if TEMPLATES_INDEX.exists():
        return TEMPLATES_INDEX
    if ROOT_INDEX.exists():
        return ROOT_INDEX
    raise HTTPException(status_code=404, detail="index.html not found.")


def validate_pbix_upload(file: UploadFile, contents: bytes) -> None:
    if not file.filename:
        raise HTTPException(status_code=400, detail="Filename cannot be empty.")
    ext = Path(file.filename).suffix.lower()
    if ext != ".pbix":
        raise HTTPException(status_code=400, detail="Only .pbix files are supported.")
    if not contents:
        raise HTTPException(status_code=400, detail="Uploaded PBIX file is empty.")
    if len(contents) > MAX_UPLOAD_SIZE_BYTES:
        raise HTTPException(
            status_code=413,
            detail=f"File too large. Maximum allowed size is {MAX_UPLOAD_SIZE_BYTES} bytes.",
        )


async def process_upload_file(
    file: UploadFile,
    screenshot: Optional[UploadFile] = None,
    base_template: Optional[UploadFile] = None,
    tmdl_metadata: Optional[UploadFile] = None,
):
    """
    Full PBIX upload pipeline with graceful fallback for layout-only PBIX files.

    Returns 422 for known PBIX structural limitations (missing DataModelSchema/Layout).
    Returns 500 only for unexpected server errors.
    """
    cleanup_old_temp_files()

    session_id = str(uuid.uuid4())
    work_dir = TEMP_DIR / session_id
    work_dir.mkdir(parents=True, exist_ok=True)

    try:
        contents = await file.read()
        validate_pbix_upload(file, contents)

        uploaded_filename = os.path.basename(file.filename)
        input_path = work_dir / uploaded_filename

        with open(input_path, "wb") as f:
            f.write(contents)
        logger.info("PBIX uploaded: %s", input_path)

        # ── Optional screenshot ──────────────────────────────────────────
        screenshot_path = None
        screenshot_filename = None
        if screenshot and screenshot.filename:
            sc_ext = Path(screenshot.filename).suffix.lower()
            if sc_ext not in (".png", ".jpg", ".jpeg", ".webp"):
                raise HTTPException(
                    status_code=400,
                    detail="Screenshot must be .png, .jpg, .jpeg, or .webp.",
                )
            screenshot_filename = os.path.basename(screenshot.filename)
            screenshot_path = work_dir / screenshot_filename
            sc_contents = await screenshot.read()
            with open(screenshot_path, "wb") as f:
                f.write(sc_contents)
            logger.info("Screenshot uploaded: %s", screenshot_path)

        # ── Optional Base Template ──────────────────────────────────────────
        base_template_path = None
        base_template_filename = None
        if base_template and base_template.filename:
            bt_ext = Path(base_template.filename).suffix.lower()
            if bt_ext not in (".xlsx", ".xlsm", ".xltx", ".xltm"):
                raise HTTPException(
                    status_code=400,
                    detail="Base template must be .xlsx, .xlsm, .xltx, or .xltm.",
                )
            base_template_filename = os.path.basename(base_template.filename)
            base_template_path = work_dir / base_template_filename
            bt_contents = await base_template.read()
            with open(base_template_path, "wb") as f:
                f.write(bt_contents)
            logger.info("Base template uploaded: %s", base_template_path)

        # ── Optional TMDL metadata file ──────────────────────────────────
        tmdl_path = None
        tmdl_filename = None
        if tmdl_metadata and tmdl_metadata.filename:
            tmdl_ext = Path(tmdl_metadata.filename).suffix.lower()
            if tmdl_ext not in (".tmdl", ".txt", ".json"):
                raise HTTPException(
                    status_code=400,
                    detail="TMDL metadata must be .tmdl, .txt, or .json.",
                )
            tmdl_filename = os.path.basename(tmdl_metadata.filename)
            tmdl_path = work_dir / tmdl_filename
            tmdl_contents = await tmdl_metadata.read()
            with open(tmdl_path, "wb") as f:
                f.write(tmdl_contents)
            logger.info("TMDL metadata uploaded: %s", tmdl_path)

        # ── Step 1: Read PBIX metadata (graceful fallback) ───────────────
        raw_metadata = read_powerbi_metadata(str(input_path))
        extraction_mode = raw_metadata.get("extraction_mode", "unknown")
        metadata_warnings = raw_metadata.get("metadata_warnings", [])

        logger.info("PBIX extraction mode: %s", extraction_mode)
        if metadata_warnings:
            for w in metadata_warnings:
                logger.warning("PBIX warning: %s", w)

        # ── Step 2: Process into chunks ──────────────────────────────────
        final_chunks = process_metadata_to_chunks(raw_metadata)

        if tmdl_path:
            try:
                apply_tmdl_metadata_to_chunks(final_chunks, tmdl_path=str(tmdl_path))
                logger.info("TMDL metadata applied before preview: %s", tmdl_path)
            except Exception as tmdl_error:
                logger.warning("Failed to apply TMDL metadata: %s", tmdl_error)
                final_chunks.setdefault("metadata_warnings", []).append(
                    f"TMDL metadata parse warning: {tmdl_error}"
                )

        # Build the screenshot-aware, HF-assisted render plan before live connection.
        try:
            build_hf_render_artifacts(
                final_chunks=final_chunks,
                screenshot_path=str(screenshot_path) if screenshot_path else None,
                output_dir=str(work_dir),
            )
            logger.info(
                "Render plan prepared with %d operation(s).",
                len(
                    (final_chunks.get("visual_render_plan") or {}).get("operations", [])
                ),
            )
        except Exception as render_plan_error:
            logger.exception("Render-plan generation failed")
            final_chunks.setdefault("metadata_warnings", []).append(
                f"Render-plan generation warning: {render_plan_error}"
            )

        if base_template_path:
            final_chunks["live_excel_analysis"] = analyze_live_excel_workbook(
                str(base_template_path)
            )
            pbix_fields = []

            for table_chunk in final_chunks.get("table_chunks", []):
                table_name = (
                    table_chunk.get("table_name")
                    or table_chunk.get("name")
                    or "UnknownTable"
                )

                for column_value in table_chunk.get("columns", []) or []:
                    if isinstance(column_value, dict):
                        column_name = column_value.get("name") or column_value.get(
                            "column_name"
                        )
                    else:
                        column_name = str(column_value)

                    if column_name:
                        pbix_fields.append(f"{table_name}[{column_name}]")

            # Also include fields found directly in visuals for layout-only PBIX files.
            for visual_chunk in final_chunks.get("visual_chunks", []) or []:
                for field_value in visual_chunk.get("uses_fields", []) or []:
                    if field_value and str(field_value) not in pbix_fields:
                        pbix_fields.append(str(field_value))

            final_chunks["excel_field_mapping"] = map_pbix_fields_to_excel_columns(
                pbix_fields,
                final_chunks["live_excel_analysis"],
            )

        # ── Step 3: Build preview ────────────────────────────────────────
        chunks_preview = build_chunks_preview(final_chunks)

        logger.info(
            "Metadata analysis keys: %s",
            list(final_chunks.get("metadata_analysis", {}).keys()),
        )
        logger.info("Preview keys: %s", list(chunks_preview.keys()))
        logger.info(
            "Metadata preview count: %s",
            chunks_preview.get("metadata_analysis_preview", {}).get("count"),
        )

        # ── Step 4: Save JSON debug output ───────────────────────────────
        json_output_path = work_dir / "converted_metadata.json"
        json_payload = {
            "session_id": session_id,
            "status": "success",
            "uploaded_filename": uploaded_filename,
            "extraction_mode": extraction_mode,
            "metadata_warnings": metadata_warnings,
            "summary": final_chunks.get("summary", {}),
            "chunks": final_chunks,
            "chunks_preview": chunks_preview,
            "validation_errors": final_chunks.get("validation_errors", []),
        }
        with open(json_output_path, "w", encoding="utf-8") as f:
            json.dump(json_payload, f, indent=2, ensure_ascii=False)

        preview_pdf_path = work_dir / "dashboard_preview.pdf"
        pdf_status = create_preview_pdf(preview_pdf_path, json_payload)
        json_payload["pdf_status"] = pdf_status
        with open(json_output_path, "w", encoding="utf-8") as f:
            json.dump(json_payload, f, indent=2, ensure_ascii=False)

        if pdf_status.get("download_ready"):
            logger.info(
                "Preview PDF created: %s pages=%s size=%s",
                preview_pdf_path,
                pdf_status.get("page_count"),
                pdf_status.get("file_size_bytes"),
            )
        else:
            logger.error("Preview PDF creation failed: %s", pdf_status.get("error"))

        # ── Step 5: Compile XLSX workbook ────────────────────────────────

        out_ext = ".xlsx"
        if base_template_filename and base_template_filename.lower().endswith(
            (".xlsm", ".xltm")
        ):
            out_ext = ".xlsm"
        xlsx_output_filename = f"excel_ready_model{out_ext}"
        xlsx_output_path = work_dir / xlsx_output_filename

        show_tech = os.getenv("SHOW_TECHNICAL_SHEETS", "false").strip().lower() in {
            "true",
            "1",
            "yes",
            "y",
            "on",
        }

        _live_enabled = os.getenv("POWERBI_LIVE_ENABLED", "false").strip().lower() in {
            "true",
            "1",
            "yes",
            "y",
            "on",
        }
        powerbi_live_config = {
            "enabled": _live_enabled,
            "workspace_id": os.getenv("POWERBI_WORKSPACE_ID", ""),
            "dataset_id": os.getenv("POWERBI_DATASET_ID", ""),
        }

        session_info = {
            "session_id": session_id,
            "uploaded_filename": uploaded_filename,
            "screenshot_filename": screenshot_filename,
            "screenshot_path": str(screenshot_path) if screenshot_path else None,
            "base_template_filename": base_template_filename,
            "base_template_path": (
                str(base_template_path) if base_template_path else None
            ),
            "tmdl_filename": tmdl_filename,
            "tmdl_path": str(tmdl_path) if tmdl_path else None,
            "has_tmdl_metadata": bool(tmdl_path),
            "has_screenshot": bool(screenshot_path),
            "has_live_excel_workbook": bool(base_template_path),
            "input_type": "PBIX",
            "show_technical_sheets": show_tech,
            "powerbi_live_config": powerbi_live_config,
        }

        try:
            compile_chunks_to_xlsx(
                final_chunks, str(xlsx_output_path), session_info=session_info
            )
        except TypeError:
            compile_chunks_to_xlsx(final_chunks, str(xlsx_output_path))

        if not xlsx_output_path.exists():
            raise RuntimeError(
                "XLSX compilation completed, but output workbook was not created."
            )

        logger.info("Excel workbook created: %s", xlsx_output_path)
        logger.info(
            "Upload processed successfully. Session: %s (mode: %s)",
            session_id,
            extraction_mode,
        )

        # ── Step 6: Return response ──────────────────────────────────────
        return JSONResponse(
            {
                "session_id": session_id,
                "status": "success",
                "extraction_mode": extraction_mode,
                "metadata_warnings": metadata_warnings,
                "has_screenshot": bool(screenshot_path),
                "has_live_excel_workbook": bool(base_template_path),
                "has_tmdl_metadata": bool(tmdl_path),
                "download_url": f"/download/{session_id}",
                "preview_download_url": (
                    f"/download-preview/{session_id}"
                    if pdf_status.get("download_ready")
                    else None
                ),
                "preview_filename": "powerbi_chunk_visualizer_preview.pdf",
                "preview_pdf_mode": "reportlab_chunk_visualizer_compact",
                "pdf_status": pdf_status,
                "download_type": out_ext.strip("."),
                "download_filename": xlsx_output_filename,
                "summary": final_chunks.get("summary", {}),
                "metadata_analysis": final_chunks.get("metadata_analysis", {}),
                "live_excel_analysis": final_chunks.get("live_excel_analysis", {}),
                "excel_field_mapping": final_chunks.get("excel_field_mapping", {}),
                "chunks_preview": chunks_preview,
                "validation_errors": final_chunks.get("validation_errors", []),
            }
        )

    except HTTPException:
        safe_delete_dir(work_dir)
        raise

    except ValueError as e:
        # Known PBIX structural limitation — return 422 Unprocessable Entity
        safe_delete_dir(work_dir)
        logger.warning("PBIX processing validation failed: %s", e)
        raise HTTPException(status_code=422, detail=f"Processing failed: {str(e)}")

    except Exception as e:
        # Unexpected server error — return 500
        safe_delete_dir(work_dir)
        logger.exception("Upload processing failed")
        raise HTTPException(status_code=500, detail=f"Processing failed: {str(e)}")


# ── Routes ───────────────────────────────────────────────────────────────────


@app.get("/", response_class=HTMLResponse)
def read_root():
    index_path = get_index_html_path()
    with open(index_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/health")
def health():
    return {
        "status": "running",
        "app": "Power BI PBIX to Excel Converter",
        "temp_dir": str(TEMP_DIR),
        "max_upload_size_bytes": MAX_UPLOAD_SIZE_BYTES,
        "cors_origins": ALLOWED_ORIGINS,
        "platform": platform.system(),
        "live_connect_available": LIVE_CONNECT_AVAILABLE,
        "live_connect_unavailable_reason": LIVE_CONNECT_UNAVAILABLE_REASON,
    }


@app.get("/hf-status")
def hf_status():
    return check_huggingface_connectivity()


@app.post("/upload")
async def upload_file(
    file: UploadFile = File(...),
    screenshot: Optional[UploadFile] = File(None),
    base_template: Optional[UploadFile] = File(None),
    tmdl_metadata: Optional[UploadFile] = File(None),
):
    return await process_upload_file(file, screenshot, base_template, tmdl_metadata)


@app.post("/upload-metadata")
async def upload_metadata(
    file: UploadFile = File(...),
    screenshot: Optional[UploadFile] = File(None),
    base_template: Optional[UploadFile] = File(None),
    tmdl_metadata: Optional[UploadFile] = File(None),
):
    return await process_upload_file(file, screenshot, base_template, tmdl_metadata)


@app.post("/upload-stream")
async def upload_stream(
    file: UploadFile = File(...),
    screenshot: Optional[UploadFile] = File(None),
    base_template: Optional[UploadFile] = File(None),
    tmdl_metadata: Optional[UploadFile] = File(None),
):
    return await process_upload_file(file, screenshot, base_template, tmdl_metadata)


@app.get("/upload")
def upload_get_help():
    return JSONResponse(
        status_code=405,
        content={
            "detail": "Use POST /upload with multipart/form-data field name 'file'. Only .pbix files are supported."
        },
    )


@app.get("/upload-metadata")
def upload_metadata_get_help():
    return JSONResponse(
        status_code=405,
        content={
            "detail": "Use POST /upload-metadata with multipart/form-data field name 'file'. Only .pbix files are supported."
        },
    )


@app.get("/upload-stream")
def upload_stream_get_help():
    return JSONResponse(
        status_code=405,
        content={
            "detail": "Use POST /upload-stream with multipart/form-data field name 'file'. Only .pbix files are supported."
        },
    )


@app.get("/download/{session_id}")
def download_excel(session_id: str):
    safe_session_id = os.path.basename(session_id)
    work_dir = TEMP_DIR / safe_session_id

    if not work_dir.exists() or not work_dir.is_dir():
        raise HTTPException(status_code=404, detail="Session not found or expired.")

    xlsx_path_xlsx = work_dir / "excel_ready_model.xlsx"
    xlsx_path_xlsm = work_dir / "excel_ready_model.xlsm"

    if xlsx_path_xlsm.exists():
        return FileResponse(
            path=str(xlsx_path_xlsm),
            filename="excel_ready_model.xlsm",
            media_type="application/vnd.ms-excel.sheet.macroEnabled.12",
        )
    elif xlsx_path_xlsx.exists():
        return FileResponse(
            path=str(xlsx_path_xlsx),
            filename="excel_ready_model.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    raise HTTPException(
        status_code=404, detail="Excel output file not found for this session."
    )


@app.get("/download-preview/{session_id}")
def download_preview_pdf(session_id: str):
    safe_session_id = os.path.basename(session_id)
    work_dir = TEMP_DIR / safe_session_id

    if not work_dir.exists() or not work_dir.is_dir():
        raise HTTPException(status_code=404, detail="Session not found or expired.")

    preview_path = work_dir / "dashboard_preview.pdf"
    validation = _validate_pdf_output(preview_path)
    if validation.get("status") != "success":
        raise HTTPException(
            status_code=404,
            detail={
                "message": "Preview PDF not found or invalid for this session.",
                "reason": validation.get("error"),
            },
        )

    return FileResponse(
        path=str(preview_path),
        filename="powerbi_chunk_visualizer_preview.pdf",
        media_type="application/pdf",
        headers={
            "Cache-Control": "no-store",
            "X-PDF-Page-Count": str(validation.get("page_count", 0)),
            "X-PDF-File-Size": str(validation.get("file_size_bytes", 0)),
        },
    )


@app.get("/download-json/{session_id}")
def download_json(session_id: str):
    safe_session_id = os.path.basename(session_id)
    work_dir = TEMP_DIR / safe_session_id

    if not work_dir.exists() or not work_dir.is_dir():
        raise HTTPException(status_code=404, detail="Session not found or expired.")

    json_path = work_dir / "converted_metadata.json"
    if not json_path.exists():
        raise HTTPException(
            status_code=404, detail="JSON output not found for this session."
        )

    return FileResponse(
        path=str(json_path),
        filename="converted_metadata.json",
        media_type="application/json",
    )


# ── Live-connect endpoints ─────────────────────────────────────────────────


def _require_live_connect():
    if (
        not LIVE_CONNECT_AVAILABLE
        or session_manager is None
        or launch_excel_for_connection is None
        or run_continue_workflow_com is None
        or create_all_visual_bindings is None
    ):
        detail = (
            "Live-connect is not available on this server. "
            "Ensure: (1) Server runs from project root directory, "
            "(2) pywin32 is installed in the virtual environment, "
            "(3) Server runs on Windows. "
            "Start with: start-server.bat or start-server.ps1"
        )
        if LIVE_CONNECT_UNAVAILABLE_REASON:
            detail += f" Reason: {LIVE_CONNECT_UNAVAILABLE_REASON}"
        raise HTTPException(status_code=503, detail=detail)


@app.post("/live-connect/start")
def live_connect_start(body: dict = Body(default={})):
    """Launch Excel visibly and wait for the user to connect a Power BI model.

    Expects JSON body: {"session_id": "<id from prior /upload call>"}
    If an active live session already exists for this upload session, returns
    its current state without launching a new Excel window.
    """
    _require_live_connect()

    upload_session_id = str(body.get("session_id") or "").strip()
    if not upload_session_id:
        raise HTTPException(status_code=400, detail="session_id is required.")

    # --- Prevent duplicate sessions for the same upload ---------------------
    existing_session = session_manager.get_active_session_by_upload_id(
        upload_session_id
    )
    if existing_session is not None:
        logger.info(
            "Live-connect: reusing existing session %s for upload %s.",
            existing_session.session_id,
            upload_session_id,
        )
        status = existing_session.to_status_dict()
        return JSONResponse(
            {
                **status,
                "upload_session_id": upload_session_id,
                "message": "Reusing existing Excel session. Excel is already open.",
            }
        )

    safe_id = os.path.basename(upload_session_id)
    work_dir = TEMP_DIR / safe_id
    chunks_path = work_dir / "converted_metadata.json"

    if not chunks_path.exists():
        raise HTTPException(
            status_code=404,
            detail=f"Upload session {upload_session_id!r} not found. Upload the PBIX first.",
        )

    try:
        with open(chunks_path, "r", encoding="utf-8") as fh:
            payload = json.load(fh)
        final_chunks = payload.get("chunks") or {}
    except Exception as exc:
        raise HTTPException(
            status_code=500, detail=f"Could not read upload metadata: {exc}"
        )

    # Derive visual bindings from the chunks
    visual_bindings = []
    try:
        visual_bindings = create_all_visual_bindings(final_chunks)
    except Exception as exc:
        logger.warning("create_all_visual_bindings failed: %s", exc)

    # Output path for the live workbook
    live_session_id = str(uuid.uuid4())
    live_out_dir = work_dir
    live_out_dir.mkdir(parents=True, exist_ok=True)
    output_path = str(live_out_dir / "excel_ready_model_live.xlsx")

    session = session_manager.create_session(
        metadata=final_chunks,
        visual_bindings=visual_bindings,
        formula_chunks=[],
        output_path=output_path,
        session_id=live_session_id,
        upload_session_id=upload_session_id,
    )

    # Store mapping so /download-live can find the file
    _LIVE_SESSION_MAP[live_session_id] = output_path

    # Dispatch Excel launch to the COM thread (non-blocking)
    session.dispatch(lambda: launch_excel_for_connection(session), wait=False)

    logger.info(
        "Live-connect session %s started for upload session %s.",
        live_session_id,
        upload_session_id,
    )

    return JSONResponse(
        {
            "session_id": live_session_id,
            "upload_session_id": upload_session_id,
            "state": "excel_launching",
            "message": (
                "Excel is opening. Once it is visible:\n"
                "1. Sign in with your Microsoft account.\n"
                "2. In Excel, go to Insert > PivotTable > From Power BI.\n"
                "3. Select the correct published semantic model.\n"
                "4. Insert the PivotTable.\n"
                "5. Return here and click 'Connection Completed'."
            ),
        }
    )


# In-memory map: live_session_id -> output file path
_LIVE_SESSION_MAP: dict = {}


@app.get("/live-connect/{session_id}/status")
def live_connect_status(session_id: str):
    """Return the current state and progress counters for a live-connect session."""
    _require_live_connect()
    safe_id = os.path.basename(session_id)
    session = session_manager.get_session(safe_id)
    if session is None:
        raise HTTPException(status_code=404, detail="Live-connect session not found.")
    return JSONResponse(session.to_status_dict())


@app.post("/live-connect/{session_id}/continue")
def live_connect_continue(session_id: str):
    """User clicked 'Connection Completed'.

    Transitions the session to detecting_connection and enqueues the compound
    detect+validate+build workflow on the COM thread.  Returns immediately with
    state=detecting_connection so the frontend can start polling /status.
    """
    _require_live_connect()
    safe_id = os.path.basename(session_id)
    session = session_manager.get_session(safe_id)
    if session is None:
        raise HTTPException(status_code=404, detail="Live-connect session not found.")

    if session.state not in (
        "waiting_for_user_connection",
        "connection_not_detected",
        "semantic_model_mismatch",
    ):
        raise HTTPException(
            status_code=409,
            detail=(
                f"Session is in state {session.state!r}. "
                "Only call /continue when Excel is waiting for connection."
            ),
        )

    with session.lock:
        if session.continue_job_enqueued:
            return JSONResponse(
                {
                    "session_id": safe_id,
                    "state": session.state,
                    "message": "Continue request already in progress for this session.",
                }
            )
        session.continue_job_enqueued = True
        session.update_state("detecting_connection")

    def _continue_workflow() -> None:
        try:
            run_continue_workflow_com(session)
        finally:
            session.continue_job_enqueued = False

    try:
        session.dispatch(_continue_workflow, wait=False)
    except Exception as exc:
        with session.lock:
            session.continue_job_enqueued = False
        logger.exception("Failed to enqueue continue workflow")
        raise HTTPException(status_code=500, detail=str(exc))

    return JSONResponse(
        {
            "session_id": safe_id,
            "state": "detecting_connection",
            "message": "Detection started. Poll /status for progress.",
        }
    )


@app.post("/live-connect/{session_id}/cancel")
def live_connect_cancel(session_id: str):
    """Cancel a live-connect session and close only its Excel instance."""
    _require_live_connect()
    safe_id = os.path.basename(session_id)
    found = session_manager.cancel_session(safe_id)
    if not found:
        raise HTTPException(status_code=404, detail="Live-connect session not found.")
    return JSONResponse({"session_id": safe_id, "state": "cancelled"})


@app.get("/live-connect/{session_id}/report")
def live_connect_report(session_id: str):
    """Return the full JSON conversion report for a completed live session."""
    _require_live_connect()
    safe_id = os.path.basename(session_id)
    session = session_manager.get_session(safe_id)
    if session is None:
        raise HTTPException(status_code=404, detail="Live-connect session not found.")
    if session.state != "completed_live" or not session.post_save_verification_passed:
        raise HTTPException(
            status_code=409,
            detail=f"Report not yet available (state: {session.state!r}).",
        )
    from fastapi.encoders import jsonable_encoder

    safe_result = {
        key: value
        for key, value in session.result.items()
        if not str(key).startswith("_") and key != "runtime_objects"
    }
    return JSONResponse(
        content=jsonable_encoder(
            {
                "session_id": safe_id,
                **safe_result,
                "warnings": list(session.warnings),
                "errors": list(session.errors),
            }
        )
    )


@app.get("/live-connect/{session_id}/download")
def download_live_excel(session_id: str):
    """Download the live-connected Excel workbook for a completed session."""
    _require_live_connect()
    safe_id = os.path.basename(session_id)
    session = session_manager.get_session(safe_id)
    if session is None:
        raise HTTPException(status_code=404, detail="Live-connect session not found.")
    if session.state != "completed_live" or not session.post_save_verification_passed:
        raise HTTPException(
            status_code=409,
            detail=(
                f"Workbook not ready or verification failed "
                f"(state: {session.state!r}, verified: {session.post_save_verification_passed})."
            ),
        )
    output_path = session.result.get("output_path") or session.output_path
    if not output_path or not Path(output_path).exists():
        raise HTTPException(
            status_code=404,
            detail="Live Excel output not found. Ensure conversion completed successfully.",
        )
    fname = Path(output_path).name
    return FileResponse(
        path=output_path,
        filename=fname,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# Keep legacy endpoint for backwards compatibility
@app.get("/download-live/{session_id}")
def download_live_excel_legacy(session_id: str):
    """Legacy download endpoint — redirects to /live-connect/{id}/download."""
    return download_live_excel(session_id)

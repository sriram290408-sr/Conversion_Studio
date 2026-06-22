"""
powerbi_live.py  –  Live Power BI Semantic Model Data Connection
================================================================

Handles authentication, DAX query generation, and live data fetching
from a Power BI Premium / Fabric semantic model via the Execute Queries
REST API.

Data priority used by the caller (app/Convertor.py):
    1. Live Power BI query result       (this module)
    2. StaticData/*.json aggregation    (Convertor.py)
    3. Uploaded CSV/XLSX               (future)
    4. Mock data fallback              (Convertor.py)

Security rules enforced here:
    - Secrets (client_secret, access_token) are NEVER logged.
    - Secrets are NEVER written to workbook sheets.
    - Only trusted DAX generated from visual chunk metadata is executed.
    - Arbitrary user DAX is never executed.
"""

import os
import re
import json
import logging
import time
from typing import Any, Dict, List, Optional

logger = logging.getLogger("powerbi_live")

# ─────────────────────────────────────────────────────────────────────────────
#  ENVIRONMENT CONFIG
# ─────────────────────────────────────────────────────────────────────────────

POWERBI_LIVE_ENABLED = os.getenv("POWERBI_LIVE_ENABLED", "false").lower() == "true"
POWERBI_TENANT_ID = os.getenv("POWERBI_TENANT_ID", "")
POWERBI_CLIENT_ID = os.getenv("POWERBI_CLIENT_ID", "")
POWERBI_CLIENT_SECRET = os.getenv("POWERBI_CLIENT_SECRET", "")
POWERBI_WORKSPACE_ID = os.getenv("POWERBI_WORKSPACE_ID", "")
POWERBI_DATASET_ID = os.getenv("POWERBI_DATASET_ID", "")
POWERBI_AUTH_MODE = os.getenv("POWERBI_AUTH_MODE", "service_principal")
POWERBI_TIMEOUT_SECONDS = int(os.getenv("POWERBI_TIMEOUT_SECONDS", "45"))
POWERBI_MAX_RETRIES = int(os.getenv("POWERBI_MAX_RETRIES", "2"))


def _build_config_from_env() -> Dict[str, Any]:
    """Build live config dict from environment variables."""
    return {
        "enabled": POWERBI_LIVE_ENABLED,
        "tenant_id": POWERBI_TENANT_ID,
        "client_id": POWERBI_CLIENT_ID,
        "client_secret": POWERBI_CLIENT_SECRET,
        "workspace_id": POWERBI_WORKSPACE_ID,
        "dataset_id": POWERBI_DATASET_ID,
        "auth_mode": POWERBI_AUTH_MODE,
        "timeout": POWERBI_TIMEOUT_SECONDS,
        "max_retries": POWERBI_MAX_RETRIES,
    }


def load_live_config(session_config: Optional[Dict] = None) -> Dict[str, Any]:
    """
    Merge session-provided config with env defaults.
    session_config values override env for workspace_id / dataset_id / enabled.
    Secrets are NEVER taken from session_config for security.
    """
    base = _build_config_from_env()
    if session_config and isinstance(session_config, dict):
        # Allow only non-secret fields from session config
        if "enabled" in session_config:
            base["enabled"] = bool(session_config["enabled"])
        if "workspace_id" in session_config and session_config["workspace_id"]:
            base["workspace_id"] = str(session_config["workspace_id"])
        if "dataset_id" in session_config and session_config["dataset_id"]:
            base["dataset_id"] = str(session_config["dataset_id"])
    return base


def _is_config_complete(config: Dict) -> bool:
    """Return True only if all required fields for live connection are present."""
    return bool(
        config.get("enabled")
        and config.get("tenant_id")
        and config.get("client_id")
        and config.get("client_secret")
        and config.get("workspace_id")
        and config.get("dataset_id")
    )


# ─────────────────────────────────────────────────────────────────────────────
#  AUTHENTICATION
# ─────────────────────────────────────────────────────────────────────────────


def get_powerbi_access_token(config: Dict) -> str:
    """
    Obtain an OAuth2 access token from Microsoft Entra ID using
    service principal (client_credentials) flow.
    Returns empty string on failure (never raises).
    """
    if not _is_config_complete(config):
        logger.info("Live Power BI config incomplete — skipping auth.")
        return ""

    tenant_id = config["tenant_id"]
    client_id = config["client_id"]
    client_secret = config["client_secret"]
    timeout = int(config.get("timeout", 30))

    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    payload = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://analysis.windows.net/powerbi/api/.default",
    }

    try:
        import requests

        resp = requests.post(token_url, data=payload, timeout=timeout)
        resp.raise_for_status()
        token = resp.json().get("access_token", "")
        if token:
            logger.info("Power BI access token obtained successfully.")
        else:
            logger.warning("Power BI auth response did not contain access_token.")
        return token
    except Exception as e:
        # Never log the secret — only log a sanitised message
        logger.warning("Power BI authentication failed: %s", type(e).__name__)
        return ""


def authenticate_powerbi_service(config: Dict) -> str:
    """Alias for get_powerbi_access_token — returns access token string."""
    return get_powerbi_access_token(config)


# ─────────────────────────────────────────────────────────────────────────────
#  DAX QUERY SAFETY HELPERS
# ─────────────────────────────────────────────────────────────────────────────


def parse_field_reference(field: str) -> Dict[str, str]:
    """
    Parse 'Table[Column]' into {'table': 'Table', 'column': 'Column'}.
    Returns empty strings if format is unexpected.
    """
    field = str(field or "").strip()
    m = re.match(r"^(.+?)\[(.+?)\]$", field)
    if m:
        return {"table": m.group(1).strip(), "column": m.group(2).strip()}
    # Bare column name
    if field:
        return {"table": "", "column": field}
    return {"table": "", "column": ""}


def format_dax_column(field: str) -> str:
    """
    Convert 'Table[Column]' → \"'Table'[Column]\" (safe DAX column reference).
    Escapes single quotes inside table names.
    """
    ref = parse_field_reference(field)
    if not ref["column"]:
        return ""
    table = ref["table"].replace("'", "''")
    column = ref["column"]
    if table:
        return f"'{table}'[{column}]"
    return f"[{column}]"


def format_dax_measure(measure: str) -> str:
    """
    Convert 'Sum of Volume' → '[Sum of Volume]' (safe DAX measure reference).
    """
    measure = str(measure or "").strip().strip("[]")
    if not measure:
        return ""
    return f"[{measure}]"


def safe_dax_alias(name: str) -> str:
    """
    Return a safe DAX column alias (strips special chars, max 64 chars).
    """
    alias = re.sub(r'["\n\r\t]', " ", str(name or "")).strip()
    return alias[:64] if alias else "Value"


# ─────────────────────────────────────────────────────────────────────────────
#  DAX QUERY BUILDER
# ─────────────────────────────────────────────────────────────────────────────


def build_dax_query_for_visual(
    visual_chunk: Dict, selected_filters: Optional[Dict] = None
) -> Optional[str]:
    """
    Build a safe DAX query from visual chunk metadata.
    Returns None if insufficient metadata to build a query.
    Only generates DAX from trusted PBIX visual chunk data.
    """
    vt = str(visual_chunk.get("visual_type", "")).lower()
    fields = visual_chunk.get("uses_fields", []) or []
    measures = visual_chunk.get("uses_measures", []) or []

    # ── Determine normalized type ─────────────────────────────────────────
    try:
        from .Convertor import normalize_visual_type
    except ImportError:
        try:
            from Convertor import normalize_visual_type
        except ImportError:
            normalize_visual_type = lambda x: "placeholder"

    norm = normalize_visual_type(vt)

    # ── KPI / card ────────────────────────────────────────────────────────
    if norm == "kpi":
        if not measures:
            return None
        mea_ref = format_dax_measure(measures[0])
        alias = safe_dax_alias(measures[0])
        return f"EVALUATE\n" f"ROW(\n" f'    "{alias}", {mea_ref}\n' f")"

    # ── Slicer — use VALUES() ─────────────────────────────────────────────
    if norm == "slicer":
        if not fields:
            return None
        col_ref = format_dax_column(fields[0])
        if not col_ref:
            return None
        return f"EVALUATE\nVALUES({col_ref})"

    # ── Charts / treemap / map — SUMMARIZECOLUMNS ────────────────────────
    if norm in (
        "line_chart",
        "column_chart",
        "bar_chart",
        "pie_chart",
        "donut_chart",
        "treemap",
        "map",
    ):
        if not fields:
            return None
        col_ref = format_dax_column(fields[0])
        if not col_ref:
            return None
        if measures:
            mea_ref = format_dax_measure(measures[0])
            alias = safe_dax_alias(measures[0])
            return (
                f"EVALUATE\n"
                f"SUMMARIZECOLUMNS(\n"
                f"    {col_ref},\n"
                f'    "{alias}", {mea_ref}\n'
                f")"
            )
        else:
            # Count rows grouped by field
            ref = parse_field_reference(fields[0])
            alias = safe_dax_alias(f"Count of {ref['column']}")
            return (
                f"EVALUATE\n"
                f"SUMMARIZECOLUMNS(\n"
                f"    {col_ref},\n"
                f'    "{alias}", COUNTROWS(VALUES({col_ref}))\n'
                f")"
            )

    # ── Table / matrix — show top 50 rows ────────────────────────────────
    if norm in ("table", "matrix"):
        if not fields:
            return None
        cols_dax = "\n    ".join(format_dax_column(f) for f in fields[:6] if f)
        if not cols_dax:
            return None
        return (
            f"EVALUATE\n"
            f"TOPN(\n"
            f"    50,\n"
            f"    SELECTCOLUMNS(\n"
            f"        DISTINCT(SELECTCOLUMNS(ALLNOBLANKROW({format_dax_column(fields[0])}), "
            f'"_key_", {format_dax_column(fields[0])})),\n'
            f"        {cols_dax}\n"
            f"    )\n"
            f")"
        )

    return None


def build_dax_filter_query(field: str) -> Optional[str]:
    """Build a VALUES() DAX query to fetch distinct values for a slicer field."""
    col_ref = format_dax_column(field)
    if not col_ref:
        return None
    return f"EVALUATE\nVALUES({col_ref})"


# ─────────────────────────────────────────────────────────────────────────────
#  DAX QUERY EXECUTION
# ─────────────────────────────────────────────────────────────────────────────


def run_powerbi_dax_query(
    access_token: str,
    workspace_id: str,
    dataset_id: str,
    dax_query: str,
    timeout: int = POWERBI_TIMEOUT_SECONDS,
) -> List[Dict]:
    """
    Execute a DAX query against a Power BI dataset using the
    Execute Queries REST API.
    Returns list of row dicts on success, empty list on failure.
    """
    if not access_token or not workspace_id or not dataset_id or not dax_query:
        return []

    url = (
        f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}"
        f"/datasets/{dataset_id}/executeQueries"
    )
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
    body = {
        "queries": [{"query": dax_query}],
        "serializerSettings": {"includeNulls": True},
    }

    try:
        import requests

        resp = requests.post(
            url, headers=headers, data=json.dumps(body), timeout=timeout
        )
        resp.raise_for_status()
        data = resp.json()
        # Navigate: results[0].tables[0].rows
        results = data.get("results", [])
        if not results:
            return []
        tables = results[0].get("tables", [])
        if not tables:
            return []
        rows = tables[0].get("rows", [])
        # Strip column name prefixes like "TableName[ColName]" → "ColName"
        cleaned = []
        for row in rows:
            clean_row = {}
            for k, v in row.items():
                # Keys look like "Date[MonthName]" → keep bare column name
                bare = re.sub(r"^[^[]+\[(.+)\]$", r"\1", k)
                clean_row[bare] = v
            cleaned.append(clean_row)
        return cleaned
    except Exception as e:
        logger.warning("DAX query execution failed: %s", type(e).__name__)
        return []


def test_powerbi_connection(config: Dict) -> Dict:
    """
    Run a minimal DAX query to verify connectivity.
    Returns status dict (never contains secrets).
    """
    result = {
        "status": "not_configured",
        "enabled": config.get("enabled", False),
        "workspace_id": config.get("workspace_id", ""),
        "dataset_id": config.get("dataset_id", ""),
        "error": None,
    }
    if not _is_config_complete(config):
        result["status"] = "not_configured"
        return result
    try:
        token = get_powerbi_access_token(config)
        if not token:
            result["status"] = "auth_failed"
            result["error"] = "Could not obtain access token."
            return result
        rows = run_powerbi_dax_query(
            token,
            config["workspace_id"],
            config["dataset_id"],
            'EVALUATE ROW("Test", 1)',
            timeout=config.get("timeout", 30),
        )
        if rows is not None:
            result["status"] = "connected"
        else:
            result["status"] = "query_failed"
            result["error"] = "DAX test query returned no result."
    except Exception as e:
        result["status"] = "error"
        result["error"] = type(e).__name__
    return result


# ─────────────────────────────────────────────────────────────────────────────
#  PER-VISUAL DATA FETCHING
# ─────────────────────────────────────────────────────────────────────────────


def fetch_live_visual_data(visual_chunk: Dict, config: Dict) -> Dict:
    """
    Fetch live tabular data for a single visual chunk.
    Returns {'headers': [...], 'rows': [...], 'source': 'live_powerbi'} or {}.
    """
    if not _is_config_complete(config):
        return {}
    dax = build_dax_query_for_visual(visual_chunk)
    if not dax:
        return {}
    token = get_powerbi_access_token(config)
    if not token:
        return {}
    rows = run_powerbi_dax_query(
        token,
        config["workspace_id"],
        config["dataset_id"],
        dax,
        timeout=config.get("timeout", POWERBI_TIMEOUT_SECONDS),
    )
    if not rows:
        return {}
    headers = list(rows[0].keys()) if rows else []
    data_rows = [[r.get(h) for h in headers] for r in rows]
    return {
        "headers": headers,
        "rows": data_rows,
        "source": "live_powerbi",
    }


def fetch_live_filter_values(field_name: str, config: Dict) -> List[str]:
    """
    Fetch distinct values for a slicer field using VALUES() DAX query.
    Always prepends 'All'. Returns ['All'] on failure.
    """
    if not _is_config_complete(config):
        return ["All"]
    dax = build_dax_filter_query(field_name)
    if not dax:
        return ["All"]
    token = get_powerbi_access_token(config)
    if not token:
        return ["All"]
    rows = run_powerbi_dax_query(
        token,
        config["workspace_id"],
        config["dataset_id"],
        dax,
        timeout=config.get("timeout", POWERBI_TIMEOUT_SECONDS),
    )
    if not rows:
        return ["All"]
    # Extract the first value from each row
    values = []
    for row in rows:
        v = list(row.values())[0] if row else None
        if v is not None and str(v).strip():
            values.append(str(v).strip())
    # Sort and prepend All
    unique_sorted = sorted(set(values))[:50]
    return ["All"] + unique_sorted


# ─────────────────────────────────────────────────────────────────────────────
#  BATCH LIVE DATA FETCHER
# ─────────────────────────────────────────────────────────────────────────────


def build_live_data_for_visuals(visual_chunks: List[Dict], config: Dict) -> Dict:
    """
    Fetch live data for all visual chunks in one pass.
    Returns:
    {
        'chart_sources': {safe_key: {title, headers, rows, source}},
        'kpi_values':    {measure_name: scalar_value},
        'filter_values': {field_key: [All, val1, ...]},
        'errors':        [str, ...]
    }
    On failure of individual visual, records error and continues.
    """
    result: Dict = {
        "chart_sources": {},
        "kpi_values": {},
        "filter_values": {},
        "errors": [],
        "live_enabled": False,
    }

    if not _is_config_complete(config):
        logger.info("Power BI live data: not configured.")
        return result

    # Get one token for all queries in this batch
    token = get_powerbi_access_token(config)
    if not token:
        result["errors"].append("Authentication failed — no token obtained.")
        return result

    result["live_enabled"] = True
    logger.info("Power BI live: querying %d visuals.", len(visual_chunks))

    try:
        from .Convertor import normalize_visual_type
    except ImportError:
        try:
            from Convertor import normalize_visual_type
        except ImportError:
            normalize_visual_type = lambda x: "placeholder"

    workspace_id = config["workspace_id"]
    dataset_id = config["dataset_id"]
    timeout = int(config.get("timeout", POWERBI_TIMEOUT_SECONDS))

    for vc in visual_chunks:
        vt = vc.get("visual_type", "")
        title = vc.get("visual_title", "") or vt or "Visual"
        safe_key = re.sub(r"[^a-z0-9_]", "_", title.lower())[:40] or "visual"
        norm = normalize_visual_type(vt)

        try:
            dax = build_dax_query_for_visual(vc)
            if not dax:
                continue

            rows = run_powerbi_dax_query(
                token, workspace_id, dataset_id, dax, timeout=timeout
            )
            if not rows:
                continue

            headers = list(rows[0].keys()) if rows else []
            data_rows = [[r.get(h) for h in headers] for r in rows]

            # ── KPI scalar ───────────────────────────────────────────
            if norm == "kpi":
                measures = vc.get("uses_measures", [])
                if measures and data_rows and data_rows[0]:
                    result["kpi_values"][measures[0]] = data_rows[0][0]

            # ── Slicer filter values ──────────────────────────────────
            elif norm == "slicer":
                fields = vc.get("uses_fields", [])
                if fields and data_rows:
                    vals = ["All"] + [
                        str(r[0]) for r in data_rows if r and r[0] is not None
                    ][:50]
                    result["filter_values"][fields[0]] = vals

            # ── Charts, treemap, map, table ───────────────────────────
            else:
                result["chart_sources"][safe_key] = {
                    "title": title,
                    "headers": headers,
                    "rows": data_rows,
                    "source": "live_powerbi",
                }

        except Exception as e:
            err_msg = f"Visual '{title}' ({vt}): {type(e).__name__}"
            logger.warning("Live data fetch error — %s", err_msg)
            result["errors"].append(err_msg)

    # ── Fetch filter values for slicer fields not already done ───────────
    for vc in visual_chunks:
        norm = normalize_visual_type(vc.get("visual_type", ""))
        fields = vc.get("uses_fields", []) or []
        if norm == "slicer" and fields and fields[0] not in result["filter_values"]:
            try:
                vals = fetch_live_filter_values(fields[0], config)
                result["filter_values"][fields[0]] = vals
            except Exception as e:
                logger.warning(
                    "Filter value fetch error for %s: %s", fields[0], type(e).__name__
                )

    logger.info(
        "Power BI live fetch complete. chart_sources=%d, kpi_values=%d, "
        "filter_values=%d, errors=%d",
        len(result["chart_sources"]),
        len(result["kpi_values"]),
        len(result["filter_values"]),
        len(result["errors"]),
    )
    return result

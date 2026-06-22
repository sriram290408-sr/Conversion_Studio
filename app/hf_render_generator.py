"""Generate constrained Excel render operations with Hugging Face.

This module is intentionally limited to presentation-only changes:
- op
- layout
- style
- title
- page_name

Semantic fields, measures, filters, connection names, formulas and source bindings
always remain owned by the deterministic render plan.
"""

from __future__ import annotations

import json
import logging
import os
import re
import time
from typing import Any, Dict, List, Optional, Tuple

import requests

logger = logging.getLogger("hf_render_generator")

ALLOWED_OPS = {
    "create_card",
    "create_kpi",
    "create_gauge",
    "create_slicer",
    "create_column_chart",
    "create_bar_chart",
    "create_line_chart",
    "create_area_chart",
    "create_pie_chart",
    "create_donut_chart",
    "create_table",
    "create_matrix",
    "create_treemap",
    "create_map",
    "create_textbox",
    "create_image",
    "create_shape",
    "create_placeholder",
}

PRESENTATION_KEYS = {
    "visual_id",
    "op",
    "page_name",
    "title",
    "layout",
    "style",
}

RETRYABLE_STATUS_CODES = {408, 409, 425, 429, 500, 502, 503, 504}


def _bool_env(name: str, default: bool = False) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return str(value).strip().lower() in {"1", "true", "yes", "on"}


def _int_env(name: str, default: int) -> int:
    try:
        return int(str(os.getenv(name, default)).strip())
    except (TypeError, ValueError):
        return default


def _float_env(name: str, default: float) -> float:
    try:
        return float(str(os.getenv(name, default)).strip())
    except (TypeError, ValueError):
        return default


def _extract_json(text: str) -> Optional[Dict[str, Any]]:
    """Extract the first valid JSON object from an HF response."""
    cleaned = re.sub(r"```(?:json)?|```", "", str(text or ""), flags=re.I).strip()

    if not cleaned:
        return None

    try:
        direct = json.loads(cleaned)
        return direct if isinstance(direct, dict) else None
    except json.JSONDecodeError:
        pass

    decoder = json.JSONDecoder()
    for index, char in enumerate(cleaned):
        if char != "{":
            continue
        try:
            value, _ = decoder.raw_decode(cleaned[index:])
            if isinstance(value, dict):
                return value
        except json.JSONDecodeError:
            continue

    return None


def _safe_operation(operation: Any) -> Optional[Dict[str, Any]]:
    """Keep only presentation fields and reject invalid operations."""
    if not isinstance(operation, dict):
        return None

    visual_id = str(operation.get("visual_id") or "").strip()
    op = str(operation.get("op") or "").strip()

    if not visual_id or op not in ALLOWED_OPS:
        return None

    safe: Dict[str, Any] = {
        key: operation.get(key)
        for key in PRESENTATION_KEYS
        if key in operation
    }

    safe["visual_id"] = visual_id
    safe["op"] = op

    if not isinstance(safe.get("layout"), dict):
        safe["layout"] = {}

    if not isinstance(safe.get("style"), dict):
        safe["style"] = {}

    safe["page_name"] = str(safe.get("page_name") or "").strip()
    safe["title"] = str(safe.get("title") or "").strip()

    return safe


def _deterministic_index(render_plan: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    result: Dict[str, Dict[str, Any]] = {}
    for operation in render_plan.get("operations", []) or []:
        if not isinstance(operation, dict):
            continue
        visual_id = str(operation.get("visual_id") or "").strip()
        if visual_id:
            result[visual_id] = dict(operation)
    return result


def _merge_with_deterministic_plan(
    render_plan: Dict[str, Any],
    hf_operations: List[Any],
) -> Tuple[List[Dict[str, Any]], List[str]]:
    """Merge HF presentation changes without allowing semantic mutation."""
    source_by_id = _deterministic_index(render_plan)
    warnings: List[str] = []
    merged: List[Dict[str, Any]] = []

    hf_by_id: Dict[str, Dict[str, Any]] = {}
    for raw in hf_operations:
        safe = _safe_operation(raw)
        if not safe:
            warnings.append("Ignored an invalid HF render operation.")
            continue

        visual_id = safe["visual_id"]
        if visual_id not in source_by_id:
            warnings.append(
                f"Ignored HF operation for unknown visual_id '{visual_id}'."
            )
            continue

        hf_by_id[visual_id] = safe

    for visual_id, deterministic in source_by_id.items():
        candidate = hf_by_id.get(visual_id)
        if not candidate:
            merged.append(deterministic)
            continue

        final = dict(deterministic)

        # Confirmed metadata visual type/op is authoritative.
        deterministic_op = str(deterministic.get("op") or "").strip()
        candidate_op = str(candidate.get("op") or "").strip()

        if deterministic_op in ALLOWED_OPS:
            final["op"] = deterministic_op
            if candidate_op and candidate_op != deterministic_op:
                warnings.append(
                    f"HF op change rejected for '{visual_id}': "
                    f"{candidate_op} -> {deterministic_op}."
                )
        elif candidate_op in ALLOWED_OPS:
            final["op"] = candidate_op

        # HF may improve only presentation fields.
        if candidate.get("page_name"):
            final["page_name"] = candidate["page_name"]

        if candidate.get("title"):
            final["title"] = candidate["title"]

        if isinstance(candidate.get("layout"), dict) and candidate["layout"]:
            final["layout"] = {
                **dict(deterministic.get("layout") or {}),
                **candidate["layout"],
            }

        if isinstance(candidate.get("style"), dict) and candidate["style"]:
            final["style"] = {
                **dict(deterministic.get("style") or {}),
                **candidate["style"],
            }

        merged.append(final)

    return merged, warnings


def _response_content(payload: Dict[str, Any]) -> str:
    """Read content from OpenAI-compatible HF router responses."""
    choices = payload.get("choices") or []
    if choices and isinstance(choices[0], dict):
        message = choices[0].get("message") or {}
        content = message.get("content", "")

        if isinstance(content, str):
            return content

        if isinstance(content, list):
            parts: List[str] = []
            for item in content:
                if isinstance(item, dict):
                    text = item.get("text") or item.get("content")
                    if text:
                        parts.append(str(text))
            return "\n".join(parts)

    generated_text = payload.get("generated_text")
    if isinstance(generated_text, str):
        return generated_text

    return ""


def _request_model(
    *,
    url: str,
    token: str,
    model: str,
    prompt: str,
    timeout: int,
    max_tokens: int,
    temperature: float,
) -> Dict[str, Any]:
    response = requests.post(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json={
            "model": model,
            "messages": [
                {
                    "role": "system",
                    "content": (
                        "Return one valid JSON object only. "
                        "Do not include markdown or explanatory text."
                    ),
                },
                {
                    "role": "user",
                    "content": prompt,
                },
            ],
            "temperature": temperature,
            "max_tokens": max_tokens,
            "response_format": {"type": "json_object"},
        },
        timeout=timeout,
    )

    if response.status_code == 400:
        # Some HF providers do not support response_format.
        response = requests.post(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json={
                "model": model,
                "messages": [
                    {
                        "role": "system",
                        "content": (
                            "Return one valid JSON object only. "
                            "Do not include markdown or explanatory text."
                        ),
                    },
                    {
                        "role": "user",
                        "content": prompt,
                    },
                ],
                "temperature": temperature,
                "max_tokens": max_tokens,
            },
            timeout=timeout,
        )

    response.raise_for_status()
    payload = response.json()

    if not isinstance(payload, dict):
        raise ValueError("HF returned a non-object response.")

    return payload


def generate_render_operations_with_hf(
    render_plan: Dict[str, Any],
) -> Dict[str, Any]:
    """Generate presentation-only render operations.

    Any HF failure, unsupported provider, quota error or malformed JSON safely
    falls back to the deterministic render plan.
    """
    deterministic_operations = list(render_plan.get("operations", []) or [])

    if not _bool_env("HF_RENDER_PLAN_ENABLED", True):
        return {
            "operations": deterministic_operations,
            "warnings": [],
            "source": "deterministic_disabled",
        }

    token = str(os.getenv("HF_API_TOKEN", "")).strip()
    if not token:
        return {
            "operations": deterministic_operations,
            "warnings": ["HF_API_TOKEN is not configured."],
            "source": "deterministic_no_token",
        }

    primary_model = str(
        os.getenv("HF_MODEL_ID", "Qwen/Qwen2.5-Coder-32B-Instruct")
    ).strip()
    fallback_model = str(
        os.getenv("HF_FALLBACK_MODEL_ID", "meta-llama/Llama-3.1-8B-Instruct")
    ).strip()

    models = [model for model in (primary_model, fallback_model) if model]
    models = list(dict.fromkeys(models))

    url = str(
        os.getenv(
            "HF_ROUTER_URL",
            "https://router.huggingface.co/v1/chat/completions",
        )
    ).strip()

    timeout = _int_env("HF_TIMEOUT_SECONDS", 60)
    max_retries = max(1, _int_env("HF_MAX_RETRIES", 2))
    max_tokens = max(500, _int_env("HF_RENDER_MAX_TOKENS", 5000))
    temperature = _float_env("HF_RENDER_TEMPERATURE", 0.05)

    prompt = (
        "You are an Excel dashboard render planner.\n"
        "Return JSON only.\n"
        "You may modify only these presentation fields: "
        "op, layout, style, title, page_name.\n"
        "Never create or modify Python, VBA, shell commands, imports, paths, "
        "URLs, formulas, table names, measure names, field names, filters, "
        "connection names, source sheets or semantic bindings.\n"
        "Never change a confirmed metadata visual type.\n"
        "A Power BI gauge must remain create_gauge.\n"
        "A visual title must never be converted into create_slicer.\n"
        "Allowed op values: "
        + ", ".join(sorted(ALLOWED_OPS))
        + '.\nOutput schema: '
        '{"operations":[{"visual_id":"","op":"","page_name":"","title":"",'
        '"layout":{},"style":{}}],"warnings":[]}.\n'
        "Use only visual_id values already present in the input plan.\n"
        "Input plan:\n"
        + json.dumps(render_plan, ensure_ascii=False, separators=(",", ":"))
    )

    errors: List[str] = []

    for model in models:
        for attempt in range(1, max_retries + 1):
            try:
                payload = _request_model(
                    url=url,
                    token=token,
                    model=model,
                    prompt=prompt,
                    timeout=timeout,
                    max_tokens=max_tokens,
                    temperature=temperature,
                )

                content = _response_content(payload)
                parsed = _extract_json(content)

                if not parsed or not isinstance(parsed.get("operations"), list):
                    raise ValueError("HF returned no valid operations JSON.")

                merged_operations, merge_warnings = _merge_with_deterministic_plan(
                    render_plan,
                    parsed["operations"],
                )

                return {
                    "operations": merged_operations,
                    "warnings": list(parsed.get("warnings") or []) + merge_warnings,
                    "source": "huggingface",
                    "model": model,
                }

            except requests.HTTPError as exc:
                status = exc.response.status_code if exc.response is not None else None
                body = ""
                if exc.response is not None:
                    try:
                        body = exc.response.text[:500]
                    except Exception:
                        body = ""

                message = (
                    f"HF model '{model}' attempt {attempt}/{max_retries} "
                    f"failed with HTTP {status}: {body or exc}"
                )
                errors.append(message)
                logger.warning(message)

                # 402 means provider credit/quota is unavailable.
                # Retrying the same model will not help.
                if status == 402:
                    break

                if status not in RETRYABLE_STATUS_CODES:
                    break

                if attempt < max_retries:
                    time.sleep(min(2 ** (attempt - 1), 4))

            except (requests.RequestException, ValueError, KeyError, TypeError) as exc:
                message = (
                    f"HF model '{model}' attempt {attempt}/{max_retries} failed: {exc}"
                )
                errors.append(message)
                logger.warning(message)

                if attempt < max_retries:
                    time.sleep(min(2 ** (attempt - 1), 4))

    logger.warning(
        "HF render generation failed; deterministic plan retained. Errors: %s",
        " | ".join(errors),
    )

    return {
        "operations": deterministic_operations,
        "warnings": errors or ["HF render generation failed."],
        "source": "deterministic_fallback",
    }


__all__ = ["generate_render_operations_with_hf"]

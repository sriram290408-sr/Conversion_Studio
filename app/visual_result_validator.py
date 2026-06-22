"""Validate native and approximate visual rendering results."""

from __future__ import annotations

from typing import Any, Dict, Iterable

SUCCESS_STATUSES = {
    "success",
    "live_approximation",
    "controlled_approximation",
}


def summarize_visual_results(results: Iterable[Dict[str, Any]]) -> Dict[str, Any]:
    items = list(results or [])
    native = 0
    approximated = 0
    failed = []

    for item in items:
        status = str(item.get("status") or "").casefold()
        rendered_as = str(item.get("rendered_as") or "").strip()
        if status == "success":
            native += 1
        elif status in {"live_approximation", "controlled_approximation"}:
            approximated += 1
        elif status not in SUCCESS_STATUSES:
            failed.append(
                {
                    "visual_id": item.get("visual_id"),
                    "error": item.get("error"),
                    "status": status,
                }
            )

    return {
        "visual_count": len(items),
        "native_visuals_created": native,
        "approximate_visuals_created": approximated,
        "failed_visuals": failed,
        "all_required_visuals_created": not failed,
    }


__all__ = ["summarize_visual_results"]

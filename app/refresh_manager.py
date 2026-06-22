"""Refresh and calculate a live-connected Excel workbook through COM."""

from __future__ import annotations

import logging
import time
from typing import Any, Callable, Iterable, Optional, Tuple

try:
    from live_binding_models import RefreshResult
except ImportError:
    from live_binding_models import RefreshResult

logger = logging.getLogger("refresh_manager")

XL_CALCULATION_DONE = 0

_RETRYABLE_COM_ERRORS = (
    "-2146777998",
    "800ac472",
    "-2147418111",
    "80010001",
    "-2147417846",
    "8001010a",
)


def _connection_name(connection: Any, index: int) -> str:
    try:
        value = str(connection.Name or "").strip()
        return value or f"Connection_{index}"
    except Exception:
        return f"Connection_{index}"


def _remaining_seconds(start_time: float, timeout_seconds: int) -> float:
    return max(0.0, float(timeout_seconds) - (time.monotonic() - start_time))


def _timed_out(start_time: float, timeout_seconds: int) -> bool:
    return _remaining_seconds(start_time, timeout_seconds) <= 0.0


def _is_retryable_com_error(exc: Exception) -> bool:
    text = str(exc).casefold()
    return any(token.casefold() in text for token in _RETRYABLE_COM_ERRORS)


def _call_with_retry(
    func: Callable[..., Any],
    *args: Any,
    retries: int = 30,
    delay: float = 0.35,
    **kwargs: Any,
) -> Any:
    last_error: Optional[Exception] = None

    for attempt in range(max(1, retries)):
        try:
            return func(*args, **kwargs)
        except Exception as exc:
            last_error = exc
            if not _is_retryable_com_error(exc) or attempt >= retries - 1:
                raise
            time.sleep(delay)

    if last_error is not None:
        raise last_error
    raise RuntimeError("COM call failed without an exception.")


def _pump_com_messages() -> None:
    """Allow Excel to finish pending COM and asynchronous refresh work."""
    try:
        import pythoncom  # type: ignore[import]

        pythoncom.PumpWaitingMessages()
    except Exception:
        pass


def _set_background_query_false(connection: Any) -> None:
    """Disable background refresh where the connection type supports it."""
    try:
        oledb = connection.OLEDBConnection
    except Exception:
        oledb = None

    if oledb is not None:
        try:
            oledb.BackgroundQuery = False
        except Exception:
            pass

    try:
        odbc = connection.ODBCConnection
    except Exception:
        odbc = None

    if odbc is not None:
        try:
            odbc.BackgroundQuery = False
        except Exception:
            pass


def _connection_refreshing(connection: Any) -> bool:
    """Return True when a workbook connection reports active refresh work."""
    for attribute in ("OLEDBConnection", "ODBCConnection"):
        try:
            provider = getattr(connection, attribute)
        except Exception:
            provider = None

        if provider is not None:
            try:
                if bool(provider.Refreshing):
                    return True
            except Exception:
                pass

    try:
        if bool(connection.Refreshing):
            return True
    except Exception:
        pass

    return False


def _pivot_cache_refreshing(cache: Any) -> bool:
    """Return True when a PivotCache reports active refresh work."""
    try:
        return bool(cache.Refreshing)
    except Exception:
        return False


def _excel_calculation_done(excel_app: Any) -> bool:
    try:
        return int(excel_app.CalculationState) == XL_CALCULATION_DONE
    except Exception:
        return True


def _collect_refresh_state(
    excel_app: Any,
    workbook: Any,
) -> Tuple[list[str], list[str], bool]:
    refreshing_connections: list[str] = []
    refreshing_caches: list[str] = []

    try:
        connection_count = int(_call_with_retry(lambda: workbook.Connections.Count))
    except Exception:
        connection_count = 0

    for index in range(1, connection_count + 1):
        try:
            connection = _call_with_retry(workbook.Connections.Item, index)
            if _connection_refreshing(connection):
                refreshing_connections.append(_connection_name(connection, index))
        except Exception:
            continue

    try:
        pivot_caches = _call_with_retry(lambda: workbook.PivotCaches())
        cache_count = int(_call_with_retry(lambda: pivot_caches.Count))
    except Exception:
        pivot_caches = None
        cache_count = 0

    if pivot_caches is not None:
        for index in range(1, cache_count + 1):
            try:
                cache = _call_with_retry(pivot_caches.Item, index)
                if _pivot_cache_refreshing(cache):
                    refreshing_caches.append(f"PivotCache_{index}")
            except Exception:
                continue

    return (
        refreshing_connections,
        refreshing_caches,
        _excel_calculation_done(excel_app),
    )


def _wait_for_refresh_idle(
    excel_app: Any,
    workbook: Any,
    start_time: float,
    timeout_seconds: int,
    poll_interval: float,
    stable_polls_required: int = 3,
) -> Tuple[bool, list[str], list[str]]:
    """Wait for connections, PivotCaches, and calculation to remain idle.

    Several consecutive idle polls are required because Excel can briefly report
    idle between OLAP refresh stages.
    """
    stable_polls = 0
    last_connections: list[str] = []
    last_caches: list[str] = []

    while not _timed_out(start_time, timeout_seconds):
        _pump_com_messages()

        connections, caches, calculation_done = _collect_refresh_state(
            excel_app,
            workbook,
        )
        last_connections = connections
        last_caches = caches

        if not connections and not caches and calculation_done:
            stable_polls += 1
            if stable_polls >= max(1, stable_polls_required):
                return True, [], []
        else:
            stable_polls = 0

        time.sleep(max(0.1, float(poll_interval)))

    return False, last_connections, last_caches


def _wait_for_excel_idle(
    excel_app: Any,
    start_time: float,
    timeout_seconds: int,
    poll_interval: float,
) -> bool:
    while not _timed_out(start_time, timeout_seconds):
        _pump_com_messages()
        if _excel_calculation_done(excel_app):
            return True
        time.sleep(max(0.1, float(poll_interval)))
    return False


def refresh_and_calculate_workbook(
    excel_app: Any,
    workbook: Any,
    timeout_seconds: int = 180,
    poll_interval: float = 0.5,
) -> RefreshResult:
    """Refresh a live Power BI workbook and return only after Excel is idle.

    Important behavior:
    - RefreshAll is called once.
    - Connections and PivotCaches are not refreshed a second time afterward.
    - Excel is required to remain idle for several consecutive polls.
    - The caller must not save when ``refresh_completed`` is False.
    """
    result = RefreshResult(
        refresh_started=True,
        timeout_seconds=max(1, int(timeout_seconds)),
    )
    start_time = time.monotonic()

    if excel_app is None:
        result.errors.append("Excel COM application is not available.")
        return result

    if workbook is None:
        result.errors.append("Excel workbook is not open.")
        return result

    try:
        try:
            connection_count = int(_call_with_retry(lambda: workbook.Connections.Count))
        except Exception:
            connection_count = 0
        result.connection_count = connection_count

        configured_connections: list[str] = []
        for index in range(1, connection_count + 1):
            try:
                connection = _call_with_retry(workbook.Connections.Item, index)
                name = _connection_name(connection, index)
                _set_background_query_false(connection)
                configured_connections.append(name)
            except Exception as exc:
                name = f"Connection_{index}"
                result.warnings.append(
                    f"Could not configure {name} for synchronous refresh: {exc}"
                )

        try:
            pivot_caches = _call_with_retry(lambda: workbook.PivotCaches())
            pivot_cache_count = int(_call_with_retry(lambda: pivot_caches.Count))
        except Exception:
            pivot_caches = None
            pivot_cache_count = 0
        result.pivot_cache_count = pivot_cache_count

        if _timed_out(start_time, result.timeout_seconds):
            result.timeout = True
            result.errors.append("Refresh timed out before RefreshAll started.")
            return result

        logger.info(
            "Triggering workbook RefreshAll for %d connection(s) and %d PivotCache(s).",
            connection_count,
            pivot_cache_count,
        )
        _call_with_retry(workbook.RefreshAll)

        # This handles Power Query and many asynchronous provider requests.
        try:
            _call_with_retry(excel_app.CalculateUntilAsyncQueriesDone)
            result.async_queries_completed = True
        except Exception as exc:
            result.warnings.append(
                f"CalculateUntilAsyncQueriesDone returned an error: {exc}"
            )

        idle, active_connections, active_caches = _wait_for_refresh_idle(
            excel_app=excel_app,
            workbook=workbook,
            start_time=start_time,
            timeout_seconds=result.timeout_seconds,
            poll_interval=poll_interval,
        )

        if not idle:
            result.timeout = True
            details: list[str] = []
            if active_connections:
                details.append("active connections: " + ", ".join(active_connections))
            if active_caches:
                details.append("active PivotCaches: " + ", ".join(active_caches))
            suffix = f" ({'; '.join(details)})" if details else ""
            result.errors.append(
                "Workbook refresh did not become idle before the timeout" + suffix + "."
            )
            return result

        # RefreshAll completed successfully. Count the workbook objects that were
        # part of that completed refresh rather than starting a second refresh.
        result.connections_refreshed = connection_count
        result.pivot_caches_refreshed = pivot_cache_count

        logger.info("Refreshing completed; calculating formulas and CUBE values.")

        try:
            _call_with_retry(excel_app.CalculateFullRebuild)
        except Exception as rebuild_error:
            result.warnings.append(
                f"CalculateFullRebuild failed; using CalculateFull: {rebuild_error}"
            )
            _call_with_retry(excel_app.CalculateFull)

        if not _wait_for_excel_idle(
            excel_app,
            start_time,
            result.timeout_seconds,
            poll_interval,
        ):
            result.timeout = True
            result.errors.append(
                "Excel calculation did not complete before the timeout."
            )
            return result

        # A final refresh-idle check catches OLAP work triggered by recalculation.
        final_idle, active_connections, active_caches = _wait_for_refresh_idle(
            excel_app=excel_app,
            workbook=workbook,
            start_time=start_time,
            timeout_seconds=result.timeout_seconds,
            poll_interval=poll_interval,
            stable_polls_required=2,
        )
        if not final_idle:
            result.timeout = True
            result.errors.append(
                "Excel started another background refresh during calculation."
            )
            return result

        result.calculation_completed = True
        result.refresh_completed = True

        logger.info(
            "Workbook refresh completed and Excel is idle. "
            "connections=%d pivot_caches=%d duration=%.2fs",
            connection_count,
            pivot_cache_count,
            time.monotonic() - start_time,
        )

    except Exception as exc:
        logger.exception("Workbook refresh failed.")
        result.errors.append(str(exc))
        result.refresh_completed = False

    finally:
        result.duration_seconds = round(time.monotonic() - start_time, 2)

    return result


__all__ = ["refresh_and_calculate_workbook"]

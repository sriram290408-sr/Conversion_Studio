"""Central COM retry helper for Power BI-to-Excel converter.

All Excel COM operations that may encounter transient RPC / Excel-busy errors
must use ``com_call`` rather than inlining retry logic.  This keeps every
module consistent and makes the retry behaviour testable in isolation.

Supported transient error codes
--------------------------------
0x800AC472   -2146777998   xlcall: Excel busy  (most common)
0x80010001   -2147418111   RPC_E_CALL_REJECTED
0x8001010A   -2147417846   RPC_E_SERVERCALL_RETRYLATER
0x80010105   -2147417851   RPC_E_SERVERFAULT (transient variant)
0x80080005   -2146959355   CO_E_SERVER_EXEC_FAILURE (COM server starting up)
"""

from __future__ import annotations

import logging
import time
from typing import Any, Callable, FrozenSet, Optional, Tuple

logger = logging.getLogger("com_retry")

# ---------------------------------------------------------------------------
# Known transient COM busy codes (decimal string representations that appear
# in win32com exception messages).
# ---------------------------------------------------------------------------
COM_BUSY_CODES: FrozenSet[str] = frozenset(
    {
        # Excel busy
        "-2146777998",
        "800ac472",
        "0x800ac472",
        # RPC call rejected
        "-2147418111",
        "80010001",
        "0x80010001",
        # RPC server call retry later
        "-2147417846",
        "8001010a",
        "0x8001010a",
        # RPC server fault (transient)
        "-2147417851",
        "80010105",
        "0x80010105",
        # COM server exec failure
        "-2146959355",
        "80080005",
        "0x80080005",
    }
)

_DEFAULT_MAX_RETRIES = 30
_DEFAULT_DELAY_SECONDS = 0.5
_DEFAULT_MAX_DELAY_SECONDS = 4.0


def _is_com_busy(exc: Exception) -> bool:
    """Return True if *exc* is a transient COM-busy error worth retrying."""
    msg = str(exc).lower()
    return any(code in msg for code in COM_BUSY_CODES)


def com_call(
    func: Callable[..., Any],
    *args: Any,
    max_retries: int = _DEFAULT_MAX_RETRIES,
    delay: float = _DEFAULT_DELAY_SECONDS,
    max_delay: float = _DEFAULT_MAX_DELAY_SECONDS,
    label: Optional[str] = None,
    **kwargs: Any,
) -> Any:
    """Call *func* with *args* / *kwargs*, retrying on transient COM errors.

    Parameters
    ----------
    func:
        Any callable (lambda, method, function).  Must be called on the
        COM-owning thread.
    max_retries:
        Maximum number of retry attempts (not counting the first call).
    delay:
        Initial sleep duration in seconds between attempts.  Doubles on each
        retry, capped at *max_delay*.
    max_delay:
        Maximum sleep between retries.
    label:
        Optional human-readable name for the operation (used in log messages).
    *args, **kwargs:
        Passed through to *func*.

    Returns
    -------
    Any
        Whatever *func* returns.

    Raises
    ------
    Exception
        Re-raises the last exception if all retries are exhausted, or
        immediately re-raises any non-COM-busy exception.
    """
    name = label or getattr(func, "__name__", repr(func))
    current_delay = delay

    for attempt in range(max_retries + 1):
        try:
            return func(*args, **kwargs)
        except Exception as exc:
            if not _is_com_busy(exc):
                # Non-transient error — fail immediately.
                raise

            if attempt >= max_retries:
                logger.error(
                    "COM call '%s' failed after %d retries: %s",
                    name,
                    max_retries,
                    exc,
                )
                raise

            logger.debug(
                "COM busy on '%s' (attempt %d/%d) — retrying in %.1fs: %s",
                name,
                attempt + 1,
                max_retries,
                current_delay,
                exc,
            )
            time.sleep(current_delay)
            current_delay = min(current_delay * 2, max_delay)

    # Should never be reached.
    raise RuntimeError(f"com_call({name}): exhausted retries without raising")


def com_getattr(obj: Any, attr: str, **retry_kwargs: Any) -> Any:
    """Read a COM property with retry on busy.

    Equivalent to ``com_call(lambda: getattr(obj, attr), ...)`` but with a
    cleaner log label.
    """
    return com_call(lambda: getattr(obj, attr), label=f"getattr({attr})", **retry_kwargs)


def com_setattr(obj: Any, attr: str, value: Any, **retry_kwargs: Any) -> None:
    """Set a COM property with retry on busy."""
    com_call(lambda: setattr(obj, attr, value), label=f"setattr({attr})", **retry_kwargs)


def com_method(obj: Any, method: str, *args: Any, **retry_kwargs: Any) -> Any:
    """Call ``obj.method(*args)`` with retry on busy.

    Note: keyword arguments for the COM method are not currently supported
    via this helper — call ``com_call(lambda: obj.method(...), ...)`` directly
    if you need them.
    """
    return com_call(getattr(obj, method), *args, label=method, **retry_kwargs)


__all__ = [
    "COM_BUSY_CODES",
    "com_call",
    "com_getattr",
    "com_setattr",
    "com_method",
]

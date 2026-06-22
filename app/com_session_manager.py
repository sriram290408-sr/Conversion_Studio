"""COM Session Manager — lifecycle tracking for interactive Excel sessions.

Each live-connect session owns:
- A UUID session identifier
- A dedicated OS thread (``com_thread``) that owns ALL COM calls for this session
- A ``work_queue`` for dispatching work items to that thread
- Pointers to the Excel app and workbook (only touched on com_thread)
- JSON-serialisable ``result`` dict updated by the thread

Thread safety
-------------
``COMSessionManager`` is a module-level singleton.  The ``sessions`` dict is
protected by ``_lock``.  Individual session state fields are only mutated from
the session's own ``com_thread`` (except for ``state`` / ``result`` which are
also read by HTTP handlers — those reads are safe because CPython's GIL makes
dict reads atomic for simple types).
"""

from __future__ import annotations

import logging
import os
import queue
import threading
import time
import uuid
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Literal, Optional

logger = logging.getLogger("com_session_manager")

# ---------------------------------------------------------------------------
# Session state literals (superset of LiveConversionResult statuses)
# ---------------------------------------------------------------------------
SessionState = Literal[
    "created",
    "excel_launching",
    "waiting_for_user_connection",
    "detecting_connection",
    "connection_not_detected",
    "connection_detected",
    "validating_semantic_model",
    "semantic_model_mismatch",
    "building",
    "refreshing",
    "saving",
    "verifying",
    "completed_live",
    "live_conversion_failed",
    "cancelled",
    "error",
]

# States where the COM thread is still running / accepting work
_ACTIVE_STATES = frozenset(
    {
        "created",
        "excel_launching",
        "waiting_for_user_connection",
        "detecting_connection",
        "connection_detected",
        "validating_semantic_model",
        "building",
        "refreshing",
        "saving",
        "verifying",
        # Retriable — Excel still open, user can fix and retry
        "connection_not_detected",
        "semantic_model_mismatch",
    }
)

# Terminal states — COM thread has exited or will exit soon
_TERMINAL_STATES = frozenset(
    {
        "completed_live",
        "live_conversion_failed",
        "cancelled",
        "error",
    }
)


# ---------------------------------------------------------------------------
# Work-item sent through work_queue
# ---------------------------------------------------------------------------
@dataclass
class _WorkItem:
    fn: Callable[[], Any]
    result_holder: Optional[List[Any]] = None  # [result] or [None, exc]
    event: Optional[threading.Event] = None  # set when fn completes


_SENTINEL = object()  # signals the COM thread to exit cleanly


# ---------------------------------------------------------------------------
# COMSession
# ---------------------------------------------------------------------------
@dataclass
class COMSession:
    """All per-session state.  Fields prefixed with ``_com_`` are only
    accessed from ``com_thread``."""

    session_id: str
    output_path: str  # where the final workbook will be written
    metadata: Dict[str, Any]  # final_chunks from the /upload step
    visual_bindings: List[Any]  # list of VisualBinding dicts
    formula_chunks: List[Any]  # CUBEVALUE formula metadata
    upload_session_id: str = ""  # upload session that spawned this live session

    state: SessionState = "created"
    result: Dict[str, Any] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)

    current_stage: str = "initializing_com"
    error_stage: Optional[str] = None
    error_type: Optional[str] = None
    error_message: Optional[str] = None
    error_trace_id: Optional[str] = None
    progress: int = 0
    message: str = ""

    created_at: float = field(default_factory=time.monotonic)
    last_activity: float = field(default_factory=time.monotonic)

    # COM objects — only valid on com_thread
    _com_excel: Any = field(default=None, repr=False)
    _com_workbook: Any = field(default=None, repr=False)
    _com_workbook_path: str = field(default="", repr=False)

    # Thread infrastructure
    com_thread: Optional[threading.Thread] = field(default=None, repr=False)
    work_queue: queue.Queue = field(default_factory=queue.Queue, repr=False)
    lock: threading.Lock = field(default_factory=threading.Lock, repr=False)
    continue_job_enqueued: bool = False

    # Progress counters (updated from com_thread; read from HTTP handlers)
    pivot_tables_created: int = 0
    pivot_tables_reused: int = 0
    pivot_charts_created: int = 0
    slicers_created: int = 0
    cube_formulas_created: int = 0
    cube_field_count: int = 0
    semantic_match_score: float = 0.0
    selected_connection_name: str = ""
    post_save_verification_passed: bool = False
    live_conversion_succeeded: bool = False
    expected_counts: Dict[str, int] = field(default_factory=dict)
    actual_counts: Dict[str, int] = field(default_factory=dict)
    runtime_objects: Dict[str, Any] = field(default_factory=dict, repr=False)

    def update_state(self, new_state: SessionState) -> None:
        """Update session state and last-activity timestamp."""
        logger.info("Session %s: %s -> %s", self.session_id, self.state, new_state)
        self.state = new_state
        self.last_activity = time.monotonic()

    def is_active(self) -> bool:
        return self.state in _ACTIVE_STATES

    def is_terminal(self) -> bool:
        return self.state in _TERMINAL_STATES

    def dispatch(
        self,
        fn: Callable[[], Any],
        wait: bool = False,
        timeout: float = 300.0,
    ) -> Any:
        """Send *fn* to the COM thread's work queue.

        If *wait* is True, blocks until *fn* completes and returns its result
        (or re-raises its exception).  If *wait* is False, returns immediately.
        """
        if self.state in _TERMINAL_STATES:
            raise RuntimeError(
                f"Session {self.session_id} is in terminal state {self.state!r}."
            )

        holder: List[Any] = []
        evt = threading.Event() if wait else None
        self.work_queue.put(_WorkItem(fn=fn, result_holder=holder, event=evt))

        if wait:
            if not evt.wait(timeout=timeout):
                raise TimeoutError(
                    f"COM work item timed out after {timeout}s "
                    f"(session {self.session_id})"
                )
            if len(holder) == 2 and isinstance(holder[1], BaseException):
                raise holder[1]
            return holder[0] if holder else None
        return None

    def to_status_dict(self) -> Dict[str, Any]:
        """Return a JSON-serialisable status snapshot."""
        return {
            "session_id": self.session_id,
            "state": self.state,
            "current_stage": self.current_stage,
            "error_stage": self.error_stage,
            "error_type": self.error_type,
            "error_message": self.error_message,
            "error_trace_id": self.error_trace_id,
            "progress": self.progress,
            "message": self.message,
            "output_available": self.state == "completed_live",
            "download_url": (
                f"/api/live-connect/{self.session_id}/download"
                if self.state == "completed_live"
                else None
            ),
            "pivot_tables_created": self.pivot_tables_created,
            "pivot_tables_reused": self.pivot_tables_reused,
            "pivot_charts_created": self.pivot_charts_created,
            "slicers_created": self.slicers_created,
            "cube_formulas_created": self.cube_formulas_created,
            "cube_field_count": self.cube_field_count,
            "semantic_match_score": self.semantic_match_score,
            "selected_connection_name": self.selected_connection_name,
            "post_save_verification_passed": self.post_save_verification_passed,
            "live_conversion_succeeded": self.live_conversion_succeeded,
            "expected_counts": dict(self.expected_counts),
            "actual_counts": dict(self.actual_counts),
            "warnings": list(self.warnings),
            "errors": list(self.errors),
            "last_activity_age_seconds": round(
                time.monotonic() - self.last_activity, 1
            ),
        }


# ---------------------------------------------------------------------------
# COM thread runner
# ---------------------------------------------------------------------------
def _com_thread_main(session: COMSession) -> None:
    """Entry point for a session's dedicated COM thread.

    Drains ``session.work_queue`` until the sentinel is received or a
    terminal error is encountered.  ``pythoncom`` is initialized/uninitialized
    here so all COM calls happen on this thread.
    """
    pythoncom = None
    try:
        import pythoncom as _pythoncom  # type: ignore[import]

        pythoncom = _pythoncom
        pythoncom.CoInitialize()
    except ImportError as exc:
        if os.name == "nt":
            logger.error("Session %s: CoInitialize failed: %s", session.session_id, exc)
            session.errors.append(f"CoInitialize failed: {exc}")
            session.update_state("error")
            return
        logger.debug(
            "pythoncom unavailable on non-Windows test host; using no-op COM apartment."
        )
    except Exception as exc:
        logger.error("Session %s: CoInitialize failed: %s", session.session_id, exc)
        session.errors.append(f"CoInitialize failed: {exc}")
        session.update_state("error")
        return

    logger.info("Session %s: COM thread started.", session.session_id)

    try:
        while True:
            try:
                item: Any = session.work_queue.get(timeout=1.0)
            except queue.Empty:
                # Keep looping so we can pick up new work items.
                continue

            if item is _SENTINEL:
                logger.info(
                    "Session %s: COM thread received sentinel — exiting.",
                    session.session_id,
                )
                break

            assert isinstance(item, _WorkItem)
            try:
                result = item.fn()
                if item.result_holder is not None:
                    item.result_holder.append(result)
            except Exception as exc:
                logger.exception(
                    "Session %s: COM work item raised: %s", session.session_id, exc
                )
                if item.result_holder is not None:
                    item.result_holder.extend([None, exc])
            finally:
                if item.event is not None:
                    item.event.set()
                session.work_queue.task_done()

            if session.is_terminal():
                # Drain remaining items quickly (they should not be queued
                # after a terminal state, but be safe).
                break
    finally:
        try:
            if pythoncom is not None:
                pythoncom.CoUninitialize()
        except Exception:
            pass
        logger.info("Session %s: COM thread exiting.", session.session_id)


# ---------------------------------------------------------------------------
# COMSessionManager singleton
# ---------------------------------------------------------------------------
class COMSessionManager:
    """Module-level singleton that owns all active COM sessions."""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._sessions: Dict[str, COMSession] = {}
        # upload_session_id -> live_session_id (for duplicate prevention)
        self._upload_to_live: Dict[str, str] = {}

    def create_session(
        self,
        *,
        metadata: Dict[str, Any],
        visual_bindings: List[Any],
        formula_chunks: List[Any],
        output_path: str,
        session_id: Optional[str] = None,
        upload_session_id: str = "",
    ) -> COMSession:
        """Create a new COM session and start its worker thread."""
        sid = session_id or str(uuid.uuid4())
        session = COMSession(
            session_id=sid,
            output_path=output_path,
            metadata=metadata,
            visual_bindings=visual_bindings,
            formula_chunks=formula_chunks,
            upload_session_id=upload_session_id,
        )

        thread = threading.Thread(
            target=_com_thread_main,
            args=(session,),
            name=f"COM-{sid[:8]}",
            daemon=True,
        )
        session.com_thread = thread

        with self._lock:
            self._sessions[sid] = session
            if upload_session_id:
                self._upload_to_live[upload_session_id] = sid

        thread.start()
        logger.info("Session %s: created and COM thread started.", sid)
        return session

    def get_active_session_by_upload_id(
        self, upload_session_id: str
    ) -> Optional[COMSession]:
        """Return the active live session for the given upload session, or None."""
        with self._lock:
            live_id = self._upload_to_live.get(upload_session_id)
            if not live_id:
                return None
            session = self._sessions.get(live_id)
            if session is None:
                return None
            # Only return non-terminal sessions (terminal = Excel closed)
            if session.is_terminal():
                return None
            return session

    def get_session(self, session_id: str) -> Optional[COMSession]:
        with self._lock:
            return self._sessions.get(session_id)

    def cancel_session(self, session_id: str) -> bool:
        """Signal the session to cancel.  Returns True if found."""
        session = self.get_session(session_id)
        if session is None:
            return False
        if session.is_terminal():
            return True  # already done

        # Import here to avoid circular imports at module load time.
        try:
            from .interactive_excel_connection import cancel_session_com
        except ImportError:
            try:
                from app.interactive_excel_connection import cancel_session_com
            except ImportError:
                cancel_session_com = None

        def _cancel() -> None:
            if cancel_session_com is not None:
                try:
                    cancel_session_com(session)
                except Exception as exc:
                    logger.error("Cancel COM cleanup error: %s", exc)
            session.update_state("cancelled")

        # Non-blocking dispatch — the thread will process it when ready.
        try:
            session.work_queue.put(_WorkItem(fn=_cancel))
            session.work_queue.put(_SENTINEL)
        except Exception as exc:
            logger.error("Could not enqueue cancel for session %s: %s", session_id, exc)
        return True

    def cleanup_expired(self, max_age_hours: float = 2.0) -> int:
        """Remove terminal sessions older than *max_age_hours*. Returns count removed."""
        cutoff = time.monotonic() - max_age_hours * 3600
        removed = 0
        with self._lock:
            expired = [
                sid
                for sid, s in self._sessions.items()
                if s.is_terminal() and s.last_activity < cutoff
            ]
            for sid in expired:
                session = self._sessions.pop(sid)
                if session.upload_session_id:
                    self._upload_to_live.pop(session.upload_session_id, None)
                removed += 1
        if removed:
            logger.info("Cleaned up %d expired COM session(s).", removed)
        return removed

    def list_sessions(self) -> List[Dict[str, Any]]:
        with self._lock:
            return [s.to_status_dict() for s in self._sessions.values()]


# Module-level singleton instance
session_manager = COMSessionManager()

__all__ = [
    "COMSession",
    "COMSessionManager",
    "SessionState",
    "session_manager",
    "_ACTIVE_STATES",
    "_TERMINAL_STATES",
    "_SENTINEL",
]

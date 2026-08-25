"""Opt-in operational and diagnostic logging for the local application."""

from __future__ import annotations

import json
import logging
import sys
import threading
from functools import wraps
from logging.handlers import RotatingFileHandler
from pathlib import Path
from time import perf_counter


LOGGER = logging.getLogger("inventory_workbench")


def configure_logging(enabled, log_path, ensure_directories):
    """Enable or disable the rotating diagnostic log chosen in Settings."""
    for handler in LOGGER.handlers[:]:
        LOGGER.removeHandler(handler)
        handler.close()
    LOGGER.propagate = False
    LOGGER.setLevel(logging.INFO)
    if not enabled:
        return
    ensure_directories()
    handler = RotatingFileHandler(log_path, maxBytes=2 * 1024 * 1024, backupCount=3, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    LOGGER.addHandler(handler)


def logging_enabled():
    return bool(LOGGER.handlers)


def _context_text(context):
    return json.dumps(context, ensure_ascii=False, default=str, sort_keys=True, separators=(",", ":"))


def log_event(message, *args):
    if logging_enabled():
        LOGGER.info(message, *args)


def log_operation(action, status, **context):
    if logging_enabled():
        LOGGER.info("operation action=%s status=%s context=%s", action, status, _context_text(context))


def log_exception(action, exc, **context):
    if logging_enabled():
        LOGGER.exception("operation action=%s status=failed error=%r context=%s", action, exc, _context_text(context))


def audited(action):
    """Record calls to user-facing service methods without logging file contents."""
    def decorate(function):
        @wraps(function)
        def wrapped(*args, **kwargs):
            started = perf_counter()
            log_operation(action, "started")
            try:
                result = function(*args, **kwargs)
            except Exception as exc:
                log_exception(action, exc, duration_ms=round((perf_counter() - started) * 1000))
                raise
            log_operation(action, "completed", duration_ms=round((perf_counter() - started) * 1000))
            return result
        return wrapped
    return decorate


def install_exception_hooks():
    """Persist otherwise unhandled main-thread and worker-thread failures."""
    original_hook = sys.excepthook
    original_thread_hook = threading.excepthook

    def handle_exception(exc_type, exc_value, traceback):
        if exc_type is not KeyboardInterrupt and logging_enabled():
            LOGGER.critical("unhandled process exception", exc_info=(exc_type, exc_value, traceback))
        original_hook(exc_type, exc_value, traceback)

    def handle_thread_exception(args):
        if args.exc_type is not KeyboardInterrupt and logging_enabled():
            LOGGER.critical(
                "unhandled thread exception thread=%s",
                args.thread.name if args.thread else "unknown",
                exc_info=(args.exc_type, args.exc_value, args.exc_traceback),
            )
        original_thread_hook(args)

    sys.excepthook = handle_exception
    threading.excepthook = handle_thread_exception

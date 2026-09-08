"""Retry helpers for flaky Excel COM calls.

Excel is a single-threaded apartment server. If it is mid-recalculation, has a
modal dialog up, or the user is editing a cell, incoming automation calls fail
immediately with "Call was rejected by callee" instead of queueing. The only
correct response is to back off and try again.
"""

from __future__ import annotations

import functools
import time
from typing import Any, Callable, TypeVar

from xlvba.core.errors import com_description, hresult_of, is_transient

T = TypeVar("T")

#: Default retry budget. 6 attempts with 0.25s base and 1.6x growth gives
#: roughly 0.25+0.4+0.64+1.02+1.64 = ~4s of total waiting, which comfortably
#: covers a user releasing a cell edit or a recalculation finishing.
DEFAULT_ATTEMPTS = 6
DEFAULT_BASE_DELAY = 0.25
DEFAULT_GROWTH = 1.6
DEFAULT_MAX_DELAY = 2.0


def com_retry(
    attempts: int = DEFAULT_ATTEMPTS,
    base_delay: float = DEFAULT_BASE_DELAY,
    growth: float = DEFAULT_GROWTH,
    max_delay: float = DEFAULT_MAX_DELAY,
) -> Callable[[Callable[..., T]], Callable[..., T]]:
    """Decorator retrying a function while Excel reports a transient COM error.

    Non-transient COM errors propagate immediately - retrying an access-denied
    or a bad-argument error just wastes time.
    """

    def decorator(func: Callable[..., T]) -> Callable[..., T]:
        @functools.wraps(func)
        def wrapper(*args: Any, **kwargs: Any) -> T:
            delay = base_delay
            last: BaseException | None = None
            for attempt in range(1, attempts + 1):
                try:
                    return func(*args, **kwargs)
                except Exception as exc:  # noqa: BLE001 - re-raised below
                    if not is_transient(exc):
                        raise
                    last = exc
                    if attempt == attempts:
                        break
                    time.sleep(delay)
                    delay = min(delay * growth, max_delay)
            assert last is not None
            raise RuntimeError(
                f"Excel stayed busy after {attempts} attempts "
                f"(hresult={hresult_of(last)}): {com_description(last)}. "
                "Close any open Excel dialog, stop cell editing, then retry."
            ) from last

        return wrapper

    return decorator


def call_with_retry(func: Callable[..., T], *args: Any, **kwargs: Any) -> T:
    """Imperative form of :func:`com_retry` for one-off calls."""
    return com_retry()(func)(*args, **kwargs)


def wait_until(
    predicate: Callable[[], bool],
    timeout: float,
    interval: float = 0.1,
) -> bool:
    """Poll ``predicate`` until it is true or ``timeout`` seconds elapse."""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            if predicate():
                return True
        except Exception:  # noqa: BLE001 - predicate may touch dying COM objects
            pass
        time.sleep(interval)
    try:
        return predicate()
    except Exception:  # noqa: BLE001
        return False

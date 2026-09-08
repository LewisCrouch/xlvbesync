"""Filesystem watcher driving hot-reload.

Design notes
------------
*Debounce*: one Ctrl+S in VS Code produces several inotify/ReadDirectoryChanges
events (truncate, write, sometimes an atomic rename). We coalesce events per
path over a short window before acting.

*Single-flight*: all COM work happens on one dedicated worker thread. Excel is
an STA server; calling it concurrently from several watchdog threads is a
reliable way to get "Call was rejected by callee" or worse. The worker owns its
own COM apartment.

*Idempotence*: content hashes (see :mod:`xlvba.core.state`) mean an editor that
rewrites a file without changing it costs nothing.
"""

from __future__ import annotations

import queue
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from watchdog.events import (
    FileSystemEvent,
    FileSystemEventHandler,
)
from watchdog.observers import Observer

from xlvba.core import console
from xlvba.core.components import FORM_BINARY_EXT, is_source_file
from xlvba.core.config import MODE_SWAP, ProjectConfig, TargetConfig
from xlvba.core.errors import XlvbaError
from xlvba.core.excel import com_apartment
from xlvba.core.state import SyncState
from xlvba.ops.inject import inject_target, remove_source
from xlvba.ops.reload import reload_target
from xlvba.ops.session import open_target

#: Event kinds we care about.
CHANGE = "change"
DELETE = "delete"


@dataclass(frozen=True)
class PendingChange:
    """A coalesced filesystem change."""

    path: Path
    kind: str


class SourceEventHandler(FileSystemEventHandler):
    """Translates watchdog events into :class:`PendingChange` items."""

    def __init__(self, config: ProjectConfig, sink: queue.Queue[PendingChange]) -> None:
        self.config = config
        self.sink = sink

    # -- helpers ----------------------------------------------------------

    def _relevant(self, raw_path: str) -> Path | None:
        path = Path(raw_path)

        # Editors write ``file.bas~``, ``file.bas.tmp``, ``4913`` etc.
        if path.name.startswith("~$") or path.name.endswith("~"):
            return None

        suffix = path.suffix.lower()
        if suffix == FORM_BINARY_EXT:
            # A .frx change must resync its .frm sibling.
            frm = path.with_suffix(".frm")
            return frm if frm.is_file() else None

        if not is_source_file(path):
            return None

        if self.config.target_for_source(path) is None:
            return None

        return path

    def _push(self, raw_path: str, kind: str) -> None:
        path = self._relevant(raw_path)
        if path is None:
            return
        self.sink.put(PendingChange(path.resolve(), kind))

    # -- watchdog hooks ---------------------------------------------------

    def on_created(self, event: FileSystemEvent) -> None:
        if not event.is_directory:
            self._push(str(event.src_path), CHANGE)

    def on_modified(self, event: FileSystemEvent) -> None:
        if not event.is_directory:
            self._push(str(event.src_path), CHANGE)

    def on_moved(self, event: FileSystemEvent) -> None:
        if event.is_directory:
            return
        # An atomic save arrives as a move onto the real filename.
        self._push(str(event.src_path), DELETE)
        dest = getattr(event, "dest_path", None)
        if dest:
            self._push(str(dest), CHANGE)

    def on_deleted(self, event: FileSystemEvent) -> None:
        if not event.is_directory:
            self._push(str(event.src_path), DELETE)


class SyncWorker:
    """Serialises all Excel work onto a single thread."""

    def __init__(
        self,
        config: ProjectConfig,
        state: SyncState,
        *,
        mode: str | None = None,
        verbose: bool = False,
    ) -> None:
        self.config = config
        self.state = state
        self.mode = mode
        self.verbose = verbose
        self.inbox: queue.Queue[PendingChange] = queue.Queue()
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None

    # -- lifecycle --------------------------------------------------------

    def start(self) -> None:
        self._thread = threading.Thread(target=self._run, name="xlvba-sync", daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop.set()
        if self._thread is not None:
            self._thread.join(timeout=5.0)

    # -- main loop --------------------------------------------------------

    def _run(self) -> None:
        # The worker thread needs its own COM apartment.
        with com_apartment():
            while not self._stop.is_set():
                batch = self._collect_batch()
                if batch:
                    self._process(batch)

    def _collect_batch(self) -> dict[Path, str]:
        """Block for one event, then drain the debounce window."""
        window = max(self.config.debounce_ms, 0) / 1000.0
        try:
            first = self.inbox.get(timeout=0.25)
        except queue.Empty:
            return {}

        batch: dict[Path, str] = {first.path: first.kind}
        deadline = time.monotonic() + window

        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                break
            try:
                change = self.inbox.get(timeout=remaining)
            except queue.Empty:
                break
            # A later CHANGE supersedes an earlier DELETE (atomic save) and
            # vice versa; last write wins.
            batch[change.path] = change.kind
            # Keep extending slightly so a burst of saves lands in one batch.
            deadline = max(deadline, time.monotonic() + window / 2)

        return batch

    def _process(self, batch: dict[Path, str]) -> None:
        by_target: dict[str, tuple[TargetConfig, list[Path], list[Path]]] = {}

        for path, kind in batch.items():
            target = self.config.target_for_source(path)
            if target is None:
                continue
            key = target.resolved_name()
            entry = by_target.setdefault(key, (target, [], []))
            # Re-check existence: a rename can leave a stale DELETE queued.
            if kind == DELETE and not path.exists():
                entry[2].append(path)
            elif path.exists():
                entry[1].append(path)

        for key, (target, changed, deleted) in by_target.items():
            try:
                self._sync_target(target, changed, deleted)
            except XlvbaError as exc:
                console.error(f"{key}: {exc}")
            except Exception as exc:  # noqa: BLE001 - watcher must survive
                console.error(f"{key}: {console.describe_exception(exc)}")

        self.state.save()

    def _sync_target(
        self,
        target: TargetConfig,
        changed: list[Path],
        deleted: list[Path],
    ) -> None:
        """Apply one target's pending changes."""
        names = ", ".join(sorted(p.name for p in changed + deleted))
        console.info(f"syncing {target.resolved_name()}: {names}")

        effective_mode = self.mode or target.mode

        # Swap mode rebuilds the whole file, so per-file granularity is moot.
        if effective_mode == MODE_SWAP:
            report = reload_target(
                self.config, target, mode=MODE_SWAP, state=self.state
            )
            from xlvba.ops.reload import print_reload

            print_reload(report, verbose=self.verbose)
            return

        with open_target(self.config, target, quiet=True) as session:
            if deleted and self.config.sync_deletions:
                for path in deleted:
                    result = remove_source(session, path, state=self.state)
                    console.detail(str(result))
            elif deleted:
                for path in deleted:
                    console.detail(
                        f"ignored deletion of {path.name} "
                        "(set sync_deletions=true to mirror deletions)"
                    )

            if changed:
                report = inject_target(session, state=self.state, only=changed)
                from xlvba.ops.inject import print_report

                print_report(report, verbose=self.verbose)


def watch(
    config: ProjectConfig,
    *,
    targets: Iterable[TargetConfig] | None = None,
    mode: str | None = None,
    verbose: bool = False,
    initial_sync: bool = False,
) -> int:
    """Watch the source tree and hot-reload on change. Blocks until Ctrl+C."""
    selected = list(targets) if targets is not None else list(config.targets)
    if not selected:
        console.error("No targets to watch.")
        return 2

    state = SyncState.load(config.root)

    src_root = config.src_root
    src_root.mkdir(parents=True, exist_ok=True)

    worker = SyncWorker(config, state, mode=mode, verbose=verbose)
    handler = SourceEventHandler(config, worker.inbox)

    observer = Observer()
    observer.schedule(handler, str(src_root), recursive=True)

    console.info(f"watching {src_root}")
    for target in selected:
        console.detail(
            f"{target.resolved_name()} -> {config.target_path(target).name} "
            f"[{mode or target.mode}]"
        )

    worker.start()
    observer.start()

    if initial_sync:
        for target in selected:
            for path in sorted(config.target_dir(target).rglob("*")):
                if path.is_file() and is_source_file(path):
                    worker.inbox.put(PendingChange(path.resolve(), CHANGE))

    console.info("ready (press Ctrl+C to stop)")

    try:
        while True:
            time.sleep(0.5)
    except KeyboardInterrupt:
        console.info("\nstopping")
    finally:
        observer.stop()
        observer.join(timeout=5.0)
        worker.stop()
        state.save()

    return 0

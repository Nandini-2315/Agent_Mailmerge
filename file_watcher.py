import os
import threading
import time
from pathlib import Path
from typing import Callable, Optional
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler, FileSystemEvent

from .logger import logger
from .excel_manager import load_config


class ExcelFileChangeHandler(FileSystemEventHandler):
    """
    Watchdog event handler with debouncing for Excel file modifications.
    Filters out temporary files created during Excel save operations (~$...).
    """
    def __init__(self, target_file: Path, debounce_seconds: float, on_change: Callable[[], None]):
        super().__init__()
        self.target_file = target_file.resolve()
        self.debounce_seconds = debounce_seconds
        self.on_change = on_change
        self._timer: Optional[threading.Timer] = None
        self._lock = threading.Lock()

    def _trigger_debounced_change(self):
        with self._lock:
            if self._timer is not None:
                self._timer.cancel()
            self._timer = threading.Timer(self.debounce_seconds, self._execute_callback)
            self._timer.daemon = True
            self._timer.start()

    def _execute_callback(self):
        logger.info(f"Excel save stabilized for {self.target_file.name}. Triggering sync...")
        try:
            self.on_change()
        except Exception as e:
            logger.error(f"Error during monitored change processing: {e}", exc_info=True)

    def _should_handle(self, event: FileSystemEvent) -> bool:
        if event.is_directory:
            return False
        
        event_path = Path(event.src_path).resolve()
        # Ignore Excel temporary lock files (e.g. ~$studentsdata.xlsx)
        if event_path.name.startswith("~$") or event_path.name.startswith(".~lock"):
            return False
            
        return event_path == self.target_file

    def on_modified(self, event: FileSystemEvent):
        if self._should_handle(event):
            logger.info(f"File modification detected on {self.target_file.name}. Debouncing for {self.debounce_seconds}s...")
            self._trigger_debounced_change()

    def on_created(self, event: FileSystemEvent):
        if self._should_handle(event):
            logger.info(f"File create/replace detected on {self.target_file.name}. Debouncing for {self.debounce_seconds}s...")
            self._trigger_debounced_change()

    def on_moved(self, event: FileSystemEvent):
        dest_path = getattr(event, "dest_path", None)
        if dest_path and Path(dest_path).resolve() == self.target_file:
            logger.info(f"Atomic file move detected for {self.target_file.name}. Debouncing for {self.debounce_seconds}s...")
            self._trigger_debounced_change()


class ExcelMonitor:
    """
    Monitors the Excel student data file continuously for changes.
    """
    def __init__(
        self,
        excel_path: Optional[Path] = None,
        debounce_seconds: Optional[float] = None,
        on_change_callback: Optional[Callable[[], None]] = None
    ):
        cfg = load_config()
        self.excel_path = Path(excel_path or cfg["paths"]["excel_file"]).resolve()
        self.debounce_seconds = debounce_seconds or cfg["watcher"].get("debounce_seconds", 2.0)
        self.on_change_callback = on_change_callback
        self.observer = Observer()

    def start(self):
        if not self.excel_path.parent.exists():
            raise FileNotFoundError(f"Directory to monitor does not exist: {self.excel_path.parent}")

        if not self.on_change_callback:
            raise ValueError("No on_change_callback provided to ExcelMonitor.")

        handler = ExcelFileChangeHandler(
            target_file=self.excel_path,
            debounce_seconds=self.debounce_seconds,
            on_change=self.on_change_callback
        )

        watch_dir = str(self.excel_path.parent)
        self.observer.schedule(handler, path=watch_dir, recursive=False)
        self.observer.start()
        logger.info(f"Started monitoring '{self.excel_path.name}' in '{watch_dir}' (debounce={self.debounce_seconds}s).")

    def stop(self):
        if self.observer.is_alive():
            self.observer.stop()
            self.observer.join()
            logger.info("Excel file monitor stopped.")

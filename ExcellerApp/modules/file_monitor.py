"""
File Monitor - Watches for file changes
Monitors Excel files for modifications
"""

import os
from pathlib import Path
from typing import Optional

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler, FileModifiedEvent, FileCreatedEvent


class ExcelFileHandler(FileSystemEventHandler):
    """Handler for Excel file events"""
    
    def __init__(self, callback):
        self.callback = callback
        self.supported_extensions = ['.xlsx', '.xls', '.csv']
    
    def on_modified(self, event):
        """Handle file modification"""
        if not event.is_directory:
            self._handle_event(event.src_path)
    
    def on_created(self, event):
        """Handle file creation"""
        if not event.is_directory:
            self._handle_event(event.src_path)
    
    def _handle_event(self, file_path: str):
        """Process file event"""
        path = Path(file_path)
        if path.suffix.lower() in self.supported_extensions:
            # Small delay to ensure file write is complete
            import time
            time.sleep(0.5)
            self.callback(file_path)


class FileMonitor:
    """Monitors files for changes"""
    
    def __init__(self):
        self.observer = Observer()
        self.watched_paths = {}
        self.file_changed = None  # Signal to emit
        
    def watch_file(self, file_path: str):
        """Start watching a file"""
        path = Path(file_path)
        
        if not path.exists():
            return False
        
        # Watch the parent directory
        watch_path = str(path.parent)
        
        if watch_path not in self.watched_paths:
            handler = ExcelFileHandler(self._on_file_changed)
            self.observer.schedule(handler, watch_path, recursive=False)
            self.watched_paths[watch_path] = []
        
        if file_path not in self.watched_paths[watch_path]:
            self.watched_paths[watch_path].append(file_path)
        
        # Start observer if not running
        if not self.observer.is_alive():
            self.observer.start()
        
        return True
    
    def stop_watching(self, file_path: str = None):
        """Stop watching a file or all files"""
        if file_path:
            path = Path(file_path)
            watch_path = str(path.parent)
            
            if watch_path in self.watched_paths:
                if file_path in self.watched_paths[watch_path]:
                    self.watched_paths[watch_path].remove(file_path)
                
                if not self.watched_paths[watch_path]:
                    del self.watched_paths[watch_path]
        else:
            # Stop all
            self.watched_paths.clear()
            if self.observer.is_alive():
                self.observer.stop()
                self.observer.join()
    
    def _on_file_changed(self, file_path: str):
        """Handle file change notification"""
        if self.file_changed and file_path in self._get_all_watched():
            self.file_changed.emit(file_path)
    
    def _get_all_watched(self):
        """Get all watched files"""
        all_files = []
        for files in self.watched_paths.values():
            all_files.extend(files)
        return all_files
    
    def is_watching(self, file_path: str = None) -> bool:
        """Check if watching a file"""
        if file_path:
            return file_path in self._get_all_watched()
        return len(self.watched_paths) > 0
    
    def __del__(self):
        """Cleanup on deletion"""
        self.stop_watching()

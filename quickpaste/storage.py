import json
import os
import shutil
import tempfile
from PyQt5 import QtCore


def save_data_atomic(data, filename):
    """Atomic file write to prevent corruption.

    Writes to a NamedTemporaryFile and then moves it into place.
    """
    temp_file = tempfile.NamedTemporaryFile(mode='w', delete=False, encoding='utf-8')
    try:
        json.dump(data, temp_file, indent=4)
        temp_file.close()
        os.makedirs(os.path.dirname(filename), exist_ok=True)
        shutil.move(temp_file.name, filename)
    except Exception:
        try:
            os.unlink(temp_file.name)
        except (OSError, IOError):
            pass
        raise


class DebouncedSaver:
    """Debounced file saving to prevent excessive writes."""

    def __init__(self, delay_ms: int = 1000):
        self.timer = QtCore.QTimer()
        self.timer.setSingleShot(True)
        self.timer.timeout.connect(self._save)
        self.pending_data = None
        self.target_file = None
        self.delay_ms = delay_ms

    def schedule_save(self, data, target_file):
        self.pending_data = data
        self.target_file = target_file
        # Restart the timer to debounce multiple calls
        self.timer.start(self.delay_ms)

    def _save(self):
        if self.pending_data is not None and self.target_file:
            save_data_atomic(self.pending_data, self.target_file)
            self.pending_data = None
            self.target_file = None


# Global debounced saver instance (kept for backward compatibility)
debounced_saver = DebouncedSaver()


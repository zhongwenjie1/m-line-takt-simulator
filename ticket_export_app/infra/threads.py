# infra/threads.py
from PySide6.QtCore import QObject, Signal, QRunnable, Slot
import traceback

class WorkerSignals(QObject):
    finished = Signal()
    error = Signal(str)
    result = Signal(object)
    progress = Signal(int)

class Worker(QRunnable):
    def __init__(self, fn, *args, **kwargs):
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

    @Slot()
    def run(self):
        try:
            res = self.fn(*self.args, **self.kwargs)
        except Exception:
            tb = traceback.format_exc()
            self.signals.error.emit(tb)
        else:
            self.signals.result.emit(res)
        finally:
            self.signals.finished.emit()
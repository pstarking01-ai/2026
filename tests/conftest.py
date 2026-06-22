import os
import sys
import pytest

# Ensure the repo root is on sys.path so modules can be imported
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

# QApplication singleton shared across all PyQt6 tests
_app = None


@pytest.fixture(scope="session", autouse=True)
def qapp():
    """Create a single QApplication instance for the entire test session."""
    global _app
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    from PyQt6.QtWidgets import QApplication

    _app = QApplication.instance() or QApplication(sys.argv)
    yield _app

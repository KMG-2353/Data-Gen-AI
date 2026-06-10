"""Make ``app`` importable when running pytest from the ``backend/`` directory.

The application code imports as ``app.<module>`` (see ``main.py``), so the
``backend/`` directory must be on ``sys.path``. Adding it here lets the test
suite run with a plain ``pytest`` invocation from ``backend/``.
"""
import os
import sys

BACKEND_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BACKEND_DIR not in sys.path:
    sys.path.insert(0, BACKEND_DIR)

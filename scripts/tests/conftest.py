# -*- coding: utf-8 -*-
"""Optional pytest support: `C:\\Python314\\python.exe -m pytest scripts\\tests`.
The dependency-free runner (run_tests.py) is the canonical gate; this just lets the
same test_* functions be collected by pytest, sharing the build-once fixture."""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, ROOT)
sys.path.insert(0, HERE)

import pytest  # noqa: E402

import harness  # noqa: E402


@pytest.fixture(scope="session")
def ctx():
    return harness.get_ctx()

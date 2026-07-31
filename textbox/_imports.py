import contextlib
import os
import sys


@contextlib.contextmanager
def without_local_module_shadowing(module_file):
    """Temporarily hide this package directory from top-level imports."""
    module_directory = os.path.realpath(os.path.dirname(module_file))
    original_path = list(sys.path)
    sys.path[:] = [
        entry
        for entry in sys.path
        if os.path.realpath(entry or os.curdir) != module_directory
    ]
    try:
        yield
    finally:
        sys.path[:] = original_path

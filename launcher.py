import os
import sys
import types

abs_app_py = os.path.abspath(os.path.join(os.path.dirname(__file__), "app.py"))

# Setup module context so the app runs as __main__
module = types.ModuleType("__main__")
module.__file__ = abs_app_py
sys.modules["__main__"] = module

# Add current directory to path
sys.path.insert(0, os.path.dirname(abs_app_py))

# Load and execute app.py with proper UTF-8 encoding
with open(abs_app_py, "r", encoding="utf-8") as f:
    code = f.read()

exec(compile(code, abs_app_py, "exec"), module.__dict__)

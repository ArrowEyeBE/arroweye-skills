"""Run the plugin's scripts with a Python that has its packages.

Some Python installs (Homebrew, many Linux distributions) refuse `pip install`
("externally-managed-environment"). check_env.py then proposes the plugin's own
virtual environment in ~/.financial-plan-builder/venv (FINANCIAL_PLAN_VENV overrides
the folder). ensure() re-runs the current script with that environment's Python when
this one lacks the packages, so every script can still be started with `python3`.
"""
import importlib.util
import os
import subprocess
import sys
from pathlib import Path

MODULES = (("openpyxl", "openpyxl"), ("docx", "python-docx"), ("matplotlib", "matplotlib"))


def venv_dir():
    return Path(os.environ.get("FINANCIAL_PLAN_VENV") or Path.home() / ".financial-plan-builder" / "venv")


def venv_python():
    return venv_dir() / ("Scripts/python.exe" if os.name == "nt" else "bin/python")


def missing_packages():
    return [pip for mod, pip in MODULES if importlib.util.find_spec(mod) is None]


def ensure():
    if not missing_packages():
        return
    py = venv_python()
    if not py.exists() or Path(sys.prefix).resolve() == venv_dir().resolve():
        return  # nothing to switch to: the import that follows reports what is missing
    args = [str(py), *sys.argv]
    if os.name == "nt":
        sys.exit(subprocess.call(args))
    os.execv(str(py), args)

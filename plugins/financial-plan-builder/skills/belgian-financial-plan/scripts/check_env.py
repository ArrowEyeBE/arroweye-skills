#!/usr/bin/env python3
"""Check that everything the financial-plan builder needs is installed.

Usage: check_env.py
Prints JSON {status: ok|missing, python, missing: [{what, install, claude_may_run}]}. Exit code 1 when
something is missing. `claude_may_run` is true for the Python packages: Claude may run that command after
the user agrees. LibreOffice and Python itself are installed by the user.

When this Python refuses `pip install` (PEP 668, "externally-managed-environment": Homebrew, many Linux
distributions), the packages go into the plugin's own virtual environment (see plugin_env.py); the
scripts switch to it by themselves.
"""
import json
import os
import platform
import shutil
import subprocess
import sys
import sysconfig
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent / "fp-model/scripts"))
from plugin_env import MODULES, missing_packages, venv_dir, venv_python  # noqa: E402


def q(path):
    return f'"{path}"'


def externally_managed():
    if sys.prefix != sys.base_prefix:  # inside a virtual environment pip is allowed
        return False
    return (Path(sysconfig.get_path("stdlib")) / "EXTERNALLY-MANAGED").exists()


def venv_ready():
    py = venv_python()
    if not py.exists():
        return False
    test = "import " + ", ".join(mod for mod, _ in MODULES)
    return subprocess.run([str(py), "-c", test], capture_output=True).returncode == 0


def libreoffice():
    for name in ("soffice", "libreoffice"):
        if shutil.which(name):
            return True
    return any(Path(p).exists() for p in ("/Applications/LibreOffice.app/Contents/MacOS/soffice",
                                          "C:/Program Files/LibreOffice/program/soffice.exe",
                                          "C:/Program Files (x86)/LibreOffice/program/soffice.exe"))


def main():
    if any(a in ("-h", "--help") for a in sys.argv[1:]):
        sys.exit(__doc__.strip())
    missing = []
    python = sys.executable
    if sys.version_info < (3, 9):
        missing.append({"what": f"Python 3.9 or newer (found {platform.python_version()})",
                        "install": "https://www.python.org/downloads/", "claude_may_run": False})
    pkgs = missing_packages()
    if pkgs and venv_ready():
        python = str(venv_python())
    elif pkgs:
        if externally_managed():
            all_pkgs = " ".join(pip for _, pip in MODULES)
            install = (f"{q(sys.executable)} -m venv {q(venv_dir())} && "
                       f"{q(venv_python())} -m pip install {all_pkgs}")
            note = (f"This Python does not allow pip to install packages; they go into the plugin's own "
                    f"environment in {venv_dir()}, which the scripts use by themselves.")
        else:
            install = f"{q(sys.executable)} -m pip install " + " ".join(pkgs)
            note = None
        item = {"what": "Python packages: " + ", ".join(pkgs), "install": install, "claude_may_run": True}
        if note:
            item["note"] = note
        missing.append(item)
    if not libreoffice():
        if platform.system() == "Darwin" and shutil.which("brew"):
            install = "brew install --cask libreoffice"
        else:
            install = "https://www.libreoffice.org/download/"
        missing.append({"what": "LibreOffice (recalculates the Excel model)", "install": install,
                        "claude_may_run": False})
    out = {"status": "missing" if missing else "ok", "python": python, "missing": missing}
    print(json.dumps(out, indent=2))
    sys.exit(1 if missing else 0)


if __name__ == "__main__":
    main()

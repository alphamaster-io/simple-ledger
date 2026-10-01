"""Recalculate an .xlsx with headless LibreOffice and save the result to a new file.

openpyxl writes formulas without cached values. Tests (and screenshot/preview
tooling) need real computed values, so we open the workbook in LibreOffice,
run calculateAll() and store a copy.

Usage:
    python tools/recalc.py input.xlsx output.xlsx [timeout_seconds]
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub RecalculateAndSave()
      ThisComponent.calculateAll()
      ThisComponent.store()
      ThisComponent.close(True)
    End Sub
</script:module>"""


class RecalcError(RuntimeError):
    pass


def _env() -> dict:
    env = os.environ.copy()
    env["SAL_USE_VCLPLUGIN"] = "svp"
    return env


def recalc(src: str | Path, dst: str | Path, timeout: int = 180) -> Path:
    """Copy ``src`` to ``dst`` and recalculate ``dst`` in place with LibreOffice."""
    if shutil.which("soffice") is None:
        raise RecalcError("LibreOffice (soffice) is not installed")
    src, dst = Path(src).resolve(), Path(dst).resolve()
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(src, dst)
    before = dst.stat().st_mtime_ns

    with tempfile.TemporaryDirectory(prefix="lo-profile-") as profile:
        url = Path(profile).as_uri()
        subprocess.run(
            ["soffice", "--headless", "--terminate_after_init", f"-env:UserInstallation={url}"],
            capture_output=True, timeout=timeout, env=_env(), check=False,
        )
        macro_dir = Path(profile) / "user" / "basic" / "Standard"
        if not macro_dir.exists():
            raise RecalcError("LibreOffice did not create a user profile")
        (macro_dir / "Module1.xba").write_text(MACRO)
        cmd = [
            "soffice", "--headless", "--norestore", f"-env:UserInstallation={url}",
            "vnd.sun.star.script:Standard.Module1.RecalculateAndSave?language=Basic&location=application",
            str(dst),
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, env=_env())
        if result.returncode != 0:
            raise RecalcError(f"LibreOffice failed: {result.stderr.strip()}")
    if dst.stat().st_mtime_ns == before:
        raise RecalcError("LibreOffice exited but did not rewrite the file")
    return dst


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(2)
    t = int(sys.argv[3]) if len(sys.argv) > 3 else 180
    print(recalc(sys.argv[1], sys.argv[2], t))

"""LibreOffice (headless) rechnet eine Kopie der Arbeitsmappe durch, um die Formeln zu pruefen und die
Ergebnisse als gespeicherte Werte uebernehmen zu koennen."""
import os
import shutil
import subprocess
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


def recalc(path, timeout=1800):
    env = dict(os.environ, SAL_USE_VCLPLUGIN="svp")
    prof = tempfile.mkdtemp(prefix="lo-prof-")
    try:
        url = Path(prof).as_uri()
        subprocess.run(["soffice", "--headless", "--terminate_after_init", f"-env:UserInstallation={url}"],
                       env=env, capture_output=True, timeout=180)
        (Path(prof) / "user/basic/Standard/Module1.xba").write_text(MACRO)
        before = os.stat(path).st_mtime_ns
        r = subprocess.run(["soffice", "--headless", "--norestore", f"-env:UserInstallation={url}",
                            "vnd.sun.star.script:Standard.Module1.RecalculateAndSave?language=Basic&location=application",
                            str(Path(path).absolute())], env=env, capture_output=True, timeout=timeout)
        if os.stat(path).st_mtime_ns == before:
            raise RuntimeError("LibreOffice hat die Datei nicht neu berechnet: " + r.stderr.decode(errors="replace")[-800:])
    finally:
        shutil.rmtree(prof, ignore_errors=True)


def convert(path, fmt, outdir, timeout=600):
    """Datei mit LibreOffice konvertieren (z. B. fmt='pdf' oder 'ods')."""
    env = dict(os.environ, SAL_USE_VCLPLUGIN="svp")
    prof = tempfile.mkdtemp(prefix="lo-prof-")
    try:
        url = Path(prof).as_uri()
        subprocess.run(["soffice", "--headless", "--norestore", f"-env:UserInstallation={url}", "--convert-to", fmt,
                        "--outdir", str(outdir), str(path)], env=env, capture_output=True, timeout=timeout)
    finally:
        shutil.rmtree(prof, ignore_errors=True)

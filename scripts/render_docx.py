from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
from pathlib import Path

from common import configure_utf8_stdio


def libreoffice_render(source: Path, output: Path, executable: str) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [executable, "--headless", "--convert-to", "pdf", "--outdir", str(output.parent), str(source)],
        check=True,
        capture_output=True,
    )
    generated = output.parent / f"{source.stem}.pdf"
    if generated != output:
        generated.replace(output)


def word_render(source: Path, output: Path) -> bool:
    powershell = shutil.which("pwsh") or shutil.which("powershell")
    if not powershell or os.name != "nt":
        return False
    script = (
        "$ErrorActionPreference='Stop';"
        "$word=New-Object -ComObject Word.Application;"
        "$word.Visible=$false;"
        "try{$doc=$word.Documents.Open($env:ODF_RENDER_INPUT,$false,$true);"
        "$doc.ExportAsFixedFormat($env:ODF_RENDER_OUTPUT,17);$doc.Close($false)}"
        "finally{$word.Quit()}"
    )
    env = os.environ.copy()
    env["ODF_RENDER_INPUT"] = str(source)
    env["ODF_RENDER_OUTPUT"] = str(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    result = subprocess.run([powershell, "-NoProfile", "-Command", script], env=env, capture_output=True)
    return result.returncode == 0 and output.exists()


def render(source: Path, output: Path) -> dict:
    executable = shutil.which("soffice") or shutil.which("libreoffice")
    if executable:
        libreoffice_render(source, output, executable)
        return {"status": "rendered", "renderer": "libreoffice", "output": str(output)}
    if word_render(source, output):
        return {"status": "rendered", "renderer": "microsoft_word", "output": str(output)}
    return {"status": "unavailable", "message": "No supported visual renderer is available; only structural validation may be reported"}


def main() -> None:
    parser = argparse.ArgumentParser(description="Optionally render a DOCX to PDF for visual inspection")
    parser.add_argument("input")
    parser.add_argument("--output")
    args = parser.parse_args()
    source = Path(args.input).expanduser().resolve()
    output = Path(args.output).expanduser().resolve() if args.output else source.with_suffix(".pdf")
    result = render(source, output)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(0 if result["status"] == "rendered" else 3)


if __name__ == "__main__":
    configure_utf8_stdio()
    main()

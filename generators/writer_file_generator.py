"""Generate a LibreOffice Writer document with a safe, manual macro.

LibreOffice document macros are embedded in the document's Basic library.  The
old implementation relied on the current GUI document, raced a background
LibreOffice process, and attempted to access the Basic API through an invalid
UNO object path.  This module uses an isolated headless LibreOffice process and
loads the requested template explicitly.

This generator intentionally accepts only inert/manual Basic sources.  It does
not embed event handlers or permit shell, network, file-changing, or process
launch behavior.  The existing Microsoft Office generator contains an
auto-open PowerShell/network payload and is not a safe parity target.
"""

from __future__ import annotations

import os
import re
import socket
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import yaml
from jinja2 import Template

from utils.xor import encrypt_file, load_xor_key


PROJECT_DIR = Path(__file__).resolve().parent.parent
CONFIG_FILE = PROJECT_DIR / "config.yml"
TEMPLATE_DIR = PROJECT_DIR / "templates"
OUTPUT_DIR = PROJECT_DIR / "output"
LIBREOFFICE_UTILS = PROJECT_DIR / "utils" / "libre_office"
WRITER_MACRO_IMPORT_FILE = LIBREOFFICE_UTILS / "writer_macro_import.py"
LIBREOFFICE_TEMPLATE_EXTENSIONS = {".odt", ".ott", ".sxw"}
LIBREOFFICE_MACRO_EXTENSIONS = {".bas"}


def _load_config() -> dict:
    with CONFIG_FILE.open("r", encoding="utf-8") as stream:
        config = yaml.safe_load(stream) or {}
    try:
        config["app"]["libre_office"]["exe"]
        config["app"]["libre_office"]["python"]
        config["app"]["macros"]
    except (KeyError, TypeError) as exc:
        raise ValueError(
            "config.yml must define app.libre_office.exe, "
            "app.libre_office.python, and app.macros"
        ) from exc
    return config


def _filename_only(value: str, description: str) -> Path:
    path = Path(value)
    if path.name != value:
        raise ValueError(f"{description} must be a filename, not a path")
    return path


def get_template_path(template_name: str) -> Path:
    path = _filename_only(template_name, "Template")
    if path.suffix.lower() not in LIBREOFFICE_TEMPLATE_EXTENSIONS:
        raise ValueError("LibreOffice templates must use .odt, .ott, or .sxw")

    template_path = TEMPLATE_DIR / "files" / "libre_office" / path.name
    if not template_path.is_file():
        raise FileNotFoundError(f"Template not found: {template_path}")
    return template_path


def get_macro_path(macro_name: str) -> Path:
    path = _filename_only(macro_name, "Macro payload")
    if path.suffix.lower() not in LIBREOFFICE_MACRO_EXTENSIONS:
        raise ValueError("LibreOffice macro payloads must use the .bas extension")

    macro_path = TEMPLATE_DIR / "macros" / "libre_office" / path.name
    if not macro_path.is_file():
        raise FileNotFoundError(f"Macro payload not found: {macro_path}")
    return macro_path


def validate_macro_source(source: str) -> None:
    """Apply a conservative guardrail for the manual training path.

    This is deliberately a deny-list in addition to the requirement that the
    file contain a Basic procedure.  It is not presented as a security sandbox;
    it prevents the known shell/network/autorun patterns from this project from
    being copied into LibreOffice documents.
    """

    if not source.strip():
        raise ValueError("LibreOffice macro payload is empty")
    if not re.search(r"(?im)^\s*(sub|function)\s+[A-Za-z_]", source):
        raise ValueError("LibreOffice macro payload must contain a Basic procedure")

    forbidden_patterns = (
        r"\b(createobject|shell|shellexecute|powershell|cmd(?:\.exe)?|curl|wget)\b",
        r"\b(winhttp|xmlhttp|http|https|ftp)\s*[:/]",
        r"\b(autoopen|autoclose|documentopen|documentclose|onload|onopen)\b",
        r"\b(filecopy|kill|mkdir|rmdir|name)\s*\(",
        r"\b(exec|execute|openurl)\s*\(",
    )
    for pattern in forbidden_patterns:
        if re.search(pattern, source, flags=re.IGNORECASE):
            raise ValueError(
                "Unsafe LibreOffice macro rejected: shell, network, "
                "file-changing, process-launch, or autorun behavior detected"
            )


def render_macro_source(macro_path: Path, config: dict) -> str:
    """Render a Basic template and reject unsafe behavior before embedding it."""

    source = macro_path.read_text(encoding="utf-8-sig")
    rendered = Template(source).render(
        agent_url=config["app"]["macros"].get("agent_url", ""),
        reverse_shell_ip=config["app"]["macros"].get("reverse_shell_ip", ""),
    )
    validate_macro_source(rendered)
    return rendered


def _free_tcp_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as listener:
        listener.bind(("127.0.0.1", 0))
        return int(listener.getsockname()[1])


def start_libreoffice_writer(executable: Path, profile_dir: Path, port: int):
    if not executable.is_file():
        raise FileNotFoundError(
            f"LibreOffice executable not found: {executable}. "
            "Update app.libre_office.exe in config.yml."
        )

    command = [
        str(executable),
        "--headless",
        "--nologo",
        "--nodefault",
        "--norestore",
        f"--env:UserInstallation={profile_dir.resolve().as_uri()}",
        f"--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext",
    ]
    return subprocess.Popen(
        command,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def run_macro_import(
    python_executable: Path,
    template_file: Path,
    macro_file: Path,
    output_file: Path,
    port: int,
) -> None:
    if not python_executable.is_file():
        raise FileNotFoundError(
            f"LibreOffice Python runtime not found: {python_executable}. "
            "Update app.libre_office.python in config.yml."
        )

    command = [
        str(python_executable),
        str(WRITER_MACRO_IMPORT_FILE),
        str(template_file),
        str(macro_file),
        str(output_file),
        str(port),
    ]
    result = subprocess.run(
        command,
        check=False,
        capture_output=True,
        text=True,
        timeout=120,
    )
    if result.stdout:
        print(result.stdout, end="")
    if result.returncode:
        detail = (
            result.stderr.strip() or result.stdout.strip() or "no diagnostic output"
        )
        raise RuntimeError(f"LibreOffice macro import failed: {detail}")


def _stop_process(process) -> None:
    if process is None or process.poll() is not None:
        return
    process.terminate()
    try:
        process.wait(timeout=10)
    except subprocess.TimeoutExpired:
        process.kill()
        process.wait(timeout=10)


def main(template_name: str, macro_name: str) -> str:
    config = _load_config()
    template_file = get_template_path(template_name)
    macro_path = get_macro_path(macro_name)
    rendered_source = render_macro_source(macro_path, config)

    output_file = OUTPUT_DIR / f"{template_file.stem}.odt"
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    libreoffice_cfg = config["app"]["libre_office"]
    executable = Path(os.path.expandvars(libreoffice_cfg["exe"]))
    python_executable = Path(os.path.expandvars(libreoffice_cfg["python"]))
    port = _free_tcp_port()

    with tempfile.TemporaryDirectory(prefix="macro-generator-libreoffice-") as runtime:
        runtime_dir = Path(runtime)
        macro_file = runtime_dir / macro_path.name
        macro_file.write_text(rendered_source, encoding="utf-8")

        process = None
        try:
            process = start_libreoffice_writer(
                executable, runtime_dir / "profile", port
            )
            run_macro_import(
                python_executable,
                template_file,
                macro_file,
                output_file,
                port,
            )
        finally:
            _stop_process(process)

    xored_path = encrypt_file(output_file, load_xor_key(PROJECT_DIR))
    print(f"XORed file: {xored_path}")
    return str(output_file)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(
            "Usage: python generators/writer_file_generator.py "
            "<template_name> <macro_name>"
        )
        sys.exit(1)

    try:
        print(f"Generated file: {main(sys.argv[1], sys.argv[2])}")
    except (OSError, RuntimeError, subprocess.SubprocessError, ValueError) as exc:
        print(f"Generation failed: {exc}", file=sys.stderr)
        sys.exit(1)

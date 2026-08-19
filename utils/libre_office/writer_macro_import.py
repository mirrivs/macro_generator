"""Embed an already-validated manual Basic macro into a Writer document.

This script runs with LibreOffice's Python runtime because that runtime provides
the UNO bridge. It intentionally uses syntax compatible with older LibreOffice
Python runtimes. It never executes the macro and never registers a document
event for it.
"""

from __future__ import print_function

import io
import os
import sys
import time

import uno


def _property(name, value):
    property_value = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    property_value.Name = name
    property_value.Value = value
    return property_value


def _connect(port):
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context
    )
    connection = (
        "uno:socket,host=127.0.0.1,port={0};urp;"
        "StarOffice.ComponentContext"
    ).format(port)
    last_error = None
    for _ in range(120):
        try:
            return resolver.resolve(connection)
        except Exception as exc:  # noqa: BLE001 - UNO raises bridge-specific errors
            last_error = exc
            time.sleep(0.25)
    error = RuntimeError("Could not connect to the LibreOffice UNO bridge")
    if last_error is not None:
        raise error
    raise error


def _basic_libraries(document):
    getter = getattr(document, "getBasicLibraries", None)
    if getter is not None:
        return getter()
    return document.BasicLibraries


def insert_macro(document, macro_path, module_name="Module1"):
    """Insert source into document.Standard without setting an autorun event."""

    with io.open(macro_path, "r", encoding="utf-8-sig") as macro_file:
        source = macro_file.read()

    libraries = _basic_libraries(document)
    if not libraries.hasByName("Standard"):
        libraries.createLibrary("Standard")
    elif hasattr(libraries, "isLibraryLoaded") and not libraries.isLibraryLoaded(
        "Standard"
    ):
        libraries.loadLibrary("Standard")

    standard = libraries.getByName("Standard")
    if standard.hasByName(module_name):
        standard.replaceByName(module_name, source)
    else:
        standard.insertByName(module_name, source)


def import_macro(template_file, macro_file, output_file, port):
    context = _connect(port)
    service_manager = context.ServiceManager
    desktop = service_manager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", context
    )
    document = None
    try:
        document = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(os.path.abspath(template_file)),
            "_blank",
            0,
            (_property("Hidden", True), _property("ReadOnly", False)),
        )
        if document is None:
            raise RuntimeError(
                "LibreOffice could not open template: {0}".format(template_file)
            )

        insert_macro(document, macro_file)
        output_directory = os.path.dirname(os.path.abspath(output_file))
        if output_directory and not os.path.isdir(output_directory):
            os.makedirs(output_directory)
        document.storeAsURL(
            uno.systemPathToFileUrl(os.path.abspath(output_file)),
            (_property("FilterName", "writer8"), _property("Overwrite", True)),
        )
    finally:
        if document is not None:
            document.close(True)


def main(argv=None):
    arguments = sys.argv[1:] if argv is None else argv
    if len(arguments) != 4:
        print(
            "Usage: writer_macro_import.py "
            "<template_file> <macro_file> <output_file> <port>",
            file=sys.stderr,
        )
        return 2

    try:
        import_macro(
            arguments[0],
            arguments[1],
            arguments[2],
            int(arguments[3]),
        )
    except (OSError, RuntimeError, ValueError) as exc:
        print("Macro import failed: {0}".format(exc), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

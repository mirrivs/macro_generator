import importlib
import sys
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent

FILE_TYPES = (
    {
        "name": "Word",
        "aliases": {"word", "doc", "msword"},
        "module": "generators.word_file_generator",
        "template_dir": PROJECT_DIR / "templates" / "files" / "ms_office",
        "template_extensions": {".doc", ".docm", ".docx"},
        "macro_dir": PROJECT_DIR / "templates" / "macros" / "ms_office",
    },
    {
        "name": "Excel",
        "aliases": {"excel", "xls", "spreadsheet"},
        "module": "generators.excel_file_generator",
        "template_dir": PROJECT_DIR / "templates" / "files" / "ms_office",
        "template_extensions": {".xls", ".xlsm", ".xlsx", ".xltm", ".xltx"},
        "macro_dir": PROJECT_DIR / "templates" / "macros" / "ms_office",
    },
    {
        "name": "PowerPoint",
        "aliases": {"powerpoint", "ppt", "presentation"},
        "module": "generators.powerpoint_file_generator",
        "template_dir": PROJECT_DIR / "templates" / "files" / "ms_office",
        "template_extensions": {".potm", ".potx", ".ppt", ".pptm", ".pptx"},
        "macro_dir": PROJECT_DIR / "templates" / "macros" / "ms_office",
    },
    {
        "name": "LibreOffice Writer",
        "aliases": {"writer", "libreoffice", "libreoffice-writer", "odt"},
        "module": "generators.writer_file_generator",
        "template_dir": PROJECT_DIR / "templates" / "files" / "libre_office",
        "template_extensions": {".odt", ".ott", ".sxw"},
        "macro_dir": PROJECT_DIR / "templates" / "macros" / "libre_office",
        "macro_extensions": {".bas"},
    },
)


class UserCancelled(Exception):
    """Raised when the user exits a selection menu."""


def list_files(directory, extensions=None):
    if not directory.is_dir():
        return []

    return sorted(
        (
            path
            for path in directory.iterdir()
            if path.is_file()
            and not path.name.startswith("~$")
            and (extensions is None or path.suffix.lower() in extensions)
        ),
        key=lambda path: path.name.lower(),
    )


def get_file_type(value):
    normalized_value = value.strip().lower()
    for file_type in FILE_TYPES:
        if (
            normalized_value == file_type["name"].lower()
            or normalized_value in file_type["aliases"]
        ):
            return file_type

    valid_types = ", ".join(file_type["name"] for file_type in FILE_TYPES)
    raise ValueError(f"Unknown file type '{value}'. Choose one of: {valid_types}")


def resolve_file(directory, filename, extensions, description):
    path = Path(filename)
    if path.name != filename:
        raise ValueError(f"{description} must be a filename, not a path")
    if extensions and path.suffix.lower() not in extensions:
        valid_extensions = ", ".join(sorted(extensions))
        raise ValueError(f"{description} must use one of: {valid_extensions}")

    resolved_path = directory / filename
    if not resolved_path.is_file():
        raise FileNotFoundError(f"{description} not found: {resolved_path}")

    return resolved_path


def resolve_template(file_type, filename):
    return resolve_file(
        file_type["template_dir"],
        filename,
        file_type["template_extensions"],
        "Template",
    )


def resolve_macro(file_type, filename):
    return resolve_file(
        file_type["macro_dir"],
        filename,
        file_type.get("macro_extensions"),
        "Macro payload",
    )


def choose_item(prompt, items):
    while True:
        print(f"\n{prompt}")
        for index, item in enumerate(items, start=1):
            print(f"  {index}. {item.name if isinstance(item, Path) else item['name']}")
        print("  q. Quit")

        selection = (
            input("Select an option: ")
            .replace("\ufeff", "")
            .replace("\xef\xbb\xbf", "")
            .strip()
            .lower()
        )
        if selection in {"q", "quit", "exit"}:
            raise UserCancelled

        try:
            index = int(selection) - 1
            return items[index]
        except (ValueError, IndexError):
            print("Please enter one of the displayed numbers, or q to quit.")


def generate(file_type, template, macro):
    generator = importlib.import_module(file_type["module"])
    output_path = generator.main(template.name, macro.name)

    if output_path:
        print(f"\nGenerated file: {output_path}")


def run_with_parameters(file_type_name, template_name, macro_name):
    file_type = get_file_type(file_type_name)
    template = resolve_template(file_type, template_name)
    macro = resolve_macro(file_type, macro_name)

    print(
        f"Generating {file_type['name']} file from "
        f"'{template.name}' with payload '{macro.name}'..."
    )
    generate(file_type, template, macro)


def run_safe_libreoffice_training(arguments):
    """Run the passive LibreOffice training-document generator."""

    generator = importlib.import_module("generators.libreoffice_training_generator")
    return generator.main(arguments)


def run_calc_webservice(arguments):
    """Run the Calc WEBSERVICE(file://) local-file-read training generator."""

    generator = importlib.import_module("generators.calc_file_generator")
    return generator.main(arguments)


def run_interactive():
    print("File Generator")
    print("==============")

    try:
        file_type = choose_item("Choose a file type:", FILE_TYPES)

        templates = list_files(
            file_type["template_dir"], file_type["template_extensions"]
        )
        if not templates:
            print(f"No templates found in {file_type['template_dir']}")
            return 1
        template = choose_item("Choose a template:", templates)

        macros = list_files(file_type["macro_dir"], file_type.get("macro_extensions"))
        if not macros:
            print(f"No macro payloads found in {file_type['macro_dir']}")
            return 1
        macro = choose_item("Choose a macro payload:", macros)

        run_with_parameters(file_type["name"], template.name, macro.name)
        return 0
    except UserCancelled:
        print("\nCancelled.")
        return 0
    except (KeyboardInterrupt, EOFError):
        print("\nCancelled.")
        return 0
    except Exception as exc:
        print(f"\nGeneration failed: {exc}")
        return 1


def print_usage():
    print("Usage:")
    print("  python main.py")
    print("  python main.py <file_type> <template_file> <macro_payload_file>")
    print(
        "  python main.py libreoffice-training [--output FILE] [--title TITLE] "
        "[--include-safe-macro]"
    )
    print(
        "  python main.py calc --output FILE.fods --files PATH [PATH ...] "
        "[--list PATHS.txt] [--mode read|exfil] [--exfil-url URL]"
    )
    print("\nExamples:")
    print("  python main.py writer Vyplatka.odt training_simulation.bas")
    print("  python main.py word Handover_Protocol.docx agent_download.txt")
    print(
        "  python main.py calc --output webservice_exercise.fods "
        '--files "C:\\Users\\demo\\secret.txt"'
    )


def main(argv=None):
    arguments = sys.argv[1:] if argv is None else argv

    if arguments and arguments[0].strip().lower() in {
        "libreoffice-training",
        "odt-training",
    }:
        return run_safe_libreoffice_training(arguments[1:])

    if arguments and arguments[0].strip().lower() in {
        "calc",
        "webservice",
        "webservice-calc",
    }:
        return run_calc_webservice(arguments[1:])

    if not arguments:
        return run_interactive()
    if arguments == ["-h"] or arguments == ["--help"]:
        print_usage()
        return 0
    if len(arguments) != 3:
        print_usage()
        return 2

    try:
        run_with_parameters(*arguments)
        return 0
    except Exception as exc:
        print(f"Generation failed: {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())

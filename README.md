# macro_generator

## Config file

#### Files

Templates are selected directly by filename from the relevant folder; they are not listed in `config.yml`.

#### Macros

Macro payloads are selected directly from `templates/macros/ms_office` or `templates/macros/libre_office`.

#### App

#### Generators

Run the interactive generator prompt with:

```text
python main.py
```

It discovers available templates and macro payloads from the folders and lets you choose both.

For a non-interactive one-liner, pass the file type, template filename, and macro payload filename:

```text
python main.py writer Vyplatka.odt test.txt
python main.py word Handover_Protocol.docx malware_download.txt
python main.py excel Weekly_Task_Tracker.xlsx malware_download.txt
python main.py powerpoint Project_Status_Report_EN.pptx malware_download.txt
```

```text
python generators/word_file_generator.py <template_name> <macro_name>
python generators/writer_file_generator.py <template_name> <macro_name>
python generators/excel_file_generator.py <template_name> <macro_name>
python generators/powerpoint_file_generator.py <template_name> <macro_name>
```

Word, Excel, and PowerPoint templates are read from `templates/files/ms_office`; Writer templates are read from `templates/files/libre_office`.


#### Setup

###### Ms Office

Must have installed ms office on computer
Before running the code you must have word app on your computer.
You must allow programmatic access to visual basic
 - Open word
 - Go to File -> Options -> Trust Center -> Trust Center Settings -> Trust access to the VBA project object model

###### Libre Office

Must have installed [LibreOffice 5.4.4](https://www.filehorse.com/download-libreoffice-64/33266/download/) office on computer - 
Add libre office to path - Ex. C:\Program Files (x86)\LibreOffice 5\program

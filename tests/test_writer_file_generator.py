import unittest
from pathlib import Path
from unittest.mock import patch

from generators import writer_file_generator as writer


class WriterGeneratorTests(unittest.TestCase):
    def test_known_template_is_resolved(self):
        template = writer.get_template_path("Vyplatka.odt")
        self.assertEqual(template.suffix.lower(), ".odt")
        self.assertTrue(template.is_file())

    def test_only_basic_payloads_are_resolved(self):
        with self.assertRaises(ValueError):
            writer.get_macro_path("test.txt")

        macro = writer.get_macro_path("training_simulation.bas")
        self.assertEqual(macro.suffix.lower(), ".bas")

    def test_training_macro_is_allowed(self):
        macro = writer.get_macro_path("training_simulation.bas")
        source = writer.render_macro_source(macro, {"app": {"macros": {}}})
        self.assertIn("Sub TrainingSimulation", source)

    def test_shell_and_autorun_sources_are_rejected(self):
        unsafe_sources = (
            "Sub Bad\n    Shell(\"powershell.exe\")\nEnd Sub",
            "Sub Bad\n    Call CreateObject(\"WinHttp.WinHttpRequest.5.1\")\nEnd Sub",
            "Sub AutoOpen\nEnd Sub",
        )
        for source in unsafe_sources:
            with self.subTest(source=source):
                with self.assertRaises(ValueError):
                    writer.validate_macro_source(source)

    def test_missing_libreoffice_is_reported_before_import(self):
        missing_executable = str(Path(__file__).parent / "does-not-exist" / "soffice.exe")
        missing_python = str(Path(__file__).parent / "does-not-exist" / "python.exe")
        config = {
            "app": {
                "libre_office": {
                    "exe": missing_executable,
                    "python": missing_python,
                },
                "macros": {},
            }
        }
        with patch.object(writer, "_load_config", return_value=config):
            with self.assertRaises(FileNotFoundError):
                writer.main("Vyplatka.odt", "training_simulation.bas")


if __name__ == "__main__":
    unittest.main()

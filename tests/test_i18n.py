import json
import tempfile
import unittest
from pathlib import Path

from docx_format_scanner import Finding, ScanResult, _format_bericht
from i18n import detect_system_language, get_translator, normalize_language
from settings import initial_language, load_language, save_language


class TranslationTests(unittest.TestCase):
    def test_language_variants_and_fallback(self):
        self.assertEqual(normalize_language("de-DE"), "de")
        self.assertEqual(normalize_language("de_AT"), "de")
        self.assertEqual(normalize_language("es-ES"), "en")
        self.assertEqual(get_translator("de")("no_file"), "Keine Datei ausgewählt")
        self.assertEqual(get_translator("fr")("language.fr"), "Français")

    def test_report_uses_selected_language(self):
        result = ScanResult(
            findings=[
                Finding(
                    kategorie="Hervorhebung",
                    wert="yellow",
                    ort="word/document.xml: Absatz 1, Lauf 1",
                    textausschnitt="café",
                )
            ]
        )
        report = _format_bericht(result, "sample.docx", translate=get_translator("fr"))
        self.assertIn("Analyse du formatage", report)
        self.assertIn("[Surbrillance]", report)
        self.assertIn("Paragraphe 1, Texte 1", report)

    def test_system_language_is_supported_or_english(self):
        self.assertIn(detect_system_language(), ("de", "en", "fr"))


class SettingsTests(unittest.TestCase):
    def test_language_round_trip_uses_english_json_key(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "settings.json"
            self.assertEqual(initial_language(path, "fr"), "fr")
            save_language("de", path)
            self.assertEqual(load_language(path), "de")
            self.assertEqual(json.loads(path.read_text(encoding="utf-8")), {"language": "de"})

    def test_invalid_settings_fall_back(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "settings.json"
            path.write_text("not json", encoding="utf-8")
            self.assertEqual(initial_language(path, "fr"), "fr")


if __name__ == "__main__":
    unittest.main()

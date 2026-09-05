"""Translations used by the graphical user interface."""

from __future__ import annotations

import locale
import os
import gettext
from collections.abc import Callable

LANGUAGES = ("de", "en", "fr")
DEFAULT_LANGUAGE = "en"

_TRANSLATIONS = {
    "en": {
        "app.title": "WordHighLightFormatInspect",
        "language": "Language",
        "language.de": "Deutsch",
        "language.en": "English",
        "language.fr": "Français",
        "no_file": "No file selected",
        "choose_file": "Choose file...",
        "scan": "Scan",
        "create_copy": "Create annotated copy...",
        "ready": "Ready.",
        "file_selected": "File selected. Please scan.",
        "findings": "{count} finding(s) found.",
        "word_file": "Select Word file",
        "word_documents": "Word documents",
        "all_files": "All files",
        "error": "Error",
        "invalid_docx": "File is not a valid .docx (not a ZIP archive).",
        "scan_error": "Scan error",
        "save_copy": "Save annotated copy as",
        "suggested_copy": "_annotated.docx",
        "same_file": "The destination file must not be identical to the original file.",
        "finished": "Finished: {count} comment(s) written to {name}.",
        "finished_title": "Finished",
        "comments_written": "{count} comment(s) were written to a new file:\n{path}\n\n"
        "The original file was not changed.",
        "copy_error": "Error creating copy",
        "report_title": "Formatting analysis: {filename}",
        "report_parts": "Inspected parts: {parts}",
        "no_findings": "No unusual background, highlighting, or color formatting found.",
        "finding_count": "[{category}] - {count} finding(s)",
        "value_count": "  Value: {value}  ({count}x)",
        "fonts": "Fonts directly set in the document: {fonts}",
        "font_hint_1": "(Several different fonts may indicate editing",
        "font_hint_2": " in different programs/systems, but are not proof.)",
        "note_1": "Note: 'Clear formatting' in Word usually removes style-based",
        "note_2": "and w:highlight formatting reliably. Directly applied",
        "note_3": "character shading (w:shd at run level) often remains",
        "note_4": "and must be removed deliberately.",
        "category.highlight": "Highlighting",
        "category.character_shading": "Character shading",
        "category.text_color": "Text color",
        "category.character_style": "Character style",
        "category.paragraph_shading": "Paragraph shading",
        "category.cell_shading": "Table-cell shading",
        "location.paragraph": "Paragraph",
        "location.run": "Run",
        "location.table": "Table",
        "location.cell": "Cell",
    },
    "de": {
        "language": "Sprache",
        "language.de": "Deutsch",
        "language.en": "English",
        "language.fr": "Français",
        "no_file": "Keine Datei ausgewählt",
        "choose_file": "Datei wählen...",
        "scan": "Scannen",
        "create_copy": "Kommentierte Kopie erstellen...",
        "ready": "Bereit.",
        "file_selected": "Datei ausgewählt. Bitte scannen.",
        "findings": "{count} Fundstelle(n) gefunden.",
        "word_file": "Word-Datei auswählen",
        "word_documents": "Word-Dokumente",
        "all_files": "Alle Dateien",
        "error": "Fehler",
        "invalid_docx": "Datei ist kein gültiges .docx (kein ZIP-Archiv).",
        "scan_error": "Fehler beim Scannen",
        "save_copy": "Kommentierte Kopie speichern unter",
        "suggested_copy": "_kommentiert.docx",
        "same_file": "Zieldatei darf nicht mit der Originaldatei identisch sein.",
        "finished": "Fertig: {count} Kommentar(e) in {name} geschrieben.",
        "finished_title": "Fertig",
        "comments_written": "{count} Kommentar(e) wurden in eine neue Datei geschrieben:\n{path}\n\n"
        "Die Original-Datei wurde nicht verändert.",
        "copy_error": "Fehler beim Erstellen der Kopie",
        "report_title": "Formatierungsanalyse: {filename}",
        "report_parts": "Untersuchte Teile: {parts}",
        "no_findings": "Keine auffälligen Hintergrund-/Hervorhebungs-/Farbformatierungen gefunden.",
        "finding_count": "[{category}] - {count} Fundstelle(n)",
        "value_count": "  Wert: {value}  ({count}x)",
        "fonts": "Im Dokument direkt gesetzte Schriftarten: {fonts}",
        "font_hint_1": "(Mehrere unterschiedliche Schriftarten können ein Indiz für Bearbeitung",
        "font_hint_2": " in verschiedenen Programmen/Systemen sein, sind aber kein Beweis.)",
        "note_1": "Hinweis: 'Formatierung löschen' in Word entfernt i.d.R. nur formatvorlagen-",
        "note_2": "basierte und w:highlight-Formatierung zuverlässig. Direkt gesetzte",
        "note_3": "Zeichen-Schattierung (w:shd auf Laufebene) bleibt dabei häufig erhalten",
        "note_4": "und muss gezielt entfernt werden.",
        "category.highlight": "Hervorhebung",
        "category.character_shading": "Zeichen-Schattierung",
        "category.text_color": "Textfarbe",
        "category.character_style": "Zeichenformatvorlage",
        "category.paragraph_shading": "Absatz-Schattierung",
        "category.cell_shading": "Tabellenzellen-Schattierung",
        "location.paragraph": "Absatz",
        "location.run": "Lauf",
        "location.table": "Tabelle",
        "location.cell": "Zelle",
    },
    "fr": {
        "language": "Langue",
        "language.de": "Deutsch",
        "language.en": "English",
        "language.fr": "Français",
        "no_file": "Aucun fichier sélectionné",
        "choose_file": "Choisir un fichier...",
        "scan": "Analyser",
        "create_copy": "Créer une copie annotée...",
        "ready": "Prêt.",
        "file_selected": "Fichier sélectionné. Veuillez lancer l'analyse.",
        "findings": "{count} résultat(s) trouvé(s).",
        "word_file": "Sélectionner un fichier Word",
        "word_documents": "Documents Word",
        "all_files": "Tous les fichiers",
        "error": "Erreur",
        "invalid_docx": "Le fichier n'est pas un .docx valide (pas une archive ZIP).",
        "scan_error": "Erreur lors de l'analyse",
        "save_copy": "Enregistrer la copie annotée sous",
        "suggested_copy": "_annoté.docx",
        "same_file": "Le fichier de destination ne doit pas être identique au fichier original.",
        "finished": "Terminé : {count} commentaire(s) écrit(s) dans {name}.",
        "finished_title": "Terminé",
        "comments_written": "{count} commentaire(s) ont été écrits dans un nouveau fichier :\n{path}\n\n"
        "Le fichier original n'a pas été modifié.",
        "copy_error": "Erreur lors de la création de la copie",
        "report_title": "Analyse du formatage : {filename}",
        "report_parts": "Parties analysées : {parts}",
        "no_findings": "Aucun formatage inhabituel d'arrière-plan, de surbrillance ou de couleur trouvé.",
        "finding_count": "[{category}] - {count} résultat(s)",
        "value_count": "  Valeur : {value}  ({count}x)",
        "fonts": "Polices définies directement dans le document : {fonts}",
        "font_hint_1": "(Plusieurs polices peuvent indiquer une modification",
        "font_hint_2": " dans différents programmes/systèmes, sans en être la preuve.)",
        "note_1": "Remarque : 'Effacer la mise en forme' dans Word supprime généralement seulement",
        "note_2": "la mise en forme basée sur les styles et w:highlight. L'ombrage",
        "note_3": "direct des caractères (w:shd au niveau du texte) reste souvent présent",
        "note_4": "et doit être supprimé explicitement.",
        "category.highlight": "Surbrillance",
        "category.character_shading": "Ombrage des caractères",
        "category.text_color": "Couleur du texte",
        "category.character_style": "Style de caractère",
        "category.paragraph_shading": "Ombrage du paragraphe",
        "category.cell_shading": "Ombrage de cellule de tableau",
        "location.paragraph": "Paragraphe",
        "location.run": "Texte",
        "location.table": "Tableau",
        "location.cell": "Cellule",
    },
}

_CATEGORY_KEYS = {
    "Hervorhebung": "category.highlight",
    "Zeichen-Schattierung": "category.character_shading",
    "Textfarbe": "category.text_color",
    "Zeichenformatvorlage": "category.character_style",
    "Absatz-Schattierung": "category.paragraph_shading",
    "Tabellenzellen-Schattierung": "category.cell_shading",
}


class CatalogTranslations(gettext.NullTranslations):
    def __init__(self, catalog: dict[str, str]) -> None:
        super().__init__()
        self.catalog = catalog

    def gettext(self, message: str) -> str:
        return self.catalog.get(message, message)


def normalize_language(language: str | None) -> str:
    if not language:
        return DEFAULT_LANGUAGE
    language = language.lower().replace("_", "-").split("-", 1)[0]
    return language if language in LANGUAGES else DEFAULT_LANGUAGE


def detect_system_language() -> str:
    language = locale.getlocale()[0]
    if not language:
        language = os.environ.get("LANGUAGE") or os.environ.get("LC_ALL") or os.environ.get("LANG")
    return normalize_language(language)


def get_translator(language: str) -> Callable[[str], str]:
    selected = normalize_language(language)
    catalog = CatalogTranslations(_TRANSLATIONS[selected])
    fallback = _TRANSLATIONS[DEFAULT_LANGUAGE]

    def translate(key: str) -> str:
        value = catalog.gettext(key)
        return fallback.get(key, key) if value == key and key not in catalog.catalog else value

    return translate


def category_key(category: str) -> str:
    return _CATEGORY_KEYS.get(category, category)


def translate_location(location: str, translate: Callable[[str], str]) -> str:
    for source, key in (
        ("Absatz", "location.paragraph"),
        ("Lauf", "location.run"),
        ("Tabelle", "location.table"),
        ("Zelle", "location.cell"),
    ):
        location = location.replace(source, translate(key))
    return location
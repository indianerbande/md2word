"""Strings that end up *inside* the generated document.

These follow the document language (``--lang``), not the language of the
program: a German document should say "Inhaltsverzeichnis", even when the
tool reports its progress in English. Terminal output — help texts, errors,
warnings — is always English and is not routed through here.

Adding a language means adding one dictionary below; anything missing falls
back to English, so a partial translation is fine.
"""

from __future__ import annotations

# Keyed by the primary subtag of the language code ("de-AT" -> "de").
_STRINGS: dict[str, dict[str, str]] = {
    "en": {
        "toc_title": "Table of Contents",
        "toc_placeholder": "Right-click here and choose "
        "“Update Field” to build the table of contents.",
        "endnotes_title": "Notes",
        "image_placeholder": "[Image: {label}]",
        "untitled": "Untitled",
        "generated_note": "Generated with md2word from {source}",
    },
    "de": {
        "toc_title": "Inhaltsverzeichnis",
        "toc_placeholder": "Rechtsklick und „Felder aktualisieren“ wählen, "
        "um das Inhaltsverzeichnis zu erzeugen.",
        "endnotes_title": "Anmerkungen",
        "image_placeholder": "[Bild: {label}]",
        "untitled": "Ohne Titel",
        "generated_note": "Erzeugt mit md2word aus {source}",
    },
    "fr": {
        "toc_title": "Table des matières",
        "toc_placeholder": "Clic droit puis « Mettre à jour les champs » "
        "pour générer la table des matières.",
        "endnotes_title": "Notes",
        "image_placeholder": "[Image : {label}]",
        "untitled": "Sans titre",
        "generated_note": "Généré avec md2word à partir de {source}",
    },
    "es": {
        "toc_title": "Índice",
        "toc_placeholder": "Haga clic con el botón derecho y elija "
        "«Actualizar campos» para generar el índice.",
        "endnotes_title": "Notas",
        "image_placeholder": "[Imagen: {label}]",
        "untitled": "Sin título",
        "generated_note": "Generado con md2word a partir de {source}",
    },
    "it": {
        "toc_title": "Indice",
        "toc_placeholder": "Fare clic con il tasto destro e scegliere "
        "«Aggiorna campo» per generare l’indice.",
        "endnotes_title": "Note",
        "image_placeholder": "[Immagine: {label}]",
        "untitled": "Senza titolo",
        "generated_note": "Generato con md2word da {source}",
    },
}

DEFAULT_LANGUAGE = "en"


def language_of(lang: str) -> str:
    """Reduces a language tag to the primary subtag: 'de-AT' -> 'de'."""
    return (lang or DEFAULT_LANGUAGE).split("-")[0].split("_")[0].lower()


def translate(lang: str, key: str, **fields: object) -> str:
    """Looks up a document string, falling back to English."""
    table = _STRINGS.get(language_of(lang), {})
    text = table.get(key) or _STRINGS[DEFAULT_LANGUAGE][key]
    return text.format(**fields) if fields else text


def available_languages() -> list[str]:
    return sorted(_STRINGS)

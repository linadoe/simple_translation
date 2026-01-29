"""
Direkt in der IDE ausfuehrbar (Run-Button / F5).
Passe die Variablen unten an und starte das Skript.
"""
import configparser

from translate_pipeline import translate_document, translate_docx

config = configparser.ConfigParser()
config.read('./config.ini')
print("Available keys in [Data]:", list(config['Data'].keys()))
# === Hier anpassen ===
INPUT_FILE = rf"{config['Data']['file_s_mnr']}\Materialnummer_5-99.docx"      # Pfad zur Eingabedatei (.doc / .docx)
OUTPUT_FILE = rf"{config['Data']['file_s_mnr']}\Materialnummer_5-99-en.docx"                 # None = automatisch <input>_en.md / _en.docx

# INPUT_FILE = r"C:\Users\ldoering\Projektdaten\test_data\test_dokument.docx"
# OUTPUT_FILE = r"C:\Users\ldoering\Projektdaten\test_data\test_dokument_en.docx"
LLM_URL = "http://localhost:11435"   # URL des lokalen LLM-Servers
MODEL = "gpt-oss:120b"              # Modellname
METHOD = "auto"                      # auto / mammoth / pypandoc / python-docx / zipfile

# Modus: "markdown" = DOCX->Markdown->Übersetzung (bisheriges Verhalten)
#         "docx"     = DOCX->Übersetzung->DOCX (Formatierung bleibt erhalten)
MODE = "docx"

if __name__ == "__main__":
    if MODE == "docx":
        translate_docx(
            input_file=INPUT_FILE,
            output_file=OUTPUT_FILE,
            llm_url=LLM_URL,
            model=MODEL,
        )
    else:
        translate_document(
            input_file=INPUT_FILE,
            output_file=OUTPUT_FILE,
            llm_url=LLM_URL,
            model=MODEL,
            conversion_method=METHOD,
        )

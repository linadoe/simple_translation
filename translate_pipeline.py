#!/usr/bin/env python3
"""
Übersetzungs-Pipeline: DOC/DOCX -> Markdown -> Englische Übersetzung (via lokalem LLM)

Konvertiert ein Word-Dokument (.doc/.docx) zu Markdown und übersetzt den Inhalt
ins Englische mithilfe eines lokal laufenden LLM. Es werden KEINE Daten ins Internet
gesendet — alle Anfragen gehen ausschließlich an den lokalen LLM-Server.
"""

import sys
import json
import re
import argparse
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError

from doc_to_md_robust import convert_doc_to_md


# ---------------------------------------------------------------------------
# Konfiguration
# ---------------------------------------------------------------------------
DEFAULT_LLM_URL = "http://localhost:11435"
DEFAULT_MODEL = "gpt-oss:120b"

# Maximale Zeichenanzahl pro Chunk für die Übersetzung.
# Größere Chunks = weniger API-Aufrufe, aber je nach Kontextfenster des Modells
# kann ein zu großer Chunk abgeschnitten werden.
MAX_CHUNK_CHARS = 3000

# Timeout in Sekunden für LLM-Anfragen
LLM_TIMEOUT = 300


# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

def _validate_localhost(url: str) -> None:
    """Stellt sicher, dass die URL auf localhost zeigt — kein Internet-Zugriff."""
    from urllib.parse import urlparse
    parsed = urlparse(url)
    hostname = parsed.hostname or ""
    allowed = {"localhost", "127.0.0.1", "::1", "[::1]"}
    if hostname not in allowed:
        raise ValueError(
            f"Sicherheitsfehler: Die URL '{url}' zeigt nicht auf localhost. "
            f"Nur lokale LLM-Server sind erlaubt (hostname={hostname})."
        )


def _call_llm(prompt: str, system_prompt: str, *,
               llm_url: str, model: str) -> str:
    """
    Sendet eine Chat-Completion-Anfrage an den lokalen LLM-Server.

    Unterstützt das OpenAI-kompatible /v1/chat/completions Endpunkt-Format,
    das von Ollama, vLLM, llama.cpp-server u.a. unterstützt wird.
    """
    _validate_localhost(llm_url)

    endpoint = f"{llm_url.rstrip('/')}/v1/chat/completions"

    payload = json.dumps({
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.1,
        "stream": False,
    }).encode("utf-8")

    req = Request(
        endpoint,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    try:
        with urlopen(req, timeout=LLM_TIMEOUT) as resp:
            body = json.loads(resp.read().decode("utf-8"))
    except HTTPError as exc:
        raise RuntimeError(
            f"LLM-Server antwortete mit HTTP {exc.code}: {exc.read().decode('utf-8', errors='replace')}"
        ) from exc
    except URLError as exc:
        raise ConnectionError(
            f"Verbindung zum LLM-Server fehlgeschlagen ({endpoint}): {exc.reason}\n"
            "Ist der lokale LLM-Server gestartet?"
        ) from exc

    # OpenAI-kompatibles Antwortformat parsen
    try:
        return body["choices"][0]["message"]["content"]
    except (KeyError, IndexError) as exc:
        raise RuntimeError(
            f"Unerwartetes Antwortformat vom LLM-Server: {json.dumps(body, indent=2)}"
        ) from exc


# ---------------------------------------------------------------------------
# Markdown-Chunking
# ---------------------------------------------------------------------------

def _split_markdown_chunks(md_text: str, max_chars: int = MAX_CHUNK_CHARS) -> list[str]:
    """
    Teilt Markdown-Text in Abschnitte auf, die das Kontextfenster des LLM
    nicht überlasten.  Versucht an Absatz- / Heading-Grenzen zu trennen.
    """
    if len(md_text) <= max_chars:
        return [md_text]

    # Aufteilen an doppelten Zeilenumbrüchen (Absatzgrenzen)
    paragraphs = re.split(r"\n{2,}", md_text)

    chunks: list[str] = []
    current_chunk = ""

    for para in paragraphs:
        candidate = (current_chunk + "\n\n" + para).strip() if current_chunk else para
        if len(candidate) <= max_chars:
            current_chunk = candidate
        else:
            if current_chunk:
                chunks.append(current_chunk)
            # Falls ein einzelner Absatz zu lang ist, nehme ihn trotzdem als eigenen Chunk
            current_chunk = para

    if current_chunk:
        chunks.append(current_chunk)

    return chunks


# ---------------------------------------------------------------------------
# Übersetzung
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = (
    "You are a professional translator. Translate the following Markdown text "
    "into English. Preserve ALL Markdown formatting exactly as-is: headings (#), "
    "bold (**), italic (*), links, tables, lists, code blocks, etc. "
    "Only translate the human-readable text content. "
    "Do NOT add any commentary, explanations, or notes — return ONLY the "
    "translated Markdown."
)


def translate_markdown(md_text: str, *,
                       llm_url: str = DEFAULT_LLM_URL,
                       model: str = DEFAULT_MODEL) -> str:
    """
    Übersetzt Markdown-Text ins Englische via lokalem LLM.

    Args:
        md_text:  Der zu übersetzende Markdown-Inhalt.
        llm_url:  Basis-URL des lokalen LLM-Servers.
        model:    Modellname.

    Returns:
        Der übersetzte Markdown-Text.
    """
    chunks = _split_markdown_chunks(md_text)
    translated_parts: list[str] = []

    total = len(chunks)
    for idx, chunk in enumerate(chunks, start=1):
        print(f"  Übersetze Abschnitt {idx}/{total} ...")
        translated = _call_llm(
            prompt=chunk,
            system_prompt=SYSTEM_PROMPT,
            llm_url=llm_url,
            model=model,
        )
        translated_parts.append(translated.strip())

    return "\n\n".join(translated_parts)


# ---------------------------------------------------------------------------
# Haupt-Pipeline
# ---------------------------------------------------------------------------

def translate_document(
    input_file: str,
    output_file: str | None = None,
    *,
    llm_url: str = DEFAULT_LLM_URL,
    model: str = DEFAULT_MODEL,
    conversion_method: str = "auto",
) -> str:
    """
    Komplette Pipeline: DOC/DOCX → Markdown → Englische Übersetzung.

    Args:
        input_file:         Pfad zur Eingabedatei (.doc / .docx).
        output_file:        Pfad für die übersetzte .md-Datei.
                            Standard: <input>_en.md
        llm_url:            Basis-URL des lokalen LLM-Servers.
        model:              Modellname des LLM.
        conversion_method:  Konvertierungsmethode für doc→md
                            ('auto', 'mammoth', 'pypandoc', 'python-docx', 'zipfile').

    Returns:
        Den übersetzten Markdown-Inhalt als String.
    """
    _validate_localhost(llm_url)

    input_path = Path(input_file)
    if not input_path.exists():
        raise FileNotFoundError(f"Datei nicht gefunden: {input_path}")

    suffix = input_path.suffix.lower()
    if suffix not in {".doc", ".docx"}:
        raise ValueError(
            f"Nicht unterstütztes Dateiformat: '{suffix}'. "
            "Bitte eine .doc oder .docx Datei angeben."
        )

    # --- Schritt 1: Konvertierung DOC/DOCX → Markdown ---
    print("=" * 60)
    print("SCHRITT 1: Konvertierung DOC/DOCX -> Markdown")
    print("=" * 60)

    # Temporäre Markdown-Datei für die Zwischenstufe
    intermediate_md = input_path.with_suffix(".md")
    md_content = convert_doc_to_md(
        input_path, intermediate_md, method=conversion_method
    )

    print(f"\nMarkdown erstellt: {intermediate_md}")
    print(f"Inhaltslänge: {len(md_content)} Zeichen\n")

    # --- Schritt 2: Übersetzung via lokalem LLM ---
    print("=" * 60)
    print("SCHRITT 2: Übersetzung ins Englische (lokaler LLM)")
    print(f"  Server: {llm_url}")
    print(f"  Modell: {model}")
    print("=" * 60)

    translated = translate_markdown(md_content, llm_url=llm_url, model=model)

    # --- Schritt 3: Ergebnis speichern ---
    if output_file is None:
        output_path = input_path.with_name(input_path.stem + "_en.md")
    else:
        output_path = Path(output_file)

    output_path.write_text(translated, encoding="utf-8")

    print("\n" + "=" * 60)
    print(f"Übersetzung abgeschlossen: {output_path}")
    print(f"Ergebnislänge: {len(translated)} Zeichen")
    print("=" * 60)

    return translated


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="DOC/DOCX → Markdown → Englische Übersetzung (lokaler LLM)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Beispiele:\n"
            "  python translate_pipeline.py dokument.docx\n"
            "  python translate_pipeline.py dokument.docx -o ergebnis_en.md\n"
            "  python translate_pipeline.py dokument.docx --model gpt-oss:120b\n"
            "  python translate_pipeline.py dokument.docx --method mammoth\n"
        ),
    )
    parser.add_argument(
        "input_file",
        help="Eingabedatei (.doc oder .docx)",
    )
    parser.add_argument(
        "-o", "--output",
        default=None,
        help="Ausgabedatei für die übersetzte Markdown-Datei (Standard: <input>_en.md)",
    )
    parser.add_argument(
        "--llm-url",
        default=DEFAULT_LLM_URL,
        help=f"URL des lokalen LLM-Servers (Standard: {DEFAULT_LLM_URL})",
    )
    parser.add_argument(
        "--model",
        default=DEFAULT_MODEL,
        help=f"LLM-Modellname (Standard: {DEFAULT_MODEL})",
    )
    parser.add_argument(
        "--method",
        default="auto",
        choices=["auto", "mammoth", "pypandoc", "python-docx", "zipfile"],
        help="Konvertierungsmethode für DOC/DOCX → Markdown (Standard: auto)",
    )

    args = parser.parse_args()

    try:
        translate_document(
            input_file=args.input_file,
            output_file=args.output,
            llm_url=args.llm_url,
            model=args.model,
            conversion_method=args.method,
        )
    except (FileNotFoundError, ValueError, ConnectionError, RuntimeError) as exc:
        print(f"\nFehler: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

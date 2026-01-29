# Simple Translation

Eine lokale Dokumenten-Konvertierungs- und Übersetzungs-Pipeline, die Word-Dokumente (.doc/.docx) ins Englische übersetzt - komplett offline und datenschutzfreundlich.

## Features

- **Datenschutz-First**: Alle Daten bleiben auf dem lokalen Rechner
- **Word zu Markdown**: Automatische Konvertierung mit mehreren Fallback-Methoden
- **Lokales LLM**: Verwendet einen lokalen LLM-Server (keine Cloud-Dienste)
- **Format-Erhaltung**: Behält Markdown-Formatierung bei der Übersetzung bei
- **Flexibel konfigurierbar**: CLI-Optionen für alle wichtigen Parameter

## Voraussetzungen

### Python-Abhängigkeiten

Für die Dokumenten-Konvertierung wird mindestens eines der folgenden Pakete empfohlen (in Reihenfolge der Präferenz):

```bash
# Empfohlen - beste Formatierung
pip install mammoth

# Alternative 1 - benötigt zusätzlich Pandoc installiert
pip install pypandoc

# Alternative 2 - reine Python-Lösung
pip install python-docx
```

Falls keine dieser Bibliotheken installiert ist, wird eine einfache Text-Extraktion als Fallback verwendet.

### Lokaler LLM-Server

Ein lokaler LLM-Server muss laufen und mit der OpenAI-API kompatibel sein. Beispiele:

- [Ollama](https://ollama.ai/)
- [vLLM](https://github.com/vllm-project/vllm)
- [llama.cpp Server](https://github.com/ggerganov/llama.cpp)

Standardmäßig wird `http://localhost:11435` verwendet.

## Verwendung

### Einfache Übersetzung

```bash
# Übersetzt dokument.docx und speichert als dokument_en.md
python translate_pipeline.py dokument.docx
```

### Mit benutzerdefinierter Ausgabedatei

```bash
python translate_pipeline.py dokument.docx -o meine_uebersetzung.md
```

### Mit spezifischem LLM-Modell

```bash
python translate_pipeline.py dokument.docx --model llama3:70b
```

### Mit benutzerdefiniertem LLM-Server

```bash
python translate_pipeline.py dokument.docx --llm-url http://localhost:11434
```

### Mit spezifischer Konvertierungsmethode

```bash
python translate_pipeline.py dokument.docx --method mammoth
```

### Alle Optionen kombiniert

```bash
python translate_pipeline.py dokument.docx \
    -o output.md \
    --llm-url http://localhost:11434 \
    --model llama3:70b \
    --method auto
```

## CLI-Optionen

| Option | Standard | Beschreibung |
|--------|----------|--------------|
| `input_file` | (erforderlich) | Pfad zur .doc oder .docx Datei |
| `-o, --output` | `<input>_en.md` | Ausgabedatei für die Übersetzung |
| `--llm-url` | `http://localhost:11435` | URL des lokalen LLM-Servers |
| `--model` | `gpt-oss:120b` | Name des LLM-Modells |
| `--method` | `auto` | Konvertierungsmethode: `auto`, `mammoth`, `pypandoc`, `python-docx`, `zipfile` |

## Nur Konvertierung (ohne Übersetzung)

Falls nur eine Word-zu-Markdown-Konvertierung gewünscht ist:

```bash
# Automatische Methodenwahl
python doc_to_md_robust.py dokument.docx

# Mit spezifischer Ausgabedatei und Methode
python doc_to_md_robust.py dokument.docx output.md mammoth
```

## Wie es funktioniert

1. **Konvertierung**: Das Word-Dokument wird in Markdown umgewandelt
2. **Chunking**: Der Markdown-Text wird in ~3000 Zeichen große Abschnitte aufgeteilt
3. **Übersetzung**: Jeder Abschnitt wird einzeln vom lokalen LLM übersetzt
4. **Zusammenführung**: Die übersetzten Abschnitte werden zu einer Datei kombiniert

## Sicherheit

Die Pipeline ist so konzipiert, dass alle Daten lokal bleiben:
- Nur localhost-Verbindungen zum LLM-Server sind erlaubt
- Keine externen API-Aufrufe
- Keine Cloud-Dienste

## Fehlerbehebung

### "Kein LLM-Server erreichbar"

Stellen Sie sicher, dass Ihr lokaler LLM-Server läuft:

```bash
# Für Ollama
ollama serve

# Überprüfen Sie die Erreichbarkeit
curl http://localhost:11435/v1/models
```

### "Keine Konvertierungsmethode verfügbar"

Installieren Sie mindestens eine der Konvertierungsbibliotheken:

```bash
pip install mammoth
```

### Formatierungsprobleme

Versuchen Sie eine andere Konvertierungsmethode:

```bash
python translate_pipeline.py dokument.docx --method pypandoc
```

## Lizenz

MIT License

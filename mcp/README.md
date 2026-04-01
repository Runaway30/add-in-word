# Word Track Changes — MCP Server

MCP server locale che permette a un **custom agent di M365 Copilot** di modificare
documenti Word applicando le modifiche come **Track Changes**, così l'utente può
accettarle o rifiutarle direttamente in Word senza creare file nuovi.

## Avvio rapido

```bash
# Installa dipendenze (una volta sola)
uv sync

# Avvia il server MCP (stdio, per Copilot Studio / Claude Desktop)
uv run word-mcp
```

## Configurazione in Claude Desktop (`claude_desktop_config.json`)

```json
{
  "mcpServers": {
    "word-track-changes": {
      "command": "uv",
      "args": ["--directory", "/percorso/a/pca-brevetti/mcp", "run", "word-mcp"]
    }
  }
}
```

## Tool esposti

| Tool | Descrizione |
|------|-------------|
| `read_document` | Legge il testo di un `.docx` per paragrafo |
| `apply_tracked_changes` | Applica sostituzioni come `w:del` + `w:ins` nel file originale |
| `get_tracked_changes_summary` | Elenca le track changes pendenti nel documento |

## Flusso tipico del custom agent

1. `read_document(file_path)` → ottieni il testo corrente
2. L'agente (LLM) decide quali sostituzioni fare in base al prompt dell'utente
3. `apply_tracked_changes(file_path, [{old_text: "...", new_text: "..."}])` → scrive le modifiche
4. L'utente apre (o ricarica) il file in Word → vede le modifiche evidenziate → accetta/rifiuta

## Note tecniche

- Le track changes sono scritte come OOXML puro (`w:ins` / `w:del`) direttamente
  sull'XML del documento, compatibili con Word, LibreOffice e altri editor OOXML.
- `python-docx` non supporta track changes nativamente: la manipolazione avviene
  via `lxml` su `paragraph._p`.
- Ogni sostituzione viene applicata alla **prima occorrenza** nel documento.
  Per più occorrenze, passa più item nella lista `replacements`.

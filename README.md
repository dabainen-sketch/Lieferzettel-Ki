# Lieferzettel KI – Streamlit

Die App kann:
- Foto hochladen
- direkt mit der Kamera fotografieren
- Weglisten mit Gemini auslesen
- die lange Weglisten-Kopfzeile beibehalten
- die Tabelle bearbeiten
- eine echte XLSX-Datei herunterladen

## Secret

In Streamlit Community Cloud unter **Advanced settings → Secrets** setzen:

```toml
GEMINI_API_KEY = "DEIN_GEMINI_KEY"
```

Den Key niemals in GitHub-Code schreiben.

## Dateien

- `streamlit_app.py`
- `requirements.txt`

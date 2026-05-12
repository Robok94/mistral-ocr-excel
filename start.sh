#!/usr/bin/env bash
set -e

# Virtuelle Umgebung anlegen (falls nicht vorhanden)
if [ ! -d ".venv" ]; then
  echo "Erstelle virtuelle Umgebung…"
  python3 -m venv .venv
fi

# Abhängigkeiten installieren
echo "Installiere Abhängigkeiten…"
.venv/bin/pip install -q -r requirements.txt

# .env anlegen falls noch nicht vorhanden
if [ ! -f ".env" ]; then
  cp .env.example .env
  echo ""
  echo "HINWEIS: .env wurde aus .env.example erstellt."
  echo "Trage deinen Mistral API-Schlüssel in .env ein"
  echo "oder gib ihn direkt im Browser-Formular an."
  echo ""
fi

echo "App läuft unter: http://127.0.0.1:5000"
.venv/bin/python app.py

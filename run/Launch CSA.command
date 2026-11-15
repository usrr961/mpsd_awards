#!/bin/zsh
cd "$(dirname "$0")/.."
export TK_SILENCE_DEPRECATION=1
exec .venv/bin/python app.py

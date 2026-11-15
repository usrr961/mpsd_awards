#!/bin/zsh
set -e
cd "$(dirname "$0")/.."

if command -v python3 >/dev/null 2>&1; then
    PYTHON=python3
else
    echo "Python 3 was not found."
    if command -v brew >/dev/null 2>&1; then
        echo "Installing Python 3 with Homebrew..."
        brew install python
        PYTHON=python3
    else
        echo "Homebrew is not installed, so Python cannot be installed automatically."
        echo "Install Python 3 from https://www.python.org/downloads/"
        echo "Then run this setup file again."
        read -r "REPLY?Press Return to close..."
        exit 1
    fi
fi

echo "Using: $($PYTHON --version)"
if [ ! -x ".venv/bin/python" ]; then
    echo "Creating the CSA virtual environment..."
    "$PYTHON" -m venv .venv
fi

echo "Installing CSA requirements..."
.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install -r requirements.txt

echo "CSA setup is complete."
echo "You can now double-click: run/Launch CSA.command"
read -r "REPLY?Press Return to close..."

#!/bin/bash
set -e

if [ ! -d ".venv" ]; then
    python -m venv .venv
fi


if ! dpkg -s python3-tk &> /dev/null; then
    sudo apt-get update
    sudo apt-get install -y python3-tk
fi

source .venv/bin/activate
pip install -r requirements.txt

#!/bin/bash
set -e

if ! dpkg -s python3-venv &> /dev/null; then
    sudo apt-get update
    sudo apt-get install -y python3-venv
fi


if [ ! -d ".venv" ]; then
    sudo python3 -m venv .venv
fi


if ! dpkg -s python3-tk &> /dev/null; then
    sudo apt-get update
    sudo apt-get install -y python3-tk
fi

source .venv/bin/activate
pip install -r requirements.txt

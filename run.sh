#!/bin/bash
set -eu pipefail
cd ~/apps/otg-env/app/
uv sync
uv run ./outlook_to_google.py > ./logs/log.txt
cd ~

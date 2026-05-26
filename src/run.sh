#!/bin/bash
set -euo pipefail
cd ~/apps/otg-env/app/
. .venv/bin/activate
python ./src/outlook_to_google.py > src/log.txt
cd ~

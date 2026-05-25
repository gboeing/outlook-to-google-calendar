#!/bin/bash
set -euo pipefail
cd ~/apps/otg-env/app/
. .venv/bin/activate
python ./outlook_to_google.py > ./logs/log.txt
cd ~

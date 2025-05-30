#!/bin/bash
set -e pipefail
cd ~/apps/otg-env/app/
. .venv/bin/activate
python ./outlook_to_google.py > ./logs/log.txt
cd ~

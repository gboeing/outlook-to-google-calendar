#!/bin/sh -e
. ~/apps/otg-env/bin/activate
cd ~/apps/otg-env/app/
python outlook_to_google.py >> ~/apps/otg-env/app/logs/log.txt
deactivate
cd ~

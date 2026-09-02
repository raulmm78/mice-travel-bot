#!/bin/zsh
cd "$(dirname "$0")"
python3 update_before_start.py
python3 process_emails.py --dashboard

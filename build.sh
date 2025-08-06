#!/bin/sh

# pyinstaller --onefile --windowed --name="fox-tools" --icon=icon.ico gui.py
pyinstaller --clean -y fox-tools.spec --log-level=DEBUG
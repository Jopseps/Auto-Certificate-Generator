#!/bin/bash
echo "==================================================="
echo "  Auto Certificate Generator - Initializing..."
echo "==================================================="
echo ""
echo "Installing requirements... (This might take a minute on the first run)"

# Handle different linux environments by falling back to --user if needed
python3 -m pip install -r requirements.txt 2>/dev/null || python3 -m pip install --user -r requirements.txt

echo ""
echo "Launching the application..."
python3 AutoCert.py

#!/bin/bash

echo "Installing Mac PST Converter dependencies..."

# Install Python dependencies
echo "Installing core Python dependencies..."
pip install -r requirements.txt

echo ""
echo "Installation complete!"
echo ""
echo "The application now includes full PST parsing support via libpff-python."
echo "If libpff-python installation failed, the app will use basic PST parsing."
echo ""
echo "To run the application:"
echo "python -m src.macpst.main"
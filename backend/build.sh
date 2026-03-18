#!/usr/bin/env bash
# Render build script for backend
set -e

echo "Installing dependencies..."
pip install --upgrade pip
pip install -r requirements.txt

echo "Build complete!"

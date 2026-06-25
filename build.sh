#!/usr/bin/env bash
set -e

echo "==> Python version: $(python --version)"

echo "==> Nâng cấp pip..."
pip install --upgrade pip

echo "==> Cài lxml bản pre-built (không cần compile)..."
pip install "lxml>=5.1.0" --prefer-binary

echo "==> Cài các package còn lại..."
pip install -r requirements.txt

echo "==> Build xong!"

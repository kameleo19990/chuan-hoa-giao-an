#!/usr/bin/env bash
set -e

echo "==> Cài thư viện hệ thống cho lxml..."
apt-get update -qq
apt-get install -y libxml2-dev libxslt-dev python3-dev

echo "==> Cài Python packages..."
pip install --upgrade pip
pip install -r requirements.txt

echo "==> Build xong!"

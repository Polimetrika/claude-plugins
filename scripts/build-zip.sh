#!/usr/bin/env bash
# ZIP плагина для загрузки в claude.ai / Claude Desktop (Customize → Plugins → Upload plugin).
set -euo pipefail
cd "$(dirname "$0")/.."
version=$(python3 -c 'import json;print(json.load(open("plugins/polimetrika-demo/.claude-plugin/plugin.json"))["version"])')
mkdir -p dist
out="dist/polimetrika-demo-${version}.zip"
rm -f "$out"
(cd plugins/polimetrika-demo && zip -qr "../../$out" . -x '*/__pycache__/*' '*.pyc')
echo "$out"

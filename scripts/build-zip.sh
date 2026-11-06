#!/usr/bin/env bash
# ZIP плагина для claude.ai / Claude Desktop (Customize → Plugins → Upload plugin) и ZIP одного скилла
# для ChatGPT (Plugins → Skills → Create → Upload): в корне — папка скилла с SKILL.md.
set -euo pipefail
cd "$(dirname "$0")/.."
version=$(python3 -c 'import json;print(json.load(open("plugins/polimetrika-demo/.claude-plugin/plugin.json"))["version"])')
mkdir -p dist
out="dist/polimetrika-demo-${version}.zip"
rm -f "$out"
(cd plugins/polimetrika-demo && zip -qr "../../$out" . -x '*/__pycache__/*' '*.pyc')
echo "$out"

skill="dist/generate-project-pack-${version}.zip"
rm -f "$skill"
(cd plugins/polimetrika-demo/skills && zip -qr "../../../$skill" generate-project-pack -x '*/__pycache__/*' '*.pyc')
echo "$skill"

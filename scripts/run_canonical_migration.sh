#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"

cd "${REPO_ROOT}"

python3 scripts/migrate_to_gsheets.py \
  --source data/sandbox.xlsm \
  --credentials google-authorized-user.json \
  --title "Planning Template"

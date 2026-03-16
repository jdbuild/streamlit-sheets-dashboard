#!/usr/bin/env bash
set -euo pipefail

PIP_BIN="${HOME}/.local/bin/pip"

if [[ ! -x "${PIP_BIN}" ]]; then
  echo "Missing ${PIP_BIN}."
  echo "Bootstrap pip first:"
  echo "  curl -fsSL https://bootstrap.pypa.io/get-pip.py -o /tmp/get-pip.py"
  echo "  python3 /tmp/get-pip.py --user --break-system-packages"
  exit 1
fi

exec "${PIP_BIN}" install --user --break-system-packages -e '.[dev]'

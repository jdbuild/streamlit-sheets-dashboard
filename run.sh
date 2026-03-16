#!/usr/bin/env bash
set -euo pipefail



# this is needed to make sure that the streamlit command is available in the PATH
export PATH="$HOME/.local/bin:$PATH"
exec streamlit run app.py "$@"

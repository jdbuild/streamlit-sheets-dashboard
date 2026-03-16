PYTHON := python3
LOCAL_BIN := $(HOME)/.local/bin
STREAMLIT := $(LOCAL_BIN)/streamlit
PYTEST := $(LOCAL_BIN)/pytest
PIP := $(LOCAL_BIN)/pip

.PHONY: install test run

install:
	$(PIP) install --user --break-system-packages -e '.[dev]'

test:
	$(PYTEST) -q

run:
	$(STREAMLIT) run app.py

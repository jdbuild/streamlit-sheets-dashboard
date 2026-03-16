from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path


_SCRIPT_PATH = Path(__file__).resolve().parents[1] / "scripts" / "migrate_to_gsheets.py"
_SPEC = spec_from_file_location("migrate_to_gsheets", _SCRIPT_PATH)
assert _SPEC is not None and _SPEC.loader is not None
_MODULE = module_from_spec(_SPEC)
_SPEC.loader.exec_module(_MODULE)
_normalize_formula = _MODULE._normalize_formula


def test_normalize_formula_uses_semicolon_argument_separator() -> None:
    formula = "=SUM(OFFSET(T20:AE20,0,0))*OFFSET($J$196,$FU20,0)"
    assert _normalize_formula(formula) == "=SUM(OFFSET(T20:AE20;0;0))*OFFSET($J$196;$FU20;0)"


def test_normalize_formula_keeps_decimal_literals_as_comma() -> None:
    formula = "=ROUND(1.25,2)"
    assert _normalize_formula(formula) == "=ROUND(1,25;2)"

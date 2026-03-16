from planning.models import TEMPLATE_SHEET
from planning.xlsm_parser import parse_local_xlsm


def test_parser_extracts_dimensions_and_blocks():
    parsed = parse_local_xlsm("data/sandbox.xlsm")
    assert TEMPLATE_SHEET in parsed.project_sheet_names
    assert len(parsed.user_dimensions) >= 2
    assert len(parsed.wp_dimensions) >= 2
    assert len(parsed.project_blocks) > 0
    first = parsed.project_blocks[0]
    assert len(first.monthly_values) == 72
    assert first.monthly_values[0][1].startswith("T")

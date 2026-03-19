from planning.xlsm_parser import parse_local_xlsm
from planning.sync_service import PlanningSyncService


def test_sync_builds_facts_and_dimensions():
    parsed = parse_local_xlsm("data/sandbox.xlsm")
    service = PlanningSyncService()
    result = service.sync_parsed_workbook(parsed)
    assert result.project_count >= 1
    assert result.fact_row_count >= 72
    df = service.analytics_monthly_capacity()
    assert not df.empty


def test_sync_allows_unassigned_helper_rows_without_unknown_user_issue():
    parsed = parse_local_xlsm("data/sandbox.xlsm")
    service = PlanningSyncService()
    service.sync_parsed_workbook(parsed)
    issues = service.issues()
    matching = issues[
        (issues["error_code"] == "UNKNOWN_USERID")
        & (issues["sheet_title"] == "P-3WinPA")
        & (issues["cell_ref"] == "B111")
    ]
    assert matching.empty


def test_sync_allows_rows_without_wp_label():
    parsed = parse_local_xlsm("data/sandbox.xlsm")
    service = PlanningSyncService()
    service.sync_parsed_workbook(parsed)
    issues = service.issues()
    matching = issues[
        (issues["error_code"] == "WP_LABEL_PARSE_ERROR")
        & (issues["sheet_title"] == "P-3WinPA")
        & (issues["cell_ref"] == "E125")
    ]
    assert matching.empty


def test_sync_ignores_rows_below_140():
    parsed = parse_local_xlsm("data/sandbox.xlsm")
    assert not any(block.sheet_title == "P-3WinPA" and block.source_row >= 140 for block in parsed.project_blocks)

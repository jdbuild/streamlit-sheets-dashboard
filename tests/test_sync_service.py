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

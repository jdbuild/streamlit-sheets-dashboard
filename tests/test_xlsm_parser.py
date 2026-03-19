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


def test_parser_reads_block_local_userid_from_following_row():
    parsed = parse_local_xlsm("data/sandbox.xlsm")
    project_block = next(block for block in parsed.project_blocks if block.sheet_title == "P-3WinPA" and block.source_row == 20)
    assert project_block.person_slot == "P1"
    assert project_block.person_name == "Eva Turk"
    assert project_block.userid == "eturk"


def test_parser_reads_second_person_block_starting_at_row_31():
    parsed = parse_local_xlsm("data/sandbox.xlsm")
    project_block = next(block for block in parsed.project_blocks if block.sheet_title == "P-3WinPA" and block.source_row == 31)
    assert project_block.person_slot == "P2"
    assert project_block.person_name == "Martin Ernst"
    assert project_block.userid == "mernst"


def test_parser_reads_multiple_wp_rows_within_person_block():
    parsed = parse_local_xlsm("data/sandbox.xlsm")
    wp_rows = [block for block in parsed.project_blocks if block.sheet_title == "P-3WinPA" and block.person_name == "Eva Turk"]
    source_rows = {block.source_row for block in wp_rows}
    assert {20, 21, 22, 23, 24}.issubset(source_rows)


def test_parser_keeps_summary_row_for_validation():
    parsed = parse_local_xlsm("data/sandbox.xlsm")
    project_block = next(block for block in parsed.project_blocks if block.sheet_title == "P-CT-Extention" and block.source_row == 31)
    jan_2026 = next(
        hours for month_date, _, hours in project_block.summary_monthly_values if month_date.year == 2026 and month_date.month == 1
    )
    jun_2026 = next(
        hours for month_date, _, hours in project_block.summary_monthly_values if month_date.year == 2026 and month_date.month == 6
    )
    assert jan_2026 == 73
    assert jun_2026 == 73


def test_parser_only_includes_projects_listed_in_budget_sheet():
    parsed = parse_local_xlsm("data/sandbox.xlsm")
    parsed_projects = {sheet_name for sheet_name in parsed.project_sheet_names if sheet_name != TEMPLATE_SHEET}
    assert "P-3WinPA" in parsed_projects
    assert "P-CT-Extention" in parsed_projects
    assert "P-GFFNOE-Competency" in parsed_projects

from __future__ import annotations

import argparse
from pathlib import Path

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from openpyxl import load_workbook

from planning.config import get_config


def _sheets_service(credentials: Credentials):
    return build("sheets", "v4", credentials=credentials, cache_discovery=False)


def _drive_service(credentials: Credentials):
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def _load_credentials(path: Path) -> Credentials:
    return Credentials.from_authorized_user_file(
        path,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )


def migrate(local_xlsm: Path, credentials_path: Path, title: str) -> str:
    credentials = _load_credentials(credentials_path)
    workbook = load_workbook(local_xlsm, data_only=False, keep_vba=True)
    drive = _drive_service(credentials)
    spreadsheet = _sheets_service(credentials).spreadsheets().create(body={"properties": {"title": title}}).execute()
    spreadsheet_id = spreadsheet["spreadsheetId"]
    sheet_metadata = _sheets_service(credentials).spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    default_sheet_id = sheet_metadata["sheets"][0]["properties"]["sheetId"]
    requests = [{"deleteSheet": {"sheetId": default_sheet_id}}]
    for sheet in workbook.worksheets:
        requests.append({"addSheet": {"properties": {"title": sheet.title}}})
    _sheets_service(credentials).spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests},
    ).execute()
    metadata = _sheets_service(credentials).spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    title_to_id = {item["properties"]["title"]: item["properties"]["sheetId"] for item in metadata["sheets"]}
    for sheet in workbook.worksheets:
        values = []
        for row in sheet.iter_rows():
            values.append([cell.value for cell in row])
        _sheets_service(credentials).spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"'{sheet.title}'!A1",
            valueInputOption="USER_ENTERED",
            body={"values": values},
        ).execute()
        requests = []
        if sheet.freeze_panes:
            frozen_cell = sheet[sheet.freeze_panes] if isinstance(sheet.freeze_panes, str) else sheet.freeze_panes
            requests.append(
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": title_to_id[sheet.title],
                            "gridProperties": {
                                "frozenRowCount": frozen_cell.row - 1,
                                "frozenColumnCount": frozen_cell.column - 1,
                            },
                        },
                        "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount",
                    }
                }
            )
        for merged_range in sheet.merged_cells.ranges:
            requests.append(
                {
                    "mergeCells": {
                        "range": _a1_to_grid_range(title_to_id[sheet.title], str(merged_range)),
                        "mergeType": "MERGE_ALL",
                    }
                }
            )
        for col_letter, dim in sheet.column_dimensions.items():
            if dim.width:
                start_index = _col_to_index(col_letter)
                requests.append(
                    {
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId": title_to_id[sheet.title],
                                "dimension": "COLUMNS",
                                "startIndex": start_index,
                                "endIndex": start_index + 1,
                            },
                            "properties": {"pixelSize": int(dim.width * 7)},
                            "fields": "pixelSize",
                        }
                    }
                )
        for row_idx, dim in sheet.row_dimensions.items():
            if dim.height:
                requests.append(
                    {
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId": title_to_id[sheet.title],
                                "dimension": "ROWS",
                                "startIndex": row_idx - 1,
                                "endIndex": row_idx,
                            },
                            "properties": {"pixelSize": int(dim.height * 1.33)},
                            "fields": "pixelSize",
                        }
                    }
                )
        if requests:
            _sheets_service(credentials).spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": requests},
            ).execute()
    return spreadsheet_id


def _col_to_index(col_letter: str) -> int:
    index = 0
    for char in col_letter:
        index = index * 26 + (ord(char.upper()) - 64)
    return index - 1


def _a1_to_grid_range(sheet_id: int, a1_range: str) -> dict:
    start, end = a1_range.split(":")
    start_col = "".join(filter(str.isalpha, start))
    start_row = int("".join(filter(str.isdigit, start)))
    end_col = "".join(filter(str.isalpha, end))
    end_row = int("".join(filter(str.isdigit, end)))
    return {
        "sheetId": sheet_id,
        "startRowIndex": start_row - 1,
        "endRowIndex": end_row,
        "startColumnIndex": _col_to_index(start_col),
        "endColumnIndex": _col_to_index(end_col) + 1,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", default="data/sandbox.xlsm")
    parser.add_argument("--credentials", required=True)
    parser.add_argument("--title", default="Planning Template")
    args = parser.parse_args()
    spreadsheet_id = migrate(Path(args.source), Path(args.credentials), args.title)
    print(spreadsheet_id)


if __name__ == "__main__":
    main()

from typing import Any, Dict, List, Optional, Sequence, Tuple

from googleapiclient.discovery import build

from google_integration.google_drive import get_gui_credentials
from sheets.gsheet_api import GSheetClient


class GoogleSheetsIntegration:
    """Thin GUI wrapper over the existing Sheets client."""

    def __init__(
        self,
        spreadsheet_id: str,
        client_secret_file: str,
        token_file: str,
    ) -> None:
        creds = get_gui_credentials(client_secret_file, token_file)
        client = GSheetClient.__new__(GSheetClient)
        client.spreadsheet_id = spreadsheet_id
        client._service = build(
            "sheets",
            "v4",
            credentials=creds,
            cache_discovery=False,
        )
        self._client = client

    def write_table_report(
        self,
        sheet_title: str,
        headers: Sequence[str],
        rows: Sequence[Sequence[Any]],
        summary: Optional[Sequence[Tuple[Any, Any]]] = None,
    ) -> None:
        headers_list = list(headers)
        rows_list = [list(row) for row in rows]
        summary_list = list(summary or [])

        nc = len(headers_list)
        nc_eff = max(nc, 2)

        row_title = 0
        row_summ_hdr = 2
        row_summ_start = 3
        row_summ_end = row_summ_start + len(summary_list)
        row_detail_hdr = row_summ_end + 1
        row_col_hdr = row_detail_hdr + 1
        row_data_start = row_col_hdr + 1
        row_data_end = row_data_start + len(rows_list)

        def _pad(values: Sequence[Any]) -> List[Any]:
            return list(values) + [""] * (nc_eff - len(values))

        all_values: List[List[Any]] = []
        all_values.append(_pad([sheet_title]))
        all_values.append(_pad([""]))
        all_values.append(_pad(["СВОДНАЯ ИНФОРМАЦИЯ"]))
        for label, value in summary_list:
            all_values.append(_pad([str(label), str(value)]))
        all_values.append(_pad([""]))
        all_values.append(_pad(["ПОДРОБНЫЙ СПИСОК"]))
        all_values.append(headers_list)
        all_values.extend(rows_list)

        self._client.update_range("Sheet1!A1", all_values)

        requests: List[Dict[str, Any]] = []
        requests.append(
            {
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": 0,
                        "dimension": "COLUMNS",
                        "startIndex": 0,
                        "endIndex": nc_eff,
                    }
                }
            }
        )
        requests.append(
            {
                "mergeCells": {
                    "range": {
                        "sheetId": 0,
                        "startRowIndex": row_title,
                        "endRowIndex": row_title + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": nc_eff,
                    },
                    "mergeType": "MERGE_ALL",
                }
            }
        )
        requests.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": 0,
                        "startRowIndex": row_title,
                        "endRowIndex": row_title + 1,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 0.17, "green": 0.24, "blue": 0.31},
                            "textFormat": {
                                "bold": True,
                                "fontSize": 13,
                                "foregroundColor": {"red": 1, "green": 1, "blue": 1},
                            },
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE",
                        }
                    },
                    "fields": (
                        "userEnteredFormat(backgroundColor,textFormat,"
                        "horizontalAlignment,verticalAlignment)"
                    ),
                }
            }
        )
        requests.append(
            {
                "mergeCells": {
                    "range": {
                        "sheetId": 0,
                        "startRowIndex": row_summ_hdr,
                        "endRowIndex": row_summ_hdr + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": nc_eff,
                    },
                    "mergeType": "MERGE_ALL",
                }
            }
        )
        requests.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": 0,
                        "startRowIndex": row_summ_hdr,
                        "endRowIndex": row_summ_hdr + 1,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 0.16, "green": 0.50, "blue": 0.73},
                            "textFormat": {
                                "bold": True,
                                "fontSize": 10,
                                "foregroundColor": {"red": 1, "green": 1, "blue": 1},
                            },
                            "horizontalAlignment": "LEFT",
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
                }
            }
        )

        for idx in range(len(summary_list)):
            row_index = row_summ_start + idx
            bg_color = (
                {"red": 0.94, "green": 0.95, "blue": 0.97}
                if idx % 2 == 0
                else {"red": 1.0, "green": 1.0, "blue": 1.0}
            )
            requests.append(
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": 0,
                            "startRowIndex": row_index,
                            "endRowIndex": row_index + 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 1,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": bg_color,
                                "textFormat": {
                                    "foregroundColor": {
                                        "red": 0.27,
                                        "green": 0.35,
                                        "blue": 0.43,
                                    }
                                },
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor,textFormat)",
                    }
                }
            )
            requests.append(
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": 0,
                            "startRowIndex": row_index,
                            "endRowIndex": row_index + 1,
                            "startColumnIndex": 1,
                            "endColumnIndex": 2,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": bg_color,
                                "textFormat": {
                                    "bold": True,
                                    "foregroundColor": {
                                        "red": 0.17,
                                        "green": 0.24,
                                        "blue": 0.31,
                                    },
                                },
                                "horizontalAlignment": "RIGHT",
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
                    }
                }
            )

        requests.append(
            {
                "mergeCells": {
                    "range": {
                        "sheetId": 0,
                        "startRowIndex": row_detail_hdr,
                        "endRowIndex": row_detail_hdr + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": nc_eff,
                    },
                    "mergeType": "MERGE_ALL",
                }
            }
        )
        requests.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": 0,
                        "startRowIndex": row_detail_hdr,
                        "endRowIndex": row_detail_hdr + 1,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 0.16, "green": 0.50, "blue": 0.73},
                            "textFormat": {
                                "bold": True,
                                "fontSize": 10,
                                "foregroundColor": {"red": 1, "green": 1, "blue": 1},
                            },
                            "horizontalAlignment": "LEFT",
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
                }
            }
        )
        requests.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": 0,
                        "startRowIndex": row_col_hdr,
                        "endRowIndex": row_col_hdr + 1,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 0.10, "green": 0.34, "blue": 0.55},
                            "textFormat": {
                                "bold": True,
                                "foregroundColor": {"red": 1, "green": 1, "blue": 1},
                            },
                            "horizontalAlignment": "CENTER",
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
                }
            }
        )

        if rows_list:
            requests.append(
                {
                    "addBanding": {
                        "bandedRange": {
                            "range": {
                                "sheetId": 0,
                                "startRowIndex": row_data_start,
                                "endRowIndex": row_data_end,
                                "startColumnIndex": 0,
                                "endColumnIndex": nc,
                            },
                            "rowProperties": {
                                "headerColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                                "firstBandColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                                "secondBandColor": {"red": 0.95, "green": 0.97, "blue": 0.99},
                            },
                        }
                    }
                }
            )

        requests.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": 0,
                        "startRowIndex": row_summ_hdr,
                        "endRowIndex": row_data_end,
                        "startColumnIndex": 0,
                        "endColumnIndex": nc_eff,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "borders": {
                                "top": {
                                    "style": "SOLID",
                                    "color": {"red": 0.8, "green": 0.8, "blue": 0.8},
                                },
                                "bottom": {
                                    "style": "SOLID",
                                    "color": {"red": 0.8, "green": 0.8, "blue": 0.8},
                                },
                                "left": {
                                    "style": "SOLID",
                                    "color": {"red": 0.8, "green": 0.8, "blue": 0.8},
                                },
                                "right": {
                                    "style": "SOLID",
                                    "color": {"red": 0.8, "green": 0.8, "blue": 0.8},
                                },
                            }
                        }
                    },
                    "fields": "userEnteredFormat.borders",
                }
            }
        )
        requests.append(
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": 0,
                        "gridProperties": {"frozenRowCount": row_col_hdr + 1},
                    },
                    "fields": "gridProperties.frozenRowCount",
                }
            }
        )

        self._client.apply_requests(requests)

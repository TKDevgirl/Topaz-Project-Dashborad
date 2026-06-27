from datetime import datetime
from io import BytesIO

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

from config import REPORT_COLUMNS, TAKENAKA_SHEETS, TRACKING_SHEETS
from compare import classify_action, should_include_tracking
from utils import base_doc_no, normalize_header


def empty_report_df():
    return pd.DataFrame(columns=REPORT_COLUMNS)


def find_header_row_and_columns(ws):
    for row in range(1, 30):
        headers = {}

        for col in range(1, ws.max_column + 1):
            value = ws.cell(row=row, column=col).value

            if value:
                headers[normalize_header(value)] = col

        if "STATUS" in headers and "INFO" in headers:
            return row, headers

    # Fallback for current Tracking_document.xlsx layout
    return 1, {
        "DOCUMENT NO": 2,
        "DOCUMENT NAME": 4,
        "STATUS": 5,
        "INFO": 6,
    }


def get_col(headers, possible_names):
    for name in possible_names:
        key = normalize_header(name)

        if key in headers:
            return headers[key]

    for key, col in headers.items():
        for name in possible_names:
            name_key = normalize_header(name)

            if name_key in key or key in name_key:
                return col

    return None


def read_takenaka(takenaka_file):
    workbook = load_workbook(takenaka_file, data_only=True)
    data = {}

    for sheet in TAKENAKA_SHEETS:
        if sheet not in workbook.sheetnames:
            continue

        ws = workbook[sheet]

        for row in range(11, ws.max_row + 1):
            doc_no = ws[f"E{row}"].value

            if not doc_no or "DETH-NSC" not in str(doc_no):
                continue

            data[base_doc_no(doc_no)] = {
                "Takenaka Sheet": sheet,
                "Takenaka Doc No": doc_no,
                "Takenaka Status 1": ws[f"AA{row}"].value,
                "Takenaka Status 2": ws[f"AB{row}"].value,
                "Takenaka Status 3": ws[f"AC{row}"].value,
            }

    return data


def generate_report(tracking_file, takenaka_file):
    takenaka_map = read_takenaka(takenaka_file)

    # data_only workbook for reading formulas/values
    read_wb = load_workbook(tracking_file, data_only=True)

    # normal workbook for export report sheet
    output_wb = load_workbook(tracking_file)

    report_sheet = "Open_On_Process_Compare"

    if report_sheet in output_wb.sheetnames:
        del output_wb[report_sheet]

    report_ws = output_wb.create_sheet(report_sheet)
    report_ws.append(REPORT_COLUMNS)

    fills = {
        "UPDATE TRACKING TO CLOSED": PatternFill(fill_type="solid", fgColor="C6EFCE"),
        "OPEN & ON PROCESS": PatternFill(fill_type="solid", fgColor="FFEB9C"),
        "OPEN": PatternFill(fill_type="solid", fgColor="BDD7EE"),
        "OVERDUE / FOLLOW UP": PatternFill(fill_type="solid", fgColor="FFC7CE"),
        "RETURNED BY NV5 / NEED RESUBMIT": PatternFill(fill_type="solid", fgColor="FFC7CE"),
        "NOT FOUND IN TAKENAKA SOURCE": PatternFill(fill_type="solid", fgColor="FFC7CE"),
        "CHECK": PatternFill(fill_type="solid", fgColor="D9EAF7"),
    }

    for cell in report_ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(fill_type="solid", fgColor="D9EAF7")

    rows = []
    total_docs = 0
    focus_docs = 0

    for sheet in TRACKING_SHEETS:
        if sheet not in read_wb.sheetnames:
            continue

        ws = read_wb[sheet]

        header_row, headers = find_header_row_and_columns(ws)

        doc_no_col = get_col(headers, ["Document No.", "Document No", "Ref No.", "Ref No", "Drawing ref No."])
        doc_name_col = get_col(headers, ["Document Name", "Equipment Name", "Description"])
        status_col = get_col(headers, ["Status"])
        info_col = get_col(headers, ["Info"])

        if not doc_no_col or not status_col or not info_col:
            continue

        for row in range(header_row + 1, ws.max_row + 1):
            doc_no = ws.cell(row=row, column=doc_no_col).value
            doc_name = ws.cell(row=row, column=doc_name_col).value if doc_name_col else ""
            tracking_status = ws.cell(row=row, column=status_col).value
            info = ws.cell(row=row, column=info_col).value

            if not doc_no or "DETH-NSC" not in str(doc_no):
                continue

            total_docs += 1

            if not should_include_tracking(tracking_status, info):
                continue

            focus_docs += 1

            source = takenaka_map.get(base_doc_no(doc_no))
            action = classify_action(source)
            checked_time = datetime.now().strftime("%d-%b-%Y %H:%M:%S")

            if source:
                row_data = [
                    sheet,
                    doc_no,
                    doc_name,
                    tracking_status,
                    info,
                    source["Takenaka Sheet"],
                    source["Takenaka Doc No"],
                    source["Takenaka Status 1"],
                    source["Takenaka Status 2"],
                    source["Takenaka Status 3"],
                    action,
                    checked_time,
                ]
            else:
                row_data = [
                    sheet,
                    doc_no,
                    doc_name,
                    tracking_status,
                    info,
                    "",
                    "",
                    "",
                    "",
                    "",
                    action,
                    checked_time,
                ]

            report_ws.append(row_data)
            report_ws[f"K{report_ws.max_row}"].fill = fills.get(action, fills["CHECK"])

            rows.append(dict(zip(REPORT_COLUMNS, row_data)))

    for col in report_ws.columns:
        max_len = 0
        col_letter = col[0].column_letter

        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))

        report_ws.column_dimensions[col_letter].width = min(max_len + 2, 55)

    output = BytesIO()
    output_wb.save(output)
    output.seek(0)

    df = pd.DataFrame(rows, columns=REPORT_COLUMNS)
    action_counts = df["Action"].value_counts().to_dict() if not df.empty else {}

    return {
        "report": output,
        "total_docs": total_docs,
        "focus_docs": focus_docs,
        "df": df,
        "action_counts": action_counts,
        "last_updated": datetime.now().strftime("%d-%b-%Y %H:%M:%S"),
    }

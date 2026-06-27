from datetime import datetime
from io import BytesIO

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

from config import DOC_CATEGORIES, REPORT_COLUMNS, TAKENAKA_SHEETS, TRACKING_SHEETS
from compare import classify_action, normalize_tracking_state, should_include_tracking
from utils import base_doc_no, contains_doc_no, extract_category, norm_text, norm_upper, normalize_header


STATUS_WORDS = [
    "OPEN",
    "CLOSE",
    "CLOSED",
    "RETURNED",
    "ON PROGRESS",
    "ON PROCESS",
    "APPROVE",
    "APPROVED",
]

INFO_WORDS = [
    "REVISED AND RESUBMIT",
    "REVISE AND RESUBMIT",
    "TTI TO CB",
    "NA",
    "NO OBJECTION",
    "APPROVE",
    "ANSWERED",
]


def empty_report_df():
    return pd.DataFrame(columns=REPORT_COLUMNS)


def find_header_row_and_columns(ws):
    best_row = 1
    best_headers = {}
    best_score = -1

    for row in range(1, 30):
        headers = {}

        for col in range(1, ws.max_column + 1):
            value = ws.cell(row=row, column=col).value

            if value:
                headers[normalize_header(value)] = col

        score = 0
        for key in headers:
            if "DOCUMENT" in key:
                score += 3
            if key in ["STATUS", "INFO", "VERSION"]:
                score += 3
            if "NAME" in key or "DESCRIPTION" in key:
                score += 2

        if score > best_score:
            best_score = score
            best_row = row
            best_headers = headers

        if "STATUS" in headers and "INFO" in headers:
            return row, headers

    return best_row, best_headers


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


def find_doc_no_in_row(ws, row):
    for col in range(1, ws.max_column + 1):
        value = ws.cell(row=row, column=col).value

        if contains_doc_no(value):
            return value, col

    return None, None


def guess_status_from_row(ws, row, status_col=None):
    if status_col:
        value = ws.cell(row=row, column=status_col).value
        if norm_text(value):
            return value

    for col in range(1, ws.max_column + 1):
        value = norm_upper(ws.cell(row=row, column=col).value)

        if value in ["OPEN", "CLOSE", "CLOSED", "RETURNED", "ON PROGRESS", "ON PROCESS", "APPROVED"]:
            return ws.cell(row=row, column=col).value

    return ""


def guess_info_from_row(ws, row, info_col=None):
    if info_col:
        value = ws.cell(row=row, column=info_col).value
        if norm_text(value):
            return value

    for col in range(1, ws.max_column + 1):
        value = norm_upper(ws.cell(row=row, column=col).value)

        if any(word in value for word in INFO_WORDS):
            return ws.cell(row=row, column=col).value

    return ""


def guess_doc_name(ws, row, doc_no_col, doc_name_col=None):
    if doc_name_col:
        value = ws.cell(row=row, column=doc_name_col).value
        if norm_text(value):
            return value

    for col in range(doc_no_col + 1, min(ws.max_column, doc_no_col + 4) + 1):
        value = ws.cell(row=row, column=col).value
        text = norm_text(value)

        if text and "DETH-NSC" not in text and norm_upper(text) not in STATUS_WORDS:
            return value

    for col in range(1, ws.max_column + 1):
        value = ws.cell(row=row, column=col).value
        text = norm_text(value)

        if len(text) > 10 and "DETH-NSC" not in text:
            return value

    return ""


def read_tracking_all(tracking_file):
    workbook = load_workbook(tracking_file, data_only=True)
    rows = []
    total_docs = 0

    for sheet in TRACKING_SHEETS:
        if sheet not in workbook.sheetnames:
            continue

        ws = workbook[sheet]
        header_row, headers = find_header_row_and_columns(ws)

        doc_no_col = get_col(headers, ["Document No.", "Document No", "Ref No.", "Ref No", "Drawing ref No."])
        doc_name_col = get_col(headers, ["Document Name", "Equipment Name", "Description"])
        category_col = get_col(headers, ["Category", "Document Category"])
        status_col = get_col(headers, ["Status"])
        info_col = get_col(headers, ["Info"])

        seen_base = set()

        for row in range(header_row + 1, ws.max_row + 1):
            doc_no = ws.cell(row=row, column=doc_no_col).value if doc_no_col else None
            doc_no_source_col = doc_no_col

            if not doc_no or "DETH-NSC" not in str(doc_no):
                doc_no, doc_no_source_col = find_doc_no_in_row(ws, row)

            if not doc_no or "DETH-NSC" not in str(doc_no):
                continue

            base_key = base_doc_no(doc_no)

            # Count base document once per sheet for dashboard summary
            if base_key in seen_base:
                continue

            seen_base.add(base_key)
            total_docs += 1

            category_value = ws.cell(row=row, column=category_col).value if category_col else ""
            category = extract_category(doc_no, category_value)

            doc_name = guess_doc_name(ws, row, doc_no_source_col or 1, doc_name_col)
            tracking_status = guess_status_from_row(ws, row, status_col)
            info = guess_info_from_row(ws, row, info_col)
            tracking_state = normalize_tracking_state(tracking_status, info)

            rows.append({
                "Tracking Sheet": sheet,
                "Document No": doc_no,
                "Base Document No": base_key,
                "Document Name": doc_name,
                "Document Category": category,
                "Tracking Status": tracking_status,
                "Info": info,
                "Tracking State": tracking_state,
            })

    return pd.DataFrame(rows), total_docs


def build_dashboard_matrix(tracking_df):
    categories = DOC_CATEGORIES
    rows = ["Open", "on progress", "Approved"]

    matrix = pd.DataFrame(0, index=rows, columns=["Total Document"] + categories)

    if tracking_df.empty:
        matrix.loc["Total", "Total Document"] = 0
        return matrix

    for _, item in tracking_df.iterrows():
        state = item["Tracking State"]
        category = item["Document Category"]

        if state not in rows:
            continue

        matrix.loc[state, "Total Document"] += 1

        if category in categories:
            matrix.loc[state, category] += 1

    total_row = pd.DataFrame(matrix.sum(axis=0)).T
    total_row.index = ["Total Document"]

    # Desired display order similar to user's screenshot:
    # Total Document / Open / on progress / Approved
    display = pd.concat([total_row, matrix.loc[["Open", "on progress", "Approved"]]])

    return display


def read_takenaka(takenaka_file):
    workbook = load_workbook(takenaka_file, data_only=True)
    data = {}

    for sheet in TAKENAKA_SHEETS:
        if sheet not in workbook.sheetnames:
            continue

        ws = workbook[sheet]

        for row in range(1, ws.max_row + 1):
            doc_no = ws[f"E{row}"].value

            if not doc_no or "DETH-NSC" not in str(doc_no):
                doc_no, _ = find_doc_no_in_row(ws, row)

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
    tracking_df, total_docs = read_tracking_all(tracking_file)
    takenaka_map = read_takenaka(takenaka_file)

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
    focus_docs = 0

    for _, item in tracking_df.iterrows():
        tracking_status = item["Tracking Status"]
        info = item["Info"]

        if not should_include_tracking(tracking_status, info):
            continue

        focus_docs += 1

        base_key = item["Base Document No"]
        source = takenaka_map.get(base_key)
        action = classify_action(source)
        checked_time = datetime.now().strftime("%d-%b-%Y %H:%M:%S")

        if source:
            row_data = [
                item["Tracking Sheet"],
                item["Document No"],
                item["Document Name"],
                item["Document Category"],
                item["Tracking Status"],
                item["Info"],
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
                item["Tracking Sheet"],
                item["Document No"],
                item["Document Name"],
                item["Document Category"],
                item["Tracking Status"],
                item["Info"],
                "",
                "",
                "",
                "",
                "",
                action,
                checked_time,
            ]

        report_ws.append(row_data)
        report_ws[f"L{report_ws.max_row}"].fill = fills.get(action, fills["CHECK"])
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
    dashboard_matrix = build_dashboard_matrix(tracking_df)

    return {
        "report": output,
        "total_docs": int(total_docs),
        "focus_docs": int(focus_docs),
        "df": df,
        "tracking_df": tracking_df,
        "dashboard_matrix": dashboard_matrix,
        "action_counts": action_counts,
        "last_updated": datetime.now().strftime("%d-%b-%Y %H:%M:%S"),
    }

from datetime import datetime
from io import BytesIO

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

from config import DASHBOARD_SHEET, REPORT_COLUMNS, TAKENAKA_SHEETS, TRACKING_SHEETS
from compare import classify_action, normalize_tracking_state, should_include_tracking
from utils import base_doc_no, contains_doc_no, extract_category, norm_text, norm_upper, normalize_header


def empty_report_df():
    return pd.DataFrame(columns=REPORT_COLUMNS)


def safe_int(value):
    try:
        if value is None or value == "":
            return 0
        return int(float(value))
    except Exception:
        return 0


def blank_dashboard_matrix():
    return pd.DataFrame([
        {"Status": "Total Document", "Total Document": 0, "MAT": 0, "MCR": 0, "MTS": 0, "CVI": 0},
        {"Status": "Open", "Total Document": 0, "MAT": 0, "MCR": 0, "MTS": 0, "CVI": 0},
        {"Status": "on progress", "Total Document": 0, "MAT": 0, "MCR": 0, "MTS": 0, "CVI": 0},
        {"Status": "Approved", "Total Document": 0, "MAT": 0, "MCR": 0, "MTS": 0, "CVI": 0},
    ])


def read_dashboard_sheet(tracking_file):
    wb = load_workbook(tracking_file, data_only=True)

    if DASHBOARD_SHEET not in wb.sheetnames:
        return blank_dashboard_matrix()

    ws = wb[DASHBOARD_SHEET]

    header_row = None
    header_cols = {}

    for row in range(1, min(ws.max_row, 40) + 1):
        temp = {}
        for col in range(1, ws.max_column + 1):
            value = norm_upper(ws.cell(row=row, column=col).value)
            if value:
                temp[value] = col

        if "TOTAL DOCUMENT" in temp and ("MAT" in temp or "MCR" in temp or "MTS" in temp or "CVI" in temp):
            header_row = row
            header_cols = temp
            break

    if not header_row:
        return blank_dashboard_matrix()

    total_col = header_cols.get("TOTAL DOCUMENT", 1)
    label_cols = list(range(1, total_col))

    def detect_label(row_num):
        for col in label_cols:
            value = norm_upper(ws.cell(row=row_num, column=col).value)

            if not value:
                continue

            if "TOTAL" in value and "DOCUMENT" in value:
                return "Total Document"
            if "SEND TO TTI" in value or "PROGRESS" in value:
                return "on progress"
            if "OPEN" in value:
                return "Open"
            if "APPROV" in value:
                return "Approved"

        return ""

    by_status = {}

    for row in range(header_row + 1, min(ws.max_row, header_row + 15) + 1):
        label = detect_label(row)
        if not label:
            continue

        by_status[label] = {
            "Status": label,
            "Total Document": safe_int(ws.cell(row=row, column=header_cols.get("TOTAL DOCUMENT", 0)).value) if header_cols.get("TOTAL DOCUMENT") else 0,
            "MAT": safe_int(ws.cell(row=row, column=header_cols.get("MAT", 0)).value) if header_cols.get("MAT") else 0,
            "MCR": safe_int(ws.cell(row=row, column=header_cols.get("MCR", 0)).value) if header_cols.get("MCR") else 0,
            "MTS": safe_int(ws.cell(row=row, column=header_cols.get("MTS", 0)).value) if header_cols.get("MTS") else 0,
            "CVI": safe_int(ws.cell(row=row, column=header_cols.get("CVI", 0)).value) if header_cols.get("CVI") else 0,
        }

    rows = []
    for status in ["Total Document", "Open", "on progress", "Approved"]:
        rows.append(by_status.get(status, {
            "Status": status,
            "Total Document": 0,
            "MAT": 0,
            "MCR": 0,
            "MTS": 0,
            "CVI": 0,
        }))

    return pd.DataFrame(rows)


def get_dashboard_value(matrix, status, column="Total Document"):
    try:
        if matrix is None or matrix.empty:
            return 0
        if "Status" not in matrix.columns or column not in matrix.columns:
            return 0

        matched = matrix.loc[matrix["Status"] == status, column]
        if matched.empty:
            return 0

        return safe_int(matched.iloc[0])
    except Exception:
        return 0


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
            if key in ["STATUS", "INFO", "INFORMATION", "VERSION"]:
                score += 3
            if "NAME" in key or "DESCRIPTION" in key:
                score += 2

        if score > best_score:
            best_score = score
            best_row = row
            best_headers = headers

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
        if any(word in value for word in ["TTI TO CB", "REVISED", "RESUBMIT", "NA", "ANSWERED"]):
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
        if text and "DETH-NSC" not in text:
            return value

    return ""


def read_tracking_all(tracking_file):
    wb = load_workbook(tracking_file, data_only=True)
    rows = []

    for sheet in TRACKING_SHEETS:
        if sheet not in wb.sheetnames:
            continue

        ws = wb[sheet]
        header_row, headers = find_header_row_and_columns(ws)

        doc_no_col = get_col(headers, ["Document No.", "Document No", "Ref No.", "Ref No", "Drawing ref No."])
        doc_name_col = get_col(headers, ["Document Name", "Equipment Name", "Description"])
        category_col = get_col(headers, ["Category", "Document Category"])
        status_col = get_col(headers, ["Status"])
        info_col = get_col(headers, ["Info", "Information"])

        seen = set()

        for row in range(header_row + 1, ws.max_row + 1):
            doc_no = ws.cell(row=row, column=doc_no_col).value if doc_no_col else None
            doc_no_source_col = doc_no_col

            if not doc_no or "DETH-NSC" not in str(doc_no):
                doc_no, doc_no_source_col = find_doc_no_in_row(ws, row)

            if not doc_no or "DETH-NSC" not in str(doc_no):
                continue

            base_key = base_doc_no(doc_no)
            if (sheet, base_key) in seen:
                continue
            seen.add((sheet, base_key))

            category_value = ws.cell(row=row, column=category_col).value if category_col else ""
            category = extract_category(doc_no, category_value)

            doc_name = guess_doc_name(ws, row, doc_no_source_col or 1, doc_name_col)
            tracking_status = guess_status_from_row(ws, row, status_col)
            info = guess_info_from_row(ws, row, info_col)

            rows.append({
                "Tracking Sheet": sheet,
                "Document No": doc_no,
                "Base Document No": base_key,
                "Document Name": doc_name,
                "Document Category": category,
                "Tracking Status": tracking_status,
                "Info": info,
                "Tracking State": normalize_tracking_state(tracking_status, info),
            })

    return pd.DataFrame(rows)


def read_takenaka(takenaka_file):
    wb = load_workbook(takenaka_file, data_only=True)
    data = {}

    for sheet in TAKENAKA_SHEETS:
        if sheet not in wb.sheetnames:
            continue

        ws = wb[sheet]

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
    dashboard_matrix = read_dashboard_sheet(tracking_file)
    tracking_df = read_tracking_all(tracking_file)
    takenaka_map = read_takenaka(takenaka_file)

    total_docs = get_dashboard_value(dashboard_matrix, "Total Document")
    open_docs = get_dashboard_value(dashboard_matrix, "Open")
    progress_docs = get_dashboard_value(dashboard_matrix, "on progress")
    focus_docs = open_docs + progress_docs

    if total_docs == 0 and not tracking_df.empty:
        total_docs = int(len(tracking_df))

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

    if not tracking_df.empty:
        for _, item in tracking_df.iterrows():
            if not should_include_tracking(item["Tracking Status"], item["Info"]):
                continue

            source = takenaka_map.get(item["Base Document No"])
            action = classify_action(source)
            checked_time = datetime.now().strftime("%d-%b-%Y %H:%M:%S")

            if source:
                row_data = [
                    item["Tracking Sheet"], item["Document No"], item["Document Name"], item["Document Category"],
                    item["Tracking Status"], item["Info"],
                    source["Takenaka Sheet"], source["Takenaka Doc No"], source["Takenaka Status 1"],
                    source["Takenaka Status 2"], source["Takenaka Status 3"], action, checked_time
                ]
            else:
                row_data = [
                    item["Tracking Sheet"], item["Document No"], item["Document Name"], item["Document Category"],
                    item["Tracking Status"], item["Info"], "", "", "", "", "", action, checked_time
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

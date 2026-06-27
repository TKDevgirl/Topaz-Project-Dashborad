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


def read_dashboard_sheet(tracking_file):
    """
    Read summary table directly from Tracking_document.xlsx > Dashboard sheet.
    This prevents mismatch between web summary and the source Excel Dashboard.
    """
    wb = load_workbook(tracking_file, data_only=True)

    if DASHBOARD_SHEET not in wb.sheetnames:
        return None

    ws = wb[DASHBOARD_SHEET]

    # Find the table area by locating headers: Total Document, MAT, MCR, MTS, CVI
    header_row = None
    header_cols = {}

    for row in range(1, min(ws.max_row, 30) + 1):
        temp = {}
        for col in range(1, ws.max_column + 1):
            value = norm_text(ws.cell(row=row, column=col).value)
            if value:
                temp[norm_upper(value)] = col

        if "TOTAL DOCUMENT" in temp and "MAT" in temp and "MCR" in temp:
            header_row = row
            header_cols = temp
            break

    if not header_row:
        return None

    row_labels = {
        "OPEN": "Open",
        "ON PROGRESS": "on progress",
        "SEND TO TTI": "on progress",
        "APPROVED": "Approved",
        "TOTAL DOCUMENT": "Total Document",
    }

    data = []
    for row in range(header_row + 1, min(ws.max_row, header_row + 10) + 1):
        label = norm_upper(ws.cell(row=row, column=header_cols["TOTAL DOCUMENT"] - 1).value)

        if not label:
            # Some files have row label in col C/D before Total Document value.
            for col in range(1, header_cols["TOTAL DOCUMENT"]):
                possible = norm_upper(ws.cell(row=row, column=col).value)
                if possible in row_labels:
                    label = possible
                    break

        if label not in row_labels:
            continue

        display_label = row_labels[label]
        item = {"Status": display_label}

        for col_name in ["TOTAL DOCUMENT", "MAT", "MCR", "MTS", "CVI"]:
            if col_name in header_cols:
                item[col_name.title() if col_name == "TOTAL DOCUMENT" else col_name] = ws.cell(
                    row=row, column=header_cols[col_name]
                ).value or 0

        data.append(item)

    if not data:
        return None

    df = pd.DataFrame(data)

    # Rename exactly for clean web display
    df = df.rename(columns={"Total Document": "Total Document"})

    # Ensure display order from user screenshot
    order = ["Total Document", "Open", "on progress", "Approved"]
    df["__order"] = df["Status"].apply(lambda x: order.index(x) if x in order else 99)
    df = df.sort_values("__order").drop(columns=["__order"])

    return df


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

    # Total docs should match Dashboard sheet if available
    if dashboard_matrix is not None and not dashboard_matrix.empty:
        total_docs = int(dashboard_matrix.loc[dashboard_matrix["Status"] == "Total Document", "Total Document"].iloc[0])
        focus_docs = int(
            dashboard_matrix.loc[dashboard_matrix["Status"].isin(["Open", "on progress"]), "Total Document"].sum()
        )
    else:
        total_docs = int(len(tracking_df))
        focus_docs = int(tracking_df["Tracking State"].isin(["Open", "on progress"]).sum()) if not tracking_df.empty else 0

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
        "total_docs": total_docs,
        "focus_docs": focus_docs,
        "df": df,
        "tracking_df": tracking_df,
        "dashboard_matrix": dashboard_matrix,
        "action_counts": action_counts,
        "last_updated": datetime.now().strftime("%d-%b-%Y %H:%M:%S"),
    }

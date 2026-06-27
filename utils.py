import re


def norm_text(value):
    return str(value or "").strip()


def norm_upper(value):
    return norm_text(value).upper()


def safe_int(value):
    try:
        if value is None or value == "":
            return 0
        return int(float(value))
    except Exception:
        return 0


def base_doc_no(doc_no):
    doc_no = norm_text(doc_no)
    parts = doc_no.split("-")
    if parts and parts[-1].isdigit() and len(parts[-1]) == 2:
        return "-".join(parts[:-1])
    return doc_no


def doc_no_key(doc_no):
    """
    Main matching key:
    - uppercase
    - remove final -00/-01 version
    - remove spaces/newlines
    """
    key = base_doc_no(doc_no).upper()
    key = re.sub(r"\s+", "", key)
    return key


def normalize_header(value):
    return str(value or "").replace("\n", " ").replace("\r", " ").strip().upper()


def contains_doc_no(value):
    return "DETH-NSC" in norm_upper(value)


def extract_category(doc_no, fallback=""):
    text = norm_upper(doc_no)
    fallback_text = norm_upper(fallback)

    for category in ["MAT", "MCR", "MTS", "CVI", "DWG", "SCH"]:
        if f"-{category}-" in text:
            return category
        if category == fallback_text:
            return category

    return fallback_text or "OTHER"

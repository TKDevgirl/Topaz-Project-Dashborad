def norm_text(value):
    return str(value or "").strip()


def norm_upper(value):
    return norm_text(value).upper()


def base_doc_no(doc_no):
    doc_no = norm_text(doc_no)
    parts = doc_no.split("-")

    if parts and parts[-1].isdigit() and len(parts[-1]) == 2:
        return "-".join(parts[:-1])

    return doc_no


def normalize_header(value):
    return (
        str(value or "")
        .replace("\n", " ")
        .replace("\r", " ")
        .strip()
        .upper()
    )


def contains_doc_no(value):
    return "DETH-NSC" in norm_upper(value)

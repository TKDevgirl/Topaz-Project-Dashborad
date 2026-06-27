from utils import norm_upper


def normalize_tracking_state(status, info):
    status_text = norm_upper(status)
    info_text = norm_upper(info)

    if "CLOSE" in status_text or "APPROVED" in status_text or "APPROVE" in status_text:
        return "Approved"

    if "ON PROGRESS" in status_text or "ON PROCESS" in status_text:
        return "on progress"

    if "ON PROGRESS" in info_text or "ON PROCESS" in info_text:
        return "on progress"

    if "OPEN" in status_text:
        return "Open"

    if "TTI TO CB" in info_text:
        return "Open"

    return "Other"


def should_include_tracking(status, info):
    return normalize_tracking_state(status, info) in ["Open", "on progress"]


def classify_action(takenaka_record):
    if not takenaka_record:
        return "NOT FOUND IN TAKENAKA SOURCE"

    s1 = norm_upper(takenaka_record.get("Takenaka Status 1"))
    s2 = norm_upper(takenaka_record.get("Takenaka Status 2"))
    s3 = norm_upper(takenaka_record.get("Takenaka Status 3"))

    if s1 in ["CLOSED", "CLOSE"]:
        return "UPDATE TRACKING TO CLOSED"

    if "RETURNED" in s1 or "RETURNED" in s2:
        return "RETURNED BY NV5 / NEED RESUBMIT"

    if "OVERDUE" in s3:
        return "OVERDUE / FOLLOW UP"

    if "OPEN" in s1 and ("ON PROCESS" in s2 or "ON PROGRESS" in s2):
        return "OPEN & ON PROCESS"

    if "OPEN" in s1:
        return "OPEN"

    return "CHECK"

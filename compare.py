from utils import norm_upper


def should_include_tracking(status, info):
    status_text = norm_upper(status)
    info_text = norm_upper(info)

    return (
        "OPEN" in status_text
        or "ON PROGRESS" in status_text
        or "ON PROCESS" in status_text
        or "ON PROGRESS" in info_text
        or "ON PROCESS" in info_text
        or "TTI TO CB" in info_text
    )


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

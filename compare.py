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
    """
    V5.1 mapping logic. No guessing from UI.
    Action is decided only from Takenaka Status 1/2/3 after document number match.
    """
    if not takenaka_record:
        return "NOT FOUND IN TAKENAKA SOURCE", "No matching document number found in Takenaka source"

    s1 = norm_upper(takenaka_record.get("Takenaka Status 1"))
    s2 = norm_upper(takenaka_record.get("Takenaka Status 2"))
    s3 = norm_upper(takenaka_record.get("Takenaka Status 3"))
    combined = " | ".join([s1, s2, s3])

    if "RETURNED" in combined or "RESUBMIT" in combined or "REVISE" in combined:
        return "RETURNED BY NV5 / NEED RESUBMIT", "Takenaka status indicates returned / revise and resubmit"

    if "OVERDUE" in combined:
        return "OVERDUE / FOLLOW UP", "Takenaka status indicates overdue"

    if "CLOSED" in s1 or s1 == "CLOSE":
        return "UPDATE TRACKING TO CLOSED", "Takenaka Status 1 is closed"

    if "OPEN" in s1 and ("ON PROCESS" in s2 or "ON PROGRESS" in s2):
        return "OPEN & ON PROCESS", "Takenaka Status 1 is open and Status 2 is on process"

    if "OPEN" in s1:
        return "OPEN", "Takenaka Status 1 is open"

    if not s1 and not s2 and not s3:
        return "CHECK", "Document matched but Takenaka status is blank"

    return "CHECK", "Document matched but status combination is not mapped"

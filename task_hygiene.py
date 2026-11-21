"""Task hygiene rules: status closure, no-target handling, duplicate merge (dry run first).

Pure functions only. The app decides when to call them; nothing here writes to the DB.
"""
import re
from collections import defaultdict

CLOSING_STATUSES = {"החלטה", "בוצע"}          # close the task
ONE_PROTOCOL_STATUS = "לידיעה"               # closes after one protocol unless it has a follow-up
NEVER_AUTO_CLOSE = {"שוטף", "מעקב"}          # stay open until a manual close
MAINTENANCE_STATUS = "מעקב"                  # status for an unclosed task without a target
APPROVERS = {"uploader", "admin"}            # who may approve closing a מעקב task
TENANT_CATEGORY = "שינוי דיירים"
TENANT_CATEGORY_ALIASES = {"שינוי דיירים", "שינויי דיירים"}

_STOP = {"של", "את", "על", "עם", "כי", "או", "גם", "זה", "זו", "יש", "לא", "כל", "אל", "in"}


def normalize_topic(topic):
    """Topic label used in the merge key; tenant-change spellings are unified."""
    t = (topic or "").strip()
    return TENANT_CATEGORY if t in TENANT_CATEGORY_ALIASES else t


def normalize_text(text):
    """Lower-case word set for similarity checks; strips punctuation and stop words."""
    words = re.findall(r"[\w֐-׿]+", (text or "").lower())
    return {w for w in words if w not in _STOP and len(w) > 1}


def similarity(a, b):
    wa, wb = normalize_text(a), normalize_text(b)
    if not wa or not wb:
        return 0.0
    return len(wa & wb) / len(wa | wb)


def merge_key(task):
    """Internal identity: sorted building tags + normalized topic + responsible."""
    buildings = tuple(sorted(task.get("buildings") or []))
    responsible = (task.get("responsible") or "").strip()
    return (buildings, normalize_topic(task.get("topic")), responsible)


def can_close(status, has_followup=False, approved_by=None):
    """True if the status allows the task to be closed by the system."""
    if status in CLOSING_STATUSES:
        return True
    if status == ONE_PROTOCOL_STATUS:
        return not has_followup
    if status == MAINTENANCE_STATUS:
        return approved_by in APPROVERS
    return False  # שוטף and anything unknown stay open


def status_for_unclosed_without_target(task):
    """Unclosed task with no target gets the מעקב status, keeping its last date."""
    if task.get("status") in CLOSING_STATUSES:
        return task["status"]
    if not (task.get("target") or "").strip():
        return MAINTENANCE_STATUS
    return task.get("status")


def find_merge_groups(tasks, threshold=0.40):
    """Group open tasks that share a merge key and are similar in text.

    Returns a list of groups (lists of task ids). Groups of one are not returned.
    Dry run only: the caller shows these groups before anything is merged.
    """
    buckets = defaultdict(list)
    for t in tasks:
        if t.get("status") in CLOSING_STATUSES:
            continue
        buckets[merge_key(t)].append(t)

    groups = []
    for members in buckets.values():
        if len(members) < 2:
            continue
        # union-find on similarity within the bucket
        parent = {m["id"]: m["id"] for m in members}

        def find(x):
            while parent[x] != x:
                parent[x] = parent[parent[x]]
                x = parent[x]
            return x

        for i, a in enumerate(members):
            for b in members[i + 1:]:
                if similarity(a.get("text"), b.get("text")) >= threshold:
                    parent[find(a["id"])] = find(b["id"])

        clusters = defaultdict(list)
        for m in members:
            clusters[find(m["id"])].append(m["id"])
        groups.extend(ids for ids in clusters.values() if len(ids) > 1)
    return groups


def merge_record(group_tasks):
    """Build one merged record from a group. Keeps the first date and sets the last update date."""
    ordered = sorted(group_tasks, key=lambda t: t.get("first_date") or "")
    first = ordered[0]
    last = ordered[-1]
    return {
        "id": first["id"],
        "merged_ids": [t["id"] for t in ordered[1:]],
        "first_date": first.get("first_date"),
        "last_update": last.get("last_date") or last.get("first_date"),
        "text": last.get("text") or first.get("text"),
        "status": last.get("status") or first.get("status"),
        "buildings": sorted(set(sum((t.get("buildings") or [] for t in ordered), []))),
    }

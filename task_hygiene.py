"""Task hygiene rules: status closure, no-target handling, duplicate merge (dry run first).

Pure functions only. The app decides when to call them; nothing here writes to the DB.
"""
import re
from collections import Counter, defaultdict

CLOSING_STATUSES = {"החלטה", "בוצע"}          # close the task
ONE_PROTOCOL_STATUS = "לידיעה"               # closes after one protocol unless it has a follow-up
NEVER_AUTO_CLOSE = {"שוטף", "מעקב"}          # stay open until a manual close
MAINTENANCE_STATUS = "מעקב"                  # status for an unclosed task without a target
APPROVERS = {"uploader", "admin"}            # who may approve closing a מעקב task
TENANT_CATEGORY = "שינוי דיירים"
TENANT_CATEGORY_ALIASES = {"שינוי דיירים", "שינויי דיירים"}

_STOP = {"של", "את", "על", "עם", "כי", "או", "גם", "זה", "זו", "יש", "כל", "אל", "in"}
MIN_EXACT_WORDS = 4  # identical repeats without topic/responsible merge only when the text has this many words
NEGATIONS = {"לא", "אין", "ללא"}  # a negation changes the meaning: "מאושרים" and "לא מאושרים" never merge


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


def _negated(text):
    """True if the text contains a negation word (לא / אין / ללא)."""
    return bool(set(re.findall(r"[\w֐-׿]+", (text or "").lower())) & NEGATIONS)


def shared_keyword(text_a, text_b, saved_kw):
    """True if one chosen keyword appears in both texts."""
    if not saved_kw:
        return False
    wa, wb = normalize_text(text_a), normalize_text(text_b)
    return any(k in wa and k in wb for k in saved_kw)


def containment(a, b):
    """Share of the shorter text's words found in the longer one. Catches a short protocol
    entry that is the start of a longer one (Jaccard alone would miss it)."""
    wa, wb = normalize_text(a), normalize_text(b)
    if not wa or not wb:
        return 0.0
    return len(wa & wb) / min(len(wa), len(wb))


def is_tenant(task):
    """Tenant-change item by topic or text, in any spelling. Never merged across apartments."""
    topic = (task.get("topic") or "")
    text = (task.get("text") or "")
    return normalize_topic(topic) == TENANT_CATEGORY or "שינוי דיירים" in topic + " " + text or "שינויי דיירים" in topic + " " + text


def exact_text(text):
    """Text for an exact-repeat check: case, spacing and punctuation ignored. Empty if no words."""
    return " ".join(re.findall(r"[\w֐-׿]+", (text or "").lower()))


def identical_group(members):
    """True if every task in the group has the same exact text (the rule for repeats).
    Such a group merges even without a topic or responsible, but never across tenant changes."""
    texts = {exact_text(m.get("text")) for m in members}
    topics = {normalize_topic(m.get("topic")) for m in members}
    if len(texts) != 1 or "" in texts or TENANT_CATEGORY in topics:
        return False
    # a short generic label ("תכנון", "אדריכלות", "חריגים מאושרים") is not a task text: a repeat
    # of it may stand for different items, so it needs at least MIN_EXACT_WORDS words
    return len(texts.pop().split()) >= MIN_EXACT_WORDS


def keywords(texts, top=8):
    """Rough keyword tags: the words that appear most often across the group's texts.
    Automatic and approximate; the supervisor can edit them."""
    counts = Counter()
    for text in texts:
        counts.update(normalize_text(text))
    return [w for w, _ in counts.most_common(top)]


def merge_key(task):
    """Internal identity: sorted building tags + normalized topic + responsible."""
    buildings = tuple(sorted(task.get("buildings") or []))
    # responsible compared by its words, so spacing, order or punctuation do not split one task
    responsible = frozenset(normalize_text(task.get("responsible")))
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


def find_merge_groups(tasks, threshold=0.40, saved_keywords=None, bonus=0.10):
    """Group open tasks that share a merge key and are similar in text.

    Returns a list of groups (lists of task ids). Groups of one are not returned.
    Dry run only: the caller shows these groups before anything is merged.
    """
    saved_kw = set()
    for words in (saved_keywords or {}).values():
        saved_kw |= set(words)
    # identical texts repeated across protocols merge whatever their topic or responsible spelling
    # (e.g. "אינסטלצי ה" vs "אינסטלציה"): count them first
    text_counts = Counter(exact_text(t.get("text")) for t in tasks
                          if t.get("status") not in CLOSING_STATUSES
                          and t.get("status") not in {ONE_PROTOCOL_STATUS, "בוטלה"}
                          and not is_tenant(t) and exact_text(t.get("text")))
    buckets = defaultdict(list)
    for t in tasks:
        if t.get("status") in CLOSING_STATUSES or t.get("status") in {ONE_PROTOCOL_STATUS, "בוטלה"}:
            continue  # לידיעה closes after one protocol; בוטלה is cancelled: neither is merged
        topic = normalize_topic(t.get("topic"))
        if is_tenant(t):
            continue  # tenant-change items stay separate, always (any apartment, any spelling)
        exact = exact_text(t.get("text"))
        if exact and text_counts[exact] > 1:
            buckets[("exact", exact)].append(t)  # identical repeat: one bucket, topic and responsible ignored
            continue
        if not topic or not (t.get("responsible") or "").strip():
            continue  # no topic or no responsible: only identical repeats (above) may merge
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
                if _negated(a.get("text")) != _negated(b.get("text")):
                    continue
                score = max(similarity(a.get("text"), b.get("text")),
                            containment(a.get("text"), b.get("text")))
                # keywords are a complement, never a gate: a shared chosen keyword adds a bonus
                if shared_keyword(a.get("text"), b.get("text"), saved_kw):
                    score += bonus
                if score >= threshold:
                    parent[find(a["id"])] = find(b["id"])

        clusters = defaultdict(list)
        for m in members:
            clusters[find(m["id"])].append(m["id"])
        groups.extend(ids for ids in clusters.values() if len(ids) > 1)
    return groups


def subject_groups(tasks, term):
    """Subject rule (section 7): open tasks whose text contains the subject term, grouped by building.
    Tenant-change items are excluded. Returns groups of task ids with two or more members.
    Dry run only: nothing is merged here. Texts are not required to match."""
    term_words = normalize_text(term)
    if not term_words:
        return []
    buckets = defaultdict(list)
    for t in tasks:
        if t.get("status") in CLOSING_STATUSES or t.get("status") in {ONE_PROTOCOL_STATUS, "בוטלה"}:
            continue
        if is_tenant(t):
            continue
        if not term_words <= normalize_text(t.get("text")):
            continue
        key = tuple(sorted(t.get("buildings") or []))
        buckets[key].append(t["id"])
    return [ids for ids in buckets.values() if len(ids) > 1]


# a sub-item number: "1.", "2 .", ".2 " or ". 2 " (no dot). A date (12.04.26) or 4.1 does not
# match, because the number must be followed by a space, or by a dot and a space.
_NUMBERED = re.compile(r"(?<!\d)(\d{1,2})\s*\.?\s+")


def split_numbered(text):
    """Split a long task text into its numbered sub-items (1. 2. 3. ...), in order.
    Only numbers that continue the sequence count, so dates and drawing numbers are skipped.
    Returns None when there are fewer than two sub-items. Dry run: nothing is written."""
    text = text or ""
    found, expect = [], 1
    for m in _NUMBERED.finditer(text):
        if int(m.group(1)) == expect:
            found.append(m)
            expect += 1
    if len(found) < 2:
        return None
    head = text[:found[0].start()].strip()
    items = []
    for i, m in enumerate(found):
        end = found[i + 1].start() if i + 1 < len(found) else len(text)
        items.append((int(m.group(1)), text[m.end():end].strip()))
    return {"head": head, "items": items}


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
        "keywords": keywords([t.get("text") or "" for t in ordered]),
    }

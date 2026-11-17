"""Extract task rows from a Hebrew execution/weekly protocol PDF.

The PDFs store Hebrew in visual order, so each cell is reversed back to
logical order while ASCII/digit runs (dates, numbers, Latin words) keep
their own direction.
"""
import hashlib
import io
import os
import re

import pdfplumber

_ASCII_RUN = re.compile(r"[0-9A-Za-z][0-9A-Za-z./:%'\"\-]*")
_MIRROR = str.maketrans({"(": ")", ")": "(", "[": "]", "]": "["})


def _fix_line(line):
    if not re.search(r"[֐-׿]", line):
        return line.strip()
    rev = line[::-1].translate(_MIRROR)
    # restore LTR runs (dates, numbers, Latin) that were reversed
    return _ASCII_RUN.sub(lambda m: m.group(0)[::-1], rev).strip()


def fix_visual(text):
    """pdfplumber returns each cell's characters in visual order. Lines inside a cell
    are already top-to-bottom, so each line is reversed on its own and the line order kept."""
    if not text:
        return ""
    lines = [_fix_line(l) for l in text.split("\n")]
    return re.sub(r"\s+", " ", " ".join(l for l in lines if l)).strip()


TARGET_TYPES = {"שוטף", "מיידי", "למעקב", "מעקב", "לידיעה", "החלטה", "החל מתאריך", "הודעה"}
_DATE_ONLY = re.compile(r"^(\d{1,2}[./]\d{1,2}[./]\d{2,4})(\s+\d{1,2}[./]\d{1,2}[./]\d{2,4})*(\s+(מיידי|שוטף|למעקב|לידיעה|החלטה|הודעה))?$")
_UNTIL = re.compile(r"^עד\s")


def subject_of(task):
    """The subject ('מרתפים', 'בטון חשוף') is written before the first colon of the task."""
    head = task.split(":", 1)[0].strip()
    return head if ":" in task and 0 < len(head) <= 30 else ""


def flags_of(row):
    flags = []
    tgt = row["target"].strip()
    if tgt and not (_DATE_ONLY.match(tgt) or tgt in TARGET_TYPES or _UNTIL.match(tgt)):
        flags.append("יעד לא מזוהה")
    if not row["task"]:
        flags.append("משימה ריקה")
    return flags


def _is_header(cells):
    joined = " ".join(c or "" for c in cells)
    return "#" in joined and "משימה" in fix_visual(joined)


_MEETING_LINE = re.compile(r"(\d{2}\.\d{2}\.\d{2})\s*:\s*\S*\s*\S*\s*\S*\s*\S*\s*\S*\s*\S*")


def meeting_date_of_page(page):
    """Each summary starts with its date: 'סיכום פגישת ביצוע dd.mm.yy' (dots) on one kind
    of report, 'תאריך הפגישה: dd/mm/yy' (slashes) on weekly reports. Returns dd.mm.yy."""
    for line in (page.extract_text() or "").splitlines():
        if "השיגפה ךיראת" in line:  # visual order of 'תאריך הפגישה:'
            m = re.search(r"(\d{2})/(\d{2})/(\d{2})", line)
            if m:
                return f"{m.group(1)}.{m.group(2)}.{m.group(3)}"
        if "סיכום פגישת" in fix_visual(line):
            m = re.search(r"(\d{2}\.\d{2}\.\d{2})", line)
            if m:
                return m.group(1)
    return None


def _row_for_box(placed, box):
    """Index of the protocol row an image belongs to: the row that contains it, else the nearest one."""
    cx, cy = (box["x0"] + box["x1"]) / 2, (box["top"] + box["bottom"]) / 2
    best = None
    for idx, bbox in placed:
        if bbox is None:
            continue
        x0, top, x1, bottom = bbox
        if x0 <= cx <= x1 and top <= cy <= bottom:
            return idx
        dist = min(abs(cy - top), abs(cy - bottom))
        if best is None or dist < best[0]:
            best = (dist, idx)
    return best[1] if best else None


def _save_crop(page, box, image_dir):
    """Render one image area of the page to a PNG file named by its content; returns the file name."""
    bbox = (max(0, box["x0"]), max(0, box["top"]), min(page.width, box["x1"]), min(page.height, box["bottom"]))
    pil = page.crop(bbox).to_image(resolution=120).original
    buf = io.BytesIO()
    pil.save(buf, "PNG")
    data = buf.getvalue()
    name = hashlib.sha1(data).hexdigest()[:16] + ".png"
    os.makedirs(image_dir, exist_ok=True)
    path = os.path.join(image_dir, name)
    if not os.path.exists(path):
        with open(path, "wb") as fh:
            fh.write(data)
    return name


def extract_rows(pdf_path, image_dir=None):
    """Return list of dicts: number, topic, task, assignee, update, target, meeting_date.
    With image_dir, each row also gets 'images': file names of the pictures placed in its area."""
    rows = []
    pictures = []  # (row index, file name, page number)
    with pdfplumber.open(pdf_path) as pdf:
        current_date = None
        for page in pdf.pages:
            current_date = meeting_date_of_page(page) or current_date
            placed = []  # (row index, bbox of the row on this page)
            for table in page.find_tables():
                data = table.extract() or []
                for i, cells in enumerate(data):
                    if not cells or _is_header(cells):
                        continue
                    if len(cells) < 6:
                        continue
                    target, update, assignee, task, topic, number = [fix_visual(c) for c in cells[:6]]
                    if not task and re.match(r"^\d+(\.\d+)?$", number or ""):
                        continue  # sub-heading row without wording (e.g. a topic label)
                    if not re.match(r"^\d+(\.\d+)?$", number or ""):
                        # continuation row or noise: attach to previous row
                        if rows and task:
                            rows[-1]["task"] += " " + task
                        continue
                    row = {
                        "number": number, "topic": topic, "task": task,
                        "assignee": assignee, "update": update, "target": target,
                    }
                    row["meeting_date"] = current_date
                    row["subject"] = subject_of(task)
                    row["flags"] = flags_of(row)
                    rows.append(row)
                    bbox = table.rows[i].bbox if i < len(table.rows) else None
                    placed.append((len(rows) - 1, bbox))
            if image_dir and placed:
                for box in page.images:
                    if box["x1"] - box["x0"] < 60 or box["bottom"] - box["top"] < 60:
                        continue  # logos, stamps and icons are not protocol pictures
                    if box["top"] < 80:
                        continue  # page header
                    idx = _row_for_box(placed, box)
                    if idx is None:
                        continue
                    pictures.append((idx, _save_crop(page, box, image_dir), page.page_number))
    # a picture repeated on several pages is a logo or a template picture, not content
    pages_of = {}
    for _, name, pno in pictures:
        pages_of.setdefault(name, set()).add(pno)
    for idx, name, _ in pictures:
        if len(pages_of[name]) <= 2:
            rows[idx].setdefault("images", [])
            if name not in rows[idx]["images"]:
                rows[idx]["images"].append(name)
    for r in rows:
        r.setdefault("images", [])
    return rows


if __name__ == "__main__":
    import json, sys
    out = extract_rows(sys.argv[1])
    print(json.dumps(out, ensure_ascii=False, indent=1)[:4000])
    print("rows:", len(out))


# header and footer tokens that share a line with item text
_NOISE = [
    re.compile(r"יום ב׳\s*\d{2}\s*נוב׳\s*[\d\s]+"),
    re.compile(r"\b(Title|Email|Created|Contact|Phone|Location|Doc|Id|No|Items|page|of|Other|reports)\b\s*[:.]?"),
    re.compile(r"[.:]?\s*\b60\.1\b"),
    re.compile(r"\b\d{9,10}\b"),
    re.compile(r"\b\d{1,2}\s+\d{1,2}\s+\d{4}\b|\b\d{1,2}\s+\d{4}\b"),  # header date fragments
    re.compile(r"[\w.]+@[\w.]+"),
    re.compile(r"זאב אלון|הלל הזקן - מרתפים נובמבר \d+"),
    re.compile(r"מושלין סימון אדריכלות וניהול"),
]
_MARK = re.compile(r"^\(\s*(\d+)\s*\)$")
_PHOTO = re.compile(r"^\d+\.jpe?g$", re.I)


def _strip_edges(text):
    """Drop number-only fragments left by the page header at the start and end of an item."""
    text = re.sub(r"[\u0000\u200e\u200f\u202a-\u202e\u2066-\u2069]", "", text)  # bidi and NUL marks
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"^(?:\d{1,4}\s+)+(?=\D)", "", text)
    text = re.sub(r"\s+(?:\d{1,4}\s*)+$", "", text)
    return text


def extract_supervision_items(pdf_path):
    """Supervision reports are text, not tables: each item is text, then '(n)', then a photo name."""
    items, buf = [], []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for raw in (page.extract_text() or "").splitlines():
                line = fix_visual(raw)
                for rx in _NOISE:
                    line = rx.sub(" ", line)
                line = re.sub(r"\s+", " ", line).strip(" .:") if line.strip() else ""
                photo = fix_visual(raw) if _PHOTO.match(fix_visual(raw)) else ""
                if photo:
                    if items:
                        items[-1]["photo"] = photo
                    continue
                if not line:
                    continue
                m = _MARK.match(line)
                if m:
                    text = _strip_edges(" ".join(buf))
                    items.append({"item_no": m.group(1), "text": text, "photo": ""})
                    buf = []
                    continue
                buf.append(line)
    return items

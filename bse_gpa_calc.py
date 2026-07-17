#!/usr/bin/env python3
"""
BSEHons (OUSL) - RC -> RC'25 Course Conversion & GPA Calculator (GENERIC)
===========================================================================

Person-agnostic version - contains NO hardcoded transcript for any specific
student. Your results come in through the window that opens when you run
this script, either by:

  (a) Importing your Marks Sheet export from the MyOUSL portal. This works
      whether the file is:
        - an ".xls" that's actually an HTML table (the most common MyOUSL
          export - handled with zero extra dependencies), or
        - a real Excel file, .xlsx or legacy binary .xls, or
        - a PDF export of the same table.
      The file type is auto-detected from its actual content, not just its
      extension, so a mislabeled file still works.
  (b) Manually adding/editing course code + grade rows in the window.
  (c) A mix of both.

Every student's elective choices can differ - nothing here assumes a fixed
set of courses. Category/Level/Credit for every course is derived directly
from the course code itself (Dept[2]+Category[1]+Level[1]+Credit[1]+Serial),
and compulsory/elective classification comes from the official RC'25 course
lists below, so it works for whatever combination of compulsory and
elective courses were actually taken.

HANDLING REPEATS / RESITS (IMPORTANT)
----------------------------------------
A Marks Sheet lists every attempt at every course. For each course code,
this script looks at ALL attempts (Pass, Repeat, Resit, Pending, Eligible -
whatever the status column says) and picks the one from the MOST RECENT
"Last Offered Year" that actually has a grade recorded (i.e. not blank/"-").
That grade is then checked against the real list of passing grades
(A+, A, A-, B+, B, B-, C+, C - per the Regulation; C-, D+, D, E, FA, RX etc.
are NOT passing grades). If it's a genuine pass, that's the result used. If
it isn't (or no attempt has a grade yet), the course is treated as not yet
completed - which is exactly right, since once someone resits and passes,
the NEWEST dated row will show the passing grade and this logic picks it up
automatically, without caring what label the "status" column used.

MULTIPLE STUDENTS / STARTING FRESH
-------------------------------------
Progress is saved to a small JSON file next to this script
(BSE_STATE_FILE, default "bse_progress_state.json"). Use the "Clear (New
Student)" button in the window to wipe the current entries and start over
for someone else, or just point BSE_STATE_FILE at a different filename per
person if you'd rather keep several students' data side by side.
"""

from dataclasses import dataclass
from collections import defaultdict
from html.parser import HTMLParser
import json
import os

# Change this if you want a different/unique save file for this person
# (or use the "Clear (New Student)" button to reset it for someone new).
BSE_STATE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bse_progress_state.json")


# ---------------------------------------------------------------------------
# 1. GRADE -> GRADE POINT VALUE (GPV) SCALE
#    (confirmed against the official Schedule 2 table)
# ---------------------------------------------------------------------------
GRADE_POINTS = {
    "A+": 4.0, "A": 4.0, "A-": 3.7,
    "B+": 3.3, "B": 3.0, "B-": 2.7,
    "C+": 2.3, "C": 2.0, "C-": 1.7,
    "D+": 1.3, "D": 1.0, "E": 0.0,
}

# Per the Regulation, only these grades "constitute Pass grades". C-, D+, D,
# E (and fail markers like FA/RX) do NOT - a course showing one of those is
# still not completed, no matter what the status column says.
PASSING_GRADES = {"A+", "A", "A-", "B+", "B", "B-", "C+", "C"}
# Non-graded pass/fail courses (e.g. FDE3023 EfIL) show "P" instead of a
# letter grade - treat that as a pass too.
PASSING_MARKERS = PASSING_GRADES | {"P", "PASS"}

# ---------------------------------------------------------------------------
# 2. OLD (RC) -> NEW (RC'25) CONVERSION MAP
#    old_code -> [new_code, ...]. Single target = one-to-one (grade carries
#    over unchanged). Multiple targets = one-to-many: the SAME grade is
#    applied to every target (e.g. one English course splitting into two,
#    or one 5-credit course splitting into two real courses + a virtual
#    credit-compensation course). "VNxxxx" targets are virtual
#    credit-compensation courses - they don't correspond to a real taught
#    course, they just top up a category's credit total.
# ---------------------------------------------------------------------------
CONVERSION_MAP = {
    "AGM3263": ["EEL3263"],
    "EEI3266": ["EEI3366"],
    "EEX3373": ["EEI3273", "VNI3173"],
    "EEX3467": ["EEI3467"],
    "LTE3401": ["LTE34SI", "LTE34SE"],
    "LTE3407": ["LTE34SI", "LTE34SE"],
    "MHZ3459": ["MHZ4359", "VNZ3159"],
    "EEI4346": ["EEI3347"],
    "EEI4366": ["EEI5486"],
    "EEX4465": ["EEI4365", "VNI4165"],
    "EEY4189": ["EER4189"],
    "MHZ4256": ["MHZ3356"],
    "CVM5402": ["CVM4402"],
    "EEI5270": ["EEI4370"],
    "EEI5280": ["EEI6280"],
    "EEI5466": ["VNI5466"],
    "EEX5563": ["EEI5263", "EEI5265", "VNI5163"],
    "MHJ5372": ["MHJ5383"],
    "LLJ3265": ["LLM5281"],
    "EEY6189": ["EER6289"],
    "EEW5811": ["EEW5611"],
    "EEX5362": ["EEI6373"],
    "EEI4369": ["EEI5369"],
    "EEX4373": ["EEI5373"],
    "EEY4489": ["EER4489"],
    "MHJ4271": ["MHJ5282"],
    "EEX5376": ["EEI5376"],
    "EEX5464": ["EEI5364", "VNI5164"],
    "EEY6689": ["EER6689"],
    "EEX6340": ["EEI4360"],
    "EEX6377": ["EEI6377"],
    "EEX6278": ["EEI6378"],
}
# Codes with no entry above are assumed unchanged between RC and RC'25.

# ---------------------------------------------------------------------------
# 3. RC'25 COURSE CLASSIFICATION (compulsory / elective) - Schedule 1(a)
# ---------------------------------------------------------------------------
COMPULSORY_CODES = {
    "EEI3262", "EEI3269", "EEI3346", "EEI3347", "EEI3366", "EEI3372",
    "EEI3467", "EEI3273", "EEL3263", "LTE34SI", "MHZ3356",
    "AGM4367", "CVM4402", "EEI4267", "EEI4360", "EEI4361", "EEI4362",
    "EEI4365", "EEI4370", "EER4189", "MHZ4359", "MHZ4377",
    "EEI5263", "EEI5265", "EEI5364", "EEI5467", "EEI5486", "EEW5611",
    "LLM5281", "MHJ5383", "MHZ5375",
    "EEI6171", "EEI6360", "EEI6373", "EEI6567", "EEM6202", "EER6289",
    "EER6689",
}

ELECTIVE_CODES = {
    "EEM3366", "EER4489", "EEI5369", "EEI5373", "EEI5376", "MHJ5282",
    "EEI6279", "EEI6280", "EEI6320", "EEI6363", "EEI6366", "EEI6369",
    "EEI6377", "EEI6378",
}

# Non-GPA courses that ALSO don't count toward category credit totals.
ZERO_WEIGHT_NON_GPA = {"LTE34SE", "MHZ2250", "FDE3020", "FDE3023"}

# Non-GPA courses that DO still count toward category credit totals (just
# excluded from the GPA average itself).
COUNTS_BUT_NOT_GPA = {"EEW5611"}

VIRTUAL_PREFIX = "VN"

# Discontinued/legacy courses that no longer appear in the RC'25 lists but
# should still be counted toward BOTH credits and GPA. Add entries here as
# needed - value is 'compulsory' or 'elective' based on how the course was
# originally classified. (Courses at Level 3 here never affect GPA anyway,
# since the official algorithm only considers Levels 4-6 - this only
# matters in practice for Level 4-6 discontinued courses.)
LEGACY_RECLASSIFY = {
    "DMM6602": "compulsory",  # Management for Engineers (old compulsory Mgmt course)
    "CSI3361": "elective",
    "ISI3376": "elective",
}

# ---------------------------------------------------------------------------
# 4. DEGREE CREDIT REQUIREMENTS (Table 1, RC'25)
# ---------------------------------------------------------------------------
CATEGORY_GROUP = {
    "I": "Industrial / Engineering Sciences",
    "S": "Industrial / Engineering Sciences",
    "M": "Management",
    "J": "General / Humanities",
    "Z": "Mathematics",
    "R": "Project",
    "L": "Language",
    "E": "Language",
    "W": "Industrial Training",
}

TABLE1_REQUIREMENTS = {
    "Industrial / Engineering Sciences": (67, 78),
    "Management": (11, 20),
    "General / Humanities": (3, 6),
    "Mathematics": (12, 12),
    "Project": (9, 13),
    "Language": (6, 6),
    "Industrial Training": (6, 6),
}
TOTAL_MIN_CREDITS = 125
GPA_CREDIT_CAP = 72

# Degree classification bands - confirmed against Regulation Sections
# 6.8-6.11 (Pass >=2.00, Second Lower >=3.00, Second Upper >=3.30, First >=3.70).
CLASSIFICATION_BANDS = [
    (3.70, 4.00, "First Class"),
    (3.30, 3.6999, "Second Class (Upper Division)"),
    (3.00, 3.2999, "Second Class (Lower Division)"),
    (2.00, 2.9999, "Pass"),
    (0.00, 1.9999, "Below Pass - no classification"),
]


def classify_gpa(gpa):
    """Classifies GPA per Regulation Sections 6.8-6.11. Important: the
    Regulation states GPA is "calculated to the second decimal place", so
    classification must use the ROUNDED value - e.g. a raw GPA of 3.2956
    rounds to 3.30 and should be classified as Second Upper, not Second
    Lower, even though 3.2956 itself is technically under 3.30."""
    if gpa is None:
        return "N/A"
    rounded = round(gpa, 2)
    for low, high, label in CLASSIFICATION_BANDS:
        if low <= rounded <= high:
            return label
    return "Unclassified"


# ---------------------------------------------------------------------------
# 5. COURSE CODE PARSER (Dept[2]+Category[1]+Level[1]+Credit[1]+Serial)
# ---------------------------------------------------------------------------
@dataclass
class CourseMeta:
    code: str
    dept: str
    category: str
    level: int
    credit: int
    serial: str
    kind: str  # 'compulsory' | 'elective' | 'virtual' | 'nongpa_zero' | 'unmapped'
    counts_for_credit: bool
    counts_for_gpa: bool


def parse_course_code(code: str) -> CourseMeta:
    code = code.strip().upper()
    if len(code) < 5:
        raise ValueError(f"Course code '{code}' is too short to parse.")
    dept = code[0:2]
    category = code[2]
    try:
        level = int(code[3])
        credit = int(code[4])
    except ValueError:
        raise ValueError(f"Could not parse level/credit digits from '{code}'.")
    serial = code[5:]

    if dept == VIRTUAL_PREFIX:
        kind = "virtual"
        counts_for_credit, counts_for_gpa = True, False
    elif code in ZERO_WEIGHT_NON_GPA:
        kind = "nongpa_zero"
        counts_for_credit, counts_for_gpa = False, False
    elif code in COUNTS_BUT_NOT_GPA:
        kind = "compulsory" if code in COMPULSORY_CODES else "unmapped"
        counts_for_credit, counts_for_gpa = True, False
    elif code in LEGACY_RECLASSIFY:
        kind = LEGACY_RECLASSIFY[code]
        counts_for_credit, counts_for_gpa = True, True
    elif code in COMPULSORY_CODES:
        kind = "compulsory"
        counts_for_credit, counts_for_gpa = True, True
    elif code in ELECTIVE_CODES:
        kind = "elective"
        counts_for_credit, counts_for_gpa = True, True
    else:
        kind = "unmapped"
        counts_for_credit, counts_for_gpa = True, False

    return CourseMeta(code, dept, category, level, credit, serial, kind,
                       counts_for_credit, counts_for_gpa)


# ---------------------------------------------------------------------------
# 6. CONVERT OLD RESULTS -> FINAL RC'25 RESULTS
# ---------------------------------------------------------------------------
def convert_results(results_old):
    """results_old: dict of (old OR already-new) course code -> grade.
    Applies CONVERSION_MAP. Returns (final_results dict, many_to_one_notes)."""
    contributions = defaultdict(list)
    for old_code, grade in results_old.items():
        targets = CONVERSION_MAP.get(old_code, [old_code])
        old_credit = parse_course_code(old_code).credit
        for new_code in targets:
            contributions[new_code].append((grade, old_credit, old_code))

    final_results = {}
    notes = []
    for new_code, contribs in contributions.items():
        if len(contribs) == 1:
            final_results[new_code] = contribs[0][0]
        else:
            total_credit = sum(c[1] for c in contribs)
            weighted_gpv = sum(GRADE_POINTS.get(g, 0.0) * cr for g, cr, _ in contribs) / total_credit
            nearest_grade = min(GRADE_POINTS.items(), key=lambda kv: abs(kv[1] - weighted_gpv))[0]
            final_results[new_code] = nearest_grade
            notes.append(f"  {new_code}: averaged from {[c[2] for c in contribs]} "
                         f"-> GPV {weighted_gpv:.2f} -> grade {nearest_grade}")
    return final_results, notes


# ---------------------------------------------------------------------------
# 7. GPA CALCULATION (official priority-order algorithm)
# ---------------------------------------------------------------------------
def compute_gpa(final_results):
    candidates = []
    skipped = []
    for code, grade in final_results.items():
        meta = parse_course_code(code)
        if not meta.counts_for_gpa or meta.level not in (4, 5, 6):
            continue
        if grade not in GRADE_POINTS:
            skipped.append((code, grade))
            continue
        candidates.append({"code": code, "credit": meta.credit, "gpv": GRADE_POINTS[grade],
                            "level": meta.level, "kind": meta.kind, "grade": grade})

    group1 = sorted([c for c in candidates if c["kind"] == "compulsory" and c["level"] in (5, 6)],
                     key=lambda x: -x["gpv"])
    group2 = sorted([c for c in candidates if c["kind"] == "elective" and c["level"] in (5, 6)],
                     key=lambda x: -x["gpv"])
    group3 = sorted([c for c in candidates if c["kind"] == "compulsory" and c["level"] == 4],
                     key=lambda x: -x["gpv"])

    total_credit, weighted_sum, used = 0.0, 0.0, []
    for c in group1 + group2 + group3:
        if total_credit >= GPA_CREDIT_CAP:
            break
        take = min(c["credit"], GPA_CREDIT_CAP - total_credit)
        weighted_sum += c["gpv"] * take
        total_credit += take
        used.append((c["code"], c["grade"], c["gpv"], take, c["credit"]))

    gpa = weighted_sum / total_credit if total_credit else None
    return gpa, total_credit, used, skipped


# ---------------------------------------------------------------------------
# 8. CATEGORY CREDIT TOTALS + EXCESS-CREDIT CAP (Regulation Section 6.8.2)
# ---------------------------------------------------------------------------
def compute_category_totals(final_results):
    totals = defaultdict(float)
    course_list = defaultdict(list)
    unmapped_flagged = []
    for code, grade in final_results.items():
        meta = parse_course_code(code)
        if not meta.counts_for_credit:
            continue
        group = CATEGORY_GROUP.get(meta.category, f"Unknown category '{meta.category}'")
        totals[group] += meta.credit
        course_list[group].append((code, meta.level, meta.credit, meta.kind == "compulsory"))
        if meta.kind == "unmapped":
            unmapped_flagged.append(code)
    return totals, course_list, unmapped_flagged


def apply_category_caps(totals, course_list):
    """Implements Regulation 6.8.2: excess credit over a category's maximum
    is trimmed, Level 3 first then 4,5,6, and NEVER from compulsory courses."""
    capped_totals, trim_notes = {}, []
    for group, raw_total in totals.items():
        max_credit = TABLE1_REQUIREMENTS.get(group, (None, None))[1]
        if max_credit is None or raw_total <= max_credit:
            capped_totals[group] = raw_total
            continue
        excess = raw_total - max_credit
        trimmable = sorted([c for c in course_list[group] if not c[3]], key=lambda c: c[1])
        remaining, trimmed_detail = excess, []
        for code, level, credit, _ in trimmable:
            if remaining <= 0:
                break
            take = min(credit, remaining)
            remaining -= take
            trimmed_detail.append((code, level, take))
        counted = raw_total - (excess - remaining)
        capped_totals[group] = counted
        if remaining > 0:
            trim_notes.append(f"  {group}: {raw_total:g} obtained, max {max_credit:g}. "
                               f"Could only trim {excess - remaining:g} of {excess:g} excess "
                               f"from non-compulsory courses - {remaining:g} credit(s) of excess "
                               f"COMPULSORY credit remain (shouldn't normally happen; verify with dept).")
        else:
            trim_notes.append(f"  {group}: {raw_total:g} obtained, capped at max {max_credit:g}. "
                               f"Trimmed {excess:g} credit(s) from: " +
                               ", ".join(f"{c} (L{l}, {t:g}cr)" for c, l, t in trimmed_detail))
    return capped_totals, trim_notes


# ---------------------------------------------------------------------------
# 9. MARKS SHEET IMPORT - auto-detects HTML(.xls)/.xlsx/legacy .xls/PDF
# ---------------------------------------------------------------------------
class _HTMLTableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows = []
        self._row, self._cell = [], []
        self._in_cell = False

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._row = []
        elif tag == "td":
            self._in_cell = True
            self._cell = []

    def handle_endtag(self, tag):
        if tag == "td":
            self._in_cell = False
            self._row.append("".join(self._cell).strip())
        elif tag == "tr":
            if self._row:
                self.rows.append(self._row)

    def handle_data(self, data):
        if self._in_cell:
            self._cell.append(data)


def _rows_from_html(path):
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        html = f.read()
    parser = _HTMLTableParser()
    parser.feed(html)
    return parser.rows


def _rows_from_xlsx(path):
    try:
        import openpyxl
    except ImportError:
        raise ImportError(
            "Reading .xlsx files needs the 'openpyxl' package. Install it with:\n"
            "    pip install openpyxl\n"
            "then try importing again."
        )
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append(["" if c is None else str(c).strip() for c in row])
    return rows


def _rows_from_legacy_xls(path):
    try:
        import xlrd
    except ImportError:
        raise ImportError(
            "This looks like a real (binary) .xls file. Reading it needs the "
            "'xlrd' package (version 1.2.0, since newer xlrd dropped .xls "
            "support). Install it with:\n"
            "    pip install xlrd==1.2.0\n"
            "or simply re-save/export the file as .xlsx and import that instead."
        )
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)
    rows = []
    for r in range(sheet.nrows):
        rows.append([str(c).strip() if c != "" else "" for c in sheet.row_values(r)])
    return rows


def _rows_from_pdf(path):
    try:
        import pdfplumber
    except ImportError:
        raise ImportError(
            "Reading .pdf marks sheets needs the 'pdfplumber' package. "
            "Install it with:\n"
            "    pip install pdfplumber\n"
            "then try importing again."
        )
    rows = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for table in (page.extract_tables() or []):
                for row in table:
                    rows.append(["" if c is None else str(c).strip() for c in row])
    return rows


def _sniff_and_get_rows(path):
    ext = os.path.splitext(path)[1].lower()
    with open(path, "rb") as f:
        head = f.read(1024)

    stripped = head.lstrip()
    if stripped[:5].lower() == b"<html" or b"<table" in head.lower():
        return _rows_from_html(path)
    if head[:4] == b"PK\x03\x04":  # zip container -> modern .xlsx
        return _rows_from_xlsx(path)
    if head[:4] == b"%PDF":
        return _rows_from_pdf(path)
    if head[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":  # OLE2 -> legacy binary .xls
        return _rows_from_legacy_xls(path)

    # Fall back to extension if content-sniffing was inconclusive
    if ext in (".html", ".htm"):
        return _rows_from_html(path)
    if ext == ".xlsx":
        return _rows_from_xlsx(path)
    if ext == ".xls":
        return _rows_from_html(path)  # MyOUSL's ".xls" is almost always HTML
    if ext == ".pdf":
        return _rows_from_pdf(path)

    raise ValueError(f"Couldn't recognise the file type of '{path}'. Expected an "
                      f"HTML/.xls/.xlsx/.pdf Marks Sheet export.")


def _process_rows(rows):
    """Shared logic for every file format: find the Course Code / Grade /
    Year columns, then for each course code take the grade from the MOST
    RECENT attempt that actually has one recorded, and classify it as
    passed/not-passed using the real passing-grade list (not just whatever
    the status column happens to say)."""
    if not rows:
        raise ValueError("No rows found - is this the right file?")

    header = [str(h).strip().lower() for h in rows[0]]

    def col_index(*names):
        for name in names:
            for i, h in enumerate(header):
                if name in h:
                    return i
        return None

    code_i = col_index("course code")
    grade_i = col_index("grade")
    year_i = col_index("last offered year", "year")

    if code_i is None or grade_i is None:
        raise ValueError("Couldn't find 'Course Code' / 'Grade' columns in this file - "
                          "the export format may be different than expected.")

    by_code = defaultdict(list)
    for row in rows[1:]:
        if code_i >= len(row):
            continue
        code = str(row[code_i]).strip().upper()
        if not code or code == "NONE":
            continue
        grade = str(row[grade_i]).strip().upper() if grade_i is not None and grade_i < len(row) else ""
        try:
            year = int(float(row[year_i])) if year_i is not None and year_i < len(row) and str(row[year_i]).strip() else 0
        except (ValueError, TypeError):
            year = 0
        by_code[code].append((year, grade))

    graded = {}
    not_yet_passed = []
    for code, attempts in by_code.items():
        # Only attempts where an actual grade was recorded (not blank/"-")
        with_grade = [(y, g) for y, g in attempts if g and g != "-"]
        if not with_grade:
            not_yet_passed.append(code)
            continue
        with_grade.sort(key=lambda a: a[0])  # by year, latest last
        latest_year, latest_grade = with_grade[-1]
        if latest_grade in PASSING_MARKERS:
            graded[code] = latest_grade if latest_grade != "PASS" else "P"
        else:
            not_yet_passed.append(code)

    return graded, not_yet_passed


def parse_marks_sheet(path):
    """Auto-detects HTML(.xls)/.xlsx/legacy-.xls/.pdf and returns
    (graded: dict code->grade, not_yet_passed: list of codes with no
    passing grade recorded yet)."""
    rows = _sniff_and_get_rows(path)
    return _process_rows(rows)


# ---------------------------------------------------------------------------
# 10. STATE PERSISTENCE (so progress survives closing the window)
# ---------------------------------------------------------------------------
def load_state():
    if not os.path.exists(BSE_STATE_FILE):
        return {}
    try:
        with open(BSE_STATE_FILE, "r") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}


def save_state(results_dict):
    with open(BSE_STATE_FILE, "w") as f:
        json.dump(results_dict, f, indent=2)


def clear_state():
    if os.path.exists(BSE_STATE_FILE):
        os.remove(BSE_STATE_FILE)


# ---------------------------------------------------------------------------
# 11. REPORT TEXT BUILDER (returns a string, so it can go to console or GUI)
# ---------------------------------------------------------------------------
def build_report(results_dict):
    out = []
    final_results, many_to_one_notes = convert_results(results_dict)

    out.append("=" * 78)
    out.append("CONVERTED RC'25 RESULTS")
    out.append("=" * 78)
    for code in sorted(final_results):
        meta = parse_course_code(code)
        tag = ""
        if meta.kind == "virtual":
            tag = "  [virtual credit-compensation course]"
        elif meta.kind == "nongpa_zero":
            tag = "  [non-GPA, doesn't count toward credit totals]"
        elif meta.kind == "unmapped":
            tag = "  [not in known RC'25 lists - verify manually]"
        elif code in LEGACY_RECLASSIFY:
            tag = "  [discontinued course, counted per manual reclassification]"
        elif code in COUNTS_BUT_NOT_GPA:
            tag = "  [counts toward credits, excluded from GPA]"
        out.append(f"  {code:10s} Grade: {final_results[code]:3s}  "
                    f"Level {meta.level}  Credit {meta.credit}  Category {meta.category}{tag}")

    if many_to_one_notes:
        out.append("\nNOTE - many-to-one conversions detected (averaged):")
        out.extend(many_to_one_notes)

    totals, course_list, unmapped_flagged = compute_category_totals(final_results)
    capped_totals, trim_notes = apply_category_caps(totals, course_list)

    out.append("\n" + "=" * 78)
    out.append("CREDIT TOTALS BY CATEGORY")
    out.append("=" * 78)
    grand_total = 0.0
    for group, (min_c, max_c) in TABLE1_REQUIREMENTS.items():
        raw = totals.get(group, 0.0)
        counted = capped_totals.get(group, raw)
        grand_total += counted
        status = "OK" if counted >= min_c else f"NEED {min_c - counted:g} MORE"
        capped_note = f"  (obtained {raw:g}, {raw - counted:g} trimmed per Reg. 6.8.2)" if counted < raw else ""
        out.append(f"  {group:35s} {counted:6g} / {min_c}-{max_c}   [{status}]{capped_note}")
    out.append(f"\n  {'TOTAL CREDITS COUNTED':35s} {grand_total:6g} / {TOTAL_MIN_CREDITS} required")

    if trim_notes:
        out.append("\n  Regulation 6.8.2 - excess credit trimming applied:")
        out.extend(trim_notes)

    if unmapped_flagged:
        out.append("\n  FLAGGED - not found in the known RC'25 course lists, category/level/")
        out.append("  credit derived purely from the code pattern. Verify manually:")
        for code in unmapped_flagged:
            out.append(f"    - {code}")

    gpa, gpa_credits, used, skipped = compute_gpa(final_results)
    out.append("\n" + "=" * 78)
    out.append("GPA CALCULATION")
    out.append("=" * 78)
    out.append("Priority order: (i) compulsory L5&6 -> (ii) elective L5&6 -> "
                "(iii) compulsory L4, highest grade first, up to 72 credits.\n")
    for code, grade, gpv, take, full_credit in used:
        partial = "" if take == full_credit else f" (partial: {take:g} of {full_credit} credits used)"
        out.append(f"  {code:10s} {grade:3s} (GPV {gpv:.1f})  credit used: {take:g}{partial}")
    out.append(f"\n  Credits used in GPA: {gpa_credits:g} / {GPA_CREDIT_CAP}")
    if gpa is not None:
        classification = classify_gpa(gpa)
        rounded_gpa = round(gpa, 2)
        out.append(f"  >>> CURRENT GPA: {rounded_gpa:.2f} <<<")
        out.append(f"  >>> PROJECTED CLASSIFICATION (if degree ended today): {classification} <<<")
        if gpa_credits < GPA_CREDIT_CAP:
            out.append(f"  (based on {gpa_credits:g} eligible credits so far - will keep updating "
                        f"as more Level 4-6 courses are completed)")
        for low, high, label in CLASSIFICATION_BANDS:
            if rounded_gpa < low:
                out.append(f"  Note: {low - rounded_gpa:.2f} more GPA points would move into '{label}'.")
                break
    else:
        out.append("  No GPA-eligible completed courses found yet.")

    if skipped:
        out.append("\n  Skipped (unrecognised grade):")
        for code, grade in skipped:
            out.append(f"    - {code}: '{grade}'")

    return "\n".join(out)


# ---------------------------------------------------------------------------
# 12. GUI
# ---------------------------------------------------------------------------
def run_gui():
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, scrolledtext
    except ImportError:
        print("Tkinter isn't available in this environment. Use the command-line "
              "fallback instead: python3 bse_gpa_calculator_generic.py --file <marks_sheet>")
        return

    state = load_state()  # code -> grade

    root = tk.Tk()
    root.title("BSEHons RC->RC'25 GPA Calculator")
    root.geometry("900x760")

    top = tk.Frame(root, padx=10, pady=8)
    top.pack(fill="x")
    tk.Label(top, text="1) Import a Marks Sheet (.xls/.xlsx/.pdf/.html), and/or  "
                        "2) add/edit rows below, then click Compute.\n"
                        "Use 'Clear (New Student)' to wipe everything and start fresh for someone else.",
              anchor="w", justify="left").pack(fill="x")

    # --- scrollable table of code/grade rows -------------------------------
    table_frame_outer = tk.Frame(root)
    table_frame_outer.pack(fill="both", expand=True, padx=10)

    canvas = tk.Canvas(table_frame_outer)
    scrollbar = tk.Scrollbar(table_frame_outer, orient="vertical", command=canvas.yview)
    table_frame = tk.Frame(canvas)
    table_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=table_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    row_widgets = {}  # id(frame) -> (code_entry, grade_entry)

    def add_row(code="", grade=""):
        r = tk.Frame(table_frame)
        r.pack(fill="x", pady=1)
        code_entry = tk.Entry(r, width=14)
        code_entry.insert(0, code)
        code_entry.pack(side="left", padx=4)
        grade_entry = tk.Entry(r, width=6)
        grade_entry.insert(0, grade)
        grade_entry.pack(side="left", padx=4)

        def remove():
            r.destroy()
            row_widgets.pop(id(r), None)

        tk.Button(r, text="Remove", command=remove, width=8).pack(side="left", padx=4)
        row_widgets[id(r)] = (code_entry, grade_entry)

    for code, grade in state.items():
        add_row(code, grade)

    # --- toolbar -------------------------------------------------------------
    toolbar = tk.Frame(root, padx=10, pady=6)
    toolbar.pack(fill="x")

    def do_import():
        path = filedialog.askopenfilename(
            title="Select the Marks Sheet export",
            filetypes=[("Marks sheet (xls/xlsx/pdf/html)", "*.xls *.xlsx *.pdf *.html *.htm"),
                       ("All files", "*.*")],
        )
        if not path:
            return
        try:
            graded, pending = parse_marks_sheet(path)
        except Exception as e:
            messagebox.showerror("Import failed", str(e))
            return
        for code, grade in graded.items():
            add_row(code, grade)
        msg = f"Imported {len(graded)} passed course(s)."
        if pending:
            msg += f"\n\n{len(pending)} course(s) have no passing grade recorded yet " \
                   f"and were skipped:\n" + ", ".join(sorted(pending))
        messagebox.showinfo("Import complete", msg)

    def do_add_blank():
        add_row("", "")

    def do_clear():
        if not messagebox.askyesno(
            "Clear everything?",
            "This clears every row in the table and deletes the saved progress "
            "file, so you can start fresh for a different student. Continue?"
        ):
            return
        for frame_id in list(row_widgets.keys()):
            pass
        for child in list(table_frame.winfo_children()):
            child.destroy()
        row_widgets.clear()
        clear_state()
        output_box.configure(state="normal")
        output_box.delete("1.0", "end")
        output_box.insert("1.0", "Cleared. Import a marks sheet or add rows above, then click "
                                  "'Compute GPA / Report'.")
        output_box.configure(state="disabled")

    output_box = None  # forward-declared, created below

    def do_compute():
        results = {}
        for code_entry, grade_entry in row_widgets.values():
            code = code_entry.get().strip().upper()
            grade = grade_entry.get().strip().upper()
            if not code or not grade:
                continue
            results[code] = grade
        if not results:
            messagebox.showwarning("Nothing to compute", "Add at least one course code + grade first.")
            return
        try:
            report = build_report(results)
        except Exception as e:
            messagebox.showerror("Error building report", str(e))
            return
        save_state(results)
        output_box.configure(state="normal")
        output_box.delete("1.0", "end")
        output_box.insert("1.0", report)
        output_box.configure(state="disabled")

    tk.Button(toolbar, text="Import Marks Sheet...", command=do_import).pack(side="left", padx=4)
    tk.Button(toolbar, text="+ Add Course Row", command=do_add_blank).pack(side="left", padx=4)
    tk.Button(toolbar, text="Compute GPA / Report", command=do_compute, bg="#dfe8ff").pack(side="left", padx=4)
    tk.Button(toolbar, text="Clear (New Student)", command=do_clear, bg="#ffe0e0").pack(side="left", padx=4)

    # --- output panel ---------------------------------------------------------
    out_frame = tk.Frame(root, padx=10, pady=6)
    out_frame.pack(fill="both", expand=True)
    tk.Label(out_frame, text="Report:", anchor="w").pack(fill="x")
    output_box = scrolledtext.ScrolledText(out_frame, height=18, font=("Courier New", 9))
    output_box.pack(fill="both", expand=True)
    output_box.insert("1.0", "Import a marks sheet or add rows above, then click "
                              "'Compute GPA / Report'.")
    output_box.configure(state="disabled")

    root.mainloop()


# ---------------------------------------------------------------------------
# 13. CLI FALLBACK (no display / headless testing)
# ---------------------------------------------------------------------------
def run_cli(marks_sheet_path=None, do_clear=False):
    if do_clear:
        clear_state()
        print("Cleared saved progress.")
    state = load_state()
    if marks_sheet_path:
        graded, pending = parse_marks_sheet(marks_sheet_path)
        state.update(graded)
        save_state(state)
        print(f"Imported {len(graded)} passed course(s) from {marks_sheet_path}.")
        if pending:
            print(f"{len(pending)} course(s) have no passing grade recorded yet: "
                  + ", ".join(sorted(pending)))
    if not state:
        print("No results yet. Run with --file <marks_sheet> to import one, "
              "or use the GUI to add rows manually.")
        return
    print()
    print(build_report(state))


if __name__ == "__main__":
    import sys
    if "--clear" in sys.argv:
        clear_state()
        sys.argv.remove("--clear")
        print("Cleared saved progress.")
    if "--file" in sys.argv:
        idx = sys.argv.index("--file")
        run_cli(sys.argv[idx + 1] if idx + 1 < len(sys.argv) else None)
    else:
        run_gui()
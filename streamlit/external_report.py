"""Build the consolidated Word report of external electors.

Takes the parsed structure that page 5 already holds — ``{sheet: {code, field,
domain, df}}`` — so the same report is produced whether the tab is reading the
submitted workbook or the database.

Word rather than PDF on purpose: ``docx2pdf`` needs a real Microsoft Word
installation and only runs on Windows, so it cannot produce anything inside the
Railway container. The document is meant to be edited before submission anyway,
and Word exports to PDF in one click.
"""

from __future__ import annotations

import io
from datetime import date

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

TITLE = "Εξωτερικοί εκλέκτορες ανά γνωστικό αντικείμενο"

BODY_FONT_PT = 6.5
HEADER_FONT_PT = 6.5

# Column widths in cm, in the order the workbook lists them. They add up to the
# 27.7 cm of a landscape A4 page inside 1 cm margins — Word ignores the widths
# entirely if the total overflows, so this must stay in budget.
COLUMN_WIDTHS = {
    "α/α": 0.7,
    "Χαρακτηρισμός": 1.6,
    "Κωδικός Χρήστη": 1.1,
    "Όνομα": 1.9,
    "Επώνυμο": 2.1,
    "Κατηγορία Χρήστη": 1.8,
    "Φορέας Χρήστη": 2.6,
    "Σχολή Χρήστη": 2.2,
    "Τμήμα/Ινστιτούτο Χρήστη": 2.2,
    "ΦΕΚ Διορισμού": 1.9,
    "Γνωστικό Αντικείμενο": 2.8,
    "Βαθμίδα": 1.8,
    "Αιτιολόγηση συνάφειας": 4.9,
}
DEFAULT_WIDTH_CM = 2.0


def _landscape_a4(document: Document) -> None:
    section = document.sections[0]
    # Swapping the page dimensions is not enough on its own — the orientation
    # flag is what Word reads when printing.
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = Cm(29.7), Cm(21.0)
    section.left_margin = section.right_margin = Cm(1.0)
    section.top_margin = section.bottom_margin = Cm(1.2)


def _set_cell(cell, text: str, *, bold: bool = False, size: float = BODY_FONT_PT) -> None:
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run("" if text is None else str(text))
    run.font.size = Pt(size)
    run.bold = bold


def _add_table(document: Document, frame, columns: list[str]):
    table = document.add_table(rows=1, cols=len(columns))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Width has to be set on every cell, not once per column: Word honours the
    # cell values and ignores a column-level width on its own. Applied as each
    # row is built, so the table is walked once rather than twice.
    widths = [Cm(COLUMN_WIDTHS.get(name, DEFAULT_WIDTH_CM)) for name in columns]

    for index, name in enumerate(columns):
        cell = table.rows[0].cells[index]
        cell.width = widths[index]
        _set_cell(cell, name, bold=True, size=HEADER_FONT_PT)
    # Repeat the header when a subject spills onto the next page
    table.rows[0]._tr.get_or_add_trPr().append(_repeat_header_element())

    for record in frame[columns].itertuples(index=False, name=None):
        cells = table.add_row().cells
        for index, value in enumerate(record):
            cells[index].width = widths[index]
            _set_cell(cells[index], value)
    return table


def _repeat_header_element():
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    element = OxmlElement("w:tblHeader")
    element.set(qn("w:val"), "true")
    return element


def _counts(frame) -> tuple[int, int]:
    if "Χαρακτηρισμός" not in frame.columns:
        return 0, 0
    values = frame["Χαρακτηρισμός"].astype(str).str.strip()
    return int((values == "ΙΔΙΟΥ").sum()), int((values == "ΣΥΝΑΦΟΥΣ").sum())


def build_report(workbook: dict[str, dict], year, source_label: str) -> bytes:
    """One Word document covering every γνωστικό αντικείμενο.

    Starts with a summary table of all subjects and their counts, then one
    landscape page per subject.
    """
    document = Document()
    _landscape_a4(document)

    heading = document.add_heading(TITLE, level=0)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run(
        f"Έτος {year} · πηγή: {source_label} · "
        f"δημιουργήθηκε {date.today().strftime('%d/%m/%Y')}"
    )
    run.italic = True

    entries = sorted(workbook.values(), key=lambda entry: int(entry["code"]))

    document.add_heading("Συγκεντρωτικός πίνακας", level=1)
    summary = document.add_table(rows=1, cols=6)
    summary.style = "Table Grid"
    for index, name in enumerate(
        ["Κωδικός", "Γνωστικό αντικείμενο", "Επιστημονικό πεδίο",
         "Ιδίου", "Συναφούς", "Σύνολο"]
    ):
        _set_cell(summary.rows[0].cells[index], name, bold=True, size=9)

    total_idiou = total_synafous = 0
    for entry in entries:
        idiou, synafous = _counts(entry["df"])
        total_idiou += idiou
        total_synafous += synafous
        cells = summary.add_row().cells
        for index, value in enumerate(
            [entry["code"], entry["field"], entry["domain"], idiou, synafous, len(entry["df"])]
        ):
            _set_cell(cells[index], value, size=9)

    cells = summary.add_row().cells
    for index, value in enumerate(
        ["", "ΣΥΝΟΛΟ", "", total_idiou, total_synafous, total_idiou + total_synafous]
    ):
        _set_cell(cells[index], value, bold=True, size=9)

    for entry in entries:
        document.add_page_break()
        document.add_heading(f"{entry['code']} — {entry['field']}", level=1)
        if entry["domain"]:
            caption = document.add_paragraph()
            run = caption.add_run(f"Επιστημονικό πεδίο: {entry['domain']}")
            run.italic = True

        frame = entry["df"]
        columns = [name for name in COLUMN_WIDTHS if name in frame.columns]
        columns += [name for name in frame.columns if name not in COLUMN_WIDTHS]
        _add_table(document, frame, columns)

    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()

"""Word output for the περιγράμματα: one course, all courses, and what changed.

Three documents, because they answer three different questions:

* :func:`render_course` — the single περίγραμμα, from the same ``docxtpl``
  template page 1 has always used. The template is read from ``files/`` rather
  than fetched from GitHub on every download: it is in the repository, and a
  download button that needs the network to work is a download button that
  fails offline.
* :func:`build_full_report` — every course of a curriculum in one file, page
  break between each. The coordinator's "give me the whole πρόγραμμα σπουδών".
* :func:`build_changes_report` — what moved in a period, read straight out of
  ``perigrammata_revisions``. Not derivable from the courses table, which only
  ever holds the present.

Word rather than PDF for the same reason as the μητρώα report: ``docx2pdf``
needs a real Word install and cannot run in the Railway container.
"""

from __future__ import annotations

import io
from datetime import datetime
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt
from docxcompose.composer import Composer
from docxtpl import DocxTemplate

from perigrammata_db import CONTENT_COLUMNS, INTEGER_COLUMNS, NUMERIC_COLUMNS

ROOT = Path(__file__).resolve().parents[1]
TEMPLATES = {
    "gr": ROOT / "files" / "perigrammata-template-gr.docx",
    "eng": ROOT / "files" / "perigrammata-template-eng.docx",
}

CHANGES_TITLE = "Μεταβολές περιγραμμάτων μαθημάτων"
BODY_FONT_PT = 8.0

# Column widths in cm for the changes table. They total 27.6 cm against the
# 27.7 cm usable on landscape A4 at 1 cm margins — Word silently ignores every
# width if the total overflows, so a new column has to take room from another.
CHANGE_COLUMNS = {
    "Πεδίο": 3.4,
    "Πριν": 9.4,
    "Μετά": 9.4,
    "Από": 3.4,
    "Ημερομηνία": 2.0,
}

# A single field can hold 4.000 characters (the weekly breakdown in `subject1`).
# Printed in full, one course's edit would run for pages and bury every other
# change. The report says what moved and roughly how; the current text is in the
# περίγραμμα itself.
MAX_CELL_CHARS = 400

_template_cache: dict[str, bytes] = {}


def _template_bytes(locale: str) -> bytes:
    """The template file, read once per process."""
    if locale not in _template_cache:
        path = TEMPLATES.get(locale)
        if path is None or not path.exists():
            raise FileNotFoundError(f"Λείπει το πρότυπο για τη γλώσσα {locale!r}: {path}")
        _template_cache[locale] = path.read_bytes()
    return _template_cache[locale]


def format_value(field: str, value: object) -> str:
    """One stored value as the Word document should print it.

    Numerics are stored as numbers so hours can be summed and εξάμηνο charted,
    but every one of them is a whole number in practice; printing ``4.0`` where
    the workbook said ``4`` was a visible wart of the Google Sheet version.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if field in NUMERIC_COLUMNS or field in INTEGER_COLUMNS:
        number = float(value)
        if number.is_integer():
            return str(int(number))
        # Greek decimal comma, to match how the workbooks are written.
        return f"{number}".replace(".", ",")
    return str(value)


def docx_context(row: dict) -> dict[str, str]:
    """A stored course as the flat, all-strings context ``docxtpl`` wants.

    Every template variable must be present: a missing key renders as an empty
    string in ``docxtpl`` only if the template guards it, and these templates
    do not.
    """
    context = {"code": "" if row.get("code") is None else str(row["code"])}
    for field in CONTENT_COLUMNS:
        context[field] = format_value(field, row.get(field))
    return context


def render_course(row: dict, locale: str = "gr") -> bytes:
    """One course's περίγραμμα as a .docx."""
    template = DocxTemplate(io.BytesIO(_template_bytes(locale)))
    template.render(docx_context(row))
    buffer = io.BytesIO()
    template.save(buffer)
    return buffer.getvalue()


def build_full_report(rows: list[dict], locale: str = "gr") -> bytes:
    """Every course in one document, one per page, in the given order.

    Each course is rendered separately and the results are merged, rather than
    looping inside a single template: the template is a whole-page form, and
    ``docxtpl`` has no way to repeat one.
    """
    if not rows:
        raise ValueError("Δεν υπάρχουν μαθήματα για την αναφορά.")

    master = Document(io.BytesIO(render_course(rows[0], locale)))
    composer = Composer(master)
    for row in rows[1:]:
        # The break goes on the master, after what has been merged so far, so
        # each περίγραμμα starts on a fresh page.
        master.add_page_break()
        composer.append(Document(io.BytesIO(render_course(row, locale))))

    buffer = io.BytesIO()
    composer.save(buffer)
    return buffer.getvalue()


# --------------------------------------------------------------------------
# The changes report
# --------------------------------------------------------------------------

def _landscape_a4(document: Document) -> None:
    section = document.sections[0]
    # Swapping the page dimensions is not enough on its own — the orientation
    # flag is what Word reads when printing.
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = Cm(29.7), Cm(21.0)
    section.left_margin = section.right_margin = Cm(1.0)
    section.top_margin = section.bottom_margin = Cm(1.2)


def _set_cell(cell, value: str, *, bold: bool = False) -> None:
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run("" if value is None else str(value))
    run.font.size = Pt(BODY_FONT_PT)
    run.bold = bold


def _shorten(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r\n", " ").replace("\n", " ").strip()
    if not text:
        return "—"
    if len(text) > MAX_CELL_CHARS:
        return text[:MAX_CELL_CHARS].rstrip() + " […]"
    return text


def build_changes_report(
    changes: pd.DataFrame,
    start: datetime,
    end: datetime,
    curriculum: int | None = None,
    names: dict[str, str] | None = None,
) -> bytes:
    """What changed in a period, grouped by course.

    ``changes`` is what :func:`perigrammata_db.changes_between` returns — one
    row per changed field. ``names`` maps course code to title, so the heading
    reads as a course rather than a code.
    """
    names = names or {}
    document = Document()
    _landscape_a4(document)

    heading = document.add_paragraph()
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = heading.add_run(CHANGES_TITLE)
    run.bold = True
    run.font.size = Pt(14)

    scope = f"Πρόγραμμα σπουδών {curriculum}" if curriculum else "Όλα τα προγράμματα σπουδών"
    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    stamp = subtitle.add_run(
        f"{scope} · περίοδος {start:%d/%m/%Y} – {end:%d/%m/%Y}"
    )
    stamp.font.size = Pt(10)

    if changes.empty:
        note = document.add_paragraph()
        note.add_run("Δεν καταγράφηκαν μεταβολές στη συγκεκριμένη περίοδο.").font.size = Pt(10)
        buffer = io.BytesIO()
        document.save(buffer)
        return buffer.getvalue()

    summary = document.add_paragraph()
    courses = changes["code"].nunique()
    summary.add_run(
        f"Σύνολο: {len(changes)} μεταβολές σε {courses} μαθήματα."
    ).font.size = Pt(10)

    for code, group in changes.groupby("code", sort=True):
        title = document.add_paragraph()
        label = names.get(code, "")
        run = title.add_run(f"{code} — {label}" if label else str(code))
        run.bold = True
        run.font.size = Pt(11)

        table = document.add_table(rows=1, cols=len(CHANGE_COLUMNS))
        table.style = "Table Grid"
        table.autofit = False
        for index, (name, width) in enumerate(CHANGE_COLUMNS.items()):
            _set_cell(table.rows[0].cells[index], name, bold=True)
            for row in table.columns[index].cells:
                row.width = Cm(width)

        for _, change in group.sort_values("edited_at").iterrows():
            cells = table.add_row().cells
            values = [
                change["label"],
                _shorten(change["before"]),
                _shorten(change["after"]),
                change["edited_by"] or "",
                pd.to_datetime(change["edited_at"]).strftime("%d/%m/%Y"),
            ]
            for index, (value, width) in enumerate(zip(values, CHANGE_COLUMNS.values())):
                _set_cell(cells[index], value)
                cells[index].width = Cm(width)

        document.add_paragraph()

    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()

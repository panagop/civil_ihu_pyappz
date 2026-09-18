import io
from dataclasses import dataclass, field

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Inches, Pt, RGBColor

DAY_NAMES = ["Δευτέρα", "Τρίτη", "Τετάρτη", "Πέμπτη", "Παρασκευή"]
TIME_SLOTS = list(range(9, 21))  # 9:00 … 20:00, each an hour-long row

HEADER_FILL = "4472C4"
CLASS_FILL = "E7EFF7"
EMPTY_FILL = "FFFFFF"

PAGE_MARGIN = Inches(0.5)
TIME_COLUMN_WIDTH = Inches(0.7)
USABLE_WIDTH = Inches(11) - 2 * PAGE_MARGIN  # landscape letter


@dataclass
class Placed:
    text: str
    start: int  # index into TIME_SLOTS
    duration: int
    col: int = 0
    span: int = 1

    @property
    def end(self) -> int:
        return self.start + self.duration


@dataclass
class DayLayout:
    classes: list[Placed] = field(default_factory=list)
    columns: int = 1


def _parse_hour(value) -> int | None:
    text = str(value)
    try:
        return int(text.split(":")[0]) if ":" in text else int(float(text))
    except ValueError:
        return None


def _class_text(row) -> str:
    def clean(column: str) -> str:
        return str(row[column]) if pd.notna(row[column]) else ""

    text = f"{clean('full_class_name')}\n{clean('instructors')}"
    if clean("room"):
        text += f"\n{clean('room')}"
    return text


def _day_classes(df_day: pd.DataFrame) -> list[Placed]:
    classes = []
    for _, row in df_day.iterrows():
        hour = _parse_hour(row["start_time"])
        if hour not in TIME_SLOTS:
            continue
        start = TIME_SLOTS.index(hour)
        duration = int(row["duration"]) if pd.notna(row["duration"]) else 1
        duration = max(1, min(duration, len(TIME_SLOTS) - start))
        classes.append(Placed(_class_text(row), start, duration))
    return classes


def layout_day(classes: list[Placed]) -> DayLayout:
    """Pack overlapping classes into side-by-side columns, as a calendar does.

    Each class takes the leftmost column that is free for its whole duration,
    then widens over the columns to its right that stay free — a class with
    nothing beside it fills the day, two concurrent ones split it.
    """
    ordered = sorted(classes, key=lambda c: (c.start, -c.duration, c.text))
    occupied: dict[tuple[int, int], Placed] = {}
    columns = 0
    for cls in ordered:
        col = 0
        while any((slot, col) in occupied for slot in range(cls.start, cls.end)):
            col += 1
        cls.col = col
        columns = max(columns, col + 1)
        for slot in range(cls.start, cls.end):
            occupied[(slot, col)] = cls

    for cls in ordered:
        span = 1
        while cls.col + span < columns and not any(
            (slot, cls.col + span) in occupied for slot in range(cls.start, cls.end)
        ):
            span += 1
        cls.span = span
        for extra in range(1, span):
            for slot in range(cls.start, cls.end):
                occupied[(slot, cls.col + extra)] = cls

    return DayLayout(ordered, max(columns, 1))


def _shade(cell, fill: str) -> None:
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), fill)
    cell._element.get_or_add_tcPr().append(shading)


def _style_cell(cell, text: str, size: int, bold: bool = False, white: bool = False) -> None:
    cell.text = text
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in paragraph.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(size)
            run.font.bold = bold
            if white:
                run.font.color.rgb = RGBColor(255, 255, 255)


ROOM_SEPARATOR = " & "


def _room_legend(df_sem: pd.DataFrame, room_names: dict[str, str]) -> str:
    """«204 = Αίθουσα 204 · ΤΣ1 = …» for the rooms this table uses, in code order."""
    codes = {
        code.strip()
        for value in df_sem["room"].dropna()
        for code in str(value).split(ROOM_SEPARATOR)
        if code.strip()
    }
    entries = [
        f"{code} = {room_names[code]}" if code in room_names else code
        for code in sorted(codes)
    ]
    return " · ".join(entries)


def _add_room_legend(doc: Document, legend: str) -> None:
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(4)
    label = paragraph.add_run("Αίθουσες: ")
    label.font.bold = True
    body = paragraph.add_run(legend)
    for run in (label, body):
        run.font.name = "Calibri"
        run.font.size = Pt(8)


def _add_semester_table(doc: Document, df_sem: pd.DataFrame) -> None:
    layouts = [
        layout_day(_day_classes(df_sem[df_sem["day"] == day])) for day in DAY_NAMES
    ]
    day_start = [1]
    for layout in layouts:
        day_start.append(day_start[-1] + layout.columns)
    total_cols = day_start[-1]

    table = doc.add_table(rows=len(TIME_SLOTS) + 1, cols=total_cols)
    table.style = "Light Grid Accent 1"
    table.autofit = False

    # Widths go on before any merge: python-docx sums the widths of the cells
    # a merge swallows, so a merged cell ends up exactly as wide as its span.
    day_width = (USABLE_WIDTH - TIME_COLUMN_WIDTH) // len(DAY_NAMES)
    widths = [TIME_COLUMN_WIDTH]
    for layout in layouts:
        widths.extend([Emu(day_width // layout.columns)] * layout.columns)
    for column, width in zip(table.columns, widths, strict=True):
        column.width = width
    for row in table.rows:
        row.height = Inches(0.4)
        for cell, width in zip(row.cells, widths, strict=True):
            cell.width = width

    header = table.rows[0]
    _style_cell(header.cells[0], "Ώρα", 10, bold=True, white=True)
    for day_idx, day_name in enumerate(DAY_NAMES):
        first, last = day_start[day_idx], day_start[day_idx + 1] - 1
        cell = header.cells[first].merge(header.cells[last]) if last > first else header.cells[first]
        _style_cell(cell, day_name, 10, bold=True, white=True)

    for slot_idx, hour in enumerate(TIME_SLOTS):
        _style_cell(table.rows[slot_idx + 1].cells[0], f"{hour}:00", 9, bold=True, white=True)

    for day_idx, layout in enumerate(layouts):
        base = day_start[day_idx]
        busy: set[tuple[int, int]] = set()
        for cls in layout.classes:
            top_left = table.rows[cls.start + 1].cells[base + cls.col]
            bottom_right = table.rows[cls.end].cells[base + cls.col + cls.span - 1]
            cell = top_left.merge(bottom_right) if bottom_right is not top_left else top_left
            _style_cell(cell, cls.text, 8)
            busy.update(
                (slot, col)
                for slot in range(cls.start, cls.end)
                for col in range(cls.col, cls.col + cls.span)
            )

        # Free cells of one hour merge across the day, so an empty hour reads
        # as one cell whatever the day's column count.
        for slot in range(len(TIME_SLOTS)):
            run_start = None
            for col in range(layout.columns + 1):
                free = col < layout.columns and (slot, col) not in busy
                if free and run_start is None:
                    run_start = col
                elif not free and run_start is not None:
                    if col - 1 > run_start:
                        row = table.rows[slot + 1]
                        row.cells[base + run_start].merge(row.cells[base + col - 1])
                    run_start = None

    # row.cells repeats a merged cell once per grid column it covers; keep the
    # wrappers alive while deduplicating, or lxml may reuse a freed proxy's id.
    unique: dict[int, tuple[int, int, object]] = {}
    for row_idx, row in enumerate(table.rows):
        for col_idx, cell in enumerate(row.cells):
            unique.setdefault(id(cell._tc), (row_idx, col_idx, cell))
    for row_idx, col_idx, cell in unique.values():
        if row_idx == 0 or col_idx == 0:
            _shade(cell, HEADER_FILL)
        else:
            _shade(cell, CLASS_FILL if cell.text.strip() else EMPTY_FILL)


def create_weekly_timetable_document(
    df: pd.DataFrame,
    period: str,
    year_label: str = "2025-2026",
    room_names: dict[str, str] | None = None,
) -> bytes:
    """Δημιουργεί Word έγγραφο με εβδομαδιαίο πρόγραμμα μαθημάτων.

    With ``room_names`` (code → name) each semester's page ends with a legend
    of the rooms it uses; the workbook-backed page writes room names into the
    cells already and passes nothing.
    """
    doc = Document()

    section = doc.sections[0]
    section.orientation = 1  # Landscape
    section.page_width = Inches(11)
    section.page_height = Inches(8.5)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(section, side, PAGE_MARGIN)

    title = doc.add_heading(f"Εβδομαδιαίο Πρόγραμμα Μαθημάτων - {period} Εξάμηνο {year_label}", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    semesters = sorted(df["semester"].unique())
    for sem_idx, semester in enumerate(semesters):
        df_sem = df[df["semester"] == semester]
        if df_sem.empty:
            continue
        heading = doc.add_heading(f"Εξάμηνο {int(semester)}", level=1)
        # A break on the heading, not a break paragraph after the table: the
        # latter lands on a fresh page when the table fills its own.
        heading.paragraph_format.page_break_before = sem_idx > 0
        _add_semester_table(doc, df_sem)
        if room_names is not None:
            legend = _room_legend(df_sem, room_names)
            if legend:
                _add_room_legend(doc, legend)

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

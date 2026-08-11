from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path
from typing import Iterable

import pythoncom
import win32com.client
from docx import Document

from .models import TableSheet


def _clean_cell_text(text: str) -> str:
    return text.replace("\r", "").replace("\x07", "").strip()


def _fallback_table_title(path: Path, index: int) -> str:
    return f"{path.stem}-{index}"


def _clean_table_title(text: str) -> str:
    title = _clean_cell_text(text).replace("\n", " ").strip()
    return title[:80]


def _docx_table_titles(doc: Document) -> dict[int, str]:
    titles: dict[int, str] = {}
    last_text = ""
    table_index = 0
    body = doc.element.body
    for child in body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            text = "".join(node.text or "" for node in child.iter() if node.tag.rsplit("}", 1)[-1] == "t")
            cleaned = _clean_table_title(text)
            if cleaned:
                last_text = cleaned
        elif tag == "tbl":
            table_index += 1
            if last_text:
                titles[table_index] = last_text
    return titles


def _word_table_title(doc, table, path: Path, index: int) -> str:
    try:
        title = _clean_table_title(table.Title)
        if title:
            return title
    except Exception:
        pass
    try:
        prefix_range = doc.Range(0, table.Range.Start)
        text = _clean_table_title(prefix_range.Paragraphs.Last.Range.Text)
        if text:
            return text
    except Exception:
        pass
    return _fallback_table_title(path, index)


def _table_to_matrix(table) -> list[list[str]]:
    rows = []
    for r in range(1, table.Rows.Count + 1):
        row = []
        for c in range(1, table.Columns.Count + 1):
            try:
                cell = table.Cell(r, c)
                text = _clean_cell_text(cell.Range.Text)
            except Exception:
                text = ""
            row.append(text)
        while row and row[-1] == "":
            row.pop()
        rows.append(row)
    while rows and not any(rows[-1]):
        rows.pop()
    return rows


@contextmanager
def _word_app(prog_id: str):
    pythoncom.CoInitialize()
    app = win32com.client.Dispatch(prog_id)
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        yield app
    finally:
        try:
            app.Quit()
        except Exception:
            pass


def extract_docx_tables(path: Path) -> list[TableSheet]:
    doc = Document(path)
    titles = _docx_table_titles(doc)
    tables: list[TableSheet] = []
    for index, table in enumerate(doc.tables, start=1):
        rows = []
        for row in table.rows:
            rows.append([cell.text.strip() for cell in row.cells])
        title = titles.get(index) or _fallback_table_title(path, index)
        tables.append(TableSheet(title=title, rows=rows, source=path, index=index))
    return tables


def extract_word_tables(path: Path) -> list[TableSheet]:
    with _word_app("Word.Application") as app:
        doc = app.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False, OpenAndRepair=True)
        try:
            tables: list[TableSheet] = []
            for index in range(1, doc.Tables.Count + 1):
                table = doc.Tables(index)
                rows = _table_to_matrix(table)
                tables.append(TableSheet(title=_word_table_title(doc, table, path, index), rows=rows, source=path, index=index))
            return tables
        finally:
            doc.Close(False)


def extract_wps_tables(path: Path) -> list[TableSheet]:
    with _word_app("kwps.Application") as app:
        doc = app.Documents.Open(str(path), False, True)
        try:
            tables: list[TableSheet] = []
            for index in range(1, doc.Tables.Count + 1):
                table = doc.Tables(index)
                rows = _table_to_matrix(table)
                tables.append(TableSheet(title=_word_table_title(doc, table, path, index), rows=rows, source=path, index=index))
            return tables
        finally:
            doc.Close(False)


def extract_tables(path: Path) -> list[TableSheet]:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return extract_docx_tables(path)
    if suffix == ".doc":
        return extract_word_tables(path)
    if suffix == ".wps":
        return extract_wps_tables(path)
    raise ValueError(f"Unsupported table source: {path}")

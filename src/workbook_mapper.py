from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any
import re
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd
from openpyxl import load_workbook


TECHNICAL_COLUMNS = ["FarmCode", "RowCode", "RowTitle", "DynRowSerial", "Dimension1"]


@dataclass(frozen=True)
class SheetTemplate:
    sheet_name: str
    title: str
    column_codes: list[str]
    column_labels: list[str]
    rows: list[dict[str, str]]


def _open_workbook(source: str | Path | bytes) -> Any:
    if isinstance(source, (str, Path)):
        return load_workbook(filename=source, read_only=True, data_only=True)
    return load_workbook(filename=BytesIO(source), read_only=True, data_only=True)


def _read_excel(source: str | Path | bytes) -> pd.DataFrame:
    archive = _open_zip_archive(source)
    shared_strings = _read_shared_strings(archive)
    sheet_xml = _read_first_sheet_xml(archive)
    rows = _parse_sheet_rows(sheet_xml, shared_strings)
    if not rows:
        return pd.DataFrame()

    headers = [_clean(value) for value in rows[0]]
    data_rows = rows[1:]

    frame = pd.DataFrame(data_rows, columns=headers)
    frame = frame.dropna(how="all").reset_index(drop=True)
    return frame


def _open_zip_archive(source: str | Path | bytes) -> zipfile.ZipFile:
    if isinstance(source, (str, Path)):
        return zipfile.ZipFile(source)
    return zipfile.ZipFile(BytesIO(source))


def _read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []

    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    values: list[str] = []
    for string_item in root.findall(".//{*}si"):
        chunks = [node.text or "" for node in string_item.findall(".//{*}t")]
        values.append("".join(chunks))
    return values


def _read_first_sheet_xml(archive: zipfile.ZipFile) -> bytes:
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    sheets = workbook_root.findall(".//{*}sheet")
    if not sheets:
        raise ValueError("Nem talalhato munkalap az omlesztett fajlban.")

    relationship_id = None
    for key, value in sheets[0].attrib.items():
        if key.endswith("id"):
            relationship_id = value
            break

    rel_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    target = None
    for relation in rel_root.findall(".//{*}Relationship"):
        if relation.attrib.get("Id") == relationship_id:
            target = relation.attrib.get("Target", "")
            break

    if not target:
        raise ValueError("Nem talalhato a bulk munkalap XML-je.")

    normalized = target.lstrip("/")
    if not normalized.startswith("xl/"):
        normalized = f"xl/{normalized.lstrip('./')}"
    return archive.read(normalized)


def _parse_sheet_rows(sheet_xml: bytes, shared_strings: list[str]) -> list[list[str]]:
    root = ET.fromstring(sheet_xml)
    parsed_rows: list[list[str]] = []

    for row in root.findall(".//{*}sheetData/{*}row"):
        values_by_index: dict[int, str] = {}
        fallback_index = 0

        for cell in row.findall("{*}c"):
            column_index = _column_index_from_ref(cell.attrib.get("r")) or fallback_index
            values_by_index[column_index] = _parse_cell_value(cell, shared_strings)
            fallback_index = max(fallback_index, column_index + 1)

        if not values_by_index:
            continue

        width = max(values_by_index) + 1
        parsed_rows.append([values_by_index.get(index, "") for index in range(width)])

    return parsed_rows


def _column_index_from_ref(reference: str | None) -> int | None:
    if not reference:
        return None

    match = re.match(r"([A-Z]+)", reference.upper())
    if not match:
        return None

    letters = match.group(1)
    index = 0
    for letter in letters:
        index = index * 26 + (ord(letter) - ord("A") + 1)
    return index - 1


def _parse_cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t", "")
    value_node = cell.find("{*}v")
    inline_node = cell.find("{*}is")

    if inline_node is not None:
        text_chunks = [node.text or "" for node in inline_node.findall(".//{*}t")]
        return "".join(text_chunks)

    raw_value = value_node.text if value_node is not None and value_node.text is not None else ""
    if cell_type == "s":
        try:
            return shared_strings[int(raw_value)]
        except (ValueError, IndexError):
            return raw_value
    return raw_value


def _clean(value: Any) -> str:
    if pd.isna(value):
        return ""
    if value is None:
        return ""
    text = str(value).strip()
    if text.casefold() == "nan":
        return ""
    return text


def _sort_key(value: str) -> tuple[int, str]:
    text = _clean(value)
    if text.isdigit():
        return (0, f"{int(text):08d}")
    return (1, text)


def load_template(source: str | Path | bytes) -> dict[str, Any]:
    workbook = _open_workbook(source)
    sheets: dict[str, SheetTemplate] = {}
    row_to_sheet: dict[str, str] = {}

    for worksheet in workbook.worksheets:
        rows = list(worksheet.iter_rows(values_only=True))
        if len(rows) < 4:
            continue

        title = _clean(rows[0][0]) or worksheet.title
        header_codes = [_clean(value) for value in rows[2]]
        header_labels = [_clean(value) for value in rows[3]]
        column_codes = header_codes[5:]
        column_labels = [
            label if label else f"Oszlop {code}"
            for code, label in zip(column_codes, header_labels[5:])
        ]

        template_rows: list[dict[str, str]] = []
        for raw_row in rows[4:]:
            row_code = _clean(raw_row[2]) if len(raw_row) > 2 else ""
            row_title = _clean(raw_row[3]) if len(raw_row) > 3 else ""
            dyn_row_serial = _clean(raw_row[4]) if len(raw_row) > 4 else ""
            if not row_code:
                continue
            template_rows.append(
                {
                    "row_code": row_code,
                    "row_title": row_title,
                    "dyn_row_serial": dyn_row_serial,
                }
            )
            row_to_sheet[row_code] = worksheet.title

        sheets[worksheet.title] = SheetTemplate(
            sheet_name=worksheet.title,
            title=title,
            column_codes=column_codes,
            column_labels=column_labels,
            rows=template_rows,
        )

    return {
        "sheets": sheets,
        "row_to_sheet": row_to_sheet,
    }


def load_bulk_data(source: str | Path | bytes) -> pd.DataFrame:
    dataframe = _read_excel(source)
    dataframe.columns = [str(column).strip().lower() for column in dataframe.columns]

    required = ["akod", "sor", "osz", "ertek", "dimenzio1", "dyn_row_serial"]
    missing = [column for column in required if column not in dataframe.columns]
    if missing:
        raise ValueError(f"Hiányzó oszlopok az ömlesztett fájlból: {', '.join(missing)}")

    prepared = dataframe[required].copy()
    for column in required:
        prepared[column] = prepared[column].map(_clean)

    prepared["sheet_name"] = prepared["sor"].map(lambda value: "")
    return prepared


def attach_sheet_names(dataframe: pd.DataFrame, row_to_sheet: dict[str, str]) -> pd.DataFrame:
    prepared = dataframe.copy()
    prepared["sheet_name"] = prepared["sor"].map(lambda value: row_to_sheet.get(value, ""))
    return prepared


def list_farms(dataframe: pd.DataFrame) -> list[str]:
    farms = sorted({value for value in dataframe["akod"].tolist() if value})
    return farms


def build_sheet_view(
    dataframe: pd.DataFrame,
    template: SheetTemplate,
    farm_code: str,
    filled_only: bool = False,
    closing_only: bool = False,
) -> pd.DataFrame:
    sheet_rows = dataframe[
        (dataframe["akod"] == farm_code) & (dataframe["sheet_name"] == template.sheet_name)
    ].copy()

    if sheet_rows.empty:
        return pd.DataFrame(columns=TECHNICAL_COLUMNS + template.column_labels)

    value_lookup = {
        (
            _clean(row.sor),
            _clean(row.dyn_row_serial),
            _clean(row.dimenzio1),
            _clean(row.osz),
        ): _clean(row.ertek)
        for row in sheet_rows.itertuples()
    }

    key_lookup: dict[str, set[tuple[str, str]]] = {}
    for row in sheet_rows.itertuples():
        key_lookup.setdefault(_clean(row.sor), set()).add(
            (_clean(row.dyn_row_serial), _clean(row.dimenzio1))
        )

    records: list[dict[str, str]] = []
    for template_row in template.rows:
        row_code = template_row["row_code"]
        row_title = template_row["row_title"]
        default_dyn = template_row["dyn_row_serial"] or "0"
        keys = sorted(key_lookup.get(row_code, {(default_dyn, "")}), key=lambda item: (_sort_key(item[0]), item[1]))

        for dyn_row_serial, dimension1 in keys:
            record = {
                "FarmCode": farm_code,
                "RowCode": row_code,
                "RowTitle": row_title,
                "DynRowSerial": dyn_row_serial or default_dyn,
                "Dimension1": dimension1,
            }
            has_value = False
            for code, label in zip(template.column_codes, template.column_labels):
                value = value_lookup.get((row_code, dyn_row_serial, dimension1, code), "")
                record[label] = value
                has_value = has_value or _has_meaningful_value(value)

            if filled_only and not has_value:
                continue
            if closing_only and not _looks_like_closing_row(row_title):
                continue
            records.append(record)

    return pd.DataFrame(records)


def _looks_like_closing_row(text: str) -> bool:
    normalized = _clean(text).casefold()
    keywords = [
        "záró",
        "záróérték",
        "záró érték",
        "záró-",
        "összesen",
        "mindösszesen",
        "év végi",
        "állomány",
    ]
    return any(keyword in normalized for keyword in keywords)


def _has_meaningful_value(value: str) -> bool:
    text = _clean(value)
    if not text:
        return False

    if text.startswith("="):
        text = text[1:].strip()

    return re.fullmatch(r"0+(?:[.,]0+)?", text) is None


def normalize_display_value(value: Any) -> str:
    text = _clean(value)
    if not text:
        return ""

    plain = text
    if plain.startswith("="):
        plain = plain[1:].strip()

    if re.fullmatch(r"0+(?:[.,]0+)?", plain):
        return ""

    return text


def summarize_template(template_bundle: dict[str, Any]) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for sheet in template_bundle["sheets"].values():
        rows.append(
            {
                "Munkalap": sheet.sheet_name,
                "Cím": sheet.title,
                "Sorok száma": len(sheet.rows),
                "Értékoszlopok": len(sheet.column_codes),
            }
        )
    return pd.DataFrame(rows).sort_values("Munkalap").reset_index(drop=True)

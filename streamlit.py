from __future__ import annotations

import re
import tempfile
from collections import OrderedDict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET

import streamlit as st

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)

DATA_START_ROW = 31
LAST_TEMPLATE_COLUMN = "U"

CONSTANTS = {
    "customer": "Cash sales",
    "deposit_account": "3102-000 Cash on Hand",
    "sales_account": "5000-000 Sales income ",
}


@dataclass
class InvoiceRow:
    invoice_no: str
    invoice_date: datetime
    item_description: str
    quantity: float
    unit_price: float
    discount: float = 0.0


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def parse_money(value: str) -> float:
    cleaned = value.replace("RM", "").replace(",", "").strip()
    return float(cleaned)


def is_money(value: str) -> bool:
    return bool(
        re.fullmatch(
            r"(?:RM)?\s*-?(?:\d{1,3}(?:,\d{3})*|\d+)(?:\.\d{1,2})",
            value.strip(),
        )
    )


def is_quantity(value: str) -> bool:
    return bool(re.fullmatch(r"\d+(?:\.\d+)?", value.strip()))


def is_service_date(value: str) -> bool:
    return bool(re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", value.strip()))


def excel_date(value: datetime) -> int:
    return (value - datetime(1899, 12, 30)).days


def number_for_excel(value: float | int) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def extract_invoice_header(text: str) -> tuple[str, datetime] | None:
    match = re.search(
        r"\bINVOICE\s*\n(\d+)\s*\nDATE\s*\n(\d{2}\.\d{2}\.\d{4})",
        text,
    )

    if not match:
        return None

    invoice_no, invoice_date = match.groups()
    return invoice_no, datetime.strptime(invoice_date, "%d.%m.%Y")


def extract_discount(text: str) -> float:
    match = re.search(
        r"DISCOUNT(?:\s+\d+%)?\s+(-?[\d,]+\.\d{2})",
        text,
        re.IGNORECASE,
    )

    if not match:
        return 0.0

    return abs(parse_money(match.group(1)))


def table_lines_from_page(text: str) -> list[str]:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    start = None

    for index in range(len(lines) - 4):
        if lines[index:index + 5] == ["ACTIVITY", "DESCRIPTION", "QTY", "RATE", "AMOUNT"]:
            start = index + 5
            break

    # For continuation page without table header
    if start is None:
        for index, line in enumerate(lines):
            if is_service_date(line):
                start = index
                break

    if start is None:
        return []

    end_markers = (
        "Terms and Conditions",
        "SUBTOTAL",
        "DISCOUNT",
        "TOTAL",
        "PAYMENT",
        "BALANCE DUE",
        "It was great working with you",
        "Estimate Summary",
    )

    end = len(lines)
    for index in range(start, len(lines)):
        if any(lines[index].startswith(marker) for marker in end_markers):
            end = index
            break

    return lines[start:end]


def activity_from_item_lines(lines: list[str]) -> str:
    item_lines = list(lines)

    if item_lines and is_service_date(item_lines[0]):
        item_lines.pop(0)

    if not item_lines:
        return ""

    if item_lines[0].startswith("Kilat Smart Equipment Set"):
        return normalize_text(" ".join(item_lines))

    activity: list[str] = []
    balance = 0
    saw_parenthesis = False

    for line in item_lines:
        activity.append(line)
        balance += line.count("(") - line.count(")")
        saw_parenthesis = saw_parenthesis or "(" in line

        if saw_parenthesis and balance <= 0:
            break

    return normalize_text(" ".join(activity))


def parse_table_items(table_lines: list[str]) -> Iterable[tuple[str, float, float]]:
    cursor = 0

    while cursor < len(table_lines):
        found = False

        for index in range(cursor, len(table_lines) - 1):
            # Normal case: QTY, RATE, AMOUNT
            if (
                index + 2 < len(table_lines)
                and is_quantity(table_lines[index])
                and is_money(table_lines[index + 1])
                and is_money(table_lines[index + 2])
            ):
                item_lines = table_lines[cursor:index]
                item_description = activity_from_item_lines(item_lines)
                quantity = parse_money(table_lines[index])
                unit_price = parse_money(table_lines[index + 1])

                if item_description:
                    yield item_description, quantity, unit_price

                cursor = index + 3
                found = True
                break

            # Void invoice case: QTY, RATE only
            if (
                is_quantity(table_lines[index])
                and is_money(table_lines[index + 1])
            ):
                item_lines = table_lines[cursor:index]
                item_description = activity_from_item_lines(item_lines)
                quantity = parse_money(table_lines[index])
                unit_price = parse_money(table_lines[index + 1])

                if item_description:
                    yield item_description, quantity, unit_price

                cursor = index + 2
                found = True
                break

        if not found:
            break


def extract_invoice_rows(pdf_path: Path) -> list[InvoiceRow]:
    if fitz is None:
        raise RuntimeError("PyMuPDF is required. Install it with: pip install pymupdf")

    records: OrderedDict[tuple[str, datetime, str], InvoiceRow] = OrderedDict()
    current_invoice: tuple[str, datetime] | None = None
    invoice_discounts: dict[str, float] = {}

    with fitz.open(pdf_path) as document:
        for page in document:
            text = page.get_text("text")

            header = extract_invoice_header(text)
            if header:
                current_invoice = header

            if current_invoice is None:
                continue

            invoice_no, invoice_date = current_invoice

            discount = extract_discount(text)
            if discount:
                invoice_discounts[invoice_no] = discount

            table_lines = table_lines_from_page(text)
            if not table_lines:
                continue

            for item_description, quantity, unit_price in parse_table_items(table_lines):
                key = (invoice_no, invoice_date, item_description)

                if key in records:
                    records[key].quantity += quantity
                    records[key].unit_price += unit_price
                else:
                    records[key] = InvoiceRow(
                        invoice_no=invoice_no,
                        invoice_date=invoice_date,
                        item_description=item_description,
                        quantity=quantity,
                        unit_price=unit_price,
                    )

    rows = list(records.values())

    seen_invoice: set[str] = set()
    for row in rows:
        discount = invoice_discounts.get(row.invoice_no, 0.0)

        if row.invoice_no not in seen_invoice:
            row.discount = discount
            seen_invoice.add(row.invoice_no)
        else:
            row.discount = 0.0

    return rows


def column_number(column_name: str) -> int:
    result = 0
    for char in column_name:
        result = result * 26 + ord(char.upper()) - 64
    return result


def column_name(column_number_value: int) -> str:
    name = ""
    while column_number_value:
        column_number_value, remainder = divmod(column_number_value - 1, 26)
        name = chr(65 + remainder) + name
    return name


def split_cell_ref(cell_ref: str) -> tuple[str, int]:
    match = re.fullmatch(r"([A-Z]+)(\d+)", cell_ref)
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    return match.group(1), int(match.group(2))


def qname(local_name: str) -> str:
    return f"{{{MAIN_NS}}}{local_name}"


def clear_cell(cell: ET.Element) -> None:
    cell.attrib.pop("t", None)
    for child in list(cell):
        cell.remove(child)


def set_string(cell: ET.Element, value: str) -> None:
    clear_cell(cell)
    cell.set("t", "inlineStr")
    inline = ET.SubElement(cell, qname("is"))
    text = ET.SubElement(inline, qname("t"))
    text.text = value


def set_number(cell: ET.Element, value: float | int) -> None:
    clear_cell(cell)
    v = ET.SubElement(cell, qname("v"))
    v.text = number_for_excel(value)


def cell_column_index(cell: ET.Element) -> int:
    col, _ = split_cell_ref(cell.attrib["r"])
    return column_number(col)


def row_number(row: ET.Element) -> int:
    return int(row.attrib["r"])


def get_or_create_row(
    sheet_data: ET.Element,
    rows_by_number: dict[int, ET.Element],
    row_num: int,
) -> ET.Element:
    row = rows_by_number.get(row_num)

    if row is not None:
        return row

    row = ET.Element(
        qname("row"),
        {"r": str(row_num), "spans": f"1:{column_number(LAST_TEMPLATE_COLUMN)}"},
    )

    insert_at = len(list(sheet_data))

    for index, existing in enumerate(list(sheet_data)):
        if row_number(existing) > row_num:
            insert_at = index
            break

    sheet_data.insert(insert_at, row)
    rows_by_number[row_num] = row
    return row


def get_or_create_cell(
    row: ET.Element,
    row_num: int,
    col_name: str,
    style_by_column: dict[str, str],
) -> ET.Element:
    target_ref = f"{col_name}{row_num}"

    for cell in row.findall(qname("c")):
        if cell.attrib.get("r") == target_ref:
            return cell

    attrs = {"r": target_ref}

    if col_name in style_by_column:
        attrs["s"] = style_by_column[col_name]

    cell = ET.Element(qname("c"), attrs)

    target_col_index = column_number(col_name)
    children = list(row)
    insert_at = len(children)

    for index, existing in enumerate(children):
        if existing.tag == qname("c") and cell_column_index(existing) > target_col_index:
            insert_at = index
            break

    row.insert(insert_at, cell)
    return cell


def write_rows_to_template(template_path: Path, output_path: Path, rows: list[InvoiceRow]) -> None:
    if not rows:
        raise RuntimeError("No invoice rows were extracted from the PDF.")

    with ZipFile(template_path, "r") as source:
        sheet_xml = source.read("xl/worksheets/sheet1.xml")
        root = ET.fromstring(sheet_xml)

        sheet_data = root.find(qname("sheetData"))
        if sheet_data is None:
            raise RuntimeError("Could not find sheet data in the Excel template.")

        rows_by_number = {
            row_number(row): row
            for row in sheet_data.findall(qname("row"))
        }

        style_by_column: dict[str, str] = {}
        sample_row = rows_by_number.get(DATA_START_ROW)

        if sample_row is not None:
            for cell in sample_row.findall(qname("c")):
                col, _ = split_cell_ref(cell.attrib["r"])
                if "s" in cell.attrib:
                    style_by_column[col] = cell.attrib["s"]

        last_col_number = column_number(LAST_TEMPLATE_COLUMN)
        last_existing_row = max(rows_by_number)

        for row_num in range(DATA_START_ROW, last_existing_row + 1):
            row = rows_by_number.get(row_num)
            if row is None:
                continue

            for cell in row.findall(qname("c")):
                col, _ = split_cell_ref(cell.attrib["r"])
                if column_number(col) <= last_col_number:
                    clear_cell(cell)

        written_invoice_headers: set[str] = set()

        for offset, invoice_row in enumerate(rows):
            row_num = DATA_START_ROW + offset
            row = get_or_create_row(sheet_data, rows_by_number, row_num)
            row.set("spans", f"1:{last_col_number}")

            should_write_invoice_header = invoice_row.invoice_no not in written_invoice_headers
            written_invoice_headers.add(invoice_row.invoice_no)

            row_values = {
                "M": CONSTANTS["sales_account"],
                "N": invoice_row.item_description,
                "O": invoice_row.quantity,
                "R": invoice_row.unit_price,
                "S": invoice_row.discount if should_write_invoice_header and invoice_row.discount else 0,
            }

            if should_write_invoice_header:
                row_values.update(
                    {
                        "B": CONSTANTS["customer"],
                        "D": invoice_row.invoice_no,
                        "E": excel_date(invoice_row.invoice_date),
                        "H": CONSTANTS["deposit_account"],
                    }
                )

            for col_index in range(1, last_col_number + 1):
                col = column_name(col_index)
                cell = get_or_create_cell(row, row_num, col, style_by_column)
                value = row_values.get(col, "")

                if isinstance(value, str):
                    if value:
                        set_string(cell, value)
                    else:
                        clear_cell(cell)
                elif value == "":
                    clear_cell(cell)
                else:
                    set_number(cell, value)

        dimension = root.find(qname("dimension"))
        if dimension is not None:
            max_row = max(last_existing_row, DATA_START_ROW + len(rows) - 1)
            dimension.set("ref", f"A1:{LAST_TEMPLATE_COLUMN}{max_row}")

        updated_sheet = ET.tostring(root, encoding="utf-8", xml_declaration=True)

        output_path.parent.mkdir(parents=True, exist_ok=True)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=output_path.parent) as temp_file:
            temp_path = Path(temp_file.name)

        try:
            with ZipFile(temp_path, "w", ZIP_DEFLATED) as target:
                for item in source.infolist():
                    data = updated_sheet if item.filename == "xl/worksheets/sheet1.xml" else source.read(item.filename)
                    target.writestr(item, data)

            temp_path.replace(output_path)

        finally:
            if temp_path.exists():
                temp_path.unlink()


def run_extraction(pdf_path: str | Path, template_path: str | Path, output_path: str | Path) -> int:
    pdf = Path(pdf_path)
    template = Path(template_path)
    output = Path(output_path)

    if not pdf.exists():
        raise FileNotFoundError(f"PDF not found: {pdf}")

    if not template.exists():
        raise FileNotFoundError(f"Excel template not found: {template}")

    rows = extract_invoice_rows(pdf)
    write_rows_to_template(template, output, rows)

    return len(rows)


st.set_page_config(page_title="PDF Invoice Extractor", layout="centered")

st.title("PDF Invoice Extractor")
st.write("Upload one or more PDF invoices and one Excel template.")

pdf_files = st.file_uploader(
    "Upload PDF invoice(s)",
    type=["pdf"],
    accept_multiple_files=True,
)

template_file = st.file_uploader("Upload Excel template", type=["xlsx"])

if st.button("Extract to Excel"):
    if not pdf_files or not template_file:
        st.error("Please upload at least one PDF and one Excel template.")
    else:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_dir_path = Path(temp_dir)
            template_path = temp_dir_path / template_file.name
            template_path.write_bytes(template_file.read())

            output_files = []

            try:
                for pdf_file in pdf_files:
                    pdf_path = temp_dir_path / pdf_file.name
                    pdf_path.write_bytes(pdf_file.read())

                    output_filename = Path(pdf_file.name).with_suffix(".xlsx").name
                    output_path = temp_dir_path / output_filename

                    count = run_extraction(pdf_path, template_path, output_path)
                    output_files.append((output_filename, output_path, count))

                st.success(f"Done. Generated {len(output_files)} Excel file(s).")

                for output_filename, output_path, count in output_files:
                    st.download_button(
                        label=f"Download {output_filename} ({count} row(s))",
                        data=output_path.read_bytes(),
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                zip_path = temp_dir_path / "all_extracted_excels.zip"

                with ZipFile(zip_path, "w", ZIP_DEFLATED) as zip_file:
                    for output_filename, output_path, _ in output_files:
                        zip_file.write(output_path, arcname=output_filename)

                st.download_button(
                    label="Download All Excels",
                    data=zip_path.read_bytes(),
                    file_name="all_extracted_excels.zip",
                    mime="application/zip",
                )

            except Exception as exc:
                st.error(f"Extraction failed: {exc}")
"""``procurement.excel_parser`` testleri.

Testler gerçek şablonu ``templates/supplier_quote_template.xlsx`` üzerinden
doğrular. Şablonun kendisi hiçbir testte değiştirilmez: değişiklik gereken
senaryolarda dosya ``tmp_path`` içine kopyalanır ve kopya düzenlenir.
"""

from __future__ import annotations

import hashlib
import shutil
from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import load_workbook

from procurement.excel_parser import (
    DuplicateHeaderError,
    ExcelParserError,
    FileTooLargeError,
    FormulaCellError,
    MissingFieldError,
    MissingHeaderError,
    MissingSheetError,
    ParsedWorkbook,
    UnsupportedFormatError,
    VatDeclarationError,
    parse_workbook,
)

TEMPLATE = Path(__file__).resolve().parents[1] / "templates" / "supplier_quote_template.xlsx"

pytestmark = pytest.mark.integration


# --------------------------------------------------------------------------- #
# Yardımcılar
# --------------------------------------------------------------------------- #


def _copy_template(tmp_path: Path, name: str = "copy.xlsx") -> Path:
    """Şablonun düzenlenebilir kopyasını üretir."""
    target = tmp_path / name
    shutil.copyfile(TEMPLATE, target)
    return target


def _approve_vat(path: Path) -> Path:
    """``Instructions!B4`` değerini ``EVET`` yapar (kopya üzerinde)."""
    workbook = load_workbook(path)
    workbook["Instructions"]["B4"] = "EVET"
    workbook.save(path)
    workbook.close()
    return path


def _approved_copy(tmp_path: Path, name: str = "approved.xlsx") -> Path:
    return _approve_vat(_copy_template(tmp_path, name))


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


# --------------------------------------------------------------------------- #
# Gerçek şablon
# --------------------------------------------------------------------------- #


def test_real_template_parses_successfully(tmp_path):
    result = parse_workbook(_approved_copy(tmp_path))

    assert isinstance(result, ParsedWorkbook)
    assert result.template.code == "RFQ-KUTUPROFIL-40X40X2"
    assert result.template.version == "1.0"
    assert result.template.vat_declaration == "EVET"


def test_real_template_rfq_fields_are_read(tmp_path):
    rfq = parse_workbook(_approved_copy(tmp_path)).rfq

    assert rfq.rfq_number == "RFQ-2026-0042"
    assert rfq.product_code == "KUTUPROFIL-40X40X2-GALV"
    assert rfq.profile_specification == "40x40x2"
    assert rfq.length_m == 6
    assert rfq.wall_thickness_mm == 2
    assert rfq.quantity_pcs == 500
    assert rfq.delivery_location == "Bursa / Nilüfer OSB"
    assert rfq.requested_delivery_days == 20
    assert rfq.currency == "TRY"
    assert "KDV hariç" in rfq.description


def test_real_template_suppliers_are_read(tmp_path):
    suppliers = parse_workbook(_approved_copy(tmp_path)).suppliers

    assert len(suppliers) == 3
    assert [s.row_number for s in suppliers] == [3, 4, 5]
    assert [s.supplier_name for s in suppliers] == [
        "Bursa Metal A.Ş.",
        "Gemlik Profil Ltd.",
        "İnegöl Çelik San.",
    ]
    assert suppliers[0].supplier_quote_number == "BM-2026-1187"
    assert suppliers[0].payment_terms == "Net 30"
    assert suppliers[0].shipping_cost_source == "Supplier Quote"
    assert suppliers[2].shipping_cost_source == "Customer Assumption"


def test_cell_value_types_are_preserved(tmp_path):
    result = parse_workbook(_approved_copy(tmp_path))

    assert isinstance(result.rfq.rfq_number, str)
    assert isinstance(result.rfq.rfq_date, (date, datetime))
    assert isinstance(result.rfq.length_m, int)
    assert isinstance(result.rfq.quantity_pcs, int)

    first = result.suppliers[0]
    assert isinstance(first.length_m, int)
    assert isinstance(first.quantity_pcs, int)
    assert isinstance(first.weight_kg_per_m, float)
    assert isinstance(first.unit_price_try_per_kg, float)
    # Şablonda 7500 olarak yazıldığı için openpyxl int okur; parser değeri
    # dönüştürmez, olduğu gibi korur.
    assert isinstance(first.shipping_cost_try, (int, float))
    assert not isinstance(first.shipping_cost_try, bool)
    assert first.shipping_cost_try == 7500
    assert isinstance(first.quote_date, (date, datetime))
    assert first.quote_date.date() == date(2026, 11, 12)


# --------------------------------------------------------------------------- #
# KDV beyanı
# --------------------------------------------------------------------------- #


def test_vat_declaration_rejected_when_hayir(tmp_path):
    """Şablon ``HAYIR`` ile gelir; beyan olmadan dosya işleme alınmamalı."""
    with pytest.raises(VatDeclarationError) as excinfo:
        parse_workbook(_copy_template(tmp_path))

    error = excinfo.value
    assert error.cell == "B4"
    assert error.sheet == "Instructions"
    assert "EVET" in str(error)


def test_vat_declaration_rejected_when_blank(tmp_path):
    path = _copy_template(tmp_path)
    workbook = load_workbook(path)
    workbook["Instructions"]["B4"] = None
    workbook.save(path)
    workbook.close()

    with pytest.raises(VatDeclarationError):
        parse_workbook(path)


def test_pristine_template_is_rejected(tmp_path):
    """Gerçek şablon dosyası da beyan yapılmadığı için reddedilir."""
    with pytest.raises(VatDeclarationError):
        parse_workbook(TEMPLATE)


# --------------------------------------------------------------------------- #
# Yapısal hatalar
# --------------------------------------------------------------------------- #


def test_missing_sheet_raises(tmp_path):
    path = _approved_copy(tmp_path)
    workbook = load_workbook(path)
    del workbook["Supplier Quotes"]
    workbook.save(path)
    workbook.close()

    with pytest.raises(MissingSheetError) as excinfo:
        parse_workbook(path)

    assert "Supplier Quotes" in str(excinfo.value)


def test_missing_header_raises(tmp_path):
    path = _approved_copy(tmp_path)
    workbook = load_workbook(path)
    workbook["Supplier Quotes"]["H2"] = None  # Shipping Cost başlığı silindi
    workbook.save(path)
    workbook.close()

    with pytest.raises(MissingHeaderError) as excinfo:
        parse_workbook(path)

    assert "Shipping Cost (TRY, excl. VAT)" in str(excinfo.value)


def test_missing_rfq_field_raises(tmp_path):
    path = _approved_copy(tmp_path)
    workbook = load_workbook(path)
    workbook["RFQ"]["A3"] = "Something Else"
    workbook.save(path)
    workbook.close()

    with pytest.raises(MissingFieldError) as excinfo:
        parse_workbook(path)

    assert "RFQ Number" in str(excinfo.value)


def test_duplicate_header_raises(tmp_path):
    path = _approved_copy(tmp_path)
    workbook = load_workbook(path)
    workbook["Supplier Quotes"]["H2"] = "Notes"  # N2 ile çakışıyor
    workbook.save(path)
    workbook.close()

    with pytest.raises(DuplicateHeaderError):
        parse_workbook(path)


def test_unsupported_format_raises(tmp_path):
    path = tmp_path / "quotes.xlsm"
    path.write_bytes(TEMPLATE.read_bytes())

    with pytest.raises(UnsupportedFormatError):
        parse_workbook(path)


def test_unreadable_file_message_hides_server_paths(tmp_path):
    """Bozuk dosya mesajı yol, geçici dizin ve teknik ayrıntı içermez."""
    path = tmp_path / "bozuk.xlsx"
    path.write_bytes(b"bu bir excel dosyasi degil")

    with pytest.raises(ExcelParserError) as excinfo:
        parse_workbook(path)

    message = str(excinfo.value)
    assert "Dosya açılamadı" in message
    assert "yeniden yükleyin" in message
    assert str(path) not in message
    assert str(tmp_path) not in message
    assert ".xlsx (" not in message
    assert "zip" not in message.lower()
    # Teknik neden hata ayıklama için korunur, kullanıcıya gösterilmez.
    assert excinfo.value.__cause__ is not None


def test_missing_file_message_hides_server_paths(tmp_path):
    """Bulunamayan dosya mesajı tam yolu kullanıcıya göstermez."""
    path = tmp_path / "yok.xlsx"

    with pytest.raises(ExcelParserError) as excinfo:
        parse_workbook(path)

    message = str(excinfo.value)
    assert "Dosya bulunamadı" in message
    assert str(path) not in message
    assert str(tmp_path) not in message


def test_file_too_large_raises(tmp_path, monkeypatch):
    import procurement.excel_parser as parser

    path = _approved_copy(tmp_path)
    monkeypatch.setattr(parser, "MAX_FILE_SIZE_BYTES", 10)

    with pytest.raises(FileTooLargeError):
        parse_workbook(path)


def test_formula_cell_rejected(tmp_path):
    """Formül hücresi güvenilir veri sayılmaz."""
    path = _approved_copy(tmp_path)
    workbook = load_workbook(path)
    workbook["Supplier Quotes"]["G3"] = "=28.5*2"
    workbook.save(path)
    workbook.close()

    with pytest.raises(FormulaCellError) as excinfo:
        parse_workbook(path)

    assert excinfo.value.cell == "G3"


def test_unknown_extra_column_ignored(tmp_path):
    """Tanımsız sütun okunmaz; formül içerse bile ayrıştırmayı durdurmaz."""
    path = _approved_copy(tmp_path)
    workbook = load_workbook(path)
    sheet = workbook["Supplier Quotes"]
    sheet["O2"] = "Calculated Total"
    sheet["O3"] = "=F3*E3*6"
    workbook.save(path)
    workbook.close()

    result = parse_workbook(path)

    assert len(result.suppliers) == 3
    assert result.suppliers[0].unit_price_try_per_kg == 28.5


# --------------------------------------------------------------------------- #
# Boş satırlar
# --------------------------------------------------------------------------- #


def test_blank_rows_are_skipped(tmp_path):
    result = parse_workbook(_approved_copy(tmp_path))

    assert len(result.suppliers) == 3
    assert [s.row_number for s in result.suppliers] == [3, 4, 5]


def test_blank_row_between_data_is_skipped(tmp_path):
    path = _approved_copy(tmp_path)
    workbook = load_workbook(path)
    sheet = workbook["Supplier Quotes"]
    sheet["A3"] = None  # 3. satır tamamen boşaltılır
    for column in range(1, 15):
        sheet.cell(row=3, column=column).value = None
    workbook.save(path)
    workbook.close()

    result = parse_workbook(path)

    assert [s.row_number for s in result.suppliers] == [4, 5]


def test_row_with_only_spaces_is_skipped(tmp_path):
    path = _approved_copy(tmp_path)
    workbook = load_workbook(path)
    sheet = workbook["Supplier Quotes"]
    for column in range(1, 15):
        sheet.cell(row=3, column=column).value = "   "
    workbook.save(path)
    workbook.close()

    result = parse_workbook(path)

    assert [s.row_number for s in result.suppliers] == [4, 5]


# --------------------------------------------------------------------------- #
# Salt okunurluk
# --------------------------------------------------------------------------- #


def test_source_file_is_not_modified(tmp_path):
    path = _approved_copy(tmp_path)
    before_hash = _sha256(path)
    before_mtime = path.stat().st_mtime_ns

    parse_workbook(path)

    assert _sha256(path) == before_hash
    assert path.stat().st_mtime_ns == before_mtime


def test_pristine_template_is_not_modified():
    before_hash = _sha256(TEMPLATE)
    before_mtime = TEMPLATE.stat().st_mtime_ns

    with pytest.raises(ExcelParserError):
        parse_workbook(TEMPLATE)

    assert _sha256(TEMPLATE) == before_hash
    assert TEMPLATE.stat().st_mtime_ns == before_mtime

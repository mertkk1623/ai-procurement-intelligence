"""``procurement.demo_data`` testleri.

Kapsam
------
* Boş şablonun resmî şablonla **aynı sözleşmeyi** taşıdığı (sayfalar, 14
  başlık, sabit ürün alanları),
* boş şablonda hiç veri satırı bulunmadığı,
* dolu örnek dosyanın doğrulamadan **hatasız** geçtiği, tüm satırların
  hesaplandığı ve AI isteğinin üretilebildiği,
* örnek verinin kasıtlı senaryoları gerçekten tetiklediği (fiyat uçurumu,
  termin darboğazı, eksik nakliye, süresi geçmiş teklif, referans dışı kg/m,
  gerekçesiz nakliye varsayımı),
* tarihlerin bugüne göre üretildiği (sabit tarih yok),
* dosya üretiminin diske yazmadığı.
"""

from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

import pytest

from procurement import ai_analysis
from procurement.calculator import calculate
from procurement.demo_data import (
    BLANK_FILENAME,
    COLUMN_GUIDE,
    DemoDataError,
    SAMPLE_QUANTITY,
    SAMPLE_QUOTES,
    build_blank_template,
    build_sample_workbook,
    sample_supplier_count,
)
from procurement.excel_parser import (
    SHEET_INSTRUCTIONS,
    SHEET_QUOTES,
    SHEET_RFQ,
    SUPPLIER_QUOTE_HEADERS,
    parse_workbook,
)
from procurement.validator import validate

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "templates" / "supplier_quote_template.xlsx"

pytestmark = pytest.mark.integration


# --------------------------------------------------------------------------- #
# Yardımcılar
# --------------------------------------------------------------------------- #


def _sample_path(tmp_path: Path, today: date | None = None) -> Path:
    path = tmp_path / "ornek.xlsx"
    path.write_bytes(build_sample_workbook(TEMPLATE, today=today))
    return path


def _blank_path(tmp_path: Path) -> Path:
    path = tmp_path / "bos.xlsx"
    path.write_bytes(build_blank_template(TEMPLATE))
    return path


def _codes(findings) -> set[str]:
    return {finding.code for finding in findings}


# --------------------------------------------------------------------------- #
# Boş şablon
# --------------------------------------------------------------------------- #


def test_blank_template_filename_is_xlsx():
    assert BLANK_FILENAME.endswith(".xlsx")


def test_blank_template_keeps_sheets_and_headers(tmp_path):
    parsed = parse_workbook(_blank_path(tmp_path), strict=False)

    assert parsed.structure.sheet_names == (
        SHEET_INSTRUCTIONS,
        SHEET_RFQ,
        SHEET_QUOTES,
    )
    assert parsed.structure.required_columns_present
    assert not parsed.structure.missing_headers
    assert not parsed.structure.unknown_headers


def test_blank_template_has_no_supplier_rows(tmp_path):
    parsed = parse_workbook(_blank_path(tmp_path), strict=False)

    assert parsed.suppliers == ()


def test_blank_template_keeps_fixed_product_fields(tmp_path):
    parsed = parse_workbook(_blank_path(tmp_path), strict=False)

    assert parsed.rfq.product_code == "KUTUPROFIL-40X40X2-GALV"
    assert parsed.rfq.profile_specification == "40x40x2"
    assert parsed.rfq.length_m == 6
    assert parsed.rfq.wall_thickness_mm == 2
    assert parsed.rfq.currency == "TRY"


def test_blank_template_leaves_user_fields_empty(tmp_path):
    parsed = parse_workbook(_blank_path(tmp_path), strict=False)

    assert parsed.rfq.rfq_number is None
    assert parsed.rfq.rfq_date is None
    assert parsed.rfq.quantity_pcs is None
    assert parsed.rfq.delivery_location is None


def test_blank_template_requires_explicit_vat_declaration(tmp_path):
    """Boş şablon da KDV beyanını kullanıcıya bırakır (HAYIR)."""
    parsed = parse_workbook(_blank_path(tmp_path), strict=False)
    validation = validate(parsed, today=date.today())

    assert validation.file_rejected
    assert "E-DECL-01" in _codes(validation.findings)


def test_blank_template_column_guide_covers_every_header(tmp_path):
    headings = {header for header, _guidance in COLUMN_GUIDE}

    assert headings == set(SUPPLIER_QUOTE_HEADERS)
    assert len(COLUMN_GUIDE) == len(SUPPLIER_QUOTE_HEADERS)
    for _header, guidance in COLUMN_GUIDE:
        assert guidance.strip()


def test_blank_template_writes_column_guide_into_instructions(tmp_path):
    from openpyxl import load_workbook

    workbook = load_workbook(_blank_path(tmp_path))
    try:
        instructions = workbook[SHEET_INSTRUCTIONS]
        text = " ".join(
            str(instructions.cell(row=row, column=column).value or "")
            for row in range(18, 45)
            for column in (1, 2)
        )
    finally:
        workbook.close()

    assert "HIZLI BAŞLANGIÇ" in text
    for header in SUPPLIER_QUOTE_HEADERS:
        assert header in text


def test_blank_template_missing_file_raises_demo_error(tmp_path):
    with pytest.raises(DemoDataError):
        build_blank_template(tmp_path / "yok.xlsx")


# --------------------------------------------------------------------------- #
# Dolu örnek dosya
# --------------------------------------------------------------------------- #


def test_sample_size_is_within_requested_range():
    assert 30 <= sample_supplier_count() <= 50
    assert sample_supplier_count() == len(SAMPLE_QUOTES)


def test_sample_quote_numbers_are_unique():
    numbers = [quote.quote_number for quote in SAMPLE_QUOTES]
    suppliers = [quote.supplier.casefold() for quote in SAMPLE_QUOTES]

    assert len(set(numbers)) == len(numbers)
    assert len(set(suppliers)) == len(suppliers)


def test_sample_workbook_parses_completely(tmp_path):
    parsed = parse_workbook(_sample_path(tmp_path), strict=False)

    assert not parsed.structure.missing_sheets
    assert parsed.structure.required_columns_present
    assert len(parsed.suppliers) == sample_supplier_count()
    assert parsed.rfq.quantity_pcs == SAMPLE_QUANTITY


def test_sample_workbook_passes_validation_without_errors(tmp_path):
    parsed = parse_workbook(_sample_path(tmp_path), strict=False)
    validation = validate(parsed, today=date.today())

    assert not validation.file_rejected
    assert validation.errors == ()
    assert parsed.rfq.rfq_number is not None


def test_sample_workbook_calculates_every_row(tmp_path):
    parsed = parse_workbook(_sample_path(tmp_path), strict=False)
    validation = validate(parsed, today=date.today())
    result = calculate(parsed, validation)

    assert len(result.rows) == sample_supplier_count()
    assert result.excluded == ()
    assert result.lowest_product_cost is not None
    assert result.lowest_total_cost is not None


def test_sample_workbook_triggers_the_intended_validation_codes(tmp_path):
    parsed = parse_workbook(_sample_path(tmp_path), strict=False)
    validation = validate(parsed, today=date.today())
    codes = _codes(validation.findings)

    # Kasıtlı senaryolar: karma ödeme, referans dışı kg/m, uzun termin,
    # eksik nakliye, gerekçesiz varsayım, süresi geçmiş teklif.
    assert "W-PAY-01" in codes
    assert "W-WT-01" in codes
    assert "W-DLV-02" in codes
    assert "I-SHIP-01" in codes
    assert "W-SRC-01" in codes
    assert "W-VAL-01" in codes


def test_sample_workbook_price_spread_and_lead_time_bottleneck():
    prices = [quote.unit_price for quote in SAMPLE_QUOTES]
    lead_times = [quote.delivery_days for quote in SAMPLE_QUOTES]

    # Fiyat uçurumu: en düşük ile en yüksek arasında en az %25 fark.
    assert max(prices) / min(prices) >= 1.25
    # Termin darboğazı: istenen 20 günün çok üzerinde bir teklif var.
    assert max(lead_times) >= 365
    assert sum(1 for days in lead_times if days <= 20) >= 5


def test_sample_workbook_has_missing_shipping_and_assumption_rows():
    missing = [quote for quote in SAMPLE_QUOTES if quote.shipping is None]
    assumptions = [
        quote
        for quote in SAMPLE_QUOTES
        if quote.shipping_source == "Customer Assumption"
    ]

    assert missing
    assert assumptions
    assert any(not quote.notes.strip() for quote in assumptions)


def test_sample_workbook_lowest_product_differs_from_lowest_total(tmp_path):
    """Ürün bedeli ile toplam maliyet işaretleri farklı satırlara düşer."""
    parsed = parse_workbook(_sample_path(tmp_path), strict=False)
    validation = validate(parsed, today=date.today())
    result = calculate(parsed, validation)

    assert result.lowest_product_cost is not None
    assert result.lowest_total_cost is not None
    assert result.lowest_product_cost.row_numbers != result.lowest_total_cost.row_numbers


def test_sample_workbook_has_one_missing_total_cost_row(tmp_path):
    parsed = parse_workbook(_sample_path(tmp_path), strict=False)
    validation = validate(parsed, today=date.today())
    result = calculate(parsed, validation)

    assert len(result.missing_data_rows) == 1


def test_sample_workbook_dates_follow_reference_day(tmp_path):
    reference = date(2030, 6, 15)
    parsed = parse_workbook(_sample_path(tmp_path, reference), strict=False)
    validation = validate(parsed, today=reference)

    assert validation.errors == ()
    assert parsed.rfq.rfq_date.date() == reference - timedelta(days=2)
    assert str(reference.year) in str(parsed.rfq.rfq_number)


def test_sample_workbook_produces_ai_request(tmp_path):
    """AI modülü örnek veriyle istek gövdesini üretebilir (ağ çağrısı yok)."""
    parsed = parse_workbook(_sample_path(tmp_path), strict=False)
    validation = validate(parsed, today=date.today())
    result = calculate(parsed, validation)

    request = ai_analysis.build_request(parsed, validation, result)

    assert ai_analysis.analysis_block_reason(validation, result) is None
    assert len(request.payload["tedarikciler"]) == sample_supplier_count()
    assert request.legend
    # Maskelenmiş etiketler gerçek tedarikçi adı içermez.
    assert all(label.startswith("Tedarikçi ") for label, _name in request.legend)


def test_sample_workbook_does_not_leak_real_names_into_ai_payload(tmp_path):
    import json

    parsed = parse_workbook(_sample_path(tmp_path), strict=False)
    validation = validate(parsed, today=date.today())
    result = calculate(parsed, validation)
    body = json.dumps(
        ai_analysis.build_request(parsed, validation, result).payload,
        ensure_ascii=False,
    )

    for quote in SAMPLE_QUOTES:
        assert quote.supplier not in body
        assert quote.quote_number not in body


def test_sample_workbook_missing_file_raises_demo_error(tmp_path):
    with pytest.raises(DemoDataError):
        build_sample_workbook(tmp_path / "yok.xlsx")


def test_demo_builders_do_not_write_files(tmp_path):
    before = {path.name for path in ROOT.iterdir()}

    build_blank_template(TEMPLATE)
    build_sample_workbook(TEMPLATE)

    after = {path.name for path in ROOT.iterdir()}
    assert after == before

"""``procurement.validator`` testleri — 46 kodlu kataloğun tamamı.

Test stratejisi
---------------
* Gerçek şablon ``tmp_path`` içine kopyalanır; asıl dosya hiç değiştirilmez.
* Her test tek bir kuralı ihlal eder ve ilgili kodu arar.
* ``CODE_TRIGGERS`` tablosu 46 kodun **tamamını** kapsar; böylece hiçbir kod
  test edilmeden kalmaz.
* Tarih bağımlılığı ``today=`` parametresiyle sabitlenir; testler sistem
  tarihine göre kırılmaz.
"""

from __future__ import annotations

import shutil
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

from procurement.excel_parser import parse_workbook
from procurement.validator import (
    ALL_CODES,
    ERROR_CODES,
    INFO_CODES,
    WARNING_CODES,
    Finding,
    Severity,
    ValidationResult,
    derive_payment_days,
    validate,
)

TEMPLATE = Path(__file__).resolve().parents[1] / "templates" / "supplier_quote_template.xlsx"

#: Testlerde sabit "bugün" — şablonun tarihleriyle uyumlu.
TODAY = date(2026, 11, 12)

pytestmark = pytest.mark.unit


# --------------------------------------------------------------------------- #
# Yardımcılar
# --------------------------------------------------------------------------- #


def _build(tmp_path: Path, mutate=None, *, name: str = "case.xlsx", approve: bool = True) -> Path:
    """Şablonun kopyasını üretir, isteğe bağlı değişiklik uygular."""
    path = tmp_path / name
    shutil.copyfile(TEMPLATE, path)

    workbook = load_workbook(path)
    try:
        if approve:
            workbook["Instructions"]["B4"] = "EVET"
        if mutate is not None:
            mutate(workbook)
    finally:
        workbook.save(path)
        workbook.close()

    return path


def _validate(path: Path) -> ValidationResult:
    return validate(parse_workbook(path, strict=False), today=TODAY)


def _codes(path: Path) -> tuple[str, ...]:
    return _validate(path).codes


def _set(sheet: str, cell: str, value) -> callable:
    def mutate(workbook) -> None:
        workbook[sheet][cell] = value

    return mutate


# --------------------------------------------------------------------------- #
# Katalog bütünlüğü
# --------------------------------------------------------------------------- #


def test_catalog_has_46_codes():
    assert len(ALL_CODES) == 46


def test_catalog_severity_distribution():
    assert len(ERROR_CODES) == 29
    assert len(WARNING_CODES) == 16
    assert len(INFO_CODES) == 1
    assert ERROR_CODES | WARNING_CODES | INFO_CODES == ALL_CODES


def test_catalog_matches_spec_codes():
    """Kod listesi ``EXCEL_TEMPLATE_SPEC.md`` §7.2 tablosuyla birebir aynı."""
    expected = {
        "E-DECL-01", "E-TPL-01", "E-TPL-02", "E-TPL-03",
        "E-SHEET-01", "E-HDR-01", "E-RFQ-01", "E-RFQ-02",
        "E-SPEC-01", "E-SPEC-02", "E-SPEC-03", "E-SPEC-04",
        "E-LEN-01", "E-LEN-02", "E-CUR-01", "E-QTY-01", "E-Q-01",
        "E-WT-01", "E-PRICE-01", "E-SHIP-01", "E-SRC-01", "E-PAY-01",
        "E-QD-01", "E-VAL-01", "E-NAME-01", "E-NAME-02", "E-DUP-01",
        "E-FMT-01", "E-FMT-02",
        "W-TPL-01", "W-SHEET-01", "W-HDR-01", "W-HDR-02",
        "W-RFQ-01", "W-RFQ-02", "W-WT-01", "W-PRICE-01", "W-SRC-01",
        "W-DLV-01", "W-DLV-02", "W-DLV-03", "W-PAY-01", "W-QD-01",
        "W-VAL-01", "W-DUP-02",
        "I-SHIP-01",
    }
    assert ALL_CODES == frozenset(expected)


# --------------------------------------------------------------------------- #
# Başarı senaryosu
# --------------------------------------------------------------------------- #


def test_clean_template_produces_no_errors(tmp_path):
    result = _validate(_build(tmp_path))

    assert result.errors == ()
    assert result.file_rejected is False
    assert result.blocked_row_numbers == ()
    assert result.is_complete is True


def test_clean_template_produces_expected_warnings_only(tmp_path):
    """Şablonun bilinçli örnekleri: karma ödeme ve aralık dışı kg/m."""
    result = _validate(_build(tmp_path))

    assert result.codes == ("W-PAY-01", "W-WT-01")
    assert len(result.warnings) == 2
    assert result.infos == ()


def test_findings_carry_code_severity_and_location(tmp_path):
    result = _validate(_build(tmp_path))
    finding = result.by_code("W-WT-01")[0]

    assert isinstance(finding, Finding)
    assert finding.code == "W-WT-01"
    assert finding.severity is Severity.WARNING
    assert finding.supplier == "İnegöl Çelik San."
    assert finding.row_number == 5
    assert finding.cell == "F5"
    assert finding.sheet == "Supplier Quotes"
    assert "2.27" in finding.message


def test_warning_does_not_block_row(tmp_path):
    """kg/m uyarısı bloklayıcı değildir (brief §6)."""
    result = _validate(_build(tmp_path))

    assert 5 not in result.blocked_row_numbers
    assert result.file_rejected is False


def test_result_contains_no_cost_or_ranking_fields():
    """Validator maliyet/sıralama/puan üretmez (görev kuralı 6)."""
    forbidden = ("cost", "total", "rank", "score", "winner", "cheapest", "best")

    for attribute in dir(ValidationResult):
        lowered = attribute.lower()
        assert not any(word in lowered for word in forbidden), attribute


# --------------------------------------------------------------------------- #
# 46 kodun tamamı için tetikleyici tablo
# --------------------------------------------------------------------------- #


def _swap_headers(workbook) -> None:
    sheet = workbook["Supplier Quotes"]
    sheet["A2"], sheet["B2"] = "Supplier Quote Number", "Supplier Name"


def _duplicate_supplier(workbook) -> None:
    workbook["Supplier Quotes"]["A4"] = "Bursa Metal A.Ş."


def _duplicate_supplier_and_number(workbook) -> None:
    sheet = workbook["Supplier Quotes"]
    sheet["A4"] = "Bursa Metal A.Ş."
    sheet["B4"] = "BM-2026-1187"


def _customer_assumption_without_note(workbook) -> None:
    sheet = workbook["Supplier Quotes"]
    sheet["I3"] = "Customer Assumption"
    sheet["N3"] = None


def _delete_rfq_sheet(workbook) -> None:
    del workbook["RFQ"]


def _add_extra_sheet(workbook) -> None:
    workbook.create_sheet("Notes")


def _blank_header(workbook) -> None:
    workbook["Supplier Quotes"]["H2"] = None


def _expired_but_ordered(workbook) -> None:
    """Geçerlilik tarihi geçmiş, ama teklif tarihinden sonra (W-VAL-01)."""
    sheet = workbook["Supplier Quotes"]
    sheet["L3"] = date(2026, 1, 5)
    sheet["M3"] = date(2026, 2, 1)


def _duplicate_header(workbook) -> None:
    """Var olan bir başlığı fazladan sütuna kopyalar (W-HDR-01)."""
    workbook["Supplier Quotes"]["O2"] = "Supplier Name"


#: (kod, şablon değişikliği). Her kod en az bir kez tetiklenir.
CODE_TRIGGERS: tuple[tuple[str, callable], ...] = (
    ("E-DECL-01", _set("Instructions", "B4", "HAYIR")),
    ("E-TPL-01", _set("Instructions", "B3", "BASKASABLON")),
    ("E-TPL-02", _set("Instructions", "B2", "2.0")),
    ("E-TPL-03", _blank_header),
    ("W-TPL-01", _set("Instructions", "B2", "1.1")),
    ("E-SHEET-01", _delete_rfq_sheet),
    ("W-SHEET-01", _add_extra_sheet),
    ("E-HDR-01", _set("Supplier Quotes", "A2", "Supplier name")),
    ("W-HDR-01", _set("Supplier Quotes", "O2", "Calculated Total")),
    ("W-HDR-02", _swap_headers),
    ("E-RFQ-01", _set("RFQ", "B10", None)),
    ("E-RFQ-02", _set("RFQ", "B9", 0)),
    ("W-RFQ-01", _set("RFQ", "B3", "TALEP-1")),
    ("W-RFQ-02", _set("RFQ", "B4", date(2027, 1, 1))),
    ("E-SPEC-01", _set("Supplier Quotes", "C3", "50x50x2")),
    ("E-SPEC-02", _set("RFQ", "B8", 3)),
    ("E-SPEC-03", _set("RFQ", "B6", "50x50x2")),
    ("E-SPEC-04", _set("RFQ", "B5", "BASKAKOD")),
    ("E-LEN-01", _set("Supplier Quotes", "D3", 12)),
    ("E-LEN-02", _set("RFQ", "B7", 12)),
    ("E-CUR-01", _set("RFQ", "B12", "USD")),
    ("E-QTY-01", _set("Supplier Quotes", "E3", 0)),
    ("E-Q-01", _set("Supplier Quotes", "E3", 400)),
    ("E-WT-01", _set("Supplier Quotes", "F3", 0)),
    ("W-WT-01", _set("Supplier Quotes", "F3", 2.90)),
    ("E-PRICE-01", _set("Supplier Quotes", "G3", 0)),
    ("W-PRICE-01", _set("Supplier Quotes", "G3", 2_000_000)),
    ("E-SHIP-01", _set("Supplier Quotes", "H3", -100)),
    ("I-SHIP-01", _set("Supplier Quotes", "H3", None)),
    ("E-SRC-01", _set("Supplier Quotes", "I3", None)),
    ("W-SRC-01", _customer_assumption_without_note),
    ("W-DLV-01", _set("Supplier Quotes", "J3", None)),
    ("W-DLV-02", _set("Supplier Quotes", "J3", 400)),
    ("W-DLV-03", _set("RFQ", "B11", None)),
    ("E-PAY-01", _set("Supplier Quotes", "K3", None)),
    ("W-PAY-01", _set("Supplier Quotes", "K3", "%50 peşin, %50 teslimatta")),
    ("E-QD-01", _set("Supplier Quotes", "L3", None)),
    ("W-QD-01", _set("Supplier Quotes", "L3", date(2027, 1, 1))),
    ("E-VAL-01", _set("Supplier Quotes", "M3", date(2026, 11, 1))),
    ("W-VAL-01", _expired_but_ordered),
    ("E-NAME-01", _set("Supplier Quotes", "A3", None)),
    ("E-NAME-02", _set("Supplier Quotes", "B3", None)),
    ("E-DUP-01", _duplicate_supplier),
    ("W-DUP-02", _duplicate_supplier_and_number),
    ("E-FMT-01", _set("Supplier Quotes", "J3", "onbes gün")),
    ("E-FMT-02", _set("Supplier Quotes", "L3", "12/11/2026")),
)


def test_trigger_table_covers_every_code():
    assert {code for code, _mutate in CODE_TRIGGERS} == ALL_CODES


@pytest.mark.parametrize(("code", "mutate"), CODE_TRIGGERS, ids=[c for c, _ in CODE_TRIGGERS])
def test_code_is_triggered(tmp_path, code, mutate):
    """Her katalog kodu gerçek çalışma akışında üretilebilir."""
    path = _build(tmp_path, mutate, name=f"{code}.xlsx")

    assert code in _codes(path)


# --------------------------------------------------------------------------- #
# Dosya düzeyi zincir (spec §2.3)
# --------------------------------------------------------------------------- #


def test_vat_declaration_short_circuits_chain(tmp_path):
    """Beyan yoksa sonraki adımlar çalışmaz; tek ve açık mesaj döner."""
    path = _build(tmp_path, approve=False)

    result = _validate(path)

    assert result.codes == ("E-DECL-01",)
    assert result.file_rejected is True
    assert "B4" == result.errors[0].cell


def test_template_code_error_stops_version_checks(tmp_path):
    path = _build(tmp_path, _set("Instructions", "B3", "BASKASABLON"))

    assert _codes(path) == ("E-TPL-01",)


def test_template_code_and_major_version_are_independent(tmp_path):
    path = _build(tmp_path, _set("Instructions", "B2", "2.0"))

    assert _codes(path) == ("E-TPL-02",)


def test_missing_column_with_correct_version_is_tpl03(tmp_path):
    path = _build(tmp_path, _blank_header)

    codes = _codes(path)
    assert "E-TPL-03" in codes
    assert "E-TPL-01" not in codes


def test_corrupt_header_is_hdr01_not_tpl03(tmp_path):
    """Hücre dolu ama metin yanlış → başlık satırı bozuk (E-HDR-01)."""
    path = _build(tmp_path, _set("Supplier Quotes", "A2", "Supplier name"))

    codes = _codes(path)
    assert "E-HDR-01" in codes
    assert "E-TPL-03" not in codes


def test_minor_version_difference_only_warns(tmp_path):
    path = _build(tmp_path, _set("Instructions", "B2", "1.1"))

    result = _validate(path)

    assert "W-TPL-01" in result.codes
    assert result.errors == ()
    assert result.file_rejected is False


def test_missing_sheet_reported_without_crashing(tmp_path):
    path = _build(tmp_path, _delete_rfq_sheet)

    result = _validate(path)

    assert "E-SHEET-01" in result.codes
    assert result.file_rejected is True


def test_missing_quotes_sheet_is_sheet_error(tmp_path):
    def mutate(workbook):
        del workbook["Supplier Quotes"]

    path = _build(tmp_path, mutate)
    codes = _codes(path)

    assert "E-SHEET-01" in codes


def test_missing_sheet_short_circuits_row_validation(tmp_path):
    """E-SHEET-01 dosyayı reddeder; satır doğrulaması çalışmaz (S5)."""
    path = _build(tmp_path, _delete_rfq_sheet)
    result = _validate(path)

    assert result.file_rejected is True
    # Şablonun örnek satırlarından gelen uyarılar da üretilmemeli.
    assert result.warnings == ()
    assert result.blocked_row_numbers == ()


def test_duplicate_header_is_reported_as_unexpected_column(tmp_path):
    """Mükerrer başlık sessizce kabul edilmez (S1)."""
    path = _build(tmp_path, _duplicate_header)
    result = _validate(path)

    findings = result.by_code("W-HDR-01")
    assert len(findings) == 1
    assert findings[0].severity is Severity.WARNING
    assert findings[0].cell == "O2"
    assert findings[0].sheet == "Supplier Quotes"
    assert "Supplier Name" in findings[0].message


def test_duplicate_header_does_not_block_file(tmp_path):
    """Mükerrer sütun veri olarak kullanılmaz; okuma sürer (spec §2.3)."""
    path = _build(tmp_path, _duplicate_header)
    result = _validate(path)

    assert result.file_rejected is False
    assert "E-TPL-03" not in result.codes
    assert "E-HDR-01" not in result.codes


def test_duplicate_header_values_are_not_read_as_data(tmp_path):
    """Mükerrer sütuna yazılan veri yok sayılır; satırlar etkilenmez."""
    def mutate(workbook):
        sheet = workbook["Supplier Quotes"]
        sheet["O2"] = "Supplier Name"
        sheet["O3"] = "Hayalet Tedarikçi"

    result = _validate(_build(tmp_path, mutate))

    assert len(result.by_code("W-HDR-01")) == 1
    assert "Hayalet Tedarikçi" not in {f.supplier for f in result.findings}


# --------------------------------------------------------------------------- #
# RFQ kuralları
# --------------------------------------------------------------------------- #


def test_rfq_required_fields_each_reported(tmp_path):
    def mutate(workbook):
        sheet = workbook["RFQ"]
        sheet["B3"] = None
        sheet["B10"] = None

    result = _validate(_build(tmp_path, mutate))
    blanks = [f for f in result.by_code("E-RFQ-01")]

    assert len(blanks) >= 2
    assert {f.cell for f in blanks} >= {"B3", "B10"}


def test_rfq_quantity_fractional_rejected(tmp_path):
    path = _build(tmp_path, _set("RFQ", "B9", 12.5))

    assert "E-RFQ-02" in _codes(path)


def test_rfq_future_date_warns(tmp_path):
    path = _build(tmp_path, _set("RFQ", "B4", date(2027, 6, 1)))
    result = _validate(path)

    assert "W-RFQ-02" in result.codes
    assert result.errors == ()


def test_rfq_missing_delivery_days_warns(tmp_path):
    path = _build(tmp_path, _set("RFQ", "B11", None))
    result = _validate(path)

    assert "W-DLV-03" in result.codes
    assert result.errors == ()


def test_rfq_unreadable_date_is_format_error(tmp_path):
    path = _build(tmp_path, _set("RFQ", "B4", "12/11/2026"))

    result = _validate(path)
    assert "E-FMT-02" in result.codes
    # RFQ zorunlu alanı: dosya düzeyi hata olmalı (S2).
    assert result.file_rejected is True
    assert result.by_code("E-FMT-02")[0].is_file_level is True


def test_rfq_unreadable_quantity_is_file_level_error(tmp_path):
    """Okunamayan RFQ miktarı dosyayı reddeder; sahipsiz hata kalmaz (S2)."""
    path = _build(tmp_path, _set("RFQ", "B9", "abc"))
    result = _validate(path)

    assert "E-RFQ-02" in result.codes
    assert result.file_rejected is True
    finding = result.by_code("E-RFQ-02")[0]
    assert finding.is_file_level is True
    assert finding.row_number is None
    assert finding.sheet == "RFQ"
    assert finding.cell == "B9"


def test_invalid_rfq_quantity_does_not_emit_generic_format_code(tmp_path):
    """B9'un kendi alan kodu vardır; E-FMT-01 çoğaltılmaz (S3)."""
    path = _build(tmp_path, _set("RFQ", "B9", "abc"))

    assert "E-FMT-01" not in _codes(path)


def test_unusable_rfq_quantity_blocks_row_level_checks(tmp_path):
    """Miktar çözülemiyorsa E-Q-01 sessizce yanlış sonuç üretmez (S3)."""
    def mutate(workbook):
        workbook["RFQ"]["B9"] = "abc"
        workbook["Supplier Quotes"]["E3"] = 999

    result = _validate(_build(tmp_path, mutate))

    assert "E-RFQ-02" in result.codes
    assert "E-Q-01" not in result.codes
    # Dosya reddedildiği için tedarikçi satırı bulguları hiç üretilmez.
    assert result.blocked_row_numbers == ()
    assert "E-RFQ-02" in {f.code for f in result.file_errors}


def test_valid_rfq_quantity_still_checks_rows(tmp_path):
    """Kontrol: miktar geçerliyken E-Q-01 çalışmaya devam eder (S3)."""
    path = _build(tmp_path, _set("Supplier Quotes", "E3", 999))
    result = _validate(path)

    assert "E-Q-01" in result.codes
    assert result.by_code("E-Q-01")[0].row_number == 3


@pytest.mark.parametrize(
    ("cell", "code"),
    [("B7", "E-LEN-02"), ("B8", "E-SPEC-02"), ("B6", "E-SPEC-03"), ("B5", "E-SPEC-04")],
)
def test_blank_fixed_rfq_field_is_error(tmp_path, cell, code):
    """Sabit alanlar boş bırakılamaz (spec §5)."""
    path = _build(tmp_path, _set("RFQ", cell, None))

    assert code in _codes(path)


# --------------------------------------------------------------------------- #
# Satır kuralları (tek tedarikçi)
# --------------------------------------------------------------------------- #


def test_missing_supplier_name_reports_name_and_row(tmp_path):
    path = _build(tmp_path, _set("Supplier Quotes", "A3", None))
    result = _validate(path)
    finding = result.by_code("E-NAME-01")[0]

    assert finding.row_number == 3
    assert finding.supplier is None
    assert finding.cell == "A3"
    assert 3 in result.blocked_row_numbers


def test_quantity_mismatch_mentions_both_values(tmp_path):
    path = _build(tmp_path, _set("Supplier Quotes", "E3", 400))
    finding = _validate(path).by_code("E-Q-01")[0]

    assert "400" in finding.message
    assert "500" in finding.message
    assert finding.supplier == "Bursa Metal A.Ş."


def test_weight_out_of_reference_range_is_warning_only(tmp_path):
    """2,27–2,51 dışı kg/m bloklamaz; hesaplama beyanla yapılır."""
    path = _build(tmp_path, _set("Supplier Quotes", "F3", 2.90))
    result = _validate(path)

    assert "W-WT-01" in result.codes
    assert "E-WT-01" not in result.codes
    assert 3 not in result.blocked_row_numbers


def test_zero_weight_is_blocking_error(tmp_path):
    path = _build(tmp_path, _set("Supplier Quotes", "F3", 0))
    result = _validate(path)

    assert "E-WT-01" in result.codes
    assert 3 in result.blocked_row_numbers
    # Satır 3 için referans aralığı uyarısı üretilmez; şablonun 5. satırındaki
    # 2,58 kg/m uyarısı bundan bağımsızdır.
    assert 3 not in {f.row_number for f in result.by_code("W-WT-01")}


@pytest.mark.parametrize("value", ["28,50 TL", "$28.50", "1.234,56", "abc"])
def test_unreadable_price_uses_price_code_only(tmp_path, value):
    """Okunamayan TL/kg alanın kendi koduyla raporlanır (spec §7.2).

    Aynı hücre için ayrıca genel ``E-FMT-01`` üretilmez.
    """
    path = _build(tmp_path, _set("Supplier Quotes", "G3", value))
    result = _validate(path)

    assert "E-PRICE-01" in result.codes
    assert "E-FMT-01" not in result.codes
    assert 3 in result.blocked_row_numbers


@pytest.mark.parametrize("value", ["abc", "2,39 TL"])
def test_unreadable_weight_uses_weight_code_only(tmp_path, value):
    """Okunamayan kg/m ``E-WT-01`` ile raporlanır; ``E-FMT-01`` çoğaltılmaz."""
    path = _build(tmp_path, _set("Supplier Quotes", "F3", value))
    result = _validate(path)

    assert "E-WT-01" in result.codes
    assert "E-FMT-01" not in result.codes
    assert 3 in result.blocked_row_numbers


def test_other_unreadable_numbers_still_use_format_code(tmp_path):
    """Alan kodu olmayan sayısal alanlar ``E-FMT-01`` kullanmaya devam eder."""
    path = _build(tmp_path, _set("Supplier Quotes", "J3", "onbes gün"))
    result = _validate(path)

    assert "E-FMT-01" in result.codes
    assert 3 in result.blocked_row_numbers


def test_comma_decimal_is_accepted(tmp_path):
    """Spec §3.4: ``28,50`` sayı olarak yorumlanabiliyorsa kabul edilir."""
    path = _build(tmp_path, _set("Supplier Quotes", "G3", "28,50"))
    result = _validate(path)

    assert "E-FMT-01" not in result.codes
    assert "E-PRICE-01" not in result.codes


def test_blank_shipping_requests_information(tmp_path):
    """Nakliye boş ≠ 0; bilgi talebi üretilir ve karşılaştırma tamamlanmaz."""
    path = _build(tmp_path, _set("Supplier Quotes", "H3", None))
    result = _validate(path)

    assert "I-SHIP-01" in result.codes
    assert result.is_complete is False
    assert result.errors == ()


def test_shipping_source_required_when_cost_present(tmp_path):
    path = _build(tmp_path, _set("Supplier Quotes", "I3", None))

    assert "E-SRC-01" in _codes(path)


def test_customer_assumption_requires_note(tmp_path):
    path = _build(tmp_path, _customer_assumption_without_note)
    result = _validate(path)
    finding = result.by_code("W-SRC-01")[0]

    assert finding.cell == "N3"
    assert finding.severity is Severity.WARNING


def test_notes_content_is_never_scanned(tmp_path):
    """Notes serbest metindir; KDV ifadeleri hata üretmez (spec §2.1.2)."""
    def mutate(workbook):
        workbook["Supplier Quotes"]["N3"] = "KDV dahil fiyat verildi, teyit edildi"

    result = _validate(_build(tmp_path, mutate))

    assert result.codes == ("W-PAY-01", "W-WT-01")


@pytest.mark.parametrize(
    ("terms", "expected_days"),
    [
        ("Peşin", 0),
        ("peşin ödeme", 0),
        ("Net 30", 30),
        ("Net 45", 45),
        ("30 gün vadeli", 30),
    ],
)
def test_simple_payment_terms_do_not_warn(tmp_path, terms, expected_days):
    path = _build(tmp_path, _set("Supplier Quotes", "K3", terms))
    result = _validate(path)

    # Şablonun 4. satırındaki karma koşul ayrı bir bulgudur; burada yalnızca
    # 3. satırın karma sayılmadığını doğruluyoruz.
    assert 3 not in {f.row_number for f in result.by_code("W-PAY-01")}
    assert derive_payment_days(terms) == expected_days


@pytest.mark.parametrize(
    "terms",
    ["%50 peşin, %50 teslimatta", "30 gün çek + 30 gün vade", "Görüşülecek"],
)
def test_complex_payment_terms_warn(tmp_path, terms):
    path = _build(tmp_path, _set("Supplier Quotes", "K3", terms))
    result = _validate(path)
    finding = result.by_code("W-PAY-01")[0]

    assert finding.severity is Severity.WARNING
    assert 3 not in result.blocked_row_numbers


def test_valid_until_before_quote_date_is_error(tmp_path):
    path = _build(tmp_path, _set("Supplier Quotes", "M3", date(2026, 11, 1)))
    result = _validate(path)

    assert "E-VAL-01" in result.codes
    assert "W-VAL-01" not in result.codes


def test_expired_valid_until_is_warning(tmp_path):
    """Teklif tarihinden sonra ama bugünden önce → yalnızca uyarı."""

    def mutate(workbook):
        sheet = workbook["Supplier Quotes"]
        sheet["L3"] = date(2026, 1, 5)
        sheet["M3"] = date(2026, 2, 1)

    result = _validate(_build(tmp_path, mutate))

    assert "W-VAL-01" in result.codes
    assert "E-VAL-01" not in result.codes
    assert result.errors == ()


def test_future_quote_date_is_warning(tmp_path):
    def mutate(workbook):
        sheet = workbook["Supplier Quotes"]
        sheet["L3"] = date(2027, 1, 1)
        sheet["M3"] = None

    result = _validate(_build(tmp_path, mutate))

    assert "W-QD-01" in result.codes
    assert result.errors == ()


# --------------------------------------------------------------------------- #
# Mükerrer tedarikçi
# --------------------------------------------------------------------------- #


def test_duplicate_supplier_blocks_both_rows(tmp_path):
    path = _build(tmp_path, _duplicate_supplier)
    result = _validate(path)
    duplicates = result.by_code("E-DUP-01")

    assert len(duplicates) == 2
    assert {f.row_number for f in duplicates} == {3, 4}
    assert {3, 4} <= set(result.blocked_row_numbers)


def test_duplicate_is_case_and_space_insensitive(tmp_path):
    def mutate(workbook):
        workbook["Supplier Quotes"]["A4"] = "  bursa metal a.ş.  "

    result = _validate(_build(tmp_path, mutate))

    assert "E-DUP-01" in result.codes


def test_same_supplier_same_quote_number_warns(tmp_path):
    path = _build(tmp_path, _duplicate_supplier_and_number)
    result = _validate(path)

    assert "W-DUP-02" in result.codes
    assert {f.row_number for f in result.by_code("W-DUP-02")} == {3, 4}


def test_different_suppliers_are_not_duplicates(tmp_path):
    result = _validate(_build(tmp_path))

    assert "E-DUP-01" not in result.codes
    assert "W-DUP-02" not in result.codes


# --------------------------------------------------------------------------- #
# Çok tedarikçili senaryolar
# --------------------------------------------------------------------------- #


def test_multiple_suppliers_are_validated_independently(tmp_path):
    """Bir satırın hatası diğer satırları etkilemez (spec §7.1)."""
    def mutate(workbook):
        sheet = workbook["Supplier Quotes"]
        sheet["G3"] = 0          # satır 3 hatalı
        sheet["G4"] = 31.50      # satır 4 geçerli

    result = _validate(_build(tmp_path, mutate))

    assert 3 in result.blocked_row_numbers
    assert 4 not in result.blocked_row_numbers
    assert 5 not in result.blocked_row_numbers
    assert result.by_code("E-PRICE-01")[0].supplier == "Bursa Metal A.Ş."


def test_findings_are_attributed_to_correct_suppliers(tmp_path):
    def mutate(workbook):
        sheet = workbook["Supplier Quotes"]
        sheet["F4"] = 0          # Gemlik
        sheet["K5"] = None       # İnegöl

    result = _validate(_build(tmp_path, mutate))

    assert result.by_code("E-WT-01")[0].supplier == "Gemlik Profil Ltd."
    assert result.by_code("E-WT-01")[0].row_number == 4
    assert result.by_code("E-PAY-01")[0].supplier == "İnegöl Çelik San."
    assert result.by_code("E-PAY-01")[0].row_number == 5


def test_error_count_scales_with_rows(tmp_path):
    def mutate(workbook):
        sheet = workbook["Supplier Quotes"]
        for row in (3, 4, 5):
            sheet[f"G{row}"] = 0

    result = _validate(_build(tmp_path, mutate))

    assert len(result.by_code("E-PRICE-01")) == 3
    assert set(result.blocked_row_numbers) == {3, 4, 5}


def test_file_level_findings_have_no_row(tmp_path):
    path = _build(tmp_path, _set("RFQ", "B12", "USD"))
    result = _validate(path)
    finding = result.by_code("E-CUR-01")[0]

    assert finding.row_number is None
    assert finding.is_file_level is True
    assert finding.sheet == "RFQ"


# --------------------------------------------------------------------------- #
# Sorumluluk sınırı
# --------------------------------------------------------------------------- #


def test_validator_does_not_compute_weight_or_cost(tmp_path):
    """Validator yalnızca doğrular; hesaplanmış alan döndürmez."""
    result = _validate(_build(tmp_path))

    for finding in result.findings:
        assert not hasattr(finding, "total_weight")
        assert not hasattr(finding, "product_cost")
        assert not hasattr(finding, "total_cost")


def test_validator_does_not_modify_source_file(tmp_path):
    import hashlib

    path = _build(tmp_path)
    before = hashlib.sha256(path.read_bytes()).hexdigest()

    _validate(path)

    assert hashlib.sha256(path.read_bytes()).hexdigest() == before

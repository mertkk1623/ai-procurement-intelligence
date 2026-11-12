"""``procurement.report`` testleri — Excel ve PDF rapor üretimi.

Kapsam
------
* Excel raporunun açılabilir ``.xlsx`` üretmesi, sayfaların ve tablo
  sütunlarının beklenen içeriği taşıması.
* PDF raporunun geçerli bir PDF (``%PDF``) üretmesi, AI yorumu varken ve
  yokken çalışması.
* Eksik nakliyenin **sıfır olarak gösterilmemesi**; ``Eksik veri`` metni.
* Türkçe karakterlerin (ş, ğ, İ, ı, ç, ö, ü) korunması.
* Boş dosya, reddedilen dosya ve karşılaştırma dışı satırlar.
* AI yorumunun raporda ayrı ve etiketli bölüm olması.
* Beklenmeyen üretim hatasının :class:`ReportError` olarak toplanması.

Testler gerçek akışı (``parse_workbook`` → ``validate`` → ``calculate``)
kullanır; rapor modülü hiçbir tutarı yeniden hesaplamaz. Bu yüzden beklenen
değerler doğrudan :class:`CalculationResult` çıktısından okunur ve raporun
**onu değiştirmeden** yansıttığı doğrulanır.
"""

from __future__ import annotations

import base64
import io
import re
import shutil
import zlib
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

from procurement.ai_analysis import AIAnalysis, AiSection, SECTION_TITLES
from procurement.calculator import CalculationResult, calculate
from procurement.excel_parser import ParsedWorkbook, parse_workbook
from procurement.presentation import MISSING_TEXT
from procurement.report import (
    EXCEL_FILENAME,
    EXCEL_MIME,
    PDF_FILENAME,
    PDF_MIME,
    ReportError,
    build_excel_report,
    build_pdf_report,
)
from procurement.validator import ValidationResult, validate

TEMPLATE = (
    Path(__file__).resolve().parents[1]
    / "templates"
    / "supplier_quote_template.xlsx"
)
TODAY = date(2026, 11, 12)

pytestmark = pytest.mark.integration

COMPARISON_SHEET = "Tedarikçi Karşılaştırma"
SUMMARY_SHEET = "Yönetici Özeti"
VALIDATION_SHEET = "Doğrulama"
AI_SHEET = "AI Yorumu"

#: Excel karşılaştırma sayfasının sütun başlıkları (sırayla).
_COMPARISON_HEADERS: tuple[str, ...] = (
    "Tedarikçi",
    "Teklif No",
    "Miktar (adet)",
    "Ağırlık (kg/m)",
    "Toplam uzunluk (m)",
    "Toplam ağırlık (kg)",
    "Birim fiyat (TL/kg)",
    "Ürün bedeli (TL)",
    "Nakliye (TL)",
    "Nakliye kaynağı",
    "Nakliye (TL/kg)",
    "Toplam tedarik maliyeti (TL)",
    "Termin (gün)",
    "Ödeme koşulları",
    "Vade (gün)",
    "En düşük ürün bedeli",
    "En düşük toplam maliyet",
)


# --------------------------------------------------------------------------- #
# Yardımcılar
# --------------------------------------------------------------------------- #


def _build(tmp_path: Path, mutate=None, *, name: str = "rapor.xlsx") -> Path:
    """Şablonu kopyalar, AI onayını verir ve senaryo değişikliğini uygular."""
    path = tmp_path / name
    shutil.copyfile(TEMPLATE, path)

    workbook = load_workbook(path)
    try:
        workbook["Instructions"]["B4"] = "EVET"
        if mutate is not None:
            mutate(workbook)
    finally:
        workbook.save(path)
        workbook.close()

    return path


def _analyse(
    tmp_path: Path, mutate=None
) -> tuple[ParsedWorkbook, ValidationResult, CalculationResult]:
    path = _build(tmp_path, mutate)
    parsed = parse_workbook(path, strict=False)
    validation = validate(parsed, today=TODAY)
    return parsed, validation, calculate(parsed, validation)


def _set(sheet: str, cell: str, value):
    def mutate(workbook) -> None:
        workbook[sheet][cell] = value

    return mutate


def _excel(parsed, validation, result, **kwargs) -> bytes:
    return build_excel_report(
        parsed,
        validation,
        result,
        file_name="teklif.xlsx",
        generated_at=TODAY,
        **kwargs,
    )


def _pdf(parsed, validation, result, **kwargs) -> bytes:
    return build_pdf_report(
        parsed,
        validation,
        result,
        file_name="teklif.xlsx",
        generated_at=TODAY,
        **kwargs,
    )


def _sheet_rows(payload: bytes, sheet: str) -> list[tuple]:
    workbook = load_workbook(io.BytesIO(payload))
    return [row for row in workbook[sheet].iter_rows(values_only=True)]


def _analysis() -> AIAnalysis:
    return AIAnalysis(
        sections=tuple(
            AiSection(key=key, title=title, body=f"{title} metni.")
            for key, title in SECTION_TITLES
        ),
        model="test-model",
        legend=(("T1", "Bursa Metal A.Ş."),),
    )


def _comparison_table(payload: bytes) -> tuple[list[str], list[tuple]]:
    """Karşılaştırma sayfasından başlık satırını ve veri satırlarını ayırır."""
    rows = _sheet_rows(payload, COMPARISON_SHEET)
    header_index = next(
        index for index, row in enumerate(rows) if row and row[0] == "Tedarikçi"
    )
    header = [str(cell) for cell in rows[header_index] if cell is not None]
    # Veri satırları miktar sütununda sayı taşır; alt not satırı taşımaz.
    data = [
        row
        for row in rows[header_index + 1 :]
        if row and isinstance(row[2], (int, float))
    ]
    return header, data


# --------------------------------------------------------------------------- #
# Excel — yapı ve içerik
# --------------------------------------------------------------------------- #


def test_excel_report_is_valid_workbook(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    payload = _excel(parsed, validation, result)

    assert payload.startswith(b"PK")  # xlsx = zip
    assert EXCEL_MIME.endswith("spreadsheetml.sheet")
    assert EXCEL_FILENAME.endswith(".xlsx")

    workbook = load_workbook(io.BytesIO(payload))
    assert workbook.sheetnames[:3] == [
        SUMMARY_SHEET,
        COMPARISON_SHEET,
        VALIDATION_SHEET,
    ]


def test_excel_report_includes_rfq_and_product_information(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    rows = _sheet_rows(_excel(parsed, validation, result), SUMMARY_SHEET)
    values = [row[1] for row in rows if row and row[0]]

    assert parsed.rfq.rfq_number in values
    assert parsed.rfq.product_code in values
    assert parsed.rfq.profile_specification in values

    labels = [row[0] for row in rows if row and row[0]]
    assert "Et kalınlığı (mm)" in labels
    assert "Uzunluk (m)" in labels
    assert "Miktar (adet)" in labels
    assert "Teslim yeri" in labels
    assert "İstenen termin (gün)" in labels


def test_excel_comparison_columns_cover_all_calculated_fields(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    header, data = _comparison_table(_excel(parsed, validation, result))

    expected = _COMPARISON_HEADERS
    assert tuple(header) == expected
    assert len(data) == len(result.rows)


def test_excel_comparison_values_come_from_calculator(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    _header, data = _comparison_table(_excel(parsed, validation, result))

    for row, calculation in zip(data, result.rows):
        assert row[0] == calculation.supplier_name
        assert row[1] == calculation.supplier_quote_number
        assert row[2] == calculation.quantity_pcs
        assert row[3] == pytest.approx(float(calculation.weight_kg_per_m))
        assert row[4] == pytest.approx(float(calculation.total_length_m))
        assert row[5] == pytest.approx(float(calculation.total_weight_kg))
        assert row[6] == pytest.approx(float(calculation.unit_price_try_per_kg))
        assert row[7] == pytest.approx(float(calculation.product_cost_try))
        assert row[8] == pytest.approx(float(calculation.shipping_cost_try))
        assert row[10] == pytest.approx(float(calculation.shipping_per_kg_try))
        assert row[11] == pytest.approx(float(calculation.total_supply_cost_try))


def test_excel_report_marks_lowest_product_and_total(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    _header, data = _comparison_table(_excel(parsed, validation, result))

    lowest_product = {
        calculation.supplier_name
        for calculation in result.rows
        if result.lowest_product_cost is not None
        and calculation.row_number in result.lowest_product_cost.row_numbers
    }
    lowest_total = {
        calculation.supplier_name
        for calculation in result.rows
        if result.lowest_total_cost is not None
        and calculation.row_number in result.lowest_total_cost.row_numbers
    }

    marked_product = {row[0] for row in data if row[15] == "✓"}
    marked_total = {row[0] for row in data if row[16] == "✓"}

    assert marked_product == lowest_product
    assert marked_total == lowest_total
    assert marked_product  # şablonda en az bir işaret beklenir


def test_excel_report_has_filter_and_frozen_header(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    workbook = load_workbook(io.BytesIO(_excel(parsed, validation, result)))
    sheet = workbook[COMPARISON_SHEET]

    assert sheet.auto_filter.ref is not None
    assert sheet.freeze_panes is not None


def test_excel_report_uses_number_formats(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    workbook = load_workbook(io.BytesIO(_excel(parsed, validation, result)))
    sheet = workbook[COMPARISON_SHEET]

    header_row = next(
        row[0].row
        for row in sheet.iter_rows(min_row=1, max_col=1)
        if row[0].value == "Tedarikçi"
    )
    price = float(parsed.suppliers[0].unit_price_try_per_kg)
    formats = {
        sheet.cell(row=header_row + 1, column=column).value: (
            sheet.cell(row=header_row + 1, column=column).number_format
        )
        for column in range(1, len(_COMPARISON_HEADERS) + 1)
    }

    assert "TL" in formats[price]
    assert formats[parsed.suppliers[0].quantity_pcs] == "#,##0"


def test_excel_report_contains_validation_findings(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    rows = _sheet_rows(_excel(parsed, validation, result), VALIDATION_SHEET)
    codes = {row[1] for row in rows if row and isinstance(row[1], str)}

    for finding in validation.findings:
        assert finding.code in codes


def test_excel_report_contains_method_notes(tmp_path):
    """KDV hariç yaklaşım ve karar sınırı her rapor türünde yazılıdır."""
    parsed, validation, result = _analyse(tmp_path)
    rows = _sheet_rows(_excel(parsed, validation, result), SUMMARY_SHEET)
    text = " ".join(str(cell) for row in rows for cell in row if cell is not None)

    assert "KDV hariç" in text
    assert "sıfır kabul edilmez" in text
    assert "Otomatik kazanan belirlenmez" in text


# --------------------------------------------------------------------------- #
# Excel — eksik veri
# --------------------------------------------------------------------------- #


def test_excel_missing_shipping_is_not_zero(tmp_path):
    parsed, validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "H3", None)
    )
    header, data = _comparison_table(_excel(parsed, validation, result))
    shipping_column = header.index("Nakliye (TL)")
    total_column = header.index("Toplam tedarik maliyeti (TL)")
    first = data[0]
    assert first[shipping_column] == MISSING_TEXT
    assert first[total_column] == MISSING_TEXT
    assert first[shipping_column] != 0
    assert first[total_column] != 0


def test_excel_missing_shipping_keeps_product_cost(tmp_path):
    """Eksik nakliye ürün bedelini silmez; yalnızca toplamı boş bırakır."""
    parsed, validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "H3", None)
    )
    header, data = _comparison_table(_excel(parsed, validation, result))
    product_column = header.index("Ürün bedeli (TL)")

    calculation = result.rows[0]
    assert calculation.product_cost_try is not None
    assert data[0][product_column] == pytest.approx(
        float(calculation.product_cost_try)
    )


def test_excel_missing_shipping_is_listed_in_validation_sheet(tmp_path):
    parsed, validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "H3", None)
    )
    rows = _sheet_rows(_excel(parsed, validation, result), VALIDATION_SHEET)
    text = " ".join(str(cell) for row in rows for cell in row if cell is not None)

    assert "Nakliye bilgisi eksik" in text
    assert "sıfır kabul edilmez" in text


def test_excel_missing_shipping_does_not_mark_total_lowest(tmp_path):
    """Nakliyesi eksik satır toplam maliyet karşılaştırmasına girmez."""
    parsed, validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "H3", None)
    )
    _header, data = _comparison_table(_excel(parsed, validation, result))

    assert result.lowest_total_cost is not None
    assert result.rows[0].row_number not in result.lowest_total_cost.row_numbers
    assert data[0][16] is None
    assert any(row[16] == "✓" for row in data)


def test_pdf_missing_shipping_builds(tmp_path):
    parsed, validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "H3", None)
    )
    payload = _pdf(parsed, validation, result)

    assert payload.startswith(b"%PDF")


# --------------------------------------------------------------------------- #
# PDF — yapı
# --------------------------------------------------------------------------- #


def test_pdf_report_is_valid_pdf(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    payload = _pdf(parsed, validation, result)

    assert payload.startswith(b"%PDF")
    assert payload.rstrip().endswith(b"%%EOF")
    assert len(payload) > 5_000
    assert PDF_MIME == "application/pdf"
    assert PDF_FILENAME.endswith(".pdf")


def test_pdf_report_without_analysis_builds(tmp_path):
    """AI analizi yoksa rapor yine üretilir; AI bölümü eklenmez."""
    parsed, validation, result = _analyse(tmp_path)
    without = _pdf(parsed, validation, result)
    with_analysis = _pdf(parsed, validation, result, analysis=_analysis())

    assert without.startswith(b"%PDF")
    assert with_analysis.startswith(b"%PDF")
    assert len(with_analysis) > len(without)


def test_pdf_report_handles_many_suppliers(tmp_path):
    """Çok satırlı rapor birden fazla sayfaya taşar ve bozulmaz."""
    def add_rows(workbook) -> None:
        sheet = workbook["Supplier Quotes"]
        base = [sheet.cell(row=3, column=column).value for column in range(1, 15)]
        for index in range(12):
            row = 6 + index
            for column, value in enumerate(base, start=1):
                sheet.cell(row=row, column=column).value = value
            sheet.cell(row=row, column=1).value = f"Ek Tedarikçi {index + 1}"
            sheet.cell(row=row, column=2).value = f"EK-{index + 1}"

    parsed, validation, result = _analyse(tmp_path, add_rows)
    payload = _pdf(parsed, validation, result)

    assert payload.startswith(b"%PDF")
    assert len(result.rows) >= 13


# --------------------------------------------------------------------------- #
# AI bölümü
# --------------------------------------------------------------------------- #


def test_excel_ai_section_is_separate_sheet(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    payload = _excel(parsed, validation, result, analysis=_analysis())
    workbook = load_workbook(io.BytesIO(payload))

    assert AI_SHEET in workbook.sheetnames
    assert workbook.sheetnames[-1] == AI_SHEET


def test_excel_without_analysis_has_no_ai_sheet(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    workbook = load_workbook(io.BytesIO(_excel(parsed, validation, result)))

    assert AI_SHEET not in workbook.sheetnames


def test_excel_ai_section_contains_all_five_sections(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    rows = _sheet_rows(
        _excel(parsed, validation, result, analysis=_analysis()), AI_SHEET
    )
    labels = [row[0] for row in rows if row and row[0]]

    for _key, title in SECTION_TITLES:
        assert title in labels
    assert len(SECTION_TITLES) == 5


def test_excel_ai_section_is_labelled_as_commentary(tmp_path):
    """AI bölümü hesaplama yapmadığını ve karar vermediğini açıkça yazar."""
    parsed, validation, result = _analyse(tmp_path)
    rows = _sheet_rows(
        _excel(parsed, validation, result, analysis=_analysis()), AI_SHEET
    )
    text = " ".join(str(cell) for row in rows for cell in row if cell is not None)

    assert "hesaplama yapmaz" in text
    assert "kazanan belirlemez" in text
    assert _analysis().model in text


def test_excel_ai_section_keeps_legend_mapping(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    rows = _sheet_rows(
        _excel(parsed, validation, result, analysis=_analysis()), AI_SHEET
    )
    values = [str(row[1]) for row in rows if row and row[0] == "Etiket eşlemesi"]

    assert values
    assert "T1 = Bursa Metal A.Ş." in values[0]


# --------------------------------------------------------------------------- #
# Türkçe karakterler
# --------------------------------------------------------------------------- #


def test_turkish_characters_survive_in_excel(tmp_path):
    parsed, validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "A3", "Şişli Çelik Öğütüm İnşaat")
    )
    rows = _sheet_rows(_excel(parsed, validation, result), COMPARISON_SHEET)
    text = " ".join(str(cell) for row in rows for cell in row if cell is not None)

    assert "Şişli Çelik Öğütüm İnşaat" in text
    assert "Ödeme koşulları" in text
    assert "Toplam tedarik maliyeti (TL)" in text


def test_turkish_characters_survive_in_pdf(tmp_path):
    """PDF gömülü yazı tipi Türkçe harfleri çizebilmelidir."""
    name = "Şş Ğğ Iı İi Çç Öö Üü Çelik"
    parsed, validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "A3", name)
    )
    payload = _pdf(parsed, validation, result)

    assert payload.startswith(b"%PDF")
    # Gömülü yazı tipi (Vera) alt kümesi belgeye eklenir.
    assert b"BitstreamVeraSans" in payload

    # ToUnicode tablosu, raporda geçen Türkçe harfleri gerçekten eşlemelidir;
    # aksi hâlde PDF görüntüleyici bu karakterleri boş kutu olarak çizer.
    codes = _pdf_text_cmaps(payload)
    for character in "ŞşğĞıİçÇöÖüÜ":
        assert ord(character) in codes, f"{character} PDF yazı tipinde yok"


def _pdf_text_cmaps(payload: bytes) -> set[int]:
    """PDF içindeki ToUnicode tablolarının eşlediği Unicode kod noktalarını toplar.

    ``pypdf``/``pdfminer`` kurulu olmadığı için PDF akışları doğrudan çözülür:
    ReportLab akışları ASCII85 + Flate ile sıkıştırır. Dönen küme, gömülü yazı
    tipinin hangi karakterleri çizebildiğini kanıtlar.
    """
    codes: set[int] = set()
    for stream in re.findall(rb"stream\r?\n(.*?)endstream", payload, re.S):
        data = stream.strip()
        if data.endswith(b"~>"):
            try:
                data = base64.a85decode(data[:-2], adobe=False)
            except ValueError:
                continue
        try:
            data = zlib.decompress(data)
        except zlib.error:
            continue
        if b"beginbfchar" not in data:
            continue
        for source, target in re.findall(
            rb"<([0-9A-Fa-f]{2,4})>\s*<([0-9A-Fa-f]{4,8})>", data
        ):
            if len(source) >= 2 and len(target) >= 4:
                codes.add(int(target[:4], 16))
    return codes


def test_pdf_does_not_need_lira_sign(tmp_path):
    """₺ glifi Vera'da yok; raporlar yalnızca 'TL' metnini kullanır."""
    parsed, validation, result = _analyse(tmp_path)
    payload = _pdf(parsed, validation, result)

    assert payload.startswith(b"%PDF")
    assert "TL".encode() in payload or b"TL" in payload

    codes = _pdf_text_cmaps(payload)
    assert ord("₺") not in codes


# --------------------------------------------------------------------------- #
# Boş / hatalı veri
# --------------------------------------------------------------------------- #


def _clear_rows(workbook) -> None:
    sheet = workbook["Supplier Quotes"]
    for row in range(3, 30):
        for column in range(1, 15):
            sheet.cell(row=row, column=column).value = None


def test_excel_report_with_no_rows(tmp_path):
    parsed, validation, result = _analyse(tmp_path, _clear_rows)
    assert not result.rows

    payload = _excel(parsed, validation, result)
    assert payload.startswith(b"PK")
    _header, data = _comparison_table(payload)
    assert data == []

    rows = _sheet_rows(payload, COMPARISON_SHEET)
    text = " ".join(str(cell) for row in rows for cell in row if cell is not None)
    assert "Hesaplanabilen teklif satırı yok." in text

def test_pdf_report_with_no_rows(tmp_path):
    parsed, validation, result = _analyse(tmp_path, _clear_rows)

    assert _pdf(parsed, validation, result).startswith(b"%PDF")


def test_report_with_rejected_file(tmp_path):
    """Reddedilen dosyada da rapor üretilir; durum açıkça yazılır."""
    parsed, validation, result = _analyse(
        tmp_path, _set("Instructions", "B4", "HAYIR")
    )
    assert validation.file_rejected

    payload = _excel(parsed, validation, result)
    rows = _sheet_rows(payload, SUMMARY_SHEET)
    text = " ".join(str(cell) for row in rows for cell in row if cell is not None)

    assert "Dosya reddedildi" in text
    assert _pdf(parsed, validation, result).startswith(b"%PDF")


def test_report_with_excluded_rows(tmp_path):
    """Hatalı satır karşılaştırmaya girmez ama raporda gerekçesiyle görünür."""
    parsed, validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "G3", "geçersiz")
    )
    assert result.excluded

    rows = _sheet_rows(_excel(parsed, validation, result), VALIDATION_SHEET)
    text = " ".join(str(cell) for row in rows for cell in row if cell is not None)

    assert result.excluded[0].supplier_name in text
    assert result.excluded[0].reason in text


def test_excel_report_survives_unusual_supplier_text(tmp_path):
    """Tedarikçi metni HTML gibi görünse bile düz metin kalır."""
    parsed, validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "A3", "<b>Kalın</b> & Şirket")
    )
    rows = _sheet_rows(_excel(parsed, validation, result), COMPARISON_SHEET)
    names = [row[0] for row in rows if row and isinstance(row[0], str)]

    assert "<b>Kalın</b> & Şirket" in names


def test_generated_at_is_used_in_report(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    rows = _sheet_rows(_excel(parsed, validation, result), SUMMARY_SHEET)
    text = " ".join(str(cell) for row in rows for cell in row if cell is not None)

    assert TODAY.strftime("%d.%m.%Y") in text


# --------------------------------------------------------------------------- #
# Hata yönetimi
# --------------------------------------------------------------------------- #


def test_excel_report_wraps_unexpected_error(monkeypatch):
    """Beklenmeyen hata teknik ayrıntı sızdırmadan ReportError olur."""
    import procurement.report as report

    def broken(*args, **kwargs):
        raise RuntimeError("iç ayrıntı: disk yolu C:/gizli")

    monkeypatch.setattr(report, "_summary_sheet", broken)

    with pytest.raises(ReportError) as error:
        report.build_excel_report(
            parsed=None,  # type: ignore[arg-type]
            validation=None,  # type: ignore[arg-type]
            result=None,  # type: ignore[arg-type]
            file_name="teklif.xlsx",
        )

    message = str(error.value)
    assert "Excel raporu oluşturulamadı" in message
    assert "C:/gizli" not in message
    assert "disk yolu" not in message


def test_pdf_report_wraps_unexpected_error(monkeypatch):
    import procurement.report as report

    def broken(*args, **kwargs):
        raise RuntimeError("iç ayrıntı: yazı tipi yolu")

    monkeypatch.setattr(report, "_pdf_meta_rows", broken)

    with pytest.raises(ReportError) as error:
        report.build_pdf_report(
            parsed=None,  # type: ignore[arg-type]
            validation=None,  # type: ignore[arg-type]
            result=None,  # type: ignore[arg-type]
            file_name="teklif.xlsx",
        )

    message = str(error.value)
    assert "PDF raporu oluşturulamadı" in message
    assert "yazı tipi yolu" not in message

"""Excel ve PDF rapor üretimi — ``PROJECT_BRIEF.md`` §9.

Sorumluluk
----------
Yüklenen teklif dosyasının **doğrulanmış ve hesaplanmış** sonuçlarını
indirilebilir Excel (``.xlsx``) ve PDF raporlarına dönüştürür.

Değişmez kurallar
-----------------
* Bu modül **hesaplamaz**. Tüm tutarlar, ağırlıklar, oranlar ve en düşük
  işaretleri :mod:`procurement.calculator` çıktısından olduğu gibi okunur;
  biçimlendirme :mod:`procurement.presentation` yardımcılarıyla yapılır.
* Eksik değer (ör. nakliye ``None``) **sıfır gösterilmez**;
  :data:`procurement.presentation.MISSING_TEXT` yazılır.
* Tüm tutarlar KDV hariç TRY'dir. Ödeme koşulları maliyet hesabına dahil
  edilmez, yalnızca bilgi olarak raporlanır.
* AI yorumu raporda **ayrı ve etiketli** bir bölümdür; varsa eklenir, yoksa
  rapor AI olmadan da üretilir. AI yorumu yeni karar veya hesaplama eklemez.
* Rapor yalnızca bellekte üretilir; dosya sistemi, veritabanı veya rapor
  geçmişi kullanılmaz.

Hata yönetimi
-------------
Beklenmeyen bir üretim hatası :class:`ReportError` içinde toplanır; arayüz
kullanıcıya anlaşılır tek bir mesaj gösterir. Teknik ayrıntı mesaja eklenmez.
"""

from __future__ import annotations

import html
import io
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

import reportlab
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

from .ai_analysis import AIAnalysis
from .calculator import CalculationResult, SupplierCalculation
from .excel_parser import ParsedWorkbook
from .presentation import (
    MISSING_SHIPPING_NOTE,
    MISSING_TEXT,
    build_comparison_rows,
    build_analysis_summary,
    excluded_row_views,
    finding_context,
    format_days,
    format_money,
    format_quantity,
    format_quantity_value,
)
from .validator import Severity, ValidationResult, derive_payment_days

__all__ = [
    "EXCEL_FILENAME",
    "EXCEL_MIME",
    "PDF_FILENAME",
    "PDF_MIME",
    "ReportError",
    "build_excel_report",
    "build_pdf_report",
]

#: Tarayıcı indirmeleri için MIME türleri ve dosya adları.
EXCEL_MIME: str = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
PDF_MIME: str = "application/pdf"
EXCEL_FILENAME: str = "tedarikci-karsilastirma-raporu.xlsx"
PDF_FILENAME: str = "tedarikci-karsilastirma-raporu.pdf"


class ReportError(Exception):
    """Rapor üretilemediğinde kullanıcıya gösterilecek hata."""


# --------------------------------------------------------------------------- #
# Ortak yardımcılar
# --------------------------------------------------------------------------- #

_SEVERITY_LABEL: dict[Severity, str] = {
    Severity.ERROR: "Hata",
    Severity.WARNING: "Uyarı",
    Severity.INFO: "Bilgi",
}

#: KDV ve karar sınırı notları her iki raporda da yer alır (brief §9).
_METHOD_NOTES: tuple[str, ...] = (
    "Tüm tutarlar KDV hariç TRY'dir; KDV eklenmez, hesaplanmaz.",
    "Nakliye bilgisi olmayan satırda toplam maliyet gösterilmez; "
    "eksik nakliye sıfır kabul edilmez.",
    "Ödeme koşulları maliyet hesabına dahil edilmez; yalnızca bilgi amaçlıdır.",
    "Otomatik kazanan belirlenmez; en düşük ürün bedeli ve en düşük toplam "
    "maliyet ayrı ayrı işaretlenir, tedarikçi sıralaması üretilmez.",
    "kg/m değeri tedarikçinin beyanıdır; hesaplama beyan edilen değerle "
    "yapılır, teorik referans ağırlık yalnızca uyarı üretir.",
)


def _number(value: Any) -> float | None:
    """Değeri Excel/PDF için sayıya çevirir; çevrilemiyorsa ``None``."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    try:
        return float(Decimal(text.replace(",", ".")))
    except (InvalidOperation, ValueError):
        return None


def _money_cell(value: Decimal | None) -> float | str:
    """Para hücresi: eksik değer metin olarak yazılır, sıfır üretilmez."""
    if value is None:
        return MISSING_TEXT
    return float(value)


def _day_cell(text: str) -> int | str:
    """Termin hücresi: sayısal metni sayıya çevirir, aksi hâlde metin kalır."""
    stripped = text.strip()
    try:
        return int(Decimal(stripped.replace(",", ".")))
    except (InvalidOperation, ValueError):
        return text


def _mark_text(value: Decimal | None, suppliers: tuple[str, ...]) -> str:
    """En düşük değeri ve ona ulaşan tedarikçileri tek satırda gösterir."""
    if value is None:
        return "—"
    names = ", ".join(suppliers) if suppliers else "—"
    return f"{format_money(value)} TL · {names}"


def _quote_field(parsed: ParsedWorkbook, row_number: int, field: str) -> Any:
    for quote in parsed.suppliers:
        if quote.row_number == row_number:
            return getattr(quote, field, None)
    return None


def _payment_text(parsed: ParsedWorkbook, row_number: int) -> str:
    terms = _quote_field(parsed, row_number, "payment_terms")
    if isinstance(terms, str) and terms.strip():
        return terms.strip()
    return MISSING_TEXT


def _comparison_rows_by_number(
    parsed: ParsedWorkbook, result: CalculationResult
) -> dict[int, Any]:
    """Sunum katmanının hazır satırlarını satır numarasına göre indeksler."""
    return {
        item.row_number: item for item in build_comparison_rows(result, parsed)
    }


def _payment_days_text(parsed: ParsedWorkbook, row_number: int) -> str:
    """Vade günü metni; karma koşullarda tek güne indirgenmez."""
    terms = _quote_field(parsed, row_number, "payment_terms")
    if not isinstance(terms, str) or not terms.strip():
        return MISSING_TEXT
    days = derive_payment_days(terms)
    return str(days) if days is not None else "Belirsiz"


def _missing_text(row: SupplierCalculation) -> str:
    """Eksik veri alanlarını okunur metne çevirir."""
    labels: list[str] = []
    for code in row.missing_data:
        if code == "shipping":
            labels.append(MISSING_SHIPPING_NOTE)
        else:
            labels.append(code)
    return ", ".join(labels)


# --------------------------------------------------------------------------- #
# Excel raporu
# --------------------------------------------------------------------------- #

_HEADER_FILL = PatternFill("solid", fgColor="0F2F5C")
_HEADER_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
_TITLE_FONT = Font(name="Calibri", size=15, bold=True, color="0F2F5C")
_SUBTITLE_FONT = Font(name="Calibri", size=9, color="667085")
_SECTION_FONT = Font(name="Calibri", size=11, bold=True, color="0F2F5C")
_BODY_FONT = Font(name="Calibri", size=10, color="101828")
_LABEL_FONT = Font(name="Calibri", size=10, bold=True, color="344054")
_MUTED_FONT = Font(name="Calibri", size=9, color="667085")
_LOWEST_FILL = PatternFill("solid", fgColor="ECFDF3")
_MISSING_FILL = PatternFill("solid", fgColor="FFFAEB")

_THIN = Side(style="thin", color="E4E7EC")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

_MONEY_FORMAT = '#,##0.00" TL"'
_WEIGHT_FORMAT = "#,##0.000"
_RATE_FORMAT = "#,##0.0000"
_QUANTITY_FORMAT = "#,##0"

#: Karşılaştırma sayfası sütunları: (başlık, genişlik, hizalama, kaydırma).
_COMPARISON_COLUMNS: tuple[tuple[str, int, str, bool], ...] = (
    ("Tedarikçi", 28, "left", True),
    ("Teklif No", 18, "left", False),
    ("Miktar (adet)", 13, "center", False),
    ("Ağırlık (kg/m)", 14, "right", False),
    ("Toplam uzunluk (m)", 17, "right", False),
    ("Toplam ağırlık (kg)", 18, "right", False),
    ("Birim fiyat (TL/kg)", 17, "right", False),
    ("Ürün bedeli (TL)", 19, "right", False),
    ("Nakliye (TL)", 16, "right", False),
    ("Nakliye kaynağı", 17, "left", True),
    ("Nakliye (TL/kg)", 16, "right", False),
    ("Toplam tedarik maliyeti (TL)", 24, "right", False),
    ("Termin (gün)", 13, "center", False),
    ("Ödeme koşulları", 30, "left", True),
    ("Vade (gün)", 13, "center", False),
    ("En düşük ürün bedeli", 18, "center", False),
    ("En düşük toplam maliyet", 21, "center", False),
)

#: En düşük değer vurgusu uygulanan sütunlar (1 tabanlı).
_LOWEST_PRODUCT_COLUMN = 8
_LOWEST_TOTAL_COLUMN = 12


def _write_title(
    sheet: Worksheet, row: int, text: str, *, span: int = 6
) -> int:
    cell = sheet.cell(row=row, column=1, value=text)
    cell.font = _TITLE_FONT
    cell.alignment = Alignment(vertical="center")
    sheet.merge_cells(
        start_row=row, start_column=1, end_row=row, end_column=span
    )
    sheet.row_dimensions[row].height = 24
    return row + 1


def _write_subtitle(
    sheet: Worksheet, row: int, text: str, *, span: int = 6
) -> int:
    cell = sheet.cell(row=row, column=1, value=text)
    cell.font = _SUBTITLE_FONT
    cell.alignment = Alignment(vertical="center", wrap_text=True)
    sheet.merge_cells(
        start_row=row, start_column=1, end_row=row, end_column=span
    )
    return row + 1


def _write_section(sheet: Worksheet, row: int, text: str) -> int:
    cell = sheet.cell(row=row, column=1, value=text)
    cell.font = _SECTION_FONT
    return row + 2


def _write_pair(
    sheet: Worksheet, row: int, label: str, value: Any
) -> int:
    label_cell = sheet.cell(row=row, column=1, value=label)
    label_cell.font = _LABEL_FONT
    value_cell = sheet.cell(row=row, column=2, value=value)
    value_cell.font = _BODY_FONT
    value_cell.alignment = Alignment(vertical="top", wrap_text=True)
    return row + 1


def _write_note(sheet: Worksheet, row: int, text: str, *, span: int = 6) -> int:
    cell = sheet.cell(row=row, column=1, value=f"• {text}")
    cell.font = _MUTED_FONT
    cell.alignment = Alignment(vertical="top", wrap_text=True)
    sheet.merge_cells(
        start_row=row, start_column=1, end_row=row, end_column=span
    )
    return row + 1


def _write_header_row(
    sheet: Worksheet, row: int, headers: tuple[str, ...]
) -> None:
    for column, title in enumerate(headers, start=1):
        cell = sheet.cell(row=row, column=column, value=title)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        cell.border = _BORDER
    sheet.row_dimensions[row].height = 32


def _write_mark_pair(
    sheet: Worksheet,
    row: int,
    label: str,
    mark: Any,
    *,
    reason: str | None = None,
) -> int:
    """En düşük değer satırını tutar biçimiyle, ulaşan tedarikçilerle yazar."""
    if mark is None:
        row = _write_pair(sheet, row, label, "—")
        if reason:
            row = _write_pair(sheet, row, "  Neden", reason)
        return row

    sheet.cell(row=row, column=1, value=label).font = _LABEL_FONT
    value_cell = sheet.cell(row=row, column=2, value=float(mark.amount_try))
    value_cell.number_format = _MONEY_FORMAT
    value_cell.font = _BODY_FONT
    row += 1
    return _write_pair(
        sheet, row, "  Ulaşan tedarikçi", ", ".join(mark.suppliers)
    )


def _summary_sheet(
    sheet: Worksheet,
    parsed: ParsedWorkbook,
    validation: ValidationResult,
    result: CalculationResult,
    *,
    file_name: str,
    generated_at: date,
) -> None:
    """Yönetici özeti: dosya, talep, göstergeler, durum ve yöntem notları."""
    sheet.title = "Yönetici Özeti"
    sheet.sheet_view.showGridLines = False

    row = _write_title(sheet, 1, "Tedarikçi Teklif Karşılaştırma Raporu")
    row = _write_subtitle(
        sheet,
        row,
        "Galvanizli kare kutu profil · 40 × 40 × 2 mm · 6 m · "
        f"KDV hariç TRY · Rapor tarihi: {generated_at.strftime('%d.%m.%Y')}",
    )
    row += 1

    row = _write_section(sheet, row, "Talep ve ürün bilgileri")
    rfq = parsed.rfq
    quantity = format_quantity_value(rfq.quantity_pcs)
    pairs: list[tuple[str, Any]] = [("Dosya", file_name)]
    if rfq.rfq_number:
        pairs.append(("Talep no", str(rfq.rfq_number)))
    if rfq.product_code:
        pairs.append(("Ürün kodu", str(rfq.product_code)))
    if rfq.profile_specification:
        pairs.append(("Ürün", str(rfq.profile_specification)))
    length = _number(rfq.length_m)
    if length is not None:
        pairs.append(("Uzunluk (m)", length))
    thickness = _number(rfq.wall_thickness_mm)
    if thickness is not None:
        pairs.append(("Et kalınlığı (mm)", thickness))
    if quantity:
        pairs.append(("Miktar (adet)", quantity))
    if rfq.delivery_location:
        pairs.append(("Teslim yeri", str(rfq.delivery_location)))
    if rfq.requested_delivery_days is not None:
        pairs.append(
            ("İstenen termin (gün)", format_days(rfq.requested_delivery_days))
        )
    if rfq.currency:
        pairs.append(("Para birimi", str(rfq.currency)))

    for label, value in pairs:
        row = _write_pair(sheet, row, label, value)
    row += 1

    row = _write_section(sheet, row, "Özet göstergeler")
    row = _write_pair(sheet, row, "Tedarikçi satırı", len(parsed.suppliers))
    row = _write_pair(sheet, row, "Hesaplanan teklif", len(result.rows))
    row = _write_pair(
        sheet, row, "Karşılaştırma dışı satır", len(result.excluded)
    )

    lowest_product = result.lowest_product_cost
    lowest_total = result.lowest_total_cost
    row = _write_mark_pair(sheet, row, "En düşük ürün bedeli", lowest_product)
    row = _write_mark_pair(
        sheet,
        row,
        "En düşük toplam maliyet",
        lowest_total,
        reason=(
            "Nakliye bilgisi eksik olduğu için hesaplanamadı."
            if result.missing_data_rows
            else None
        ),
    )
    row += 1

    summary = build_analysis_summary(
        parsed, validation, result, file_name=file_name
    )
    row = _write_section(sheet, row, "Durum")
    row = _write_pair(sheet, row, "Sonuç", summary.status_label)
    row = _write_pair(sheet, row, "Açıklama", summary.status_detail)
    row += 1

    row = _write_section(sheet, row, "Yöntem ve varsayımlar")
    for note in _METHOD_NOTES:
        row = _write_note(sheet, row, note)

    sheet.column_dimensions["A"].width = 30
    sheet.column_dimensions["B"].width = 78


def _comparison_sheet(
    sheet: Worksheet,
    parsed: ParsedWorkbook,
    result: CalculationResult,
) -> None:
    """Tedarikçi bazında karşılaştırma tablosu (filtreli, dondurulmuş başlık)."""
    sheet.title = "Tedarikçi Karşılaştırma"
    sheet.sheet_view.showGridLines = False

    row = _write_title(
        sheet,
        1,
        "Tedarikçi Karşılaştırması",
        span=len(_COMPARISON_COLUMNS),
    )
    row = _write_subtitle(
        sheet,
        row,
        "Satırlar dosyadaki sırada gösterilir; sıralama veya puanlama "
        "yapılmaz. Eksik değerler sıfır değil, 'Eksik veri' olarak yazılır.",
        span=len(_COMPARISON_COLUMNS),
    )
    row += 1

    header_row = row
    _write_header_row(
        sheet,
        header_row,
        tuple(title for title, _w, _a, _wrap in _COMPARISON_COLUMNS),
    )
    for index, (_title, width, _align, _wrap) in enumerate(
        _COMPARISON_COLUMNS, start=1
    ):
        sheet.column_dimensions[get_column_letter(index)].width = width

    comparison = _comparison_rows_by_number(parsed, result)
    lowest_product = (
        set(result.lowest_product_cost.row_numbers)
        if result.lowest_product_cost is not None
        else set()
    )
    lowest_total = (
        set(result.lowest_total_cost.row_numbers)
        if result.lowest_total_cost is not None
        else set()
    )

    data_row = header_row + 1
    for calculation in result.rows:
        view = comparison.get(calculation.row_number)
        values: list[tuple[Any, str | None]] = [
            (
                calculation.supplier_name
                or f"Satır {calculation.row_number}",
                None,
            ),
            (calculation.supplier_quote_number or "", None),
            (calculation.quantity_pcs, _QUANTITY_FORMAT),
            (float(calculation.weight_kg_per_m), _WEIGHT_FORMAT),
            (float(calculation.total_length_m), _WEIGHT_FORMAT),
            (float(calculation.total_weight_kg), _WEIGHT_FORMAT),
            (float(calculation.unit_price_try_per_kg), _MONEY_FORMAT),
            (_money_cell(calculation.product_cost_try), _MONEY_FORMAT),
            (_money_cell(calculation.shipping_cost_try), _MONEY_FORMAT),
            (calculation.shipping_cost_source or "", None),
            (
                _money_cell(calculation.shipping_per_kg_try),
                _RATE_FORMAT,
            ),
            (_money_cell(calculation.total_supply_cost_try), _MONEY_FORMAT),
            (
                _day_cell(view.delivery_days)
                if view is not None
                else MISSING_TEXT,
                _QUANTITY_FORMAT,
            ),
            (
                _payment_text(parsed, calculation.row_number),
                None,
            ),
            (
                _day_cell(_payment_days_text(parsed, calculation.row_number)),
                _QUANTITY_FORMAT,
            ),
            ("✓" if calculation.row_number in lowest_product else "", None),
            ("✓" if calculation.row_number in lowest_total else "", None),
        ]

        for column, (value, number_format) in enumerate(values, start=1):
            _title, _width, align, wrap = _COMPARISON_COLUMNS[column - 1]
            cell = sheet.cell(row=data_row, column=column, value=value)
            cell.font = _BODY_FONT
            cell.border = _BORDER
            cell.alignment = Alignment(
                vertical="top", wrap_text=wrap, horizontal=align
            )
            if number_format is not None:
                cell.number_format = number_format
            if value == MISSING_TEXT:
                cell.fill = _MISSING_FILL
            elif column == _LOWEST_PRODUCT_COLUMN:
                if calculation.row_number in lowest_product:
                    cell.fill = _LOWEST_FILL
            elif column == _LOWEST_TOTAL_COLUMN:
                if calculation.row_number in lowest_total:
                    cell.fill = _LOWEST_FILL

        sheet.row_dimensions[data_row].height = 26
        data_row += 1

    if data_row == header_row + 1:
        cell = sheet.cell(
            row=data_row,
            column=1,
            value="Hesaplanabilen teklif satırı yok.",
        )
        cell.font = _MUTED_FONT

    sheet.auto_filter.ref = (
        f"A{header_row}:{get_column_letter(len(_COMPARISON_COLUMNS))}"
        f"{max(data_row - 1, header_row)}"
    )
    sheet.freeze_panes = f"A{header_row + 1}"
    sheet.cell(
        row=data_row + 1,
        column=1,
        value=(
            "Yeşil hücre: Python hesaplamasındaki en düşük değer. "
            "Bu bir sıralama değildir; eşitlikte tüm satırlar işaretlenir."
        ),
    ).font = _MUTED_FONT


def _validation_sheet(
    sheet: Worksheet,
    validation: ValidationResult,
    result: CalculationResult,
) -> None:
    """Eksik veriler, karşılaştırma dışı satırlar ve doğrulama bulguları."""
    sheet.title = "Doğrulama"
    sheet.sheet_view.showGridLines = False
    span = 5

    row = _write_title(sheet, 1, "Doğrulama ve Eksik Veriler", span=span)
    row = _write_subtitle(
        sheet,
        row,
        "Hatalar satırı karşılaştırmadan çıkarır; uyarılar hesaplamayı "
        "durdurmaz. Eksik nakliye sıfır kabul edilmez.",
        span=span,
    )
    row += 1

    row = _write_section(sheet, row, "Eksik veri içeren satırlar")
    if result.missing_data_rows:
        _write_header_row(
            sheet, row, ("Satır", "Tedarikçi", "Eksik alan")
        )
        row += 1
        for calculation in result.rows:
            if not calculation.missing_data:
                continue
            sheet.cell(row=row, column=1, value=calculation.row_number)
            sheet.cell(
                row=row,
                column=2,
                value=calculation.supplier_name
                or f"Satır {calculation.row_number}",
            )
            sheet.cell(row=row, column=3, value=_missing_text(calculation))
            for column in range(1, 4):
                cell = sheet.cell(row=row, column=column)
                cell.font = _BODY_FONT
                cell.border = _BORDER
                cell.fill = _MISSING_FILL
            row += 1
    else:
        sheet.cell(row=row, column=1, value="Eksik veri yok.").font = _BODY_FONT
        row += 1
    row += 1

    row = _write_section(sheet, row, "Karşılaştırma dışı satırlar")
    excluded = excluded_row_views(result)
    if excluded:
        _write_header_row(
            sheet, row, ("Satır", "Tedarikçi", "Gerekçe", "Kodlar")
        )
        row += 1
        for item in excluded:
            values = (
                item.row_number,
                item.supplier_name,
                item.reason,
                item.codes_text,
            )
            for column, value in enumerate(values, start=1):
                cell = sheet.cell(row=row, column=column, value=value)
                cell.font = _BODY_FONT
                cell.border = _BORDER
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            row += 1
    else:
        sheet.cell(
            row=row, column=1, value="Karşılaştırma dışı satır yok."
        ).font = _BODY_FONT
        row += 1
    row += 1

    row = _write_section(sheet, row, "Doğrulama bulguları")
    findings = list(validation.findings)
    if findings:
        _write_header_row(
            sheet, row, ("Önem", "Kod", "Tedarikçi / hücre", "Açıklama")
        )
        row += 1
        for finding in findings:
            values = (
                _SEVERITY_LABEL.get(finding.severity, finding.severity.value),
                finding.code,
                finding_context(finding),
                finding.message,
            )
            for column, value in enumerate(values, start=1):
                cell = sheet.cell(row=row, column=column, value=value)
                cell.font = _BODY_FONT
                cell.border = _BORDER
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            row += 1
    else:
        sheet.cell(
            row=row, column=1, value="Doğrulama bulgusu yok."
        ).font = _BODY_FONT
        row += 1
    row += 1

    row = _write_section(sheet, row, "Yöntem ve varsayımlar")
    for note in _METHOD_NOTES:
        row = _write_note(sheet, row, note, span=span)

    for column, width in (("A", 13), ("B", 22), ("C", 30), ("D", 70)):
        sheet.column_dimensions[column].width = width


def build_excel_report(
    parsed: ParsedWorkbook,
    validation: ValidationResult,
    result: CalculationResult,
    *,
    file_name: str,
    analysis: AIAnalysis | None = None,
    generated_at: date | None = None,
) -> bytes:
    """Excel raporunu ``.xlsx`` baytları olarak üretir.

    Args:
        parsed: Ayrıştırılmış dosya.
        validation: Doğrulama sonucu.
        result: Hesaplama sonucu — tüm tutarlar buradan okunur.
        file_name: Raporda gösterilecek dosya adı.
        analysis: Varsa AI yorumu **ayrı sayfada** eklenir. Yoksa rapor
            AI bölümü olmadan üretilir.
        generated_at: Rapor tarihi (varsayılan: bugün).

    Raises:
        ReportError: Rapor üretilemezse.
    """
    try:
        workbook = Workbook()
        summary = workbook.active
        assert summary is not None
        created = generated_at or date.today()

        _summary_sheet(
            summary,
            parsed,
            validation,
            result,
            file_name=file_name,
            generated_at=created,
        )
        _comparison_sheet(workbook.create_sheet(), parsed, result)
        _validation_sheet(workbook.create_sheet(), validation, result)
        if analysis is not None:
            _ai_sheet(workbook.create_sheet(), analysis)

        buffer = io.BytesIO()
        workbook.save(buffer)
        return buffer.getvalue()
    except ReportError:
        raise
    except Exception as error:  # pragma: no cover - beklenmeyen üretim hatası
        raise ReportError(
            "Excel raporu oluşturulamadı. Lütfen tekrar deneyin."
        ) from error


def _ai_sheet(sheet: Worksheet, analysis: AIAnalysis) -> None:
    """AI yorumunu ayrı ve etiketli sayfa olarak yazar."""
    sheet.title = "AI Yorumu"
    sheet.sheet_view.showGridLines = False

    row = _write_title(sheet, 1, "AI Yorumu")
    row = _write_subtitle(
        sheet,
        row,
        "Yalnızca doğrulanmış sonuçların yorumudur; hesaplama yapmaz, "
        "kazanan belirlemez. Sayısal ifadeler kaynak verilerle "
        f"karşılaştırılmıştır. Model: {analysis.model}",
    )
    row += 1

    if analysis.legend:
        legend = ", ".join(
            f"{label} = {name}" for label, name in analysis.legend
        )
        row = _write_pair(sheet, row, "Etiket eşlemesi", legend)
        row += 1

    for section in analysis.sections:
        row = _write_section(sheet, row, section.title)
        cell = sheet.cell(row=row, column=1, value=section.body)
        cell.font = _BODY_FONT
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        sheet.merge_cells(
            start_row=row, start_column=1, end_row=row, end_column=6
        )
        sheet.row_dimensions[row].height = max(
            15, 15 * (len(section.body) // 90 + 1)
        )
        row += 2

    sheet.column_dimensions["A"].width = 24
    for column in ("B", "C", "D", "E", "F"):
        sheet.column_dimensions[column].width = 18


# --------------------------------------------------------------------------- #
# PDF raporu
# --------------------------------------------------------------------------- #

_FONT = "Vera"
_FONT_BOLD = "Vera-Bold"
_fonts_registered = False


def _register_fonts() -> None:
    """Türkçe karakterleri destekleyen gömülü yazı tipini kaydeder.

    ReportLab'ın varsayılan Helvetica'sı ``ş``/``ğ``/``İ`` gibi karakterleri
    çizemez; paketle gelen Vera yazı tipi Türkçe harfleri içerir.
    """
    global _fonts_registered
    if _fonts_registered:
        return
    fonts_dir = Path(reportlab.__file__).parent / "fonts"
    pdfmetrics.registerFont(TTFont(_FONT, str(fonts_dir / "Vera.ttf")))
    pdfmetrics.registerFont(TTFont(_FONT_BOLD, str(fonts_dir / "VeraBd.ttf")))
    pdfmetrics.registerFontFamily(
        _FONT,
        normal=_FONT,
        bold=_FONT_BOLD,
        italic=_FONT,
        boldItalic=_FONT_BOLD,
    )
    _fonts_registered = True


def _styles() -> dict[str, ParagraphStyle]:
    return {
        "title": ParagraphStyle(
            "ReportTitle",
            fontName=_FONT_BOLD,
            fontSize=18,
            leading=22,
            textColor=colors.HexColor("#0F2F5C"),
            spaceAfter=2,
        ),
        "subtitle": ParagraphStyle(
            "ReportSubtitle",
            fontName=_FONT,
            fontSize=8.5,
            leading=12,
            textColor=colors.HexColor("#667085"),
            spaceAfter=10,
        ),
        "h2": ParagraphStyle(
            "SectionHeading",
            fontName=_FONT_BOLD,
            fontSize=11.5,
            leading=15,
            textColor=colors.HexColor("#0F2F5C"),
            spaceBefore=12,
            spaceAfter=6,
        ),
        "body": ParagraphStyle(
            "Body",
            fontName=_FONT,
            fontSize=9.5,
            leading=14,
            textColor=colors.HexColor("#101828"),
            spaceAfter=4,
        ),
        "note": ParagraphStyle(
            "Note",
            fontName=_FONT,
            fontSize=8,
            leading=11.5,
            textColor=colors.HexColor("#667085"),
            spaceAfter=2,
        ),
        "cell": ParagraphStyle(
            "Cell",
            fontName=_FONT,
            fontSize=7.5,
            leading=9.5,
            textColor=colors.HexColor("#101828"),
        ),
        "cell_head": ParagraphStyle(
            "CellHead",
            fontName=_FONT_BOLD,
            fontSize=7.5,
            leading=9.5,
            textColor=colors.white,
        ),
    }


def _p(text: Any, style: ParagraphStyle) -> Paragraph:
    """Kullanıcı verisini güvenli biçimde paragrafa çevirir."""
    return Paragraph(html.escape(str(text)), style)


class _NumberedCanvas(canvas.Canvas):
    """Toplam sayfa sayısını bilen altbilgi için iki geçişli canvas."""

    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self._saved_page_states: list[dict[str, Any]] = []

    def showPage(self) -> None:  # noqa: N802 - ReportLab API
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self) -> None:
        total = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self._draw_footer(total)
            super().showPage()
        super().save()

    def _draw_footer(self, total: int) -> None:
        self.setStrokeColor(colors.HexColor("#E4E7EC"))
        self.setLineWidth(0.5)
        self.line(20 * mm, 15 * mm, A4[0] - 20 * mm, 15 * mm)
        self.setFont(_FONT, 7.5)
        self.setFillColor(colors.HexColor("#667085"))
        self.drawString(
            20 * mm, 10 * mm, "AI Procurement Intelligence · KDV hariç TRY"
        )
        self.drawRightString(
            A4[0] - 20 * mm,
            10 * mm,
            f"Sayfa {self._pageNumber} / {total}",
        )


def _pdf_table(
    data: list[list[Paragraph]],
    widths: list[float],
    *,
    highlight_cells: tuple[tuple[int, int], ...] = (),
) -> Table:
    """Başlığı tekrarlanan, ince çizgili rapor tablosu.

    ``highlight_cells`` ``(sütun, satır)`` çiftleridir; vurgu yalnızca en
    düşük değerin bulunduğu hücreye uygulanır.
    """
    table = Table(data, colWidths=widths, repeatRows=1, hAlign="LEFT")
    style: list[tuple[Any, ...]] = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F2F5C")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#E4E7EC")),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    for row_index in range(1, len(data)):
        if row_index % 2 == 0:
            style.append(
                (
                    "BACKGROUND",
                    (0, row_index),
                    (-1, row_index),
                    colors.HexColor("#FAFBFC"),
                )
            )
    for column, row_index in highlight_cells:
        style.append(
            (
                "BACKGROUND",
                (column, row_index),
                (column, row_index),
                colors.HexColor("#ECFDF3"),
            )
        )
    table.setStyle(TableStyle(style))
    return table


def _pdf_meta_rows(
    parsed: ParsedWorkbook,
    validation: ValidationResult,
    result: CalculationResult,
    *,
    file_name: str,
    generated_at: date,
) -> list[tuple[str, str]]:
    rfq = parsed.rfq
    summary = build_analysis_summary(
        parsed, validation, result, file_name=file_name
    )
    rows: list[tuple[str, str]] = [
        ("Rapor tarihi", generated_at.strftime("%d.%m.%Y")),
        ("Dosya", file_name),
    ]
    if rfq.rfq_number:
        rows.append(("Talep no", str(rfq.rfq_number)))
    if rfq.profile_specification:
        rows.append(("Ürün", str(rfq.profile_specification)))
    length = _number(rfq.length_m)
    if length is not None:
        rows.append(("Uzunluk", f"{format_quantity(int(length))} m"))
    thickness = _number(rfq.wall_thickness_mm)
    if thickness is not None:
        rows.append(("Et kalınlığı", f"{format_quantity(int(thickness))} mm"))
    quantity = format_quantity_value(rfq.quantity_pcs)
    if quantity:
        rows.append(("Miktar", f"{quantity} adet"))
    if rfq.delivery_location:
        rows.append(("Teslim yeri", str(rfq.delivery_location)))
    if rfq.requested_delivery_days is not None:
        rows.append(
            ("İstenen termin", f"{format_days(rfq.requested_delivery_days)} gün")
        )
    rows.append(("Durum", summary.status_label))
    return rows


def _pdf_comparison_table(
    parsed: ParsedWorkbook,
    result: CalculationResult,
    styles: dict[str, ParagraphStyle],
) -> list[Any]:
    headers = (
        "Tedarikçi",
        "Miktar",
        "Ağırlık (kg)",
        "TL/kg",
        "Ürün bedeli",
        "Nakliye",
        "Toplam maliyet",
        "Termin",
        "Ödeme koşulları",
    )
    widths = [
        26 * mm,
        12 * mm,
        17 * mm,
        15 * mm,
        21 * mm,
        19 * mm,
        21 * mm,
        11 * mm,
        27 * mm,
    ]

    data: list[list[Paragraph]] = [
        [_p(title, styles["cell_head"]) for title in headers]
    ]
    views = _comparison_rows_by_number(parsed, result)
    highlights: list[tuple[int, int]] = []

    for index, calculation in enumerate(result.rows, start=1):
        view = views.get(calculation.row_number)
        values = (
            calculation.supplier_name or f"Satır {calculation.row_number}",
            view.quantity if view is not None else MISSING_TEXT,
            view.weight if view is not None else MISSING_TEXT,
            view.unit_price if view is not None else MISSING_TEXT,
            view.product_cost if view is not None else MISSING_TEXT,
            view.shipping_cost if view is not None else MISSING_TEXT,
            view.total_cost if view is not None else MISSING_TEXT,
            view.delivery_days if view is not None else MISSING_TEXT,
            view.payment_terms if view is not None else MISSING_TEXT,
        )
        data.append([_p(value, styles["cell"]) for value in values])
        if view is not None and view.is_lowest_product:
            highlights.append((4, index))
        if view is not None and view.is_lowest_total:
            highlights.append((6, index))

    if len(data) == 1:
        data.append(
            [_p("Hesaplanabilen teklif satırı yok.", styles["cell"])]
            + [_p("", styles["cell"]) for _ in range(len(headers) - 1)]
        )

    story: list[Any] = [_pdf_table(data, widths, highlight_cells=tuple(highlights))]
    story.append(
        _p(
            "Yeşil hücre: Python hesaplamasındaki en düşük değer. "
            "Sıralama değildir; eşitlikte tüm satırlar işaretlenir. "
            "'Eksik veri' sıfır anlamına gelmez.",
            styles["note"],
        )
    )
    return story


def _pdf_delivery_table(
    parsed: ParsedWorkbook,
    result: CalculationResult,
    styles: dict[str, ParagraphStyle],
) -> list[Any]:
    headers = ("Tedarikçi", "Termin", "İstenen", "Ödeme koşulu", "Vade")
    widths = [35 * mm, 17 * mm, 20 * mm, 75 * mm, 21 * mm]
    requested = format_days(parsed.rfq.requested_delivery_days)
    views = _comparison_rows_by_number(parsed, result)

    data: list[list[Paragraph]] = [
        [_p(title, styles["cell_head"]) for title in headers]
    ]
    for calculation in result.rows:
        view = views.get(calculation.row_number)
        values = (
            calculation.supplier_name or f"Satır {calculation.row_number}",
            view.delivery_days if view is not None else MISSING_TEXT,
            requested,
            view.payment_terms if view is not None else MISSING_TEXT,
            _payment_days_text(parsed, calculation.row_number),
        )
        data.append([_p(value, styles["cell"]) for value in values])

    if len(data) == 1:
        data.append(
            [_p("Hesaplanabilen teklif satırı yok.", styles["cell"])]
            + [_p("", styles["cell"]) for _ in range(len(headers) - 1)]
        )

    if len(data) == 1:
        data.append([_p("Hesaplanabilen teklif satırı yok.", styles["cell"])] + [
            _p("", styles["cell"]) for _ in range(len(headers) - 1)
        ])

    story: list[Any] = [_pdf_table(data, widths)]
    story.append(
        _p(
            "Termin tedarikçinin beyan ettiği teslim süresidir. Vade günü "
            "yalnızca bilgi amaçlıdır; karma ödeme koşulları tek güne "
            "indirgenmez ve 'Belirsiz' olarak gösterilir.",
            styles["note"],
        )
    )
    return story


def _pdf_findings(
    validation: ValidationResult,
    result: CalculationResult,
    styles: dict[str, ParagraphStyle],
) -> list[Any]:
    story: list[Any] = []

    missing = result.missing_data_rows
    if missing:
        story.append(
            _p("Eksik veri içeren satırlar", styles["body"])
        )
        for calculation in result.rows:
            if not calculation.missing_data:
                continue
            story.append(
                _p(
                    f"• Satır {calculation.row_number} · "
                    f"{calculation.supplier_name or '—'} · "
                    f"{_missing_text(calculation)}. Toplam maliyet "
                    "gösterilmedi; eksik nakliye sıfır kabul edilmedi.",
                    styles["note"],
                )
            )
        story.append(Spacer(1, 4))

    if result.excluded:
        story.append(_p("Karşılaştırma dışı satırlar", styles["body"]))
        for item in result.excluded:
            codes = f" ({', '.join(item.codes)})" if item.codes else ""
            story.append(
                _p(
                    f"• Satır {item.row_number} · "
                    f"{item.supplier_name or '—'} · {item.reason}{codes}",
                    styles["note"],
                )
            )
        story.append(Spacer(1, 4))

    if validation.findings:
        story.append(_p("Doğrulama bulguları", styles["body"]))
        for finding in validation.findings:
            story.append(
                _p(
                    f"• [{finding.code}] "
                    f"{_SEVERITY_LABEL.get(finding.severity, '')} · "
                    f"{finding_context(finding) or 'Dosya'} · "
                    f"{finding.message}",
                    styles["note"],
                )
            )
    elif not missing and not result.excluded:
        story.append(_p("Doğrulama bulgusu yok.", styles["body"]))
    return story


def _pdf_ai_section(
    analysis: AIAnalysis, styles: dict[str, ParagraphStyle]
) -> list[Any]:
    story: list[Any] = [
        _p("AI Yorumu", styles["h2"]),
        _p(
            "Yalnızca doğrulanmış sonuçların yorumudur; hesaplama yapmaz, "
            "tedarikçi sıralamaz ve kazanan belirlemez. Sayısal ifadeler "
            f"kaynak verilerle karşılaştırılmıştır. Model: {analysis.model}",
            styles["note"],
        ),
    ]
    if analysis.legend:
        legend = ", ".join(
            f"{label} = {name}" for label, name in analysis.legend
        )
        story.append(_p(f"Etiket eşlemesi: {legend}", styles["note"]))
    story.append(Spacer(1, 4))

    for section in analysis.sections:
        story.append(_p(section.title, styles["body"]))
        for paragraph in section.body.split("\n"):
            if paragraph.strip():
                story.append(_p(paragraph.strip(), styles["body"]))
        story.append(Spacer(1, 2))
    return story


def build_pdf_report(
    parsed: ParsedWorkbook,
    validation: ValidationResult,
    result: CalculationResult,
    *,
    file_name: str,
    analysis: AIAnalysis | None = None,
    generated_at: date | None = None,
) -> bytes:
    """PDF raporunu bayt olarak üretir.

    Args:
        parsed: Ayrıştırılmış dosya.
        validation: Doğrulama sonucu.
        result: Hesaplama sonucu — tüm tutarlar buradan okunur.
        file_name: Kapakta gösterilecek dosya adı.
        analysis: Varsa AI yorumu ayrı ve etiketli bölüm olarak eklenir.
        generated_at: Rapor tarihi (varsayılan: bugün).

    Raises:
        ReportError: Rapor üretilemezse.
    """
    try:
        _register_fonts()
        styles = _styles()
        created = generated_at or date.today()

        buffer = io.BytesIO()
        document = BaseDocTemplate(
            buffer,
            pagesize=A4,
            leftMargin=20 * mm,
            rightMargin=20 * mm,
            topMargin=18 * mm,
            bottomMargin=20 * mm,
            title="Tedarikçi Teklif Karşılaştırma Raporu",
            author="AI Procurement Intelligence",
        )
        frame = Frame(
            document.leftMargin,
            document.bottomMargin,
            document.width,
            document.height,
            id="body",
            leftPadding=0,
            rightPadding=0,
            topPadding=0,
            bottomPadding=0,
        )
        document.addPageTemplates(
            [PageTemplate(id="main", frames=[frame])]
        )

        story: list[Any] = [
            _p("Tedarikçi Teklif Karşılaştırma Raporu", styles["title"]),
            _p(
                "Galvanizli kare kutu profil · 40 × 40 × 2 mm · 6 m · "
                "KDV hariç TRY",
                styles["subtitle"],
            ),
        ]

        meta = _pdf_meta_rows(
            parsed,
            validation,
            result,
            file_name=file_name,
            generated_at=created,
        )
        meta_data: list[list[Paragraph]] = [
            [
                _p("Rapor bilgisi", styles["cell_head"]),
                _p("Değer", styles["cell_head"]),
            ]
        ]
        for label, value in meta:
            meta_data.append(
                [
                    _p(label, styles["cell"]),
                    _p(value, styles["cell"]),
                ]
            )
        story.append(
            _pdf_table(meta_data, [45 * mm, 125 * mm])
        )

        summary = build_analysis_summary(
            parsed, validation, result, file_name=file_name
        )
        story.append(
            _p("Yönetici Özeti", styles["h2"])
        )
        story.append(_p(summary.status_detail, styles["body"]))
        story.append(
            _p(
                f"Tedarikçi satırı: {len(parsed.suppliers)} · "
                f"Hesaplanan teklif: {len(result.rows)} · "
                f"Karşılaştırma dışı: {len(result.excluded)}",
                styles["body"],
            )
        )
        product_mark = result.lowest_product_cost
        story.append(
            _p(
                "En düşük ürün bedeli: "
                + (
                    _mark_text(product_mark.amount_try, product_mark.suppliers)
                    if product_mark is not None
                    else "—"
                ),
                styles["body"],
            )
        )
        total_mark = result.lowest_total_cost
        if total_mark is not None:
            total_text = _mark_text(total_mark.amount_try, total_mark.suppliers)
        elif result.missing_data_rows:
            total_text = (
                "Nakliye bilgisi eksik olduğu için hesaplanamadı."
            )
        else:
            total_text = "—"
        story.append(
            _p(f"En düşük toplam tedarik maliyeti: {total_text}", styles["body"])
        )

        story.append(_p("Tedarikçi Karşılaştırması", styles["h2"]))
        story.extend(_pdf_comparison_table(parsed, result, styles))

        story.append(_p("Maliyet Değerlendirmesi", styles["h2"]))
        if product_mark is not None:
            story.append(
                _p(
                    "Ürün bedeli karşılaştırması: en düşük değer "
                    f"{format_money(product_mark.amount_try)} TL "
                    f"({', '.join(product_mark.suppliers)}).",
                    styles["body"],
                )
            )
        if total_mark is not None:
            story.append(
                _p(
                    "Toplam tedarik maliyeti karşılaştırması: en düşük değer "
                    f"{format_money(total_mark.amount_try)} TL "
                    f"({', '.join(total_mark.suppliers)}).",
                    styles["body"],
                )
            )
        else:
            story.append(
                _p(
                    "Toplam tedarik maliyeti için en düşük değer "
                    "hesaplanamadı: en az bir satırda nakliye bilgisi eksik. "
                    "Eksik nakliye sıfır kabul edilmez.",
                    styles["body"],
                )
            )
        for calculation in result.rows:
            if calculation.shipping_cost_source:
                story.append(
                    _p(
                        f"• {calculation.supplier_name or '—'}: nakliye "
                        f"kaynağı '{calculation.shipping_cost_source}'.",
                        styles["note"],
                    )
                )

        story.append(_p("Teslimat ve Ödeme Koşulları", styles["h2"]))
        story.extend(_pdf_delivery_table(parsed, result, styles))

        story.append(_p("Veri Eksiklikleri ve Doğrulama Uyarıları", styles["h2"]))
        story.extend(_pdf_findings(validation, result, styles))

        story.append(_p("Yöntem ve Varsayımlar", styles["h2"]))
        for note in _METHOD_NOTES:
            story.append(_p(f"• {note}", styles["note"]))

        if analysis is not None:
            story.extend(_pdf_ai_section(analysis, styles))

        document.build(story, canvasmaker=_NumberedCanvas)
        return buffer.getvalue()
    except ReportError:
        raise
    except Exception as error:  # pragma: no cover - beklenmeyen üretim hatası
        raise ReportError(
            "PDF raporu oluşturulamadı. Lütfen tekrar deneyin."
        ) from error

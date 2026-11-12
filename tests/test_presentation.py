"""``procurement.presentation`` testleri — görünüm modeli ve biçimlendirme.

Kapsam
------
Bu modül hesaplama yapmaz; calculator çıktısını sunuma hazırlar. Testler
gerçek akışla (``parse_workbook`` → ``validate`` → ``calculate``) üretilmiş
sonuçları kullanır, böylece gösterim ile motor arasındaki sözleşme sınanır.

Biçimlendirme testleri elle yazılmış beklenen metinlerle karşılaştırılır;
motorun kendi çıktısı kendine doğrulatılmaz.
"""

from __future__ import annotations

import shutil
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import load_workbook

from procurement.calculator import CalculationResult, calculate
from procurement.excel_parser import ParsedWorkbook, parse_workbook
from procurement.presentation import (
    LOWEST_BADGE,
    MISSING_TEXT,
    build_comparison_rows,
    build_cost_figure,
    build_kpis,
    comparison_table_html,
    excluded_row_views,
    finding_context,
    format_days,
    format_money,
    format_quantity,
    format_rate,
    format_weight,
    group_findings,
    rejection_hints,
)
from procurement.validator import ValidationResult, validate

TEMPLATE = Path(__file__).resolve().parents[1] / "templates" / "supplier_quote_template.xlsx"
TODAY = date(2026, 11, 12)

pytestmark = pytest.mark.unit


# --------------------------------------------------------------------------- #
# Yardımcılar
# --------------------------------------------------------------------------- #


def _build(tmp_path: Path, mutate=None, *, name: str = "view.xlsx") -> Path:
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
    tmp_path: Path, mutate=None, *, name: str = "view.xlsx"
) -> tuple[ParsedWorkbook, ValidationResult, CalculationResult]:
    path = _build(tmp_path, mutate, name=name)
    parsed = parse_workbook(path, strict=False)
    validation = validate(parsed, today=TODAY)
    return parsed, validation, calculate(parsed, validation)


def _set(sheet: str, cell: str, value):
    def mutate(workbook) -> None:
        workbook[sheet][cell] = value

    return mutate


# --------------------------------------------------------------------------- #
# Sayı biçimlendirme — Türkçe ondalık/ayırıcı
# --------------------------------------------------------------------------- #


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (Decimal("204345.00"), "204.345,00"),
        (Decimal("211845.00"), "211.845,00"),
        (Decimal("0.03"), "0,03"),
        (Decimal("1234567.5"), "1.234.567,50"),
        (Decimal("999.99"), "999,99"),
        (Decimal("1000000"), "1.000.000,00"),
        (Decimal("0"), "0,00"),
    ],
)
def test_format_money_uses_turkish_separators(value, expected):
    """Tutar binlik ``.`` ve ondalık ``,`` ile 2 basamak gösterilir."""
    assert format_money(value) == expected


def test_format_money_missing_is_not_zero():
    """Eksik tutar sıfır gösterilmez; açıkça "Eksik veri" yazılır."""
    assert format_money(None) == MISSING_TEXT
    assert MISSING_TEXT != "0,00"
    assert MISSING_TEXT != "0"


def test_format_weight_uses_three_decimals():
    assert format_weight(Decimal("7170.000")) == "7.170,000"
    assert format_weight(Decimal("97.999986")) == "98,000"
    assert format_weight(None) == MISSING_TEXT


def test_format_rate_uses_four_decimals():
    assert format_rate(Decimal("1.0460")) == "1,0460"
    assert format_rate(Decimal("0.1667")) == "0,1667"
    assert format_rate(None) == MISSING_TEXT


def test_format_quantity_uses_thousands_separator():
    assert format_quantity(500) == "500"
    assert format_quantity(1_000_000) == "1.000.000"


def test_format_days_renders_integer_without_decimals():
    assert format_days(15) == "15"
    assert format_days(15.0) == "15"
    assert format_days(None) == MISSING_TEXT


def test_format_days_keeps_non_integer_value():
    assert format_days(2.5) == "2,5"


def test_format_days_passes_through_text():
    """Okunamayan termin sıfır veya hata değil, okunan metin olarak gösterilir."""
    assert format_days("yaklaşık 2 hafta") == "yaklaşık 2 hafta"


def test_formatting_rounds_only_for_display():
    """Gösterim yuvarlaması yarım yukarı; hesap değeri motorda kalır."""
    assert format_money(Decimal("0.025")) == "0,03"
    assert format_money(Decimal("0.0249")) == "0,02"


# --------------------------------------------------------------------------- #
# Karşılaştırma tablosu
# --------------------------------------------------------------------------- #


def test_comparison_rows_follow_source_order(tmp_path):
    parsed, _validation, result = _analyse(tmp_path)
    rows = build_comparison_rows(result, parsed)

    assert [row.row_number for row in rows] == [3, 4, 5]
    assert [row.supplier_name for row in rows] == [
        "Bursa Metal A.Ş.",
        "Gemlik Profil Ltd.",
        "İnegöl Çelik San.",
    ]


def test_comparison_row_values_match_calculator(tmp_path):
    """Tablodaki tutarlar motorun ürettiği değerlerin biçimlenmiş hâlidir."""
    parsed, _validation, result = _analyse(tmp_path)
    rows = {row.row_number: row for row in build_comparison_rows(result, parsed)}
    computed = {row.row_number: row for row in result.rows}

    for number, view in rows.items():
        source = computed[number]
        assert view.product_cost == format_money(source.product_cost_try)
        assert view.total_cost == format_money(source.total_supply_cost_try)
        assert view.shipping_cost == format_money(source.shipping_cost_try)
        assert view.weight == format_weight(source.total_weight_kg)
        assert view.unit_price == format_money(source.unit_price_try_per_kg)
        assert view.shipping_per_kg == format_rate(source.shipping_per_kg_try)
        assert view.quantity == format_quantity(source.quantity_pcs)


def test_comparison_row_shows_payment_terms_verbatim(tmp_path):
    """Ödeme koşulu orijinal metin olarak gösterilir; yeniden yazılmaz."""
    parsed, _validation, result = _analyse(tmp_path)
    rows = {row.row_number: row for row in build_comparison_rows(result, parsed)}

    assert rows[4].payment_terms == "%50 peşin, %50 teslimatta"
    assert rows[5].payment_terms == "Peşin"


def test_comparison_row_shows_delivery_days(tmp_path):
    parsed, _validation, result = _analyse(tmp_path)
    rows = {row.row_number: row for row in build_comparison_rows(result, parsed)}

    assert rows[3].delivery_days == "15"
    assert rows[4].delivery_days == "25"
    assert rows[5].delivery_days == "18"


def test_missing_shipping_cells_are_marked_not_zero(tmp_path):
    """Nakliyesi eksik satırda nakliye, toplam ve nakliye/kg "Eksik veri" olur."""
    parsed, _validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "H3", None)
    )
    rows = {row.row_number: row for row in build_comparison_rows(result, parsed)}

    assert rows[3].shipping_cost == MISSING_TEXT
    assert rows[3].total_cost == MISSING_TEXT
    assert rows[3].shipping_per_kg == MISSING_TEXT
    # Ürün bedeli nakliyeden bağımsız olduğu için yine gösterilir.
    assert rows[3].product_cost == "204.345,00"


def test_explicit_zero_shipping_is_shown_as_zero(tmp_path):
    """Açıkça girilen 0 nakliye "0,00" olarak gösterilir, eksik sayılmaz."""
    parsed, _validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "H3", 0)
    )
    row = next(
        item for item in build_comparison_rows(result, parsed) if item.row_number == 3
    )

    assert row.shipping_cost == "0,00"
    assert row.total_cost == "204.345,00"
    assert row.shipping_per_kg == "0,0000"
    assert row.is_complete is True


def test_missing_row_carries_note(tmp_path):
    parsed, _validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "H3", None)
    )
    rows = {row.row_number: row for row in build_comparison_rows(result, parsed)}

    assert rows[3].missing_note == "Nakliye bilgisi eksik"
    assert rows[4].missing_note == ""


def test_lowest_flags_come_from_calculator_marks(tmp_path):
    """İşaretler motordan okunur; sunum katmanı kendi karşılaştırmasını yapmaz."""
    parsed, _validation, result = _analyse(tmp_path)
    rows = {row.row_number: row for row in build_comparison_rows(result, parsed)}

    assert rows[3].is_lowest_product is True
    assert rows[3].is_lowest_total is True
    assert rows[3].is_lowest is True
    assert rows[4].is_lowest is False
    assert rows[5].is_lowest is False


def test_lowest_flags_track_separate_marks(tmp_path):
    """Ürün bedeli ve toplam maliyet farklı satırları işaretleyebilir."""
    parsed, _validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "H3", 40_000)
    )
    rows = {row.row_number: row for row in build_comparison_rows(result, parsed)}

    assert rows[3].is_lowest_product is True
    assert rows[3].is_lowest_total is False
    assert rows[4].is_lowest_product is False
    assert rows[4].is_lowest_total is True


def test_ties_mark_every_row(tmp_path):
    def mutate(workbook) -> None:
        sheet = workbook["Supplier Quotes"]
        sheet["F4"] = 2.39
        sheet["G4"] = 28.5
        sheet["H4"] = 7500

    parsed, _validation, result = _analyse(tmp_path, mutate)
    rows = {row.row_number: row for row in build_comparison_rows(result, parsed)}

    assert rows[3].is_lowest_total is True
    assert rows[4].is_lowest_total is True
    assert rows[5].is_lowest_total is False


# --------------------------------------------------------------------------- #
# Tablo HTML
# --------------------------------------------------------------------------- #


def test_table_html_contains_every_supplier(tmp_path):
    parsed, _validation, result = _analyse(tmp_path)
    html = comparison_table_html(build_comparison_rows(result, parsed))

    for name in ("Bursa Metal A.Ş.", "Gemlik Profil Ltd.", "İnegöl Çelik San."):
        assert name in html
    assert html.count("<tr>") == 4  # 1 başlık + 3 veri satırı


def test_table_html_marks_lowest_cell_with_badge(tmp_path):
    parsed, _validation, result = _analyse(tmp_path)
    html = comparison_table_html(build_comparison_rows(result, parsed))

    assert "is-lowest" in html
    assert LOWEST_BADGE in html


def test_table_html_marks_missing_cells(tmp_path):
    parsed, _validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "H3", None)
    )
    html = comparison_table_html(build_comparison_rows(result, parsed))

    assert "is-missing" in html
    assert MISSING_TEXT in html


def test_table_html_escapes_user_text(tmp_path):
    """Dosyadaki metin HTML olarak yorumlanmaz; kaçışlanır."""
    def mutate(workbook) -> None:
        sheet = workbook["Supplier Quotes"]
        sheet["A3"] = "<script>alert(1)</script>"
        sheet["K3"] = "<b>Peşin</b>"

    parsed, _validation, result = _analyse(tmp_path, mutate)
    html = comparison_table_html(build_comparison_rows(result, parsed))

    assert "<script>" not in html
    assert "&lt;script&gt;" in html
    assert "<b>Peşin</b>" not in html


def test_table_html_empty_rows_returns_empty_string():
    assert comparison_table_html(()) == ""


# --------------------------------------------------------------------------- #
# Özet kartları
# --------------------------------------------------------------------------- #


def test_kpi_cards_report_counts(tmp_path):
    parsed, _validation, result = _analyse(tmp_path)
    cards = build_kpis(parsed, result)

    assert len(cards) == 4
    assert cards[0].value == "3"  # tedarikçi sayısı
    assert cards[1].value == "3"  # hesaplanabilen teklif


def test_kpi_cards_report_lowest_amounts(tmp_path):
    parsed, _validation, result = _analyse(tmp_path)
    cards = {card.label: card for card in build_kpis(parsed, result)}

    assert cards["En düşük ürün bedeli"].value == "204.345,00 TL"
    assert cards["En düşük toplam tedarik maliyeti"].value == "211.845,00 TL"
    assert "Bursa Metal A.Ş." in cards["En düşük ürün bedeli"].caption


def test_kpi_caption_lists_all_tied_suppliers(tmp_path):
    def mutate(workbook) -> None:
        sheet = workbook["Supplier Quotes"]
        sheet["F4"] = 2.39
        sheet["G4"] = 28.5
        sheet["H4"] = 7500

    parsed, _validation, result = _analyse(tmp_path, mutate)
    card = next(
        c for c in build_kpis(parsed, result) if c.label == "En düşük toplam tedarik maliyeti"
    )

    assert "Bursa Metal A.Ş." in card.caption
    assert "Gemlik Profil Ltd." in card.caption


def test_kpi_counts_excluded_rows_in_caption(tmp_path):
    parsed, _validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "G3", 0)
    )
    card = next(
        c for c in build_kpis(parsed, result) if c.label == "Hesaplanabilen teklif"
    )

    assert card.value == "2"
    assert "1 satır" in card.caption


def test_kpi_shows_dash_when_no_total_can_be_computed(tmp_path):
    """Hiçbir satırda toplam maliyet yoksa kart sıfır değil "—" gösterir."""

    def mutate(workbook) -> None:
        sheet = workbook["Supplier Quotes"]
        for row in (3, 4, 5):
            sheet[f"H{row}"] = None

    parsed, _validation, result = _analyse(tmp_path, mutate)
    cards = {card.label: card for card in build_kpis(parsed, result)}

    assert cards["En düşük toplam tedarik maliyeti"].value == "—"
    assert "eksik" in cards["En düşük toplam tedarik maliyeti"].caption.lower()


def test_kpi_product_mark_survives_missing_shipping(tmp_path):
    """Nakliye eksik olsa da ürün bedeli en düşüğü gösterilir."""
    parsed, _validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "H3", None)
    )
    cards = {card.label: card for card in build_kpis(parsed, result)}

    assert cards["En düşük ürün bedeli"].value == "204.345,00 TL"
    assert cards["En düşük toplam tedarik maliyeti"].value == "213.793,00 TL"


# --------------------------------------------------------------------------- #
# Grafikler
# --------------------------------------------------------------------------- #


def test_product_chart_includes_every_computed_row(tmp_path):
    _parsed, _validation, result = _analyse(tmp_path)
    figure = build_cost_figure(result, "product")

    assert figure is not None
    assert list(figure.data[0].x) == [204345.0, 208593.0, 215946.0]


def test_total_chart_excludes_missing_total(tmp_path):
    """Toplam maliyeti hesaplanamayan satır grafiğe dahil edilmez."""
    _parsed, _validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "H3", None)
    )
    figure = build_cost_figure(result, "total")

    assert figure is not None
    assert len(figure.data[0].x) == 2
    assert 204345.0 not in [round(value) for value in figure.data[0].x]
    assert list(figure.data[0].y) == ["Gemlik Profil Ltd.", "İnegöl Çelik San."]


def test_product_chart_still_includes_row_with_missing_shipping(tmp_path):
    _parsed, _validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "H3", None)
    )
    figure = build_cost_figure(result, "product")

    assert figure is not None
    assert "Bursa Metal A.Ş." in list(figure.data[0].y)


def test_chart_highlights_lowest_value(tmp_path):
    from procurement.presentation import CHART_COLOR_BASE, CHART_COLOR_LOWEST

    _parsed, _validation, result = _analyse(tmp_path)
    figure = build_cost_figure(result, "total")
    colors = list(figure.data[0].marker.color)

    assert colors[0] == CHART_COLOR_LOWEST
    assert colors[1] == CHART_COLOR_BASE
    assert colors[2] == CHART_COLOR_BASE


def test_chart_keeps_source_order(tmp_path):
    """Grafik sıralama üretmez; çubuk sırası Excel sırasıdır."""
    _parsed, _validation, result = _analyse(tmp_path)
    figure = build_cost_figure(result, "product")

    assert list(figure.data[0].y) == [
        "Bursa Metal A.Ş.",
        "Gemlik Profil Ltd.",
        "İnegöl Çelik San.",
    ]


def test_chart_returns_none_when_no_candidate(tmp_path):
    """Grafik için değer yoksa çizim üretilmez (boş grafik gösterilmez)."""
    parsed, validation, result = _analyse(
        tmp_path, _set("Instructions", "B4", "HAYIR")
    )
    assert result.rows == ()
    assert build_cost_figure(result, "product") is None
    assert build_cost_figure(result, "total") is None


# --------------------------------------------------------------------------- #
# Bulgular
# --------------------------------------------------------------------------- #


def test_group_findings_separates_severities(tmp_path):
    _parsed, validation, _result = _analyse(tmp_path)
    groups = group_findings(validation)

    assert groups.errors == ()
    assert len(groups.warnings) == 2
    assert groups.infos == ()
    assert groups.has_any is True


def test_group_findings_collects_info_requests(tmp_path):
    _parsed, validation, _result = _analyse(
        tmp_path, _set("Supplier Quotes", "H3", None)
    )
    groups = group_findings(validation)

    assert len(groups.infos) == 1
    assert groups.infos[0].code == "I-SHIP-01"


def test_group_findings_without_findings_is_empty(tmp_path):
    """Bir satırdaki uyarı kaldırıldığında o bulgu grubu boşalır."""
    parsed, validation, _result = _analyse(
        tmp_path, _set("Supplier Quotes", "F5", 2.39)
    )
    groups = group_findings(validation)

    assert [f.code for f in groups.warnings] == ["W-PAY-01"]
    assert groups.errors == ()


def test_finding_context_includes_supplier_and_cell(tmp_path):
    _parsed, validation, _result = _analyse(tmp_path)
    warning = next(f for f in validation.warnings if f.code == "W-WT-01")

    context = finding_context(warning)

    assert "İnegöl Çelik San." in context
    assert "Supplier Quotes!F5" in context


def test_finding_context_handles_file_level_findings(tmp_path):
    _parsed, validation, _result = _analyse(
        tmp_path, _set("Instructions", "B4", "HAYIR")
    )
    finding = validation.file_errors[0]

    context = finding_context(finding)

    assert "Instructions" in context
    assert "B4" in context


def test_rejection_hints_explain_file_rejection(tmp_path):
    _parsed, validation, _result = _analyse(
        tmp_path, _set("Instructions", "B4", "HAYIR")
    )
    hints = rejection_hints(validation)

    assert hints
    assert any("EVET" in hint for hint in hints)


def test_rejection_hints_empty_when_file_accepted(tmp_path):
    _parsed, validation, _result = _analyse(tmp_path)

    assert validation.file_rejected is False
    assert rejection_hints(validation) == ()


def test_excluded_row_views_carry_codes(tmp_path):
    _parsed, _validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "G3", 0)
    )
    views = excluded_row_views(result)

    assert len(views) == 1
    assert views[0].row_number == 3
    assert "E-PRICE-01" in views[0].codes
    assert views[0].codes_text == "E-PRICE-01"


def test_file_rejection_excludes_every_row(tmp_path):
    _parsed, _validation, result = _analyse(
        tmp_path, _set("Instructions", "B4", "HAYIR")
    )
    views = excluded_row_views(result)

    assert result.file_rejected is True
    assert {view.row_number for view in views} == {3, 4, 5}


def test_kpi_computed_caption_handles_empty_file(tmp_path):
    """Teklif satırı yoksa kart bunu açıkça söyler, 'hepsi alındı' demez."""
    def mutate(workbook) -> None:
        sheet = workbook["Supplier Quotes"]
        for row in range(3, 6):
            for column in "ABCDEFGHIJKLMN":
                sheet[f"{column}{row}"] = None

    parsed, _validation, result = _analyse(tmp_path, mutate)
    cards = {card.label: card for card in build_kpis(parsed, result)}

    assert cards["Tedarikçi sayısı"].value == "0"
    assert cards["Hesaplanabilen teklif"].value == "0"
    assert cards["Hesaplanabilen teklif"].caption == "Dosyada teklif satırı yok"


def test_comparison_rows_empty_when_no_quote_rows(tmp_path):
    def mutate(workbook) -> None:
        sheet = workbook["Supplier Quotes"]
        for row in range(3, 6):
            for column in "ABCDEFGHIJKLMN":
                sheet[f"{column}{row}"] = None

    parsed, _validation, result = _analyse(tmp_path, mutate)

    assert build_comparison_rows(result, parsed) == ()
    assert comparison_table_html(()) == ""
    assert build_cost_figure(result, "product") is None


# --------------------------------------------------------------------------- #
# Sınırlar
# --------------------------------------------------------------------------- #


def test_presentation_exposes_no_ranking_or_scoring(tmp_path):
    """Sunum katmanı sıralama/puan/kazanan alanı eklemez."""
    parsed, _validation, result = _analyse(tmp_path)
    row = build_comparison_rows(result, parsed)[0]

    forbidden = ("rank", "score", "winner", "weighted", "best", "recommend")
    for attribute in dir(row):
        assert not any(word in attribute.lower() for word in forbidden)


def test_comparison_rows_do_not_recalculate(tmp_path):
    """Aynı girdi aynı çıktıyı verir; sunum katmanı durumsuzdur."""
    parsed, _validation, result = _analyse(tmp_path)

    first = build_comparison_rows(result, parsed)
    second = build_comparison_rows(result, parsed)

    assert first == second

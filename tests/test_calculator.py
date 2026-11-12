"""``procurement.calculator`` testleri — formüller, Decimal hassasiyeti, eksik veri.

Test stratejisi
---------------
* Gerçek şablon ``tmp_path`` içine kopyalanır; asıl dosya değiştirilmez.
* Hesaplama her zaman gerçek akışla çalıştırılır:
  ``parse_workbook(strict=False)`` → ``validate`` → ``calculate``.
  Böylece "yalnızca Validator'ın kabul ettiği veri" kuralı gerçekten sınanır.
* Beklenen değerler testte **ellerinle** hesaplanır, motorun çıktısıyla
  karşılaştırılır; motorun kendi sonucunu kendine doğrulattırmayız.
* Tarih bağımlılığı ``today=`` ile sabitlenir.
"""

from __future__ import annotations

import shutil
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import load_workbook

from procurement.calculator import (
    MISSING_SHIPPING,
    CalculationResult,
    ExcludedQuote,
    SupplierCalculation,
    calculate,
)
from procurement.excel_parser import parse_workbook
from procurement.validator import validate

TEMPLATE = Path(__file__).resolve().parents[1] / "templates" / "supplier_quote_template.xlsx"

#: Şablonun tarihleriyle uyumlu sabit "bugün".
TODAY = date(2026, 11, 12)

pytestmark = pytest.mark.unit


# --------------------------------------------------------------------------- #
# Yardımcılar
# --------------------------------------------------------------------------- #


def _build(tmp_path: Path, mutate=None, *, name: str = "calc.xlsx") -> Path:
    """Şablonun kopyasını üretir, isteğe bağlı değişiklik uygular."""
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


def _calculate(path: Path) -> CalculationResult:
    parsed = parse_workbook(path, strict=False)
    return calculate(parsed, validate(parsed, today=TODAY))


def _set(sheet: str, cell: str, value):
    def mutate(workbook) -> None:
        workbook[sheet][cell] = value

    return mutate


def _row(result: CalculationResult, row_number: int) -> SupplierCalculation:
    row = result.by_row(row_number)
    assert row is not None, f"satır {row_number} hesaplanmadı"
    return row


# --------------------------------------------------------------------------- #
# Formüller — tek tedarikçi, el ile doğrulanmış değerler
# --------------------------------------------------------------------------- #


def test_total_length_is_quantity_times_six(tmp_path):
    """Toplam uzunluk (m) = adet × 6."""
    result = _calculate(_build(tmp_path))
    row = _row(result, 3)

    assert row.quantity_pcs == 500
    assert row.total_length_m == Decimal("3000.000")


def test_total_weight_uses_declared_weight(tmp_path):
    """Toplam ağırlık = adet × 6 × tedarikçinin beyan ettiği kg/m.

    3000 m × 2,39 kg/m = 7170 kg. Referans teorik ağırlık (2,39) bu satırda
    beyanla aynı olduğu için ayrı bir testte ayrıştırılır.
    """
    result = _calculate(_build(tmp_path))
    row = _row(result, 3)

    assert row.weight_kg_per_m == Decimal("2.390")
    assert row.total_weight_kg == Decimal("7170.000")


def test_declared_weight_is_used_not_reference_weight(tmp_path):
    """Beyan ile referans farklıysa hesaplama beyanla yapılır (brief §6).

    Satır 4: 2,31 kg/m beyan. Referans 2,39 kullanılsaydı 7170 kg çıkardı;
    doğru sonuç 3000 × 2,31 = 6930 kg.
    """
    result = _calculate(_build(tmp_path))
    row = _row(result, 4)

    assert row.weight_kg_per_m == Decimal("2.310")
    assert row.total_weight_kg == Decimal("6930.000")
    assert row.total_weight_kg != Decimal("7170.000")


def test_out_of_reference_range_still_calculates_with_declared_value(tmp_path):
    """Referans aralık dışı kg/m (2,58) uyarı üretir ama hesaba girer (brief §6)."""
    result = _calculate(_build(tmp_path))
    row = _row(result, 5)

    assert row.weight_kg_per_m == Decimal("2.580")
    assert row.total_weight_kg == Decimal("7740.000")
    assert row.is_complete is True


def test_product_cost_is_weight_times_unit_price(tmp_path):
    """Ürün bedeli = toplam ağırlık × TL/kg. 7170 × 28,50 = 204.345,00."""
    result = _calculate(_build(tmp_path))
    row = _row(result, 3)

    assert row.unit_price_try_per_kg == Decimal("28.50")
    assert row.product_cost_try == Decimal("204345.00")


def test_shipping_per_kg_is_shipping_over_weight(tmp_path):
    """Kg başına nakliye = nakliye ÷ toplam ağırlık. 7500 ÷ 7170 = 1,04602…"""
    result = _calculate(_build(tmp_path))
    row = _row(result, 3)

    assert row.shipping_cost_try == Decimal("7500.00")
    assert row.shipping_per_kg_try == Decimal("1.0460")


def test_total_supply_cost_is_product_plus_shipping(tmp_path):
    """Toplam tedarik maliyeti = ürün bedeli + nakliye. 204.345 + 7500 = 211.845."""
    result = _calculate(_build(tmp_path))
    row = _row(result, 3)

    assert row.total_supply_cost_try == Decimal("211845.00")


def test_all_formulas_hold_together_for_every_row(tmp_path):
    """Üç şablon satırında formüller tutarlı biçimde uygulanır."""
    result = _calculate(_build(tmp_path))

    # (satır, adet, kg/m, TL/kg, nakliye, beklenen ağırlık, beklenen ürün bedeli)
    expected = [
        (3, 500, Decimal("2.39"), Decimal("28.50"), Decimal("7500"),
         Decimal("7170.000"), Decimal("204345.00")),
        (4, 500, Decimal("2.31"), Decimal("30.10"), Decimal("5200"),
         Decimal("6930.000"), Decimal("208593.00")),
        (5, 500, Decimal("2.58"), Decimal("27.90"), Decimal("9100"),
         Decimal("7740.000"), Decimal("215946.00")),
    ]

    for row_no, qty, kg_m, price, shipping, weight, product in expected:
        row = _row(result, row_no)
        assert row.quantity_pcs == qty
        assert row.total_length_m == Decimal(qty) * 6
        assert row.total_weight_kg == weight
        assert row.product_cost_try == product
        assert row.total_supply_cost_try == product + shipping


# --------------------------------------------------------------------------- #
# Decimal hassasiyeti ve yuvarlama
# --------------------------------------------------------------------------- #


def test_amounts_are_decimal_not_float(tmp_path):
    """Para alanları ``Decimal``; ikili kayan nokta artığı taşımaz."""
    result = _calculate(_build(tmp_path))
    row = _row(result, 3)

    for value in (
        row.product_cost_try,
        row.shipping_cost_try,
        row.total_supply_cost_try,
        row.shipping_per_kg_try,
        row.unit_price_try_per_kg,
        row.weight_kg_per_m,
        row.total_weight_kg,
    ):
        assert isinstance(value, Decimal)


def test_decimal_conversion_is_exact_for_float_inputs(tmp_path):
    """``Decimal(2.39)`` yerine ``Decimal('2.39')`` kullanıldığı doğrulanır.

    Excel'den float gelen 28.50 değeri tam 28,50 olmalı; ``Decimal(28.5)``
    kullanılsaydı da tam olurdu, bu yüzden 0,1 + 0,2 sınıfı hata aranır:
    fiyat 0.1 ve ağırlık 0.2 gibi değerlerde artık birikmemeli.

    Not: RFQ miktarı satır miktarıyla eşleşmelidir (``E-Q-01``), bu yüzden
    ikisi birlikte değiştirilir.
    """
    def mutate(workbook):
        workbook["RFQ"]["B9"] = 1
        sheet = workbook["Supplier Quotes"]
        sheet["E3"] = 1
        sheet["F3"] = 0.2
        sheet["G3"] = 0.1
        sheet["H3"] = 0

    result = _calculate(_build(tmp_path, mutate))
    row = _row(result, 3)

    # 1 × 6 × 0,2 = 1,2 kg → önce 0,1 artığıyla 1,2000000000000002 olmamalı.
    assert row.total_weight_kg == Decimal("1.200")
    assert row.product_cost_try == Decimal("0.12")


def test_money_rounds_half_up(tmp_path):
    """TRY tutarlar 2 ondalığa yuvarlanır.

    1 adet × 6 m × 1,0000833 kg/m = 6,0004998 kg; × 1,00 TL/kg =
    6,0004998 TL → 6,00.
    """
    def mutate(workbook):
        workbook["RFQ"]["B9"] = 1
        sheet = workbook["Supplier Quotes"]
        sheet["E3"] = 1
        sheet["F3"] = 1.0000833
        sheet["G3"] = 1
        sheet["H3"] = 0

    result = _calculate(_build(tmp_path, mutate))
    row = _row(result, 3)

    assert row.product_cost_try == Decimal("6.00")


def test_quantize_uses_round_half_up_not_half_even():
    """``_quantize`` tam yarıda yukarı yuvarlar (banker's rounding değil).

    Gerçek akışta tam .005 sınırına ulaşmak yapay olduğu için yuvarlama
    yardımcısı doğrudan sınanır: 0,025 → 0,03 (``ROUND_HALF_UP``) iken
    ``ROUND_HALF_EVEN`` 0,02 verirdi.
    """
    from procurement.calculator import _quantize

    quantum = Decimal("0.01")

    assert _quantize(Decimal("0.025"), quantum) == Decimal("0.03")
    assert _quantize(Decimal("0.015"), quantum) == Decimal("0.02")
    assert _quantize(Decimal("0.045"), quantum) == Decimal("0.05")
    # Aşağı yuvarlanması gereken değer yukarı çıkmaz.
    assert _quantize(Decimal("0.0249"), quantum) == Decimal("0.02")


def test_intermediate_values_are_not_rounded(tmp_path):
    """Ara adımda yuvarlama yapılmaz; ürün bedeli tam ağırlıktan üretilir.

    kg/m = 2,333333 ve adet = 7 → ağırlık 97,999986 kg (gösterim 98,000),
    ancak ürün bedeli yuvarlanmış ağırlıktan değil tam değerden hesaplanır.
    """
    def mutate(workbook):
        workbook["RFQ"]["B9"] = 7
        sheet = workbook["Supplier Quotes"]
        sheet["E3"] = 7
        sheet["F3"] = 2.333333
        sheet["G3"] = 10
        sheet["H3"] = 0

    result = _calculate(_build(tmp_path, mutate))
    row = _row(result, 3)

    # 7 × 6 = 42 m; 42 × 2,333333 = 97,999986 kg; × 10 = 979,99986 → 980,00
    assert row.total_weight_kg == Decimal("98.000")
    assert row.product_cost_try == Decimal("980.00")


def test_shipping_per_kg_rounds_to_four_decimals(tmp_path):
    """Kg başına nakliye 4 ondalığa yuvarlanır (birim maliyet hassasiyeti)."""
    def mutate(workbook):
        workbook["RFQ"]["B9"] = 1
        sheet = workbook["Supplier Quotes"]
        sheet["E3"] = 1
        sheet["F3"] = 1
        sheet["G3"] = 10
        sheet["H3"] = 1

    result = _calculate(_build(tmp_path, mutate))
    row = _row(result, 3)

    # 6 kg'a 1 TL nakliye → 1/6 = 0,16666… → 0,1667
    assert row.shipping_per_kg_try == Decimal("0.1667")


def test_large_quantity_stays_exact(tmp_path):
    """Büyük miktarda Decimal hassasiyeti korunur."""
    def mutate(workbook):
        workbook["RFQ"]["B9"] = 1_000_000
        sheet = workbook["Supplier Quotes"]
        sheet["E3"] = 1_000_000
        sheet["F3"] = 2.39
        sheet["G3"] = 28.5
        sheet["H3"] = 0

    result = _calculate(_build(tmp_path, mutate))
    row = _row(result, 3)

    # 6.000.000 m × 2,39 = 14.340.000 kg; × 28,50 = 408.690.000,00
    assert row.total_weight_kg == Decimal("14340000.000")
    assert row.product_cost_try == Decimal("408690000.00")


# --------------------------------------------------------------------------- #
# Eksik nakliye — sıfır kabul edilmez
# --------------------------------------------------------------------------- #


def test_missing_shipping_is_not_treated_as_zero(tmp_path):
    """Nakliye boş ≠ 0: toplam maliyet hesaplanmaz, eksik olarak işaretlenir."""
    result = _calculate(_build(tmp_path, _set("Supplier Quotes", "H3", None)))
    row = _row(result, 3)

    assert row.shipping_cost_try is None
    assert row.total_supply_cost_try is None
    assert row.shipping_per_kg_try is None
    assert row.total_supply_cost_try != Decimal("204345.00")
    assert MISSING_SHIPPING in row.missing_data
    assert row.is_complete is False


def test_missing_shipping_still_computes_product_cost(tmp_path):
    """Ürün bedeli nakliyeden bağımsızdır ve yine hesaplanır."""
    result = _calculate(_build(tmp_path, _set("Supplier Quotes", "H3", None)))
    row = _row(result, 3)

    assert row.product_cost_try == Decimal("204345.00")
    assert row.total_weight_kg == Decimal("7170.000")


def test_missing_shipping_marks_result_incomplete(tmp_path):
    """Eksik nakliye karşılaştırmayı tamamlanmış saydırmaz (brief §5.4)."""
    result = _calculate(_build(tmp_path, _set("Supplier Quotes", "H3", None)))

    assert result.is_complete is False
    assert result.missing_data_rows == (3,)


def test_missing_shipping_row_lists_in_missing_data(tmp_path):
    def mutate(workbook):
        sheet = workbook["Supplier Quotes"]
        sheet["H3"] = None
        sheet["H4"] = None

    result = _calculate(_build(tmp_path, mutate))

    assert result.missing_data_rows == (3, 4)
    assert _row(result, 5).is_complete is True


def test_zero_shipping_is_a_real_value(tmp_path):
    """Açıkça girilen 0 nakliye geçerlidir; eksik veri sayılmaz."""
    result = _calculate(_build(tmp_path, _set("Supplier Quotes", "H3", 0)))
    row = _row(result, 3)

    assert row.shipping_cost_try == Decimal("0.00")
    assert row.total_supply_cost_try == Decimal("204345.00")
    assert row.shipping_per_kg_try == Decimal("0.0000")
    assert row.is_complete is True


def test_shipping_without_source_is_excluded_by_validator(tmp_path):
    """Nakliye dolu ama kaynak geçersizse satır Validator tarafından bloklanır."""
    result = _calculate(_build(tmp_path, _set("Supplier Quotes", "I3", None)))

    excluded = {item.row_number for item in result.excluded}
    assert 3 in excluded
    assert result.by_row(3) is None


# --------------------------------------------------------------------------- #
# Sıfır / negatif / geçersiz değerler
# --------------------------------------------------------------------------- #


@pytest.mark.parametrize("cell", ["E3", "F3", "G3"])
def test_zero_value_row_is_excluded(tmp_path, cell):
    """Sıfır miktar, kg/m veya TL/kg hesaplamaya alınmaz."""
    result = _calculate(_build(tmp_path, _set("Supplier Quotes", cell, 0)))

    assert result.by_row(3) is None
    assert 3 in {item.row_number for item in result.excluded}


@pytest.mark.parametrize("cell", ["E3", "F3", "G3"])
def test_negative_value_row_is_excluded(tmp_path, cell):
    """Negatif miktar, kg/m veya TL/kg hesaplamaya alınmaz."""
    result = _calculate(_build(tmp_path, _set("Supplier Quotes", cell, -5)))

    assert result.by_row(3) is None
    assert 3 in {item.row_number for item in result.excluded}


def test_negative_shipping_row_is_excluded(tmp_path):
    result = _calculate(_build(tmp_path, _set("Supplier Quotes", "H3", -100)))

    assert result.by_row(3) is None
    assert 3 in {item.row_number for item in result.excluded}


def test_excluded_row_carries_validation_codes(tmp_path):
    """Hesaplamaya alınmayan satır gerekçesini doğrulama kodlarıyla taşır."""
    result = _calculate(_build(tmp_path, _set("Supplier Quotes", "G3", 0)))

    excluded = next(item for item in result.excluded if item.row_number == 3)
    assert isinstance(excluded, ExcludedQuote)
    assert "E-PRICE-01" in excluded.codes
    assert excluded.supplier_name == "Bursa Metal A.Ş."


def test_fractional_quantity_is_excluded(tmp_path):
    """Adet tam sayı olmalıdır."""
    result = _calculate(_build(tmp_path, _set("Supplier Quotes", "E3", 12.5)))

    assert result.by_row(3) is None


def test_unreadable_number_is_excluded(tmp_path):
    """Okunamayan TL/kg hesaplamaya alınmaz."""
    result = _calculate(_build(tmp_path, _set("Supplier Quotes", "G3", "abc")))

    assert result.by_row(3) is None


# --------------------------------------------------------------------------- #
# Validator kapısı — motor kendi kuralını üretmez
# --------------------------------------------------------------------------- #


def test_file_level_error_blocks_all_calculations(tmp_path):
    """Dosya reddedildiyse hiçbir satır hesaplanmaz."""
    result = _calculate(_build(tmp_path, _set("Instructions", "B4", "HAYIR")))

    assert result.file_rejected is True
    assert result.rows == ()
    assert result.lowest_product_cost is None
    assert result.lowest_total_cost is None
    assert result.is_complete is False


def test_file_level_error_lists_every_row_as_excluded(tmp_path):
    result = _calculate(_build(tmp_path, _set("RFQ", "B12", "USD")))

    assert result.file_rejected is True
    assert {item.row_number for item in result.excluded} == {3, 4, 5}


def test_warning_does_not_exclude_row(tmp_path):
    """Uyarı (referans dışı kg/m, karma ödeme) satırı düşürmez."""
    result = _calculate(_build(tmp_path))

    # Şablonun kendi uyarıları: satır 4 karma ödeme, satır 5 aralık dışı kg/m.
    assert result.by_row(4) is not None
    assert result.by_row(5) is not None
    assert result.file_rejected is False


def test_one_bad_row_does_not_affect_others(tmp_path):
    """Satır hatası yalnızca o satırı etkiler (spec §7.1)."""
    result = _calculate(_build(tmp_path, _set("Supplier Quotes", "G4", 0)))

    assert result.by_row(3) is not None
    assert result.by_row(4) is None
    assert result.by_row(5) is not None
    assert result.excluded[0].row_number == 4


def test_duplicate_supplier_rows_are_both_excluded(tmp_path):
    """Mükerrer tedarikçide her iki satır da karşılaştırmaya alınmaz."""
    def mutate(workbook):
        workbook["Supplier Quotes"]["A4"] = "Bursa Metal A.Ş."

    result = _calculate(_build(tmp_path, mutate))

    assert result.by_row(3) is None
    assert result.by_row(4) is None
    assert {3, 4} <= {item.row_number for item in result.excluded}


def test_calculator_does_not_revalidate(tmp_path):
    """Motor, Validator'ın kabul ettiği satırı kendi kuralıyla elemez.

    Sayısal olarak ``weight_kg_per_m`` bakımından geçerli ama referans aralık
    dışı olan satır (2,58) hesaplanır; motor ek doğrulama yapmaz.
    """
    result = _calculate(_build(tmp_path))

    assert result.by_row(5) is not None
    assert _row(result, 5).weight_kg_per_m == Decimal("2.580")


# --------------------------------------------------------------------------- #
# En düşük maliyet işaretleri
# --------------------------------------------------------------------------- #


def test_lowest_product_cost_is_marked(tmp_path):
    """En düşük ürün bedeli ayrı işaretlenir."""
    result = _calculate(_build(tmp_path))

    assert result.lowest_product_cost is not None
    assert result.lowest_product_cost.amount_try == Decimal("204345.00")
    assert result.lowest_product_cost.row_numbers == (3,)


def test_lowest_total_cost_is_marked(tmp_path):
    """En düşük toplam tedarik maliyeti ayrı işaretlenir."""
    result = _calculate(_build(tmp_path))

    assert result.lowest_total_cost is not None
    assert result.lowest_total_cost.amount_try == Decimal("211845.00")
    assert result.lowest_total_cost.row_numbers == (3,)


def test_product_and_total_marks_are_independent(tmp_path):
    """İki işaret ayrı hesaplanır ve farklı satırları gösterebilir (brief §6).

    Satır 3 en ucuz ürün bedeline sahiptir. Nakliyesi çok yüksek yapılırsa
    toplam maliyet sıralaması değişir; iki işaret ayrışır.
    """
    def mutate(workbook):
        workbook["Supplier Quotes"]["H3"] = 40_000

    result = _calculate(_build(tmp_path, mutate))

    assert result.lowest_product_cost.row_numbers == (3,)
    assert result.lowest_total_cost.row_numbers == (4,)
    assert result.lowest_product_cost.row_numbers != result.lowest_total_cost.row_numbers


def test_lowest_marks_exclude_incomplete_rows(tmp_path):
    """Eksik nakliyeli satır toplam maliyet işaretine aday olamaz."""
    result = _calculate(_build(tmp_path, _set("Supplier Quotes", "H3", None)))

    assert result.lowest_total_cost is not None
    assert 3 not in result.lowest_total_cost.row_numbers
    # Ürün bedeli nakliyeden bağımsız olduğu için aday kalır.
    assert result.lowest_product_cost.row_numbers == (3,)


def test_lowest_mark_lists_all_ties(tmp_path):
    """Eşit en düşük tutarda tüm satırlar listelenir; keyfî seçim yapılmaz."""
    def mutate(workbook):
        sheet = workbook["Supplier Quotes"]
        sheet["F4"] = 2.39
        sheet["G4"] = 28.5
        sheet["H4"] = 7500

    result = _calculate(_build(tmp_path, mutate))

    assert result.lowest_product_cost.row_numbers == (3, 4)
    assert len(result.lowest_product_cost.suppliers) == 2


def test_single_supplier_gets_both_marks(tmp_path):
    """Tek tedarikçide iki işaret de aynı satırı gösterir."""
    def mutate(workbook):
        sheet = workbook["Supplier Quotes"]
        for row in (4, 5):
            for column in "ABCDEFGHIJKLMN":
                sheet[f"{column}{row}"] = None

    result = _calculate(_build(tmp_path, mutate))

    assert len(result.rows) == 1
    assert result.lowest_product_cost.row_numbers == (3,)
    assert result.lowest_total_cost.row_numbers == (3,)


# --------------------------------------------------------------------------- #
# Sınırlar: sıralama / puan / kazanan yok
# --------------------------------------------------------------------------- #


def test_rows_keep_source_order(tmp_path):
    """Satırlar kaynak (Excel) sırasında döner; maliyete göre sıralanmaz."""
    result = _calculate(_build(tmp_path))

    assert [row.row_number for row in result.rows] == [3, 4, 5]
    # En pahalı toplam maliyet son satırdadır; yine de son sırada kalır.
    assert result.rows[-1].total_supply_cost_try == Decimal("225046.00")


def test_result_has_no_ranking_or_scoring_fields():
    """Ağırlıklı puan, skor, kazanan veya sıra alanı yoktur (brief §6)."""
    forbidden = ("rank", "score", "winner", "weighted", "cheapest", "best", "recommend")

    for attribute in dir(CalculationResult):
        lowered = attribute.lower()
        assert not any(word in lowered for word in forbidden), attribute

    for attribute in dir(SupplierCalculation):
        lowered = attribute.lower()
        assert not any(word in lowered for word in forbidden), attribute


def test_payment_terms_are_not_part_of_calculation(tmp_path):
    """Ödeme koşulu değişse de maliyet değişmez (brief §6)."""
    def mutate(workbook):
        workbook["Supplier Quotes"]["K3"] = "%50 peşin, %50 teslimatta"

    baseline = _row(_calculate(_build(tmp_path)), 3)
    changed = _row(
        _calculate(_build(tmp_path, mutate, name="pay.xlsx")), 3
    )

    assert changed.product_cost_try == baseline.product_cost_try
    assert changed.total_supply_cost_try == baseline.total_supply_cost_try
    assert changed.shipping_per_kg_try == baseline.shipping_per_kg_try


def test_vat_is_not_applied(tmp_path):
    """KDV eklenmez: toplam maliyet tam olarak ürün bedeli + nakliyedir."""
    row = _row(_calculate(_build(tmp_path)), 3)

    assert row.total_supply_cost_try == row.product_cost_try + row.shipping_cost_try
    # KDV'li bir tutar (×1,20) üretilmediği doğrulanır.
    assert row.total_supply_cost_try != (row.product_cost_try * Decimal("1.20")).quantize(
        Decimal("0.01")
    )


def test_calculator_does_not_modify_source_file(tmp_path):
    """Hesaplama dosyaya yazmaz."""
    import hashlib

    path = _build(tmp_path)
    before = hashlib.sha256(path.read_bytes()).hexdigest()

    _calculate(path)

    assert hashlib.sha256(path.read_bytes()).hexdigest() == before


# --------------------------------------------------------------------------- #
# Çok tedarikçili senaryolar
# --------------------------------------------------------------------------- #


def test_many_suppliers_are_all_calculated(tmp_path):
    """Beş tedarikçi bağımsız hesaplanır."""
    def mutate(workbook):
        sheet = workbook["Supplier Quotes"]
        extra = [
            ("Ege Profil A.Ş.", "EP-1", 2.40, 35.00, 4000),
            ("Marmara Çelik", "MC-2", 2.45, 29.75, 6000),
        ]
        for offset, (name, quote_no, kg_m, price, shipping) in enumerate(extra):
            row = 6 + offset
            sheet[f"A{row}"] = name
            sheet[f"B{row}"] = quote_no
            sheet[f"C{row}"] = "40x40x2"
            sheet[f"D{row}"] = 6
            sheet[f"E{row}"] = 500
            sheet[f"F{row}"] = kg_m
            sheet[f"G{row}"] = price
            sheet[f"H{row}"] = shipping
            sheet[f"I{row}"] = "Supplier Quote"
            sheet[f"J{row}"] = 20
            sheet[f"K{row}"] = "Peşin"
            sheet[f"L{row}"] = date(2026, 11, 12)
            sheet[f"M{row}"] = date(2026, 12, 31)

    result = _calculate(_build(tmp_path, mutate))

    assert len(result.rows) == 5
    assert [row.row_number for row in result.rows] == [3, 4, 5, 6, 7]
    assert result.is_complete is True
    # Satır 6: 3000 × 2,40 = 7200 kg; × 35,00 = 252.000,00
    assert _row(result, 6).product_cost_try == Decimal("252000.00")
    # Satır 7: 3000 × 2,45 = 7350 kg; × 29,75 = 218.662,50
    assert _row(result, 7).product_cost_try == Decimal("218662.50")


def test_lowest_total_cost_can_come_from_any_supplier(tmp_path):
    """Toplam maliyeti en düşük tedarikçi, ürün bedeli en düşük olandan farklı olabilir."""
    def mutate(workbook):
        sheet = workbook["Supplier Quotes"]
        sheet["A6"] = "Düşük Nakliye Ltd."
        sheet["B6"] = "DN-1"
        sheet["C6"] = "40x40x2"
        sheet["D6"] = 6
        sheet["E6"] = 500
        sheet["F6"] = 2.39
        sheet["G6"] = 28.6
        sheet["H6"] = 100
        sheet["I6"] = "Supplier Quote"
        sheet["J6"] = 20
        sheet["K6"] = "Peşin"
        sheet["L6"] = date(2026, 11, 12)
        sheet["M6"] = date(2026, 12, 31)

    result = _calculate(_build(tmp_path, mutate))

    # Satır 3 ürün bedeli en düşük; satır 6 toplam maliyet en düşük.
    assert result.lowest_product_cost.row_numbers == (3,)
    assert result.lowest_total_cost.row_numbers == (6,)
    # 3000 × 2,39 = 7170 kg; × 28,60 = 205.062,00; + 100 = 205.162,00
    assert result.lowest_total_cost.amount_try == Decimal("205162.00")


def test_mixed_complete_and_incomplete_rows(tmp_path):
    """Bir satırda eksik nakliye varken diğerleri tamamdır."""
    result = _calculate(_build(tmp_path, _set("Supplier Quotes", "H4", None)))

    assert result.is_complete is False
    assert result.missing_data_rows == (4,)
    assert _row(result, 3).is_complete is True
    assert _row(result, 5).is_complete is True
    assert result.lowest_total_cost.row_numbers == (3,)


def test_duplicate_quote_number_is_only_a_warning(tmp_path):
    """Aynı teklif numarası uyarıdır; hesaplama sürer."""
    def mutate(workbook):
        workbook["Supplier Quotes"]["A4"] = "Gemlik Profil Ltd."
        workbook["Supplier Quotes"]["B4"] = "BM-2026-1187"

    result = _calculate(_build(tmp_path, mutate))

    # Farklı tedarikçi + aynı teklif no → W-DUP-02, satırlar hesaplanır.
    assert len(result.rows) == 3

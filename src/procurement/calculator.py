"""Maliyet hesaplama motoru — ``PROJECT_BRIEF.md`` §6, ``EXCEL_TEMPLATE_SPEC.md`` §10.

Sorumluluk
----------
Yalnızca **Validator tarafından kabul edilen** satırlar için maliyet alanlarını
üretir. Hesaplama tamamen Python tarafında ve deterministiktir; Excel'den
hesaplanmış hiçbir değer okunmaz.

Formüller (spec §10):

1. Toplam uzunluk (m) = ``Quantity (pcs)`` × 6
2. Toplam ağırlık (kg) = ``Quantity (pcs)`` × 6 × ``Weight (kg/m)``
3. Ürün bedeli (TRY, KDV hariç) = toplam ağırlık × ``Unit Price (TRY/kg)``
4. Kg başına nakliye (TRY/kg) = nakliye ÷ toplam ağırlık
5. Toplam tedarik maliyeti (TRY) = ürün bedeli + nakliye

Değişmez kurallar
-----------------
* Ağırlık **tedarikçinin beyan ettiği kg/m** ile hesaplanır. Referans teorik
  ağırlık (2,39 kg/m) yalnızca Validator'ın uyarı üretmesinde kullanılır;
  burada hiçbir düzeltme yapılmaz.
* Nakliye bilgisi eksikse **sıfır kabul edilmez**. Ürün bedeli hesaplanır,
  ancak toplam tedarik maliyeti ve kg başına nakliye ``None`` kalır ve satır
  ``missing_data`` ile işaretlenir.
* **KDV hesaplanmaz, eklenmez, çıkarılmaz.** Tüm tutarlar KDV hariç TRY'dir.
* **Ödeme koşulları maliyet hesabına dahil edilmez**; vade günü yalnızca
  karşılaştırma bilgisidir.
* En düşük ürün bedeli ve en düşük toplam tedarik maliyeti **ayrı ayrı**
  işaretlenir. Otomatik kazanan, ağırlıklı puan, bileşik skor veya tedarikçi
  sıralama algoritması **yoktur**; satırlar kaynak (Excel) sırasında döner.

Para birimi
-----------
Tutarlar :class:`decimal.Decimal` ile hesaplanır. Ara adımlarda yuvarlama
yapılmaz; ürün bedeli yuvarlanmamış toplam ağırlıktan üretilir, böylece
hatalar birikmez. Yalnızca çıktı alanları yuvarlanır:

* TRY tutarlar → 2 ondalık (``ROUND_HALF_UP``)
* Toplam ağırlık/uzunluk → 3 ondalık
* Kg başına nakliye → 4 ondalık
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal
from typing import Any, Final

from .excel_parser import ParsedWorkbook
from .validator import ValidationResult

__all__ = [
    "MISSING_SHIPPING",
    "CalculationResult",
    "CostMark",
    "ExcludedQuote",
    "SupplierCalculation",
    "calculate",
]

#: Profil uzunluğu sabittir; Validator satır uzunluğunu 6 m ile eşleştirir.
PROFILE_LENGTH_M: Final[Decimal] = Decimal(6)

MONEY_QUANTUM: Final[Decimal] = Decimal("0.01")
WEIGHT_QUANTUM: Final[Decimal] = Decimal("0.001")
RATE_QUANTUM: Final[Decimal] = Decimal("0.0001")

#: ``missing_data`` içinde kullanılan eksik-veri işareti.
MISSING_SHIPPING: Final[str] = "shipping"


# --------------------------------------------------------------------------- #
# Sayı dönüşümü
# --------------------------------------------------------------------------- #


def _decimal(value: Any) -> Decimal | None:
    """Değeri kayıpsız ``Decimal``'e çevirir; çevrilemiyorsa ``None``.

    ``float`` değerler ``str()`` üzerinden çevrilir: ``Decimal(2.39)`` ikili
    kayan nokta artığını taşırken ``Decimal("2.39")`` tam olarak 2,39 olur.
    """
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, int):
        return Decimal(value)
    if isinstance(value, float):
        if not math.isfinite(value):
            return None
        return Decimal(str(value))
    return None


def _positive_int(value: Any) -> int | None:
    """Pozitif tam sayıya çevirir; değilse ``None``."""
    number = _decimal(value)
    if number is None or number <= 0 or number != number.to_integral_value():
        return None
    return int(number)


def _quantize(value: Decimal, quantum: Decimal) -> Decimal:
    return value.quantize(quantum, rounding=ROUND_HALF_UP)


# --------------------------------------------------------------------------- #
# Sonuç modeli
# --------------------------------------------------------------------------- #


@dataclass(frozen=True, slots=True)
class SupplierCalculation:
    """Tek bir tedarikçi satırının hesaplanmış maliyet dökümü.

    Tüm tutarlar KDV hariç TRY'dir. ``shipping_cost_try`` ``None`` ise nakliye
    bilgisi yoktur: sıfır **değildir** ve ``total_supply_cost_try`` de ``None``
    kalır. Bu durum :attr:`missing_data` içinde açıkça işaretlenir.
    """

    row_number: int
    supplier_name: str
    supplier_quote_number: str | None
    quantity_pcs: int
    weight_kg_per_m: Decimal
    unit_price_try_per_kg: Decimal
    total_length_m: Decimal
    total_weight_kg: Decimal
    product_cost_try: Decimal
    shipping_cost_try: Decimal | None
    shipping_cost_source: str | None
    shipping_per_kg_try: Decimal | None
    total_supply_cost_try: Decimal | None
    missing_data: tuple[str, ...] = ()

    @property
    def is_complete(self) -> bool:
        """Bu satır için tüm maliyet alanları hesaplanabildi mi."""
        return not self.missing_data


@dataclass(frozen=True, slots=True)
class ExcludedQuote:
    """Hesaplamaya alınmayan satır ve gerekçesi."""

    row_number: int
    supplier_name: str | None
    reason: str
    codes: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class CostMark:
    """En düşük değer işareti.

    Bir sıralama değildir: yalnızca en düşük tutarı ve o tutara ulaşan
    satırları taşır. Eşitlik varsa tüm satırlar listelenir.
    """

    amount_try: Decimal
    row_numbers: tuple[int, ...]
    suppliers: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class CalculationResult:
    """Hesaplama çıktısı.

    ``rows`` yalnızca hesaplamaya uygun satırları **kaynak sırasında** taşır;
    sıralama, puan, kazanan veya öneri içermez.
    """

    rows: tuple[SupplierCalculation, ...]
    excluded: tuple[ExcludedQuote, ...]
    file_rejected: bool
    lowest_product_cost: CostMark | None
    lowest_total_cost: CostMark | None

    @property
    def is_complete(self) -> bool:
        """Karşılaştırma için gerekli tüm veri hesaplanabildi mi.

        Dosya reddedilmediyse, en az bir satır varsa ve hiçbir satırda eksik
        veri yoksa tamamdır. Eksik nakliye karşılaştırmayı tamamlanmış
        saydırmaz (brief §5.4).
        """
        return (
            not self.file_rejected
            and bool(self.rows)
            and all(row.is_complete for row in self.rows)
        )

    @property
    def missing_data_rows(self) -> tuple[int, ...]:
        """Eksik veri nedeniyle tamamlanamayan satırlar."""
        return tuple(row.row_number for row in self.rows if row.missing_data)

    def by_row(self, row_number: int) -> SupplierCalculation | None:
        for row in self.rows:
            if row.row_number == row_number:
                return row
        return None


# --------------------------------------------------------------------------- #
# Hesaplama
# --------------------------------------------------------------------------- #


def _blocking_codes(quote_row: int, validation: ValidationResult) -> tuple[str, ...]:
    """Satırı bloklayan doğrulama kodları; blok yoksa boş demet."""
    return tuple(
        f.code
        for f in validation.errors
        if f.row_number == quote_row and not f.is_file_level
    )


def _calculate_row(
    quote: Any,
) -> tuple[SupplierCalculation | None, str | None]:
    """Tek satırı hesaplar.

    Returns:
        ``(hesap, None)`` başarılı; ``(None, neden)`` girdi kullanılamıyor.
    """
    quantity = _positive_int(quote.quantity_pcs)
    if quantity is None:
        return None, "Miktar pozitif tam sayı olarak kullanılamıyor."

    weight_per_m = _decimal(quote.weight_kg_per_m)
    if weight_per_m is None or weight_per_m <= 0:
        return None, "kg/m pozitif sayı olarak kullanılamıyor."

    unit_price = _decimal(quote.unit_price_try_per_kg)
    if unit_price is None or unit_price <= 0:
        return None, "TL/kg pozitif sayı olarak kullanılamıyor."

    total_length = Decimal(quantity) * PROFILE_LENGTH_M
    total_weight = total_length * weight_per_m
    product_cost = total_weight * unit_price

    shipping = _decimal(quote.shipping_cost_try)
    missing: list[str] = []

    if shipping is None:
        # Nakliye boş ≠ 0: toplam maliyet hesaplanamaz, eksik olarak işaretlenir.
        missing.append(MISSING_SHIPPING)
        shipping_per_kg = None
        total_supply_cost = None
    elif shipping < 0:
        return None, "Nakliye tutarı negatif olamaz."
    else:
        shipping_per_kg = _quantize(shipping / total_weight, RATE_QUANTUM)
        total_supply_cost = _quantize(product_cost + shipping, MONEY_QUANTUM)

    return SupplierCalculation(
        row_number=quote.row_number,
        supplier_name=quote.supplier_name or "",
        supplier_quote_number=quote.supplier_quote_number,
        quantity_pcs=quantity,
        weight_kg_per_m=_quantize(weight_per_m, WEIGHT_QUANTUM),
        unit_price_try_per_kg=_quantize(unit_price, MONEY_QUANTUM),
        total_length_m=_quantize(total_length, WEIGHT_QUANTUM),
        total_weight_kg=_quantize(total_weight, WEIGHT_QUANTUM),
        product_cost_try=_quantize(product_cost, MONEY_QUANTUM),
        shipping_cost_try=(
            None if shipping is None else _quantize(shipping, MONEY_QUANTUM)
        ),
        shipping_cost_source=quote.shipping_cost_source,
        shipping_per_kg_try=shipping_per_kg,
        total_supply_cost_try=total_supply_cost,
        missing_data=tuple(missing),
    ), None


def _mark_lowest(
    rows: tuple[SupplierCalculation, ...], attribute: str
) -> CostMark | None:
    """Verilen alanda en düşük değeri işaretler; sıralama üretmez."""
    candidates = [
        (row, getattr(row, attribute))
        for row in rows
        if getattr(row, attribute) is not None
    ]
    if not candidates:
        return None

    lowest = min(amount for _row, amount in candidates)
    hits = [row for row, amount in candidates if amount == lowest]
    return CostMark(
        amount_try=lowest,
        row_numbers=tuple(sorted(row.row_number for row in hits)),
        suppliers=tuple(row.supplier_name for row in hits),
    )


def calculate(
    parsed: ParsedWorkbook, validation: ValidationResult
) -> CalculationResult:
    """Doğrulanmış teklifler için maliyet alanlarını üretir.

    Args:
        parsed: :func:`procurement.excel_parser.parse_workbook` çıktısı.
        validation: Aynı dosya için :func:`procurement.validator.validate`
            çıktısı. Hesaplamaya uygunluk **yalnızca** bu sonuca göre belirlenir;
            motor kendi doğrulama kuralını üretmez.

    Returns:
        Hesaplanmış satırlar (kaynak sırasında), hesaplamaya alınmayan satırlar
        ve en düşük ürün bedeli / en düşük toplam maliyet işaretleri.

    Not:
        Dosya düzeyinde hata varsa (``validation.file_rejected``) hiçbir satır
        hesaplanmaz: veri güvenilir okunamamıştır.
    """
    if validation.file_rejected:
        excluded = tuple(
            ExcludedQuote(
                row_number=quote.row_number,
                supplier_name=quote.supplier_name,
                reason="Dosya düzeyinde doğrulama hatası; dosya reddedildi.",
            )
            for quote in parsed.suppliers
        )
        return CalculationResult(
            rows=(),
            excluded=excluded,
            file_rejected=True,
            lowest_product_cost=None,
            lowest_total_cost=None,
        )

    rows: list[SupplierCalculation] = []
    excluded: list[ExcludedQuote] = []

    for quote in parsed.suppliers:
        codes = _blocking_codes(quote.row_number, validation)
        if codes:
            excluded.append(
                ExcludedQuote(
                    row_number=quote.row_number,
                    supplier_name=quote.supplier_name,
                    reason="Satır doğrulama hatası içeriyor; karşılaştırmaya alınmaz.",
                    codes=codes,
                )
            )
            continue

        calculation, reason = _calculate_row(quote)
        if calculation is None:
            excluded.append(
                ExcludedQuote(
                    row_number=quote.row_number,
                    supplier_name=quote.supplier_name,
                    reason=reason or "Girdi kullanılamıyor.",
                )
            )
            continue

        rows.append(calculation)

    computed = tuple(rows)
    return CalculationResult(
        rows=computed,
        excluded=tuple(excluded),
        file_rejected=False,
        lowest_product_cost=_mark_lowest(computed, "product_cost_try"),
        lowest_total_cost=_mark_lowest(computed, "total_supply_cost_try"),
    )

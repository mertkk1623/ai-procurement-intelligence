"""Sunum katmanı — Streamlit'ten bağımsız görünüm modelleri.

Sorumluluk
----------
Arayüz (:mod:`app`) yalnızca bu modülün ürettiği hazır görünümü çizer.
Biçimlendirme, işaretleme ve etiketleme mantığı burada durur; böylece
Streamlit çalıştırmadan test edilebilir (brief §8 mimari ilkesi).

Bu modül **hesaplamaz**. Ürün bedeli, toplam maliyet, en düşük işaretleri
ve hesaplamaya uygunluk kararı :mod:`procurement.calculator` çıktısından
okunur; burada yeniden üretilmez. Tedarikçi sıralaması, puanlama veya
kazanan belirleme yoktur — satırlar kaynak (Excel) sırasında kalır.

Sunum kuralları
---------------
* Tutarlar Türkçe biçimde gösterilir: binlik ayırıcı ``.``, ondalık ``,``.
* Eksik maliyet **sıfır gösterilmez**; :data:`MISSING_TEXT` yazılır.
* En düşük ürün bedeli ve en düşük toplam maliyet ayrı ayrı işaretlenir;
  eşitlikte tüm satırlar işaretlenir.
"""

from __future__ import annotations

import html
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from typing import Literal

import plotly.graph_objects as go

from .calculator import CalculationResult, SupplierCalculation
from .excel_parser import ParsedWorkbook
from .validator import Finding, Severity, ValidationResult

__all__ = [
    "CHART_COLOR_BASE",
    "CHART_COLOR_LOWEST",
    "LOWEST_BADGE",
    "MISSING_SHIPPING_NOTE",
    "MISSING_TEXT",
    "ComparisonRow",
    "ExcludedRowView",
    "FindingsView",
    "KpiCard",
    "build_comparison_rows",
    "build_cost_figure",
    "build_kpis",
    "comparison_table_html",
    "excluded_row_views",
    "finding_context",
    "format_days",
    "format_money",
    "format_quantity",
    "format_rate",
    "format_weight",
    "group_findings",
    "rejection_hints",
]

#: Eksik maliyet hücresinin metni — sıfır yerine kullanılır.
MISSING_TEXT: str = "Eksik veri"

#: Nakliye bilgisi olmayan satır için kısa etiket.
MISSING_SHIPPING_NOTE: str = "Nakliye bilgisi eksik"

#: En düşük değer rozeti.
LOWEST_BADGE: str = "En düşük"

#: Grafiklerde vurgu renkleri (en düşük değer yeşil).
CHART_COLOR_BASE: str = "#64748b"
CHART_COLOR_LOWEST: str = "#0b7a55"

Metric = Literal["product", "total"]

_METRIC_ATTR: dict[str, str] = {
    "product": "product_cost_try",
    "total": "total_supply_cost_try",
}

_METRIC_LABEL: dict[str, str] = {
    "product": "Ürün bedeli (TL)",
    "total": "Toplam tedarik maliyeti (TL)",
}


# --------------------------------------------------------------------------- #
# Sayı biçimlendirme
# --------------------------------------------------------------------------- #


def _format_number(value: Decimal, places: int) -> str:
    """Türkçe ondalık biçim: ``204345.5`` → ``"204.345,50"``.

    Yuvarlama yalnızca gösterim içindir; hesap değerleri motorda zaten
    kuantize edilmiştir. Buradaki yuvarlama arayüz tutarlılığı içindir.
    """
    quantum = Decimal(1).scaleb(-places)
    rounded = value.quantize(quantum, rounding=ROUND_HALF_UP)
    text = f"{rounded:,.{places}f}"
    return text.replace(",", "\0").replace(".", ",").replace("\0", ".")


def format_money(value: Decimal | None) -> str:
    """TRY tutarını 2 ondalıkla gösterir; ``None`` → "Eksik veri"."""
    if value is None:
        return MISSING_TEXT
    return _format_number(value, 2)


def format_weight(value: Decimal | None) -> str:
    """Toplam ağırlığı 3 ondalıkla gösterir."""
    if value is None:
        return MISSING_TEXT
    return _format_number(value, 3)


def format_rate(value: Decimal | None) -> str:
    """Kg başına nakliyeyi 4 ondalıkla gösterir."""
    if value is None:
        return MISSING_TEXT
    return _format_number(value, 4)


def format_quantity(value: int) -> str:
    """Adet değerini binlik ayırıcıyla gösterir."""
    return f"{value:,}".replace(",", ".")


def format_days(value: float | int | None) -> str:
    """Termin gününü gösterir; boşsa veya sayı değilse okunur metin döner."""
    if value is None:
        return MISSING_TEXT
    try:
        number = Decimal(str(value))
    except (InvalidOperation, ValueError):
        text = str(value).strip()
        return text or MISSING_TEXT
    if not number.is_finite():
        return str(value)
    if number != number.to_integral_value():
        return _format_number(number, 1)
    return str(int(number))


# --------------------------------------------------------------------------- #
# Karşılaştırma tablosu
# --------------------------------------------------------------------------- #


@dataclass(frozen=True, slots=True)
class ComparisonRow:
    """Tek bir tedarikçinin tablo satırı — tüm değerler gösterime hazır metin.

    ``is_lowest_product`` / ``is_lowest_total`` motorun işaretlerinden
    türetilir; burada yeniden hesaplama yapılmaz.
    """

    row_number: int
    supplier_name: str
    supplier_quote_number: str | None
    quantity: str
    weight: str
    unit_price: str
    product_cost: str
    shipping_cost: str
    total_cost: str
    shipping_per_kg: str
    delivery_days: str
    payment_terms: str
    is_lowest_product: bool
    is_lowest_total: bool
    is_complete: bool
    missing_note: str

    @property
    def is_lowest(self) -> bool:
        """En az bir maliyet alanında en düşük değere sahip mi."""
        return self.is_lowest_product or self.is_lowest_total


#: Tablo sütunları: (alan, başlık, css sınıfı).
_COLUMNS: tuple[tuple[str, str, str], ...] = (
    ("supplier", "Tedarikçi", "text"),
    ("quantity", "Miktar (adet)", "num"),
    ("weight", "Toplam ağırlık (kg)", "num"),
    ("unit_price", "Birim fiyat (TL/kg)", "num"),
    ("product", "Ürün bedeli (TL)", "num"),
    ("shipping", "Nakliye (TL)", "num"),
    ("total", "Toplam tedarik maliyeti (TL)", "num"),
    ("shipping_per_kg", "Nakliye (TL/kg)", "num"),
    ("delivery", "Termin (gün)", "num"),
    ("payment", "Ödeme koşulları", "text"),
)


def build_comparison_rows(
    result: CalculationResult, parsed: ParsedWorkbook
) -> tuple[ComparisonRow, ...]:
    """Hesaplanmış satırları tablo satırlarına dönüştürür (kaynak sırasında).

    Termin ve ödeme koşulu motorda hesaplanmaz; Excel'den okunan teklif
    satırından alınır. Ödeme koşulu **orijinal metin** olarak gösterilir.
    """
    quotes = {quote.row_number: quote for quote in parsed.suppliers}
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

    rows: list[ComparisonRow] = []
    for row in result.rows:
        quote = quotes.get(row.row_number)
        payment_terms = getattr(quote, "payment_terms", None)
        rows.append(
            ComparisonRow(
                row_number=row.row_number,
                supplier_name=row.supplier_name or f"Satır {row.row_number}",
                supplier_quote_number=row.supplier_quote_number,
                quantity=format_quantity(row.quantity_pcs),
                weight=format_weight(row.total_weight_kg),
                unit_price=format_money(row.unit_price_try_per_kg),
                product_cost=format_money(row.product_cost_try),
                shipping_cost=format_money(row.shipping_cost_try),
                total_cost=format_money(row.total_supply_cost_try),
                shipping_per_kg=format_rate(row.shipping_per_kg_try),
                delivery_days=format_days(getattr(quote, "delivery_days", None)),
                payment_terms=(
                    payment_terms.strip()
                    if isinstance(payment_terms, str) and payment_terms.strip()
                    else MISSING_TEXT
                ),
                is_lowest_product=row.row_number in lowest_product,
                is_lowest_total=row.row_number in lowest_total,
                is_complete=row.is_complete,
                missing_note=MISSING_SHIPPING_NOTE if row.missing_data else "",
            )
        )
    return tuple(rows)


def _cell(value: str, css: str, *, lowest: bool = False) -> str:
    """Tablo hücresi; en düşük değer rozetle ve vurgu sınıfıyla işaretlenir."""
    classes = [css]
    if lowest:
        classes.append("is-lowest")
    if value == MISSING_TEXT:
        classes.append("is-missing")
    badge = f'<span class="badge">{LOWEST_BADGE}</span>' if lowest else ""
    return f'<td class="{" ".join(classes)}">{html.escape(value)}{badge}</td>'


def _supplier_cell(row: ComparisonRow) -> str:
    name = html.escape(row.supplier_name)
    meta = [f"Satır {row.row_number}"]
    if row.supplier_quote_number:
        meta.insert(0, html.escape(row.supplier_quote_number))
    note = (
        f'<span class="tag tag-warn">{html.escape(row.missing_note)}</span>'
        if row.missing_note
        else ""
    )
    return (
        '<td class="text"><span class="supplier-cell">'
        f'<span class="supplier-name">{name}</span>'
        f'<span class="supplier-meta">{" · ".join(meta)}</span>'
        f"{note}</span></td>"
    )


def comparison_table_html(rows: tuple[ComparisonRow, ...]) -> str:
    """Karşılaştırma tablosunu HTML olarak üretir.

    Tüm metin alanları kaçışlanır (``html.escape``): yüklenen dosyadaki
    tedarikçi adı veya ödeme koşulu biçimlendirme olarak yorumlanmaz.
    """
    if not rows:
        return ""

    head = "".join(
        f'<th class="{css}">{html.escape(label)}</th>' for _key, label, css in _COLUMNS
    )
    body: list[str] = []
    for row in rows:
        body.append(
            "<tr>"
            + _supplier_cell(row)
            + _cell(row.quantity, "num")
            + _cell(row.weight, "num")
            + _cell(row.unit_price, "num")
            + _cell(row.product_cost, "num", lowest=row.is_lowest_product)
            + _cell(row.shipping_cost, "num")
            + _cell(row.total_cost, "num", lowest=row.is_lowest_total)
            + _cell(row.shipping_per_kg, "num")
            + _cell(row.delivery_days, "num")
            + _cell(row.payment_terms, "text")
            + "</tr>"
        )

    return (
        '<div class="table-card"><div class="table-scroll"><table class="cmp">'
        f"<thead><tr>{head}</tr></thead>"
        f"<tbody>{''.join(body)}</tbody>"
        "</table></div></div>"
    )


# --------------------------------------------------------------------------- #
# Özet kartları
# --------------------------------------------------------------------------- #


@dataclass(frozen=True, slots=True)
class KpiCard:
    """Genel özet kartı."""

    label: str
    value: str
    caption: str = ""
    emphasis: bool = False


def _mark_caption(suppliers: tuple[str, ...]) -> str:
    """En düşük değere ulaşan tedarikçi(ler)i tek satırda listeler.

    Tedarikçi adı dosyadan geldiği için kaçışlanır; kart HTML olarak
    çizildiğinde biçimlendirme olarak yorumlanmaz.
    """
    return " · ".join(html.escape(name) for name in suppliers)


def build_kpis(
    parsed: ParsedWorkbook, result: CalculationResult
) -> tuple[KpiCard, ...]:
    """Dört özet kartını üretir.

    En düşük maliyetler motordan okunur; hiçbir satır hesaplanamadıysa
    değer "—" olur ve kart nedeni açıklar.
    """
    supplier_count = len(parsed.suppliers)
    rfq_number = parsed.rfq.rfq_number
    excluded_count = len(result.excluded)

    if not supplier_count:
        computed_caption = "Dosyada teklif satırı yok"
    elif excluded_count:
        computed_caption = f"{excluded_count} satır karşılaştırmaya alınmadı"
    else:
        computed_caption = "Tüm satırlar karşılaştırmaya alındı"

    if result.lowest_product_cost is not None:
        product_value = f"{format_money(result.lowest_product_cost.amount_try)} TL"
        product_caption = _mark_caption(result.lowest_product_cost.suppliers)
    else:
        product_value = "—"
        product_caption = "Hesaplanabilen satır yok"

    if result.lowest_total_cost is not None:
        total_value = f"{format_money(result.lowest_total_cost.amount_try)} TL"
        total_caption = _mark_caption(result.lowest_total_cost.suppliers)
    else:
        total_value = "—"
        if result.missing_data_rows:
            total_caption = (
                "Nakliye bilgisi eksik olduğu için hesaplanamadı"
            )
        else:
            total_caption = "Hesaplanabilen satır yok"

    return (
        KpiCard(
            label="Tedarikçi sayısı",
            value=str(supplier_count),
            caption=(
                f"Talep {html.escape(str(rfq_number))}"
                if rfq_number
                else "Teklif dosyası"
            ),
        ),
        KpiCard(
            label="Hesaplanabilen teklif",
            value=str(len(result.rows)),
            caption=computed_caption,
        ),
        KpiCard(
            label="En düşük ürün bedeli",
            value=product_value,
            caption=product_caption,
            emphasis=True,
        ),
        KpiCard(
            label="En düşük toplam tedarik maliyeti",
            value=total_value,
            caption=total_caption,
            emphasis=True,
        ),
    )


# --------------------------------------------------------------------------- #
# Grafikler
# --------------------------------------------------------------------------- #


def _chart_rows(
    result: CalculationResult, metric: str
) -> list[tuple[SupplierCalculation, Decimal]]:
    """Grafiğe girecek satırlar: değeri ``None`` olanlar dahil edilmez."""
    attribute = _METRIC_ATTR[metric]
    rows: list[tuple[SupplierCalculation, Decimal]] = []
    for row in result.rows:
        value = getattr(row, attribute)
        if value is not None:
            rows.append((row, value))
    return rows


def build_cost_figure(result: CalculationResult, metric: Metric) -> go.Figure | None:
    """Maliyet karşılaştırma grafiği (yatay çubuk).

    Satırlar kaynak sırasında kalır — grafik bir sıralama üretmez. En düşük
    değer yeşil vurgulanır. Toplam maliyeti hesaplanamayan (nakliyesi eksik)
    satırlar bu grafiğe **dahil edilmez**.
    """
    rows = _chart_rows(result, metric)
    if not rows:
        return None

    lowest = min(value for _row, value in rows)
    labels = [row.supplier_name or f"Satır {row.row_number}" for row, _v in rows]
    values = [float(value) for _row, value in rows]
    colors = [
        CHART_COLOR_LOWEST if value == lowest else CHART_COLOR_BASE
        for _row, value in rows
    ]
    texts = [f"{format_money(value)} TL" for _row, value in rows]

    # Uzun tedarikçi adları kırpılmasın: sol kenar en uzun etikete göre açılır.
    left_margin = min(240, max(96, max(len(label) for label in labels) * 7))
    right_margin = max(96, max(len(text) for text in texts) * 7)

    figure = go.Figure(
        go.Bar(
            x=values,
            y=labels,
            orientation="h",
            marker={"color": colors, "line": {"width": 0}},
            text=texts,
            textposition="outside",
            cliponaxis=False,
            hovertemplate="%{y}<br>%{text}<extra></extra>",
        )
    )
    figure.update_layout(
        title=None,
        margin={"l": left_margin, "r": right_margin, "t": 6, "b": 6},
        height=max(190, 54 * len(rows) + 40),
        bargap=0.5,
        showlegend=False,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Inter, sans-serif", "size": 12, "color": "#475467"},
        xaxis={
            "showgrid": False,
            "zeroline": False,
            "showticklabels": False,
            "range": [0, max(values) * 1.34],
        },
        yaxis={
            "autorange": "reversed",
            "showgrid": False,
            "tickfont": {"size": 12, "color": "#101828"},
            "automargin": False,
        },
        hoverlabel={"font": {"family": "Inter, sans-serif", "size": 12}},
    )
    return figure


def figure_title(metric: Metric) -> str:
    """Grafik başlığı."""
    return _METRIC_LABEL[metric]


# --------------------------------------------------------------------------- #
# Doğrulama bulguları
# --------------------------------------------------------------------------- #


@dataclass(frozen=True, slots=True)
class ExcludedRowView:
    """Karşılaştırmaya alınmayan satırın gösterim bilgisi."""

    row_number: int
    supplier_name: str
    reason: str
    codes: tuple[str, ...]

    @property
    def codes_text(self) -> str:
        return ", ".join(self.codes)


@dataclass(frozen=True, slots=True)
class FindingsView:
    """Bulguların önem düzeyine göre ayrılmış hâli."""

    errors: tuple[Finding, ...]
    warnings: tuple[Finding, ...]
    infos: tuple[Finding, ...]

    @property
    def has_any(self) -> bool:
        return bool(self.errors or self.warnings or self.infos)


def group_findings(validation: ValidationResult) -> FindingsView:
    """Bulguları hata / uyarı / bilgi talebi olarak ayırır."""
    return FindingsView(
        errors=tuple(f for f in validation.findings if f.severity is Severity.ERROR),
        warnings=tuple(
            f for f in validation.findings if f.severity is Severity.WARNING
        ),
        infos=tuple(f for f in validation.findings if f.severity is Severity.INFO),
    )


def finding_context(finding: Finding) -> str:
    """Bulguyu tedarikçi ve Excel hücresiyle ilişkilendiren kısa metin."""
    parts: list[str] = []
    if finding.supplier:
        parts.append(finding.supplier)
    if finding.cell and finding.sheet:
        parts.append(f"{finding.sheet}!{finding.cell}")
    elif finding.cell:
        parts.append(finding.cell)
    elif finding.row_number is not None:
        parts.append(f"Satır {finding.row_number}")
    elif finding.sheet:
        parts.append(finding.sheet)
    return " · ".join(parts)


def excluded_row_views(result: CalculationResult) -> tuple[ExcludedRowView, ...]:
    """Hesaplamaya alınmayan satırları gösterim nesnelerine çevirir."""
    return tuple(
        ExcludedRowView(
            row_number=item.row_number,
            supplier_name=item.supplier_name or f"Satır {item.row_number}",
            reason=item.reason,
            codes=item.codes,
        )
        for item in result.excluded
    )


#: Dosya reddi nedenine göre kullanıcıya önerilen düzeltme.
_REJECTION_HINTS: dict[str, str] = {
    "E-DECL-01": (
        "Instructions sayfasındaki B4 (VAT Declaration) hücresine EVET yazın; "
        "KDV beyanı onaylanmadan dosya işleme alınmaz."
    ),
    "E-TPL-01": "Güncel teklif şablonunu kullanın; şablon kodunu değiştirmeyin.",
    "E-TPL-02": "Güncel şablon sürümünü kullanın; sürüm alanını değiştirmeyin.",
    "E-TPL-03": "Şablondan silinen sütun başlıklarını geri ekleyin.",
    "E-SHEET-01": (
        "Zorunlu sayfaları (Instructions, RFQ, Supplier Quotes) geri ekleyin."
    ),
    "E-HDR-01": (
        "Supplier Quotes sayfasının 2. satırındaki başlıkları şablonla "
        "birebir aynı hale getirin."
    ),
    "E-RFQ-01": "RFQ sayfasındaki zorunlu alanları doldurun.",
    "E-RFQ-02": "RFQ alanlarının etiketlerini şablondaki adlarla eşleştirin.",
    "E-SPEC-02": "Ürün kodunu şablondaki sabit değerle eşleştirin.",
    "E-SPEC-03": "Ürün kesitini 40x40x2 olarak bırakın.",
    "E-SPEC-04": "Et kalınlığını 2 mm olarak bırakın.",
    "E-LEN-02": "Profil uzunluğunu 6 m olarak bırakın.",
    "E-CUR-01": "Para birimini TRY olarak bırakın; çoklu para birimi desteklenmez.",
}


def rejection_hints(validation: ValidationResult) -> tuple[str, ...]:
    """Dosya reddinde geçerli olan düzeltme önerilerini döndürür."""
    codes = {finding.code for finding in validation.file_errors}
    return tuple(hint for code, hint in _REJECTION_HINTS.items() if code in codes)

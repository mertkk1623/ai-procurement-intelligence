"""DeepSeek destekli satın alma yorumu — yalnızca doğrulanmış sonuçları yorumlar.

Sorumluluk
----------
:func:`build_request` parser/validator/calculator çıktısından **maskelenmiş**
ve yalnızca hesaplanmış değerleri içeren bir istek gövdesi üretir.
:func:`analyze` bu gövdeyi OpenAI uyumlu Deploy/DeepSeek ucuna gönderir ve
yanıtı beş başlıklı yapıya çevirir.

Değişmez kurallar (``PROJECT_BRIEF.md`` §6–§7)
----------------------------------------------
* Hiçbir maliyet burada hesaplanmaz: tüm tutarlar
  :mod:`procurement.calculator` çıktısından **olduğu gibi** okunur.
* AI'a ham Excel içeriği gönderilmez; notlar, teklif numaraları ve gerçek
  tedarikçi adları gönderilmez (varsayılan maskeleme açık).
* AI yanıtındaki sayısal ifadeler kaynak verilerle karşılaştırılır. Kaynakta
  karşılığı olmayan bir tutar/oran/ağırlık/termin bulunursa analiz
  **gösterilmez** (:class:`AIAnalysisError`, ``kind="unverified"``).
* Dosya reddedilmişse veya hesaplanabilir satır yoksa analiz çalıştırılmaz.
* API anahtarı yalnızca ortamdan (``DEPLOY_API_KEY``) ya da ``.env``
  dosyasından okunur; koda yazılmaz. Anahtar, istek gövdesi ve yanıt
  loglanmaz.
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from typing import Any

from .calculator import CalculationResult
from .excel_parser import ParsedWorkbook, SupplierQuote
from .presentation import (
    format_days,
    format_money,
    format_quantity,
    format_quantity_value,
    format_rate,
    format_weight,
)
from .validator import ValidationResult, derive_payment_days

__all__ = [
    "API_KEY_ENV",
    "DEFAULT_BASE_URL",
    "DEFAULT_MODEL",
    "SECTION_TITLES",
    "AIAnalysis",
    "AIAnalysisError",
    "AiRequest",
    "AiSection",
    "FigureSet",
    "analysis_block_reason",
    "analyze",
    "api_key",
    "build_request",
    "unverified_figures",
]

#: API anahtarının okunduğu ortam değişkeni (brief §7).
API_KEY_ENV: str = "DEPLOY_API_KEY"

#: Deploy/DeepSeek uç noktası — ``opencode.json`` ile aynı.
DEFAULT_BASE_URL: str = "https://api-cdn.deploy.com/api/v1"
BASE_URL_ENV: str = "DEPLOY_BASE_URL"

#: Varsayılan model — ``opencode.json`` ile aynı.
DEFAULT_MODEL: str = "deepseek-ai/deepseek-v4.1-flash-1017"
MODEL_ENV: str = "DEPLOY_MODEL"

#: Ağ zaman aşımı (saniye). Model bir "reasoning" modeli olduğu için
#: ölçülü bir üst sınır tanınır; arayüz bu süre boyunca spinner gösterir.
DEFAULT_TIMEOUT: float = 90.0

#: Yanıttaki bölüm sırası ve başlıkları.
SECTION_TITLES: tuple[tuple[str, str], ...] = (
    ("executive_summary", "Executive Summary"),
    ("cost_analysis", "Cost Analysis"),
    ("delivery_payment", "Delivery & Payment Evaluation"),
    ("risks_missing", "Risks & Missing Information"),
    ("recommendations", "Procurement Recommendations"),
)

#: Yanıt doğrulamasında sabit yüzde tolerans **kullanılmaz**: iddia, kaynak
#: değerin iddianın kendi ondalık hassasiyetinde yuvarlanmış hâliyle
#: karşılaştırılır (``_claim_matches``). Böylece "211.845" doğru değeri
#: kabul edilir, "226.100" ise 225.046,00 için reddedilir.


# --------------------------------------------------------------------------- #
# Hata modeli
# --------------------------------------------------------------------------- #


class AIAnalysisError(Exception):
    """AI analizi üretilemediğinde fırlatılır.

    Attributes:
        kind: Makine tarafından okunabilir neden. Arayüz buna göre mesaj
            gösterir; ayrıntılı API yanıtı kullanıcıya taşınmaz.
        details: ``unverified`` türünde, kaynakla eşleşmeyen ifadeler.
    """

    def __init__(
        self,
        message: str,
        *,
        kind: str = "service",
        details: tuple[str, ...] = (),
    ) -> None:
        super().__init__(message)
        self.kind = kind
        self.details = details


# --------------------------------------------------------------------------- #
# İstek modeli
# --------------------------------------------------------------------------- #


@dataclass(frozen=True, slots=True)
class SupplierFigures:
    """Tek tedarikçinin kaynak sayıları — metrik bazında ayrık.

    Değerler bilinçli olarak tek havuzda toplanmaz: bir tedarikçinin
    nakliye tutarının başka bir tedarikçiye veya vade gününün teslimat
    terminine atfedilmesi bu ayrım sayesinde yakalanır (brief §7).

    ``shipping`` / ``total_cost`` ``None`` ise veri **yoktur**; sıfır
    değildir. ``payment_days`` teslimat termininden ayrı tutulur.
    """

    label: str
    product_cost: Decimal | None
    unit_price: Decimal | None
    weight_per_m: Decimal | None
    total_weight: Decimal | None
    shipping: Decimal | None
    shipping_rate: Decimal | None
    total_cost: Decimal | None
    delivery_days: Decimal | None
    payment_days: Decimal | None
    payment_percents: tuple[Decimal, ...]


@dataclass(frozen=True, slots=True)
class FigureSet:
    """Yanıt doğrulamasında kabul edilen kaynak sayılar.

    Sayılar tedarikçi bazında tutulur (:class:`SupplierFigures`). Serbest
    metinlerde geçen referans değerler (ör. doğrulama mesajındaki ağırlık
    aralığı ``2.27–2.51``) yalnızca kendi metriğinde ve yalnızca metin
    gerçekten o referansa atıf yaptığında kabul edilir.
    """

    suppliers: tuple[SupplierFigures, ...]
    rfq_days: Decimal | None = None
    lowest_product_labels: tuple[str, ...] = ()
    lowest_product_amount: Decimal | None = None
    lowest_total_labels: tuple[str, ...] = ()
    lowest_total_amount: Decimal | None = None
    reference_weights: tuple[Decimal, ...] = ()

    def by_label(self, label: str) -> SupplierFigures | None:
        for supplier in self.suppliers:
            if supplier.label == label:
                return supplier
        return None


@dataclass(frozen=True, slots=True)
class AiRequest:
    """AI'a gönderilecek maskelenmiş gövde ve doğrulama referansları.

    Attributes:
        payload: JSON'a çevrilebilir, yalnızca hesaplanmış alanları taşıyan
            sözlük. Ham Excel verisi, notlar ve gerçek tedarikçi adları yoktur.
        legend: ``(maskeli etiket, gerçek tedarikçi adı)`` çiftleri; yalnızca
            arayüzde gösterilir, API'a gönderilmez.
        figures: Yanıttaki sayısal ifadelerin karşılaştırılacağı kaynak küme.
    """

    payload: dict[str, Any]
    legend: tuple[tuple[str, str], ...]
    figures: FigureSet


@dataclass(frozen=True, slots=True)
class AiSection:
    """AI yorumunun tek bölümü."""

    key: str
    title: str
    body: str


@dataclass(frozen=True, slots=True)
class AIAnalysis:
    """Doğrulanmış AI yorumu."""

    sections: tuple[AiSection, ...]
    model: str
    legend: tuple[tuple[str, str], ...]


# --------------------------------------------------------------------------- #
# Yapılandırma
# --------------------------------------------------------------------------- #


def _load_dotenv() -> None:
    """``.env`` varsa okur; ortam değişkenini ezmez."""
    try:
        from dotenv import load_dotenv
    except ImportError:  # pragma: no cover - bağımlılık requirements'ta
        return
    load_dotenv(override=False)


def api_key() -> str | None:
    """Ortamdaki (veya ``.env`` içindeki) API anahtarını döndürür.

    Anahtar asla loglanmaz ve istisna mesajına eklenmez.
    """
    _load_dotenv()
    value = os.environ.get(API_KEY_ENV, "").strip()
    return value or None


def _base_url() -> str:
    return os.environ.get(BASE_URL_ENV, "").strip() or DEFAULT_BASE_URL


def _model_name() -> str:
    return os.environ.get(MODEL_ENV, "").strip() or DEFAULT_MODEL


def _create_client(key: str, timeout: float):
    """OpenAI uyumlu istemciyi üretir (yalnızca gerçek çağrıda)."""
    from openai import OpenAI

    return OpenAI(
        api_key=key,
        base_url=_base_url(),
        timeout=timeout,
        max_retries=1,
    )


# --------------------------------------------------------------------------- #
# Ön koşullar
# --------------------------------------------------------------------------- #


def analysis_block_reason(
    validation: ValidationResult, result: CalculationResult
) -> str | None:
    """AI analizi çalıştırılamıyorsa gerekçesini döndürür.

    Reddedilen dosyada ve hesaplanabilir satır yokken analiz yapılmaz;
    eksik nakliye analizi **engellemez**, yalnızca veride işaretlenir.
    """
    if validation.file_rejected:
        return "Dosya reddedildiği için AI analizi çalıştırılmaz."
    if not result.rows:
        return "Karşılaştırılabilir teklif satırı olmadığı için AI analizi çalıştırılmaz."
    return None


# --------------------------------------------------------------------------- #
# İstek gövdesi
# --------------------------------------------------------------------------- #


def _masked_labels(count: int) -> tuple[str, ...]:
    return tuple(f"Tedarikçi {index}" for index in range(1, count + 1))


def _mask_text(text: str, mapping: dict[str, str]) -> str:
    """Serbest metindeki gerçek adları maskeli etiketle değiştirir.

    Adlar **en uzun önce** ve büyük/küçük harf duyarsız değiştirilir;
    aksi hâlde ``Bursa Metal`` girdisi ``Bursa Metal A.Ş.`` adının bir
    parçasını maskeleyip şirket ekini ve grup ilişkisini sızdırırdı.
    """
    masked = text
    for real in sorted(mapping, key=len, reverse=True):
        label = mapping[real]
        if not real or real.casefold() not in masked.casefold():
            continue
        masked = re.sub(
            re.escape(real), label.replace("\\", r"\\"), masked, flags=re.IGNORECASE
        )
    return masked


_NUMBER = re.compile(r"\d{1,3}(?:\.\d{3})+(?:,\d+)?|\d+(?:[.,]\d+)?")


def _shipping_source(value: str | None) -> str | None:
    """Teklif kaynağını kullanıcıya anlaşılır biçimde etiketler."""
    if not value:
        return None
    normalized = value.strip().lower()
    if "assumption" in normalized or "varsay" in normalized:
        return "Kullanıcı varsayımı"
    if "quote" in normalized or "teklif" in normalized:
        return "Tedarikçi teklifi"
    return value.strip()


def _days_value(value: object) -> Decimal | None:
    if value is None or isinstance(value, bool):
        return None
    try:
        return Decimal(str(value).strip().replace(",", "."))
    except (InvalidOperation, ValueError):
        return None


def _days_text(value: object) -> str | None:
    """Termin değerini metne çevirir; veri yoksa ``None`` döndürür.

    Gönderimde ``"Eksik veri"`` metni yerine açık ``null`` kullanılır;
    AI'ın bunu bir gün değeri sanması engellenir.
    """
    if _days_value(value) is None:
        return None
    return format_days(value)


_MISSING_LABELS: dict[str, str] = {
    "shipping": "Nakliye bilgisi eksik",
}


def _supplier_view(
    row: Any,
    quote: SupplierQuote | None,
    label: str,
) -> dict[str, Any]:
    """Tek tedarikçinin AI'a gönderilecek maskelenmiş görünümü."""
    shipping = row.shipping_cost_try
    total = row.total_supply_cost_try
    view: dict[str, Any] = {
        "tedarikci": label,
        "miktar": f"{format_quantity(row.quantity_pcs)} adet",
        "kg_m": f"{format_weight(row.weight_kg_per_m)} kg/m",
        "birim_fiyat": f"{format_money(row.unit_price_try_per_kg)} TL/kg",
        "toplam_agirlik": f"{format_weight(row.total_weight_kg)} kg",
        "urun_bedeli": f"{format_money(row.product_cost_try)} TL",
        "nakliye": None if shipping is None else f"{format_money(shipping)} TL",
        "nakliye_kaynagi": _shipping_source(row.shipping_cost_source),
        "nakliye_tl_kg": (
            None
            if row.shipping_per_kg_try is None
            else f"{format_rate(row.shipping_per_kg_try)} TL/kg"
        ),
        "toplam_tedarik_maliyeti": (
            None if total is None else f"{format_money(total)} TL"
        ),
        "termin": (
            _days_text(getattr(quote, "delivery_days", None))
            if quote is not None
            else None
        ),
        "odeme_kosullari": (
            quote.payment_terms.strip()
            if quote is not None
            and isinstance(quote.payment_terms, str)
            and quote.payment_terms.strip()
            else "Belirtilmemiş"
        ),
        "vade_gunu": (
            derive_payment_days(getattr(quote, "payment_terms", None))
            if quote is not None
            else None
        ),
        "eksik_veri": [
            _MISSING_LABELS.get(code, code) for code in row.missing_data
        ],
    }
    return view


def _mark_view(mark, labels: dict[int, str]) -> dict[str, Any] | None:
    if mark is None:
        return None
    return {
        "tutar": f"{format_money(mark.amount_try)} TL",
        "tedarikciler": [labels[row] for row in mark.row_numbers if row in labels],
        "esitlik": len(mark.row_numbers) > 1,
    }


def _figures(
    rows: tuple[Any, ...],
    parsed: ParsedWorkbook,
    result: CalculationResult,
    *,
    labels: dict[int, str],
    reference_weights: tuple[Decimal, ...] = (),
) -> FigureSet:
    """Yanıt doğrulaması için tedarikçi bazlı kaynak sayıları toplar.

    Her metrik kendi tedarikçisinin kaydında tutulur; global bir sayı
    havuzu oluşturulmaz. ``lowest_*`` işaretleri doğrudan
    :class:`~procurement.calculator.CostMark` nesnelerinden okunur.
    """
    quotes = {quote.row_number: quote for quote in parsed.suppliers}
    suppliers: list[SupplierFigures] = []

    for row in rows:
        quote = quotes.get(row.row_number)
        payment_terms = getattr(quote, "payment_terms", None)
        percent_text = payment_terms if isinstance(payment_terms, str) else ""
        suppliers.append(
            SupplierFigures(
                label=labels[row.row_number],
                product_cost=row.product_cost_try,
                unit_price=row.unit_price_try_per_kg,
                weight_per_m=row.weight_kg_per_m,
                total_weight=row.total_weight_kg,
                shipping=row.shipping_cost_try,
                shipping_rate=row.shipping_per_kg_try,
                total_cost=row.total_supply_cost_try,
                delivery_days=_days_value(getattr(quote, "delivery_days", None)),
                payment_days=(
                    Decimal(days)
                    if (days := derive_payment_days(payment_terms)) is not None
                    else None
                ),
                payment_percents=_percent_values(percent_text),
            )
        )

    def mark(mark_obj) -> tuple[tuple[str, ...], Decimal | None]:
        if mark_obj is None:
            return (), None
        return (
            tuple(labels[row] for row in mark_obj.row_numbers if row in labels),
            mark_obj.amount_try,
        )

    low_product_labels, low_product_amount = mark(result.lowest_product_cost)
    low_total_labels, low_total_amount = mark(result.lowest_total_cost)

    return FigureSet(
        suppliers=tuple(suppliers),
        rfq_days=_days_value(parsed.rfq.requested_delivery_days),
        lowest_product_labels=low_product_labels,
        lowest_product_amount=low_product_amount,
        lowest_total_labels=low_total_labels,
        lowest_total_amount=low_total_amount,
        reference_weights=reference_weights,
    )


def build_request(
    parsed: ParsedWorkbook,
    validation: ValidationResult,
    result: CalculationResult,
) -> AiRequest:
    """Maskelenmiş AI isteğini üretir.

    Yalnızca hesaplanmış ve doğrulanmış alanlar gönderilir. Tedarikçi adları
    ``Tedarikçi N`` etiketleriyle değiştirilir; teklif numaraları, notlar ve
    dosya adı gönderilmez (brief §7).
    """
    quotes = {quote.row_number: quote for quote in parsed.suppliers}
    labels = _masked_labels(len(result.rows))
    label_by_row = {
        row.row_number: label for row, label in zip(result.rows, labels)
    }
    # Karşılaştırmaya alınmayan satırlar da maskelenir: doğrulama bulguları
    # ve dışlama gerekçeleri gerçek adı taşıyabilir (brief §7).
    excluded_by_row = {
        excluded.row_number: f"Karşılaştırma dışı {index}"
        for index, excluded in enumerate(result.excluded, start=1)
    }
    label_by_row.update(excluded_by_row)

    name_mapping = {
        row.supplier_name: label_by_row[row.row_number]
        for row in result.rows
        if row.supplier_name
    }
    name_mapping.update(
        {
            excluded.supplier_name: excluded_by_row[excluded.row_number]
            for excluded in result.excluded
            if excluded.supplier_name
        }
    )
    # Eşleşmeyen kalan adlar (ör. yalnızca dosya düzeyinde geçen bir ad)
    # kimlik sızdırmasın diye nötr etikete çevrilir.
    for quote in parsed.suppliers:
        name = quote.supplier_name
        if name and name not in name_mapping:
            name_mapping[name] = "Tedarikçi (belirtilmemiş)"

    def masked_name(value: str | None) -> str | None:
        if not value:
            return None
        return name_mapping.get(value) or _mask_text(value, name_mapping)

    suppliers = [
        _supplier_view(row, quotes.get(row.row_number), label)
        for row, label in zip(result.rows, labels)
    ]

    rfq = parsed.rfq
    payload: dict[str, Any] = {
        "talep": {
            "urun": "Galvanizli kare kutu profil 40 × 40 × 2 mm, 6 m",
            "miktar": (
                f"{format_quantity_value(rfq.quantity_pcs)} adet"
                if format_quantity_value(rfq.quantity_pcs)
                else None
            ),
            "teslim_yeri": rfq.delivery_location,
            "istenen_termin": _days_text(rfq.requested_delivery_days),
            "para_birimi": "TRY (KDV hariç)",
        },
        "tedarikciler": suppliers,
        "en_dusuk": {
            "urun_bedeli": _mark_view(result.lowest_product_cost, label_by_row),
            "toplam_tedarik_maliyeti": _mark_view(
                result.lowest_total_cost, label_by_row
            ),
        },
        "dogrulama_bulgulari": [
            {
                "kod": finding.code,
                "seviye": finding.severity.value,
                "tedarikci": masked_name(finding.supplier),
                "mesaj": _mask_text(finding.message, name_mapping),
            }
            for finding in validation.findings
        ],
        "karsilastirma_disi": [
            {
                "tedarikci": masked_name(excluded.supplier_name)
                or "Bilinmeyen tedarikçi",
                "neden": _mask_text(excluded.reason, name_mapping),
                "kodlar": list(excluded.codes),
            }
            for excluded in result.excluded
        ],
    }

    legend = tuple(
        (label_by_row[row.row_number], row.supplier_name or "Belirtilmemiş")
        for row in result.rows
    )
    # Yalnızca doğrulama mesajlarındaki **referans aralığı** ayrıca kaynak
    # sayılır (ör. "referans 2.27–2.51"); başka serbest metin sayıları
    # ilgisiz metriklerde kullanılmaz.
    reference_weights: list[Decimal] = []
    for finding in validation.findings:
        text = _mask_text(finding.message, name_mapping)
        match = _REFERENCE_CONTEXT.search(text)
        if match is None:
            continue
        for number in _NUMBER.finditer(text[match.end() :]):
            value = _to_decimal(number.group(0))
            if value is not None:
                reference_weights.append(value)

    return AiRequest(
        payload=payload,
        legend=legend,
        figures=_figures(
            result.rows,
            parsed,
            result,
            labels=label_by_row,
            reference_weights=tuple(dict.fromkeys(reference_weights)),
        ),
    )


# --------------------------------------------------------------------------- #
# İstem
# --------------------------------------------------------------------------- #

_SYSTEM_PROMPT = """\
Sen deneyimli bir satın alma analistisin. Sana yalnızca Python tarafından \
doğrulanmış ve hesaplanmış teklif özeti JSON olarak verilir. Görevin bu veriyi \
Türkçe, profesyonel ve ölçülü bir dille yorumlamak.

Kesin kurallar:
1. Yalnızca verilen JSON içindeki sayıları kullan. Yeni fiyat, ağırlık, \
termin, tutar veya oran üretme.
2. Aritmetik yapma. Fark, yüzde, ortalama, toplam ya da birim maliyet \
hesaplama; karşılaştırmayı "daha yüksek", "daha düşük", "yakın" gibi nitel \
ifadelerle yap.
3. Tutarları, oranları ve ağırlıkları verilen biçimde aynen yaz \
(örnek: 204.345,00 TL, 28,50 TL/kg, 2,390 kg/m, 15 gün).
4. Tedarikçi sıralaması, puanlama, kazanan ilan etme veya "en iyi tedarikçi" \
önerisi verme. Seçim kullanıcıya aittir.
5. "eksik_veri" listesi dolu olan satırlarda nakliye ve toplam tedarik \
maliyeti bilinmiyor demektir; bunları sıfır kabul etme ve hesaplama.
6. Veride olmayan bilgi uydurma; bilgi yoksa açıkça "veri yok" de.
7. `dogrulama_bulgulari` ve `karsilastirma_disi` alanlarını risk ve eksik \
bilgi değerlendirmesinde kullan.
8. Yalnızca şu JSON nesnesini döndür; her değer düz metin olsun (en fazla 90 \
kelime): {"executive_summary": "", "cost_analysis": "", "delivery_payment": "", \
"risks_missing": "", "recommendations": ""}

Bölümlerin kapsamı:
- executive_summary: genel satın alma değerlendirmesi.
- cost_analysis: ürün bedeli ve toplam tedarik maliyeti farkları, nakliyenin \
maliyete etkisi.
- delivery_payment: teslimat süreleri ve ödeme koşullarının değerlendirmesi.
- risks_missing: eksik veri, belirsizlik ve doğrulama bulguları.
- recommendations: satın alma kararında dikkat edilmesi gereken noktalar; \
kazanan seçmeden."""


# --------------------------------------------------------------------------- #
# Yanıt ayrıştırma ve doğrulama
# --------------------------------------------------------------------------- #

_FENCE = re.compile(r"^\s*```(?:json)?\s*|\s*```\s*$", re.IGNORECASE)

_MONEY_CLAIM = re.compile(
    r"(?<![\d.,])(\d{1,3}(?:\.\d{3})+(?:,\d+)?|\d+,\d+|\d+)(?![\d.,])"
    r"\s*(?:TL|TRY|₺)(?!\s*/\s*kg)",
    re.IGNORECASE,
)
_RATE_CLAIM = re.compile(
    r"(?<![\d.,])(\d{1,3}(?:\.\d{3})+(?:,\d+)?|\d+,\d+|\d+|"
    r"\d{1,3}(?:\.\d{3})*\.\d+)"
    r"\s*TL\s*/\s*kg",
    re.IGNORECASE,
)
_WEIGHT_CLAIM = re.compile(
    r"(?<![\d.,])(\d{1,3}(?:\.\d{3})+(?:,\d+)?|\d+,\d+|\d+|"
    r"\d{1,3}(?:\.\d{3})*\.\d+)"
    r"\s*kg\s*/\s*m",
    re.IGNORECASE,
)
_KG_CLAIM = re.compile(
    r"(?<![\d.,])(\d{1,3}(?:\.\d{3})+(?:,\d+)?|\d+,\d+|\d+)"
    r"\s*kg(?!\s*/\s*m)",
    re.IGNORECASE,
)
_DAYS_CLAIM = re.compile(r"(?<![\d.,])(\d{1,4})\s*(?:iş\s*)?gün", re.IGNORECASE)
# Türkçe gösterimde işaret sayıdan önce de sonra da yazılabilir: %50 / 50 %.
_PERCENT_BEFORE = re.compile(
    r"%\s*(\d{1,3}(?:\.\d{3})*(?:,\d+)?)", re.IGNORECASE
)
_PERCENT_AFTER = re.compile(
    r"(?<![\w.,])(\d{1,3}(?:\.\d{3})*(?:,\d+)?)\s*%", re.IGNORECASE
)
# İngilizce ondalık/binlik biçimi (``999,999.00``) Türkçe okunamaz; sessizce
# yanlış yorumlamak yerine güvenli tarafta kalıp reddedilir.
_ENGLISH_MONEY = re.compile(
    r"(?<![\d.,])\d{1,3}(?:,\d{3})+(?:\.\d+)?\s*(?:TL|TRY|₺)", re.IGNORECASE
)


def _to_decimal(raw: str) -> Decimal | None:
    """Türkçe sayı gösterimini Decimal'e çevirir.

    Türkçe biçimde ``.`` binlik, ``,`` ondalıktır: ``211.845,00`` → 211845.00.
    Nokta yalnızca üçlü gruplar hâlinde binlik sayılır; ``2.390`` bu yüzden
    2390 olur ve ``2,390`` (2,39) ile karışmaz.
    """
    text = raw.strip()
    if not text:
        return None
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(?:\.\d{3})+", text):
        text = text.replace(".", "")
    try:
        return Decimal(text)
    except (InvalidOperation, ValueError):
        return None


def _decimal_places(raw: str) -> int:
    """İfadedeki ondalık hane sayısı (Türkçe biçime göre)."""
    text = raw.strip()
    if "," in text:
        return len(text.rsplit(",", 1)[1])
    return 0


def _claim_matches(claim: Decimal, raw: str, value: Decimal | None) -> bool:
    """İddia, kaynak değerin iddia hassasiyetinde yuvarlanmış hâli mi?

    Sabit bir yüzde tolerans yerine iddianın kendi hassasiyeti kullanılır:
    ``211.845`` ile ``211845.00`` eşleşir, ``226.100`` ile ``225046.00``
    eşleşmez.
    """
    if value is None:
        return False
    quantum = Decimal(1).scaleb(-_decimal_places(raw))
    return value.quantize(quantum, rounding=ROUND_HALF_UP) == claim


def _percent_values(text: str) -> tuple[Decimal, ...]:
    """Metindeki yüzdeleri (``%50`` ve ``50 %``) sayı olarak döndürür."""
    found: dict[Decimal, None] = {}
    for pattern in (_PERCENT_BEFORE, _PERCENT_AFTER):
        for match in pattern.finditer(text):
            value = _to_decimal(match.group(1))
            if value is not None:
                found[value] = None
    return tuple(found)


def _segment_scope(segment: str, labels: tuple[str, ...]) -> tuple[str, ...]:
    """Metin parçasında anılan tedarikçi etiketleri.

    Model etiketleri sık sık ``T1``/``T2`` diye kısalttığı için bu yazım da
    tanınır; aksi hâlde atıf denetimi sessizce tüm tedarikçilere düşerdi.
    """
    found: dict[str, None] = {}
    for label in labels:
        if re.search(rf"{re.escape(label)}\b", segment, re.IGNORECASE):
            found[label] = None
            continue
        index = label.rsplit(" ", 1)[-1]
        if index.isdigit() and re.search(
            rf"\bT{index}\b", segment, re.IGNORECASE
        ):
            found[label] = None
    return tuple(found)


def _distance(label: str, segment: str, position: int) -> int:
    """Etiketin ``position`` noktasına en yakın geçişinin uzaklığı."""
    best: int | None = None
    patterns = [re.escape(label)]
    index = label.rsplit(" ", 1)[-1]
    if index.isdigit():
        patterns.append(rf"T{index}\b")
    for pattern in patterns:
        for match in re.finditer(pattern, segment, re.IGNORECASE):
            gap = abs(match.start() - position)
            if best is None or gap < best:
                best = gap
    return best if best is not None else len(segment)


def _field_values(
    suppliers: tuple[SupplierFigures, ...], *fields: str
) -> tuple[Decimal, ...]:
    values: dict[Decimal, None] = {}
    for supplier in suppliers:
        for field in fields:
            value = getattr(supplier, field)
            if isinstance(value, Decimal):
                values[value] = None
    return tuple(values)


def _percent_scope(
    suppliers: tuple[SupplierFigures, ...],
) -> tuple[Decimal, ...]:
    """Tedariçilerin ödeme koşullarındaki yüzdeler (yalnızca bu metrikte)."""
    values: dict[Decimal, None] = {}
    for supplier in suppliers:
        for value in supplier.payment_percents:
            values[value] = None
    return tuple(values)


_ZERO_SHIPPING = re.compile(
    r"nakliye\w*\s*(?:tutarı\s*)?(?:sıfır|yok\s+say|yoktur|\b0(?:[.,]0+)?\b)"
    r"|sıfır\s+nakliye"
    r"|nakliyesiz",
    re.IGNORECASE,
)
_THOUSANDS_CLAIM = re.compile(
    r"(?<![\d.,])(\d{1,3})\s*bin\s*(?:TL|TRY|₺)", re.IGNORECASE
)
_REFERENCE_CONTEXT = re.compile(r"referans|aral[ıi]ğ", re.IGNORECASE)
# Gün iddiasının yerel bağlamı: istenen termin ≠ ödeme vadesi.
# Sayıdan hemen önce gelen bağlam (ör. "istenen termin ", "vade ").
# Türkçe ``İ`` (U+0130) Python'un harf duyarsız eşleşmesine girmediği için
# açıkça sınıf olarak yazılır.
_REQUEST_WORDS_END = re.compile(
    r"(?:[iİıI]stenen|talep|rfq)\b[^,;.!?]{0,25}$", re.IGNORECASE
)
_PAYMENT_WORDS = re.compile(r"vade|[öÖoO]deme", re.IGNORECASE)
_DELIVERY_WORDS = re.compile(r"termin|teslim", re.IGNORECASE)
_LOWEST = re.compile(r"en\s+d[üu]ş[üu]k|en\s+iyi|en\s+uygun", re.IGNORECASE)
# "en düşük değil", "en düşük olmayan" gibi olumsuz ifadeler atıf değildir.
_NEGATED_LOWEST = re.compile(
    r"en\s+d[üu]ş[üu]k\w*\s*[^.;]{0,30}?değil"
    r"|(?:değil|olmayan)\w*\s*[^.;]{0,20}en\s+d[üu]ş[üu]k",
    re.IGNORECASE,
)
_SENTENCE_SPLIT = re.compile(r"(?<=[.!?;])\s+|\n+")
# ``W-WT-01`` gibi bulgu kodlarının sonundaki sayı iddia değildir.
_CODE_PREFIX = re.compile(r"[A-Za-z]-+$")


def _inside_code(segment: str, start: int) -> bool:
    """Eşleşme bir bulgu kodunun parçasıysa (``W-WT-01``) ``True``."""
    return bool(_CODE_PREFIX.search(segment[:start]))


def _check_figures(
    segment: str,
    suppliers: tuple[SupplierFigures, ...],
    figures: FigureSet,
    *,
    attributed: bool,
) -> tuple[tuple[str, ...], tuple[str, ...]]:
    """Bir metin parçasındaki çelişkileri döndürür: (sayı, atıf notu).

    ``attributed`` yalnızca parçada tedarikçi adı **açıkça** geçtiğinde
    doğrudur. Atıf gerektiren kurallar (eksik nakliye, "en düşük") ancak o
    zaman uygulanır; aksi hâlde tedarikçi belirtmeyen genel bir cümle
    yanlış alarm üretirdi.
    """
    bad: dict[str, None] = {}
    notes: dict[str, None] = {}

    def check(pattern, allowed: tuple[Decimal, ...], suffix: str) -> None:
        for match in pattern.finditer(segment):
            raw = match.group(1)
            # ``W-WT-01`` gibi bulgu kodlarının parçası sayı iddiası değildir.
            if _inside_code(segment, match.start(1)):
                continue
            claim = _to_decimal(raw)
            if claim is None:
                continue
            if any(_claim_matches(claim, raw, value) for value in allowed):
                continue
            bad[f"{raw.strip()} {suffix}"] = None

    money_fields = ("product_cost", "shipping", "total_cost")
    check(_MONEY_CLAIM, _field_values(suppliers, *money_fields), "TL")
    check(
        _RATE_CLAIM,
        _field_values(suppliers, "unit_price", "shipping_rate"),
        "TL/kg",
    )

    weight_scope = _field_values(suppliers, "weight_per_m")
    if _REFERENCE_CONTEXT.search(segment):
        weight_scope += figures.reference_weights
    check(_WEIGHT_CLAIM, weight_scope, "kg/m")
    check(_KG_CLAIM, _field_values(suppliers, "total_weight"), "kg")

    # Gün değerleri, sayının **kendi yerel bağlamına** göre ayrılır:
    # istenen termin (RFQ) ≠ tedarikçi teslim termini ≠ ödeme vadesi.
    previous_end = 0
    for match in _DAYS_CLAIM.finditer(segment):
        if _inside_code(segment, match.start(1)):
            continue
        raw = match.group(1)
        claim = _to_decimal(raw)
        if claim is None:
            continue
        # Bağlam yalnızca sayıdan **önceki** kısa pencereden okunur ve bir
        # önceki gün sayısında kesilir. İleriye dönük kelimeler (ör. "25 gün
        # ile istenen sürenin üzerinde") sayının metriğini değil,
        # karşılaştırmayı anlatır.
        start = match.start(1)
        back = segment[max(previous_end, start - 40) : start]
        forward = segment[match.end(1) : match.end(1) + 20]
        previous_end = match.end(1)
        if _PAYMENT_WORDS.search(back) or re.match(
            r"\s*(?:vade\w*|ödeme\w*)", forward, re.IGNORECASE
        ):
            # Vade ile teslimat termini aynı cümlede geçebilir; ikisi de
            # kaynak sayılır, aksi hâlde doğru değer reddedilirdi.
            day_scope = _field_values(suppliers, "payment_days")
            if _DELIVERY_WORDS.search(segment):
                day_scope += _field_values(suppliers, "delivery_days")
        elif _REQUEST_WORDS_END.search(back):
            # "İstenen (termin) 20 gün" bir tedarikçi termini değil, RFQ değeridir.
            day_scope = (figures.rfq_days,) if figures.rfq_days is not None else ()
        else:
            # Tedarikçi termini; "20 gün termin" istenen termini de
            # kastedebilir, bu yüzden RFQ değeri de kabul edilir.
            day_scope = _field_values(suppliers, "delivery_days")
            if figures.rfq_days is not None:
                day_scope += (figures.rfq_days,)
        if any(_claim_matches(claim, raw, value) for value in day_scope):
            continue
        bad[f"{raw.strip()} gün"] = None

    percent_scope = _percent_scope(suppliers)
    label_before = re.compile(r"(?:tedarikçi|t)\s*$", re.IGNORECASE)
    for pattern in (_PERCENT_BEFORE, _PERCENT_AFTER):
        for match in pattern.finditer(segment):
            raw = match.group(1)
            # "Tedarikçi 2 %50" ifadesindeki 2 bir yüzde değil, etiket
            # numarasıdır; aksi hâlde uydurma "%2" alarmı oluşur.
            if label_before.search(segment[: match.start(1)]):
                continue
            claim = _to_decimal(raw)
            if claim is None:
                continue
            if any(_claim_matches(claim, raw, value) for value in percent_scope):
                continue
            bad[f"%{raw.strip()}"] = None

    # Türkçe biçimde okunamayan İngilizce tutarlar güvenli tarafta reddedilir.
    for match in _ENGLISH_MONEY.finditer(segment):
        bad[match.group(0).strip()] = None

    # "212 bin TL" gibi yuvarlanmış ifadeler bin hassasiyetinde kabul edilir.
    for match in _THOUSANDS_CLAIM.finditer(segment):
        claim = Decimal(match.group(1)) * 1000
        allowed = _field_values(suppliers, *money_fields)
        if not any(
            value.quantize(Decimal("1E3"), rounding=ROUND_HALF_UP) == claim
            for value in allowed
        ):
            bad[match.group(0).strip()] = None

    # R3: nakliyesi eksik satıra sıfır nakliye veya hesaplanmış toplam atfı.
    if attributed and any(
        s.shipping is None or s.total_cost is None for s in suppliers
    ):
        if _ZERO_SHIPPING.search(segment):
            notes["nakliyesi eksik tedarikçi için sıfır/yok sayılan nakliye"] = None
        if re.search(r"toplam", segment, re.IGNORECASE) and _MONEY_CLAIM.search(
            segment
        ):
            notes["nakliyesi eksik tedarikçi için hesaplanmış toplam"] = None

    # R1/R4: "en düşük" atfı CostMark etiketleriyle doğrulanır.
    lowest = _LOWEST.search(segment)
    if (
        attributed
        and lowest
        and not _NEGATED_LOWEST.search(segment)
        # Nakliye karşılaştırması ("nakliye ... en düşük") bir CostMark
        # metriği değildir; en düşük nakliye işareti yoktur.
        and "nakliye" not in segment.casefold()
    ):
        if re.search(r"toplam|maliyet", segment, re.IGNORECASE):
            allowed = figures.lowest_total_labels
        elif re.search(r"ür[üu]n", segment, re.IGNORECASE):
            allowed = figures.lowest_product_labels
        else:
            # Ölçüt belirsizse (ör. "en düşük birim fiyat") CostMark ile
            # karşılaştırma yapılmaz; yanlış alarm üretmekten kaçınılır.
            allowed = None
        if allowed:
            # Atıf, "en düşük" ifadesine **en yakın** tedarikçiye kurulur;
            # aksi hâlde bir liste cümlesindeki tüm etiketler alarm verirdi.
            named = sorted(
                _segment_scope(segment, tuple(s.label for s in suppliers)),
                key=lambda label: _distance(label, segment, lowest.start()),
            )
            if named and named[0] not in allowed:
                notes[f"en düşük atfı: {named[0]}"] = None

    return tuple(bad), tuple(notes)


def unverified_figures(text: str, figures: FigureSet) -> tuple[str, ...]:
    """Yanıttaki kaynakla çelişen sayısal/atıf ifadelerini döndürür.

    Her sayı, yalnızca **aynı cümlede anılan tedarikçinin** kendi metriğiyle
    karşılaştırılır. Böylece bir tedarikçinin tutarı, termini veya nakliyesi
    başkasına atfedilemez. Tedarikçi adı geçmeyen cümlede sayılar tüm
    tedarikçilerin aynı metriğine karşı denenir; atıf gerektiren kurallar
    (eksik nakliye, "en düşük") ise yalnızca ad geçen cümlede uygulanır.
    Yalnızca birimiyle yazılmış ifadeler denetlenir; "3 tedarikçi" gibi
    serbest sayılar iddia sayılmaz.
    """
    labels = tuple(supplier.label for supplier in figures.suppliers)
    found: dict[str, None] = {}
    for segment in _SENTENCE_SPLIT.split(text):
        if not segment.strip():
            continue
        # Atıf yalnızca ad **aynı cümlede** geçtiğinde kurulur. Cümleler arası
        # taşıma, adsız genel bir cümleyi (ör. "en düşük toplam maliyet ...")
        # önceki tedarikçiye bağlayıp yanlış alarm üretirdi.
        hits = _segment_scope(segment, labels)
        if hits:
            active = tuple(
                supplier
                for supplier in (figures.by_label(label) for label in hits)
                if supplier is not None
            )
        else:
            active = figures.suppliers
        bad, notes = _check_figures(
            segment, active, figures, attributed=bool(hits)
        )
        for item in bad + notes:
            found[item] = None
    return tuple(found)


def _section_bodies(raw: object) -> dict[str, str]:
    if not isinstance(raw, dict):
        raise AIAnalysisError(
            "AI yanıtı beklenen yapıda değil. Lütfen yeniden deneyin.",
            kind="format",
        )

    bodies: dict[str, str] = {}
    for key, title in SECTION_TITLES:
        value = raw.get(key)
        if isinstance(value, (list, tuple)):
            value = "\n".join(str(item).strip() for item in value if str(item).strip())
        if not isinstance(value, str) or not value.strip():
            raise AIAnalysisError(
                f"AI yanıtında '{title}' bölümü boş. Lütfen yeniden deneyin.",
                kind="format",
            )
        bodies[key] = value.strip()
    return bodies


def _parse_response(content: str) -> dict[str, str]:
    """Yanıt metnini JSON'a çevirir; kod bloğu sarmalını tolere eder."""
    cleaned = _FENCE.sub("", content.strip())
    try:
        raw = json.loads(cleaned)
    except json.JSONDecodeError as error:
        raise AIAnalysisError(
            "AI yanıtı okunamadı (geçersiz JSON). Lütfen yeniden deneyin.",
            kind="format",
        ) from error
    return _section_bodies(raw)


# --------------------------------------------------------------------------- #
# Analiz
# --------------------------------------------------------------------------- #


def _content_of(response: Any) -> str:
    choices = getattr(response, "choices", None) or ()
    if not choices:
        raise AIAnalysisError(
            "AI servisi boş yanıt döndürdü. Lütfen yeniden deneyin.",
            kind="empty",
        )
    message = getattr(choices[0], "message", None)
    content = getattr(message, "content", None)
    if not isinstance(content, str) or not content.strip():
        raise AIAnalysisError(
            "AI servisi boş yanıt döndürdü. Lütfen yeniden deneyin.",
            kind="empty",
        )
    return content


def analyze(
    parsed: ParsedWorkbook,
    validation: ValidationResult,
    result: CalculationResult,
    *,
    client: Any | None = None,
    timeout: float = DEFAULT_TIMEOUT,
) -> AIAnalysis:
    """Doğrulanmış sonuçları AI ile yorumlar.

    Args:
        parsed: Parser çıktısı (yalnızca maskelenmiş alanlar kullanılır).
        validation: Doğrulama sonucu; kapı koşulu ve bulgular buradan gelir.
        result: Hesaplama sonucu; tüm tutarlar buradan **okunur**.
        client: Test veya özel kurulum için OpenAI uyumlu istemci.
        timeout: İstemci üretilirken kullanılan ağ zaman aşımı.

    Raises:
        AIAnalysisError: Analiz yapılamıyor, servis hata verdi veya yanıt
            kaynak verilerle çelişiyor.
    """
    reason = analysis_block_reason(validation, result)
    if reason is not None:
        raise AIAnalysisError(reason, kind="unavailable")

    request = build_request(parsed, validation, result)

    if client is None:
        key = api_key()
        if key is None:
            raise AIAnalysisError(
                "AI analizi için API anahtarı tanımlı değil "
                f"({API_KEY_ENV}). Anahtar tanımlandıktan sonra tekrar deneyin.",
                kind="config",
            )
        client = _create_client(key, timeout)

    import openai

    try:
        response = client.chat.completions.create(
            model=_model_name(),
            messages=[
                {"role": "system", "content": _SYSTEM_PROMPT},
                {
                    "role": "user",
                    "content": json.dumps(
                        request.payload, ensure_ascii=False, indent=2
                    ),
                },
            ],
            temperature=0.2,
            # Model bir "reasoning" modeli: düşünme jetonları da bu bütçeden
            # harcanır. Küçük bir sınır içeriği boş bırakır.
            max_tokens=4000,
        )
    except openai.APITimeoutError as error:
        raise AIAnalysisError(
            "AI servisi zaman aşımına uğradı. Lütfen tekrar deneyin.",
            kind="timeout",
        ) from error
    except openai.AuthenticationError as error:
        raise AIAnalysisError(
            "AI servisi kimlik doğrulamayı reddetti. API anahtarını kontrol edin.",
            kind="auth",
        ) from error
    except openai.RateLimitError as error:
        raise AIAnalysisError(
            "AI servisi istek sınırına ulaşıldığını bildirdi. Kısa bir süre "
            "sonra tekrar deneyin.",
            kind="rate_limit",
        ) from error
    except openai.APIConnectionError as error:
        raise AIAnalysisError(
            "AI servisine bağlanılamadı. Ağ bağlantısını kontrol edip tekrar "
            "deneyin.",
            kind="connection",
        ) from error
    except openai.APIStatusError as error:
        raise AIAnalysisError(
            "AI servisi isteği tamamlayamadı. Lütfen tekrar deneyin.",
            kind="service",
        ) from error
    except AIAnalysisError:
        raise
    except Exception as error:  # pragma: no cover - beklenmeyen istemci hatası
        raise AIAnalysisError(
            "AI analizi sırasında beklenmeyen bir hata oluştu.",
            kind="service",
        ) from error

    bodies = _parse_response(_content_of(response))

    text = "\n".join(bodies[key] for key, _title in SECTION_TITLES)
    unverified = unverified_figures(text, request.figures)
    if unverified:
        raise AIAnalysisError(
            "AI yanıtında kaynak verilerle eşleşmeyen sayısal ifadeler var; "
            "analiz gösterilmedi. Lütfen yeniden deneyin.",
            kind="unverified",
            details=unverified,
        )

    sections = tuple(
        AiSection(key=key, title=title, body=bodies[key])
        for key, title in SECTION_TITLES
    )
    return AIAnalysis(
        sections=sections,
        model=_model_name(),
        legend=request.legend,
    )

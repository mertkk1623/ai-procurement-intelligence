"""Doğrulama modülü — ``EXCEL_TEMPLATE_SPEC.md`` §7 kataloğunun tamamı.

Sorumluluk ayrımı
-----------------
* :mod:`procurement.excel_parser` dosyayı **okur** ve yapıyı **tespit eder**.
* Bu modül **doğrular** ve bulgu üretir; 46 kodun 29'u hata, 16'sı uyarı,
  1'i bilgi talebidir.

Yapmadıkları (bilinçli sınır):

* maliyet hesabı (toplam ağırlık, ürün bedeli, nakliye/kg, toplam maliyet),
* tedarikçi sıralaması, puanlama, kazanan belirleme,
* dosyaya yazma, hücre düzeltme, varsayılan değer atama.

kg/m referans aralığı (2,27–2,51) yalnızca **uyarı** üretir; hesaplama her
zaman tedarikçinin beyan ettiği ağırlıkla yapılır (Karar: brief §6).

Önem sırası ve kısa devre
-------------------------
Katalogdaki dosya düzeyi zincir spec §2.3'teki sırayı izler:
``E-DECL-01`` → ``E-TPL-01`` → ``E-TPL-02`` → ``E-TPL-03`` → ``W-TPL-01``.
Bu zincirde ilk hata sonraki adımı durdurur. Zorunlu başlıklar eksikse
satır düzeyi doğrulama çalıştırılmaz: veri güvenilir okunamaz.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal
from enum import Enum
from typing import Any

from openpyxl.utils import get_column_letter

from .excel_parser import (
    EXPECTED_TEMPLATE_CODE,
    SHEET_INSTRUCTIONS,
    SHEET_QUOTES,
    SHEET_RFQ,
    SUPPLIER_QUOTE_HEADERS,
    VAT_DECLARATION_YES,
    ParsedWorkbook,
)

__all__ = [
    "ALL_CODES",
    "ERROR_CODES",
    "INFO_CODES",
    "WARNING_CODES",
    "Finding",
    "Severity",
    "ValidationResult",
    "derive_payment_days",
    "validate",
]

# --------------------------------------------------------------------------- #
# Sabitler — spec §5, §6, §7
# --------------------------------------------------------------------------- #

EXPECTED_TEMPLATE_VERSION: str = "1.0"
EXPECTED_PROFILE_SPEC: str = "40x40x2"
EXPECTED_PRODUCT_CODE: str = "KUTUPROFIL-40X40X2-GALV"
EXPECTED_LENGTH_M: int = 6
EXPECTED_WALL_THICKNESS_MM: int = 2
EXPECTED_CURRENCY: str = "TRY"

WEIGHT_REFERENCE_MIN: float = 2.27
WEIGHT_REFERENCE_MAX: float = 2.51

MAX_QUANTITY: int = 1_000_000
PRICE_UPPER_BOUND: float = 1_000_000.0
MAX_DELIVERY_DAYS: int = 365

RFQ_NUMBER_PATTERN: re.Pattern[str] = re.compile(r"^RFQ-\d{4}-\d{3,4}$")

SHIPPING_SOURCES: tuple[str, ...] = ("Supplier Quote", "Customer Assumption")

#: Katalogdaki 46 kod — spec §7.2.
ERROR_CODES: frozenset[str] = frozenset(
    {
        "E-DECL-01",
        "E-TPL-01",
        "E-TPL-02",
        "E-TPL-03",
        "E-SHEET-01",
        "E-HDR-01",
        "E-RFQ-01",
        "E-RFQ-02",
        "E-SPEC-01",
        "E-SPEC-02",
        "E-SPEC-03",
        "E-SPEC-04",
        "E-LEN-01",
        "E-LEN-02",
        "E-CUR-01",
        "E-QTY-01",
        "E-Q-01",
        "E-WT-01",
        "E-PRICE-01",
        "E-SHIP-01",
        "E-SRC-01",
        "E-PAY-01",
        "E-QD-01",
        "E-VAL-01",
        "E-NAME-01",
        "E-NAME-02",
        "E-DUP-01",
        "E-FMT-01",
        "E-FMT-02",
    }
)

WARNING_CODES: frozenset[str] = frozenset(
    {
        "W-TPL-01",
        "W-SHEET-01",
        "W-HDR-01",
        "W-HDR-02",
        "W-RFQ-01",
        "W-RFQ-02",
        "W-WT-01",
        "W-PRICE-01",
        "W-SRC-01",
        "W-DLV-01",
        "W-DLV-02",
        "W-DLV-03",
        "W-PAY-01",
        "W-QD-01",
        "W-VAL-01",
        "W-DUP-02",
    }
)

INFO_CODES: frozenset[str] = frozenset({"I-SHIP-01"})

ALL_CODES: frozenset[str] = ERROR_CODES | WARNING_CODES | INFO_CODES

#: Dosya düzeyi hata kodları — bulgu satır numarası taşımaz (spec §7.1).
FILE_LEVEL_CODES: frozenset[str] = frozenset(
    {
        "E-DECL-01",
        "E-TPL-01",
        "E-TPL-02",
        "E-TPL-03",
        "E-SHEET-01",
        "E-HDR-01",
        "E-RFQ-01",
        "E-RFQ-02",
        "E-SPEC-02",
        "E-SPEC-03",
        "E-SPEC-04",
        "E-LEN-02",
        "E-CUR-01",
    }
)

_CURRENCY_TOKENS: tuple[str, ...] = ("₺", "$", "€", "£", "¥", "tl", "try", "usd", "eur", "gbp")
_DATE_FORMATS: tuple[str, ...] = ("%Y-%m-%d", "%d.%m.%Y")


# --------------------------------------------------------------------------- #
# Bulgu modeli
# --------------------------------------------------------------------------- #


class Severity(str, Enum):
    """Bulguların önem düzeyi — spec §7.1."""

    ERROR = "error"
    WARNING = "warning"
    INFO = "info"


def _severity_of(code: str) -> Severity:
    if code.startswith("E-"):
        return Severity.ERROR
    if code.startswith("W-"):
        return Severity.WARNING
    return Severity.INFO


@dataclass(frozen=True, slots=True)
class Finding:
    """Tek bir doğrulama bulgusu.

    Attributes:
        code: Katalog kodu (ör. ``E-WT-01``).
        severity: ``E-`` hata, ``W-`` uyarı, ``I-`` bilgi talebi.
        message: Kullanıcıya gösterilecek açıklama.
        supplier: İlgili tedarikçi adı; dosya düzeyi bulgularda ``None``.
        row_number: Excel satır numarası; dosya düzeyinde ``None``.
        cell: Hücre adresi (ör. ``F5``); bilinmiyorsa ``None``.
        sheet: İlgili sayfa adı.
        file_level: Kodun varsayılan kapsamını geçersiz kılar. ``None`` ise
            :data:`FILE_LEVEL_CODES` esas alınır. ``E-FMT-01``/``E-FMT-02``
            gibi kodlar hem RFQ (dosya) hem tedarikçi satırı (satır) düzeyinde
            geçebildiği için kapsam bulgu bazında belirtilebilir.
    """

    code: str
    severity: Severity
    message: str
    supplier: str | None = None
    row_number: int | None = None
    cell: str | None = None
    sheet: str | None = None
    file_level: bool | None = None

    @property
    def is_file_level(self) -> bool:
        if self.file_level is not None:
            return self.file_level
        return self.code in FILE_LEVEL_CODES


@dataclass(frozen=True, slots=True)
class ValidationResult:
    """Doğrulama çıktısı.

    Sıralama, puan veya kazanan bilgisi **içermez**; yalnızca bulgular ve
    bunlardan türetilen durum bilgisi.
    """

    findings: tuple[Finding, ...]

    @property
    def errors(self) -> tuple[Finding, ...]:
        return tuple(f for f in self.findings if f.severity is Severity.ERROR)

    @property
    def warnings(self) -> tuple[Finding, ...]:
        return tuple(f for f in self.findings if f.severity is Severity.WARNING)

    @property
    def infos(self) -> tuple[Finding, ...]:
        return tuple(f for f in self.findings if f.severity is Severity.INFO)

    @property
    def file_errors(self) -> tuple[Finding, ...]:
        return tuple(f for f in self.errors if f.is_file_level)

    @property
    def file_rejected(self) -> bool:
        """Dosya düzeyinde en az bir hata var mı."""
        return bool(self.file_errors)

    @property
    def blocked_row_numbers(self) -> tuple[int, ...]:
        """Satır düzeyi hata alan satırlar — karşılaştırmaya alınmaz."""
        rows = {
            f.row_number
            for f in self.errors
            if f.row_number is not None and not f.is_file_level
        }
        return tuple(sorted(rows))

    @property
    def is_complete(self) -> bool:
        """Hata yok ve bilgi talebi yok — karşılaştırma tamamlanmış sayılır."""
        return not self.errors and not self.infos

    @property
    def codes(self) -> tuple[str, ...]:
        """Üretilen kodlar, ilk görülme sırasıyla."""
        return tuple(dict.fromkeys(f.code for f in self.findings))

    def by_code(self, code: str) -> tuple[Finding, ...]:
        return tuple(f for f in self.findings if f.code == code)


def _finding(
    code: str,
    message: str,
    *,
    supplier: str | None = None,
    row_number: int | None = None,
    cell: str | None = None,
    sheet: str | None = None,
    file_level: bool | None = None,
) -> Finding:
    return Finding(
        code=code,
        severity=_severity_of(code),
        message=message,
        supplier=supplier,
        row_number=row_number,
        cell=cell,
        sheet=sheet,
        file_level=file_level,
    )


# --------------------------------------------------------------------------- #
# Hücre değeri yorumlama
# --------------------------------------------------------------------------- #


def _coerce_number(value: Any) -> tuple[float | None, str | None]:
    """Sayıya çevirir; çevrilemiyorsa nedeni döndürür.

    Spec §3.4: ``28,50`` gibi virgüllü ondalık, sayı olarak yorumlanabiliyorsa
    kabul edilir. Para simgesi içeren veya binlik/ondalık ayırıcısı karışmış
    metin hata üretir; sessiz tahmin yapılmaz.

    Returns:
        ``(sayı, None)`` başarılı; ``(None, neden)`` biçim hatası;
        ``(None, None)`` hücre boş.
    """
    if value is None:
        return None, None
    if isinstance(value, bool):
        return None, "mantıksal değer sayı olarak kullanılamaz"
    if isinstance(value, (int, float)):
        return float(value), None
    if isinstance(value, Decimal):
        return float(value), None
    if not isinstance(value, str):
        return None, f"{type(value).__name__} tipi sayı değil"

    raw = value.strip()
    if not raw:
        return None, None

    lowered = raw.casefold()
    if any(token in lowered for token in _CURRENCY_TOKENS):
        return None, "para simgesi/birimi içeriyor"
    if "." in raw and "," in raw:
        return None, "binlik ve ondalık ayırıcı karışık"
    if raw.count(",") > 1:
        return None, "birden fazla ondalık ayırıcı"

    candidate = raw.replace(",", ".") if "," in raw else raw
    try:
        return float(candidate), None
    except ValueError:
        return None, "sayı olarak okunamadı"


def _coerce_date(value: Any) -> tuple[date | None, str | None]:
    """Tarihe çevirir; ``yyyy-mm-dd`` ve ``dd.mm.yyyy`` kabul edilir (spec §3.4)."""
    if value is None:
        return None, None
    if isinstance(value, datetime):
        return value.date(), None
    if isinstance(value, date):
        return value, None
    if not isinstance(value, str):
        return None, f"{type(value).__name__} tipi tarih değil"

    raw = value.strip()
    if not raw:
        return None, None
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(raw, fmt).date(), None
        except ValueError:
            continue
    return None, "tarih olarak okunamadı (yyyy-mm-dd veya dd.mm.yyyy beklenir)"


def _normalize_spec(value: Any) -> str:
    """Ürün kimliği karşılaştırması için normalizasyon — spec §7.3.

    ``×`` → ``x``, boşluk silme, ``mm`` sonekini atma, Türkçe ``İ/I`` katlama.
    Yalnızca karşılaştırma içindir; raporda kullanıcının yazdığı değer kalır.
    """
    if not isinstance(value, str):
        return ""
    text = value.strip().casefold().replace("×", "x")
    text = "".join(text.split())
    if text.endswith("mm"):
        text = text[:-2]
    return text.translate(str.maketrans({"ı": "i", "i": "i"}))


def _is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _as_int(value: float) -> int:
    return int(value)


# --------------------------------------------------------------------------- #
# Ödeme koşulu → vade günü (spec §8)
# --------------------------------------------------------------------------- #


def derive_payment_days(text: str | None) -> int | None:
    """Ödeme koşulundan vade günü türetir; karma/belirsizde ``None``.

    Sistem belirsiz veya karma koşulda gün **uydurmaz** (spec §8, kural 1).
    """
    if not text:
        return None
    normalized = " ".join(text.casefold().split())
    if not normalized:
        return None

    # Karma / çok dilimli işaretleri: tek bir güne indirgenemez.
    if any(marker in normalized for marker in ("%", "+", "/", "ve", "&")):
        return None

    if normalized in {"peşin", "peşin ödeme", "pesin", "cash", "nakit"}:
        return 0

    match = re.search(r"\bnet\s*(\d{1,3})\b", normalized)
    if match:
        return int(match.group(1))

    match = re.search(r"\b(\d{1,3})\s*g[üu]n", normalized)
    if match:
        return int(match.group(1))

    if "peşin" in normalized or "cash" in normalized:
        return 0

    return None


# --------------------------------------------------------------------------- #
# Yardımcı: hücre adresi
# --------------------------------------------------------------------------- #


def _column_letter(parsed: ParsedWorkbook, header: str) -> str:
    columns = parsed.structure.header_columns
    index = columns.get(header)
    if index is None:
        index = SUPPLIER_QUOTE_HEADERS.index(header) + 1
    return get_column_letter(index)


def _cell_ref(parsed: ParsedWorkbook, header: str, row: int) -> str:
    return f"{_column_letter(parsed, header)}{row}"


# --------------------------------------------------------------------------- #
# 1. Sayfa yapısı
# --------------------------------------------------------------------------- #


def _check_sheets(parsed: ParsedWorkbook) -> list[Finding]:
    findings: list[Finding] = []

    for missing in parsed.structure.missing_sheets:
        findings.append(
            _finding(
                "E-SHEET-01",
                f"Zorunlu sayfa bulunamadı: '{missing}'. "
                f"Dosyadaki sayfalar: {', '.join(parsed.structure.sheet_names) or '(yok)'}.",
                sheet=missing,
            )
        )

    for extra in parsed.structure.extra_sheets:
        findings.append(
            _finding(
                "W-SHEET-01",
                f"Fazladan sayfa yok sayıldı: '{extra}'.",
                sheet=extra,
            )
        )

    return findings


# --------------------------------------------------------------------------- #
# 2. Şablon kimliği ve KDV beyanı (spec §2.3 zinciri)
# --------------------------------------------------------------------------- #


def _parse_version(raw: Any) -> tuple[int, int] | None:
    if not isinstance(raw, str):
        return None
    match = re.fullmatch(r"\s*(\d+)\.(\d+)\s*", raw)
    if not match:
        return None
    return int(match.group(1)), int(match.group(2))


def _header_defect_sets(
    parsed: ParsedWorkbook,
) -> tuple[tuple[str, ...], tuple[tuple[str, str], ...]]:
    """Başlık kusurlarını iki kümeye ayırır.

    Returns:
        ``(eksik_sütunlar, bozuk_başlıklar)``.

        * *Eksik sütun*: başlık hücresi tamamen boş — sütun yok (E-TPL-03).
        * *Bozuk başlık*: hücre dolu ama beklenen metne eşleşmiyor (E-HDR-01).
    """
    structure = parsed.structure
    misspelled_expected = {expected for _label, expected, _col in structure.misspelled_headers}

    absent = tuple(
        name for name in structure.missing_headers if name not in misspelled_expected
    )
    corrupt = tuple(
        (expected, f"{get_column_letter(column)}2")
        for _label, expected, column in structure.misspelled_headers
    )
    return absent, corrupt


def _check_template(parsed: ParsedWorkbook) -> list[Finding]:
    findings: list[Finding] = []
    template = parsed.template

    # (0) KDV beyanı — dosya işleme alınmadan önce. İlk hata zinciri durdurur.
    declared = template.vat_declaration
    value = declared.strip().upper() if isinstance(declared, str) else ""
    if value != VAT_DECLARATION_YES:
        shown = "(boş)" if _is_blank(declared) else str(declared)
        findings.append(
            _finding(
                "E-DECL-01",
                "KDV beyanı onaylanmadı; dosya işleme alınamaz. "
                f"Instructions!B4 değeri '{VAT_DECLARATION_YES}' olmalıdır (okunan: {shown}).",
                cell="B4",
                sheet=SHEET_INSTRUCTIONS,
            )
        )
        return findings

    # (1) Şablon kodu.
    if template.code != EXPECTED_TEMPLATE_CODE:
        shown = "(boş)" if _is_blank(template.code) else str(template.code)
        findings.append(
            _finding(
                "E-TPL-01",
                f"Şablon kodu uyumsuz: beklenen '{EXPECTED_TEMPLATE_CODE}', okunan '{shown}'.",
                cell="B3",
                sheet=SHEET_INSTRUCTIONS,
            )
        )
        return findings  # zincir durur

    # (2) Major sürüm.
    version = _parse_version(template.version)
    expected_major, expected_minor = _parse_version(EXPECTED_TEMPLATE_VERSION) or (1, 0)
    if version is None:
        findings.append(
            _finding(
                "E-TPL-02",
                f"Şablon sürümü okunamadı: {template.version!r}. "
                f"Beklenen biçim MAJOR.MINOR (ör. '{EXPECTED_TEMPLATE_VERSION}').",
                cell="B2",
                sheet=SHEET_INSTRUCTIONS,
            )
        )
        return findings

    if version[0] != expected_major:
        findings.append(
            _finding(
                "E-TPL-02",
                f"Şablon major sürümü uyumsuz: beklenen {expected_major}.x, "
                f"okunan {template.version}.",
                cell="B2",
                sheet=SHEET_INSTRUCTIONS,
            )
        )
        return findings

    # (3) Zorunlu sütunların varlığı.
    # Ayrım: hücre tamamen boşsa sütun "yok"tur (E-TPL-03). Hücre dolu ama
    # beklenen metin değilse başlık satırı "bozuk"tur (E-HDR-01) ve bu
    # E-TPL-03'ten önce raporlanır (spec §7.1).
    structure = parsed.structure
    absent, corrupt = _header_defect_sets(parsed)

    if SHEET_QUOTES in structure.sheet_names and absent:
        findings.append(
            _finding(
                "E-TPL-03",
                "Aynı major sürümde zorunlu sütun(lar) eksik: "
                f"{', '.join(absent)}. Dosya reddedilir; varsayılan değer doldurulmaz.",
                cell="A2",
                sheet=SHEET_QUOTES,
            )
        )
        return findings

    if corrupt:
        for name, cell in corrupt:
            findings.append(
                _finding(
                    "E-HDR-01",
                    f"Başlık satırı bozuk: '{name}' beklenen konumda değil veya yazımı farklı.",
                    cell=cell,
                    sheet=SHEET_QUOTES,
                )
            )
        return findings

    # (4) Minor sürüm farkı.
    if version[1] != expected_minor and structure.required_columns_present:
        findings.append(
            _finding(
                "W-TPL-01",
                f"Şablon minor sürümü farklı: okunan {template.version}, "
                f"beklenen {EXPECTED_TEMPLATE_VERSION}. Zorunlu sütunlar mevcut; okuma sürer.",
                cell="B2",
                sheet=SHEET_INSTRUCTIONS,
            )
        )

    return findings


# --------------------------------------------------------------------------- #
# 3. Başlık satırı
# --------------------------------------------------------------------------- #


def _check_headers(parsed: ParsedWorkbook) -> list[Finding]:
    findings: list[Finding] = []
    structure = parsed.structure

    if SHEET_QUOTES not in structure.sheet_names:
        return findings

    for label, column in structure.unknown_headers:
        findings.append(
            _finding(
                "W-HDR-01",
                f"Beklenmeyen sütun okunmadı: '{label}'. "
                "Yalnızca tanımlı 14 başlık veri olarak kullanılır.",
                cell=f"{get_column_letter(column)}2",
                sheet=SHEET_QUOTES,
            )
        )

    # Mükerrer başlık: aynı ad ikinci kez geçiyorsa fazladan sütun veri olarak
    # kullanılmaz ve tanımlı 14 başlık dışında kalan sütun gibi bildirilir
    # (spec §2.3 kural 3, §7.2 W-HDR-01). Sessizce kabul edilmez.
    for label, column in structure.duplicate_headers:
        findings.append(
            _finding(
                "W-HDR-01",
                f"Beklenmeyen sütun okunmadı: '{label}' (başlık {column}. sütunda "
                f"mükerrer). Yalnızca tanımlı 14 başlık veri olarak kullanılır.",
                cell=f"{get_column_letter(column)}2",
                sheet=SHEET_QUOTES,
            )
        )

    # Sıra kontrolü yalnızca tüm başlıklar birebir mevcutsa anlamlıdır.
    if structure.required_columns_present:
        canonical = {name: index + 1 for index, name in enumerate(SUPPLIER_QUOTE_HEADERS)}
        moved = [
            (name, column)
            for name, column in structure.header_order
            if canonical.get(name) != column
        ]
        if moved:
            findings.append(
                _finding(
                    "W-HDR-02",
                    "Sütun sırası beklenenden farklı: "
                    + ", ".join(f"'{name}' → {get_column_letter(column)}" for name, column in moved)
                    + ". Başlık adına göre eşleştirme yapıldı.",
                    cell="A2",
                    sheet=SHEET_QUOTES,
                )
            )

    return findings


# --------------------------------------------------------------------------- #
# 4. RFQ
# --------------------------------------------------------------------------- #


def _rfq_order_quantity(parsed: ParsedWorkbook) -> tuple[float | None, str | None]:
    """RFQ sipariş miktarını çözer.

    Returns:
        ``(geçerli_miktar, sorun)``. Miktar geçerliyse ``(miktar, None)``;
        değilse ``(None, neden)``. Tek kaynak: ``_check_rfq`` (E-RFQ-02) ve
        ``_check_suppliers`` (E-Q-01) aynı kararı kullanır, böylece miktar
        tutarlılığı denetimi sessizce farklı yorumlanmaz.
    """
    quantity, error = _coerce_number(parsed.rfq.quantity_pcs)
    if error is not None:
        return None, f"sayı olarak okunamadı ({error})"
    if quantity is None:
        return None, "boş"
    if quantity <= 0 or not float(quantity).is_integer() or quantity > MAX_QUANTITY:
        return None, repr(parsed.rfq.quantity_pcs)
    return quantity, None


def _check_rfq(parsed: ParsedWorkbook, today: date) -> list[Finding]:
    findings: list[Finding] = []
    rfq = parsed.rfq
    sheet = SHEET_RFQ

    # E-RFQ-01 — zorunlu alanlar (B3, B4, B10)
    required = (
        ("B3", "RFQ Number", rfq.rfq_number),
        ("B4", "RFQ Date", rfq.rfq_date),
        ("B10", "Delivery Location", rfq.delivery_location),
    )
    for cell, label, value in required:
        if _is_blank(value):
            findings.append(
                _finding(
                    "E-RFQ-01",
                    f"Zorunlu RFQ alanı boş: '{label}'. Varsayılan değer atanmaz.",
                    cell=cell,
                    sheet=sheet,
                )
            )

    # E-FMT-02 — RFQ tarihi okunamıyorsa. RFQ zorunlu bir alan olduğu için bu
    # bulgu dosya düzeyindedir (spec §7.1: eksik/geçersiz zorunlu RFQ alanı).
    rfq_date, date_error = _coerce_date(rfq.rfq_date)
    if date_error:
        findings.append(
            _finding(
                "E-FMT-02",
                f"RFQ Date okunamadı: {date_error}.",
                cell="B4",
                sheet=sheet,
                file_level=True,
            )
        )

    # W-RFQ-02 — ileri tarih (1 günden fazla)
    if rfq_date is not None and rfq_date > today + timedelta(days=1):
        findings.append(
            _finding(
                "W-RFQ-02",
                f"RFQ Date ileri tarih: {rfq_date.isoformat()} (bugün {today.isoformat()}).",
                cell="B4",
                sheet=sheet,
            )
        )

    # E-RFQ-02 — sipariş miktarı. Pozitif tam sayı değilse (metin dahil) dosya
    # düzeyinde hatadır (spec §7.2, FILE_LEVEL_CODES). Okunamayan değer için
    # ayrıca E-FMT-01 üretilmez: B9'un kendi alan kodu vardır.
    _quantity, quantity_problem = _rfq_order_quantity(parsed)
    if quantity_problem is not None:
        findings.append(
            _finding(
                "E-RFQ-02",
                f"Sipariş miktarı geçersiz: {quantity_problem}. "
                f"1–{MAX_QUANTITY} arası pozitif tam sayı beklenir. "
                "Tedarikçi satırlarının miktar tutarlılığı denetlenemez.",
                cell="B9",
                sheet=sheet,
            )
        )

    # W-RFQ-01 — talep numarası deseni
    if isinstance(rfq.rfq_number, str) and rfq.rfq_number.strip():
        if not RFQ_NUMBER_PATTERN.match(rfq.rfq_number.strip()):
            findings.append(
                _finding(
                    "W-RFQ-01",
                    f"RFQ Number önerilen desene uymuyor: '{rfq.rfq_number}'. "
                    "Önerilen desen: RFQ-YYYY-NNN(N).",
                    cell="B3",
                    sheet=sheet,
                )
            )

    # Sabit ürün kimliği alanları
    if _normalize_spec(rfq.profile_specification) != EXPECTED_PROFILE_SPEC:
        findings.append(
            _finding(
                "E-SPEC-03",
                f"RFQ Profile Specification sabit değere uymuyor: "
                f"{rfq.profile_specification!r} (beklenen '{EXPECTED_PROFILE_SPEC}').",
                cell="B6",
                sheet=sheet,
            )
        )

    wall, wall_error = _coerce_number(rfq.wall_thickness_mm)
    if wall_error or wall is None or wall != EXPECTED_WALL_THICKNESS_MM:
        findings.append(
            _finding(
                "E-SPEC-02",
                f"RFQ Wall Thickness sabit değere uymuyor: "
                f"{rfq.wall_thickness_mm!r} (beklenen {EXPECTED_WALL_THICKNESS_MM} mm).",
                cell="B8",
                sheet=sheet,
            )
        )

    if rfq.product_code != EXPECTED_PRODUCT_CODE:
        shown = "(boş)" if _is_blank(rfq.product_code) else str(rfq.product_code)
        findings.append(
            _finding(
                "E-SPEC-04",
                f"RFQ Product Code sabit değere uymuyor: '{shown}' "
                f"(beklenen '{EXPECTED_PRODUCT_CODE}').",
                cell="B5",
                sheet=sheet,
            )
        )

    length, length_error = _coerce_number(rfq.length_m)
    if length_error or length is None or length != EXPECTED_LENGTH_M:
        findings.append(
            _finding(
                "E-LEN-02",
                f"RFQ Length (m) sabit değere uymuyor: "
                f"{rfq.length_m!r} (beklenen {EXPECTED_LENGTH_M} m).",
                cell="B7",
                sheet=sheet,
            )
        )

    if rfq.currency != EXPECTED_CURRENCY:
        shown = "(boş)" if _is_blank(rfq.currency) else str(rfq.currency)
        findings.append(
            _finding(
                "E-CUR-01",
                f"RFQ Currency sabit değere uymuyor: '{shown}' "
                f"(beklenen '{EXPECTED_CURRENCY}').",
                cell="B12",
                sheet=sheet,
            )
        )

    # W-DLV-03 — istenen teslimat süresi boş
    if _is_blank(rfq.requested_delivery_days):
        findings.append(
            _finding(
                "W-DLV-03",
                "RFQ Requested Delivery Days boş; karşılaştırmada '—' gösterilecek.",
                cell="B11",
                sheet=sheet,
            )
        )

    return findings


# --------------------------------------------------------------------------- #
# 5. Tedarikçi satırları
# --------------------------------------------------------------------------- #


def _check_supplier_row(
    parsed: ParsedWorkbook,
    quote: Any,
    *,
    rfq_quantity: float | None,
    today: date,
) -> list[Finding]:
    findings: list[Finding] = []
    row = quote.row_number
    supplier = quote.supplier_name if isinstance(quote.supplier_name, str) else None
    sheet = SHEET_QUOTES

    def add(code: str, message: str, header: str) -> None:
        findings.append(
            _finding(
                code,
                message,
                supplier=supplier,
                row_number=row,
                cell=_cell_ref(parsed, header, row),
                sheet=sheet,
            )
        )

    # --- Kimlik: A, B ---
    if _is_blank(quote.supplier_name):
        add("E-NAME-01", "Tedarikçi adı boş.", "Supplier Name")
    if _is_blank(quote.supplier_quote_number):
        add("E-NAME-02", "Teklif numarası boş.", "Supplier Quote Number")

    # --- Ürün kimliği: C, D ---
    if _normalize_spec(quote.profile_specification) != EXPECTED_PROFILE_SPEC:
        add(
            "E-SPEC-01",
            f"Profil ölçüsü sabit değere uymuyor: {quote.profile_specification!r} "
            f"(beklenen '{EXPECTED_PROFILE_SPEC}'). Satır karşılaştırmaya alınmaz.",
            "Profile Specification",
        )

    length, length_error = _coerce_number(quote.length_m)
    if length_error:
        add("E-FMT-01", f"Length (m) sayı olarak okunamadı: {length_error}.", "Length (m)")
    elif length is None:
        add("E-LEN-01", "Length (m) boş; sabit değer 6 m beklenir.", "Length (m)")
    elif length != EXPECTED_LENGTH_M:
        add(
            "E-LEN-01",
            f"Profil uzunluğu sabit değere uymuyor: {quote.length_m!r} "
            f"(beklenen {EXPECTED_LENGTH_M} m). Satır karşılaştırmaya alınmaz.",
            "Length (m)",
        )

    # --- Miktar: E ---
    quantity, quantity_error = _coerce_number(quote.quantity_pcs)
    if quantity_error:
        add("E-FMT-01", f"Quantity (pcs) sayı olarak okunamadı: {quantity_error}.", "Quantity (pcs)")
    elif quantity is None:
        add("E-QTY-01", "Miktar boş; pozitif tam sayı beklenir.", "Quantity (pcs)")
    elif quantity <= 0 or not float(quantity).is_integer() or quantity > MAX_QUANTITY:
        add(
            "E-QTY-01",
            f"Miktar geçersiz: {quote.quantity_pcs!r}. "
            f"1–{MAX_QUANTITY} arası pozitif tam sayı beklenir.",
            "Quantity (pcs)",
        )
    elif rfq_quantity is not None and quantity != rfq_quantity:
        add(
            "E-Q-01",
            f"Miktar RFQ miktarıyla uyuşmuyor: satır {_as_int(quantity)}, "
            f"RFQ {_as_int(rfq_quantity)}.",
            "Quantity (pcs)",
        )

    # --- Ağırlık: F ---
    # Okunamayan kg/m alanın kendi koduyla raporlanır (E-WT-01: "≤ 0 veya sayı
    # değil", spec §7.2). Aynı hücre için ayrıca E-FMT-01 üretilmez.
    weight, weight_error = _coerce_number(quote.weight_kg_per_m)
    if weight_error:
        add(
            "E-WT-01",
            f"kg/m sayı olarak okunamadı ({weight_error}); pozitif sayı beklenir "
            f"(okunan: {quote.weight_kg_per_m!r}).",
            "Weight (kg/m)",
        )
    elif weight is None:
        add("E-WT-01", "kg/m boş; pozitif sayı beklenir.", "Weight (kg/m)")
    elif weight <= 0:
        add("E-WT-01", f"kg/m pozitif olmalıdır (okunan: {weight:g}).", "Weight (kg/m)")
    elif weight < WEIGHT_REFERENCE_MIN or weight > WEIGHT_REFERENCE_MAX:
        add(
            "W-WT-01",
            f"Beyan edilen kg/m referans aralığın dışında: {weight:g} "
            f"(referans {WEIGHT_REFERENCE_MIN:.2f}–{WEIGHT_REFERENCE_MAX:.2f}). "
            "Karşılaştırma sürer; hesaplama beyan edilen değerle yapılır.",
            "Weight (kg/m)",
        )

    # --- Birim fiyat: G ---
    # Okunamayan TL/kg alanın kendi koduyla raporlanır (E-PRICE-01, spec §7.2).
    price, price_error = _coerce_number(quote.unit_price_try_per_kg)
    if price_error:
        add(
            "E-PRICE-01",
            f"TL/kg sayı olarak okunamadı ({price_error}); pozitif sayı beklenir "
            f"(okunan: {quote.unit_price_try_per_kg!r}).",
            "Unit Price (TRY/kg, excl. VAT)",
        )
    elif price is None:
        add("E-PRICE-01", "TL/kg boş; pozitif sayı beklenir.", "Unit Price (TRY/kg, excl. VAT)")
    elif price <= 0:
        add("E-PRICE-01", f"TL/kg pozitif olmalıdır (okunan: {price:g}).", "Unit Price (TRY/kg, excl. VAT)")
    elif price > PRICE_UPPER_BOUND:
        add(
            "W-PRICE-01",
            f"TL/kg makul üst sınırın üzerinde: {price:,.2f} "
            f"(olası birim hatası, sınır {PRICE_UPPER_BOUND:,.0f}).",
            "Unit Price (TRY/kg, excl. VAT)",
        )

    # --- Nakliye: H, I ---
    shipping, shipping_error = _coerce_number(quote.shipping_cost_try)
    if shipping_error:
        add("E-FMT-01", f"Shipping Cost sayı olarak okunamadı: {shipping_error}.", "Shipping Cost (TRY, excl. VAT)")
    elif shipping is None:
        add(
            "I-SHIP-01",
            "Nakliye tutarı boş. Sıfır kabul edilmez; tutar veya açık varsayım girilmelidir.",
            "Shipping Cost (TRY, excl. VAT)",
        )
    elif shipping < 0:
        add(
            "E-SHIP-01",
            f"Nakliye tutarı negatif olamaz (okunan: {shipping:g}).",
            "Shipping Cost (TRY, excl. VAT)",
        )

    source = quote.shipping_cost_source
    source_text = source.strip() if isinstance(source, str) else ""
    source_known = source_text.casefold() in {s.casefold() for s in SHIPPING_SOURCES}

    if shipping is not None and not source_known:
        add(
            "E-SRC-01",
            f"Nakliye tutarı dolu iken kaynak boş veya geçersiz: {source!r}. "
            f"Beklenen: {' veya '.join(SHIPPING_SOURCES)}.",
            "Shipping Cost Source",
        )

    if source_known and source_text.casefold() == "customer assumption" and _is_blank(quote.notes):
        add(
            "W-SRC-01",
            "Nakliye 'Customer Assumption' olarak işaretlenmiş ancak Notes alanında "
            "gerekçe yok. Varsayımın gerekçesi yazılmalıdır.",
            "Notes",
        )

    # --- Termin: J ---
    days, days_error = _coerce_number(quote.delivery_days)
    if days_error:
        add("E-FMT-01", f"Delivery Days sayı olarak okunamadı: {days_error}.", "Delivery Days")
    elif days is None:
        add(
            "W-DLV-01",
            "Termin (Delivery Days) boş; karşılaştırmada '—' gösterilecek.",
            "Delivery Days",
        )
    elif days > MAX_DELIVERY_DAYS:
        add(
            "W-DLV-02",
            f"Termin {days:g} gün; {MAX_DELIVERY_DAYS} günün üzerinde.",
            "Delivery Days",
        )

    # --- Ödeme koşulu: K ---
    if _is_blank(quote.payment_terms):
        add("E-PAY-01", "Ödeme koşulu boş.", "Payment Terms")
    elif derive_payment_days(quote.payment_terms) is None:
        add(
            "W-PAY-01",
            f"Ödeme koşulu karma/belirsiz: '{quote.payment_terms}'. "
            "Vade günü türetilemedi; tek güne indirgenmez.",
            "Payment Terms",
        )

    # --- Tarihler: L, M ---
    quote_date, quote_date_error = _coerce_date(quote.quote_date)
    if quote_date_error:
        add("E-FMT-02", f"Quote Date okunamadı: {quote_date_error}.", "Quote Date")
    elif quote_date is None:
        add("E-QD-01", "Teklif tarihi boş; zorunludur.", "Quote Date")
    elif quote_date > today:
        add(
            "W-QD-01",
            f"Teklif tarihi gelecekte: {quote_date.isoformat()} (bugün {today.isoformat()}).",
            "Quote Date",
        )

    valid_until, valid_error = _coerce_date(quote.quote_valid_until)
    if valid_error:
        add("E-FMT-02", f"Quote Valid Until okunamadı: {valid_error}.", "Quote Valid Until")
    elif valid_until is not None:
        if quote_date is not None and valid_until < quote_date:
            add(
                "E-VAL-01",
                f"Geçerlilik tarihi teklif tarihinden önce: "
                f"{valid_until.isoformat()} < {quote_date.isoformat()}.",
                "Quote Valid Until",
            )
        elif valid_until < today:
            add(
                "W-VAL-01",
                f"Teklif geçerlilik tarihi geçmiş: {valid_until.isoformat()} "
                f"(bugün {today.isoformat()}). Karşılaştırmaya dahil edilir.",
                "Quote Valid Until",
            )

    return findings


def _check_duplicates(
    parsed: ParsedWorkbook, suppliers: tuple[Any, ...]
) -> list[Finding]:
    """E-DUP-01 ve W-DUP-02 — ad, boşluk kırpılıp büyük/küçük harf duyarsız."""
    findings: list[Finding] = []

    groups: dict[str, list[Any]] = {}
    for quote in suppliers:
        name = quote.supplier_name
        if not isinstance(name, str) or not name.strip():
            continue
        groups.setdefault(" ".join(name.casefold().split()), []).append(quote)

    for quotes in groups.values():
        if len(quotes) < 2:
            continue
        for quote in quotes:
            findings.append(
                _finding(
                    "E-DUP-01",
                    f"Aynı tedarikçi {len(quotes)} satırda tanımlı: "
                    f"satırlar {', '.join(str(q.row_number) for q in quotes)}. "
                    "Her iki satır da karşılaştırmaya alınmaz; tek satır bırakılmalıdır.",
                    supplier=quote.supplier_name,
                    row_number=quote.row_number,
                    cell=_cell_ref(parsed, "Supplier Name", quote.row_number),
                    sheet=SHEET_QUOTES,
                )
            )

        quote_groups: dict[str, list[Any]] = {}
        for quote in quotes:
            number = quote.supplier_quote_number
            if not isinstance(number, str) or not number.strip():
                continue
            quote_groups.setdefault(" ".join(number.casefold().split()), []).append(quote)

        for same in quote_groups.values():
            if len(same) < 2:
                continue
            for quote in same:
                findings.append(
                    _finding(
                        "W-DUP-02",
                        f"Aynı tedarikçi ve aynı teklif numarası tekrar ediyor: "
                        f"'{quote.supplier_quote_number}' "
                        f"(satırlar {', '.join(str(q.row_number) for q in same)}).",
                        supplier=quote.supplier_name,
                        row_number=quote.row_number,
                        cell=_cell_ref(parsed, "Supplier Quote Number", quote.row_number),
                        sheet=SHEET_QUOTES,
                    )
                )

    return findings


def _check_suppliers(parsed: ParsedWorkbook, today: date) -> list[Finding]:
    findings: list[Finding] = []

    # RFQ miktarı geçersizse E-RFQ-02 zaten dosya düzeyinde raporlanmıştır ve
    # dosya reddedildiği için buraya hiç gelinmez. Yine de tek kaynak kullanılır:
    # miktar çözülemezse E-Q-01 sessizce yanlış sonuç üretmek yerine atlanır.
    rfq_quantity, _problem = _rfq_order_quantity(parsed)

    for quote in parsed.suppliers:
        findings.extend(
            _check_supplier_row(
                parsed, quote, rfq_quantity=rfq_quantity, today=today
            )
        )

    findings.extend(_check_duplicates(parsed, parsed.suppliers))
    return findings


# --------------------------------------------------------------------------- #
# Genel giriş noktası
# --------------------------------------------------------------------------- #


def validate(parsed: ParsedWorkbook, *, today: date | None = None) -> ValidationResult:
    """``ParsedWorkbook`` üzerinde 46 kodlu katalogu uygular.

    Args:
        parsed: :func:`procurement.excel_parser.parse_workbook` çıktısı.
            Yapısal kodların üretilebilmesi için ``strict=False`` ile
            ayrıştırılmış olması önerilir.
        today: Tarih karşılaştırmalarında "bugün" olarak kullanılacak gün.
            Verilmezse sistem tarihi kullanılır; testlerde sabitlenebilir.

    Returns:
        Bulguları önem düzeyine göre sınıflandırılmış :class:`ValidationResult`.
        Sonuç sıralama, puan veya kazanan bilgisi içermez.
    """
    reference_day = today or date.today()
    findings: list[Finding] = []

    findings.extend(_check_sheets(parsed))

    structure = parsed.structure
    instructions_present = SHEET_INSTRUCTIONS in structure.sheet_names
    quotes_present = SHEET_QUOTES in structure.sheet_names
    rfq_present = SHEET_RFQ in structure.sheet_names

    # Dosya düzeyi zinciri (spec §2.3): sayfalar → beyan → kod → major sürüm
    # → zorunlu sütun → minor sürüm. İlk hata sonraki adımları durdurur.
    if instructions_present:
        findings.extend(_check_template(parsed))

    # Eksik zorunlu sayfa veya şablon zinciri hatası dosyayı reddeder; bu
    # durumda veri güvenilir okunamaz ve satır doğrulaması çalıştırılmaz.
    chain_rejected = any(
        f.code in {"E-SHEET-01", "E-DECL-01", "E-TPL-01", "E-TPL-02", "E-TPL-03"}
        for f in findings
    )
    if quotes_present and not chain_rejected:
        findings.extend(_check_headers(parsed))

    if rfq_present and not chain_rejected:
        findings.extend(_check_rfq(parsed, reference_day))

    # Satır düzeyi doğrulama; dosya reddedilmişse veya zorunlu başlıklar
    # eksikse çalıştırılmaz (spec §11.3: varsayılan değer doldurulmaz).
    if (
        quotes_present
        and not chain_rejected
        and not any(f.code == "E-HDR-01" for f in findings)
    ):
        findings.extend(_check_suppliers(parsed, reference_day))

    return ValidationResult(findings=tuple(findings))

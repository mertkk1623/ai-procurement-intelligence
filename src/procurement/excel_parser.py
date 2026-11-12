"""Excel teklif şablonunu salt okunur biçimde ayrıştıran modül.

Kapsam
------
Bu modül yalnızca *okur*: dosyayı açar, yapısal beklentileri (sayfalar,
başlıklar, RFQ alan adları) denetler ve hücre değerlerini Python veri
nesnelerine aktarır.

Yapmadıkları (bilinçli olarak dışarıda bırakılmıştır):

* maliyet hesaplama,
* 46 kodlu hata/uyarı kataloğunun doğrulanması (ayrı modül),
* şablon kodunun/sürümünün iş kuralı olarak kabulü,
* Excel dosyasına yazma — modül hiçbir koşulda ``save()`` çağırmaz.

Okunan değerler dönüştürülmez: ``int`` tam sayı, ``float`` ondalık,
``datetime`` tarih olarak korunur. Yalnızca tamamen boş veya yalnızca
boşluktan oluşan hücreler ``None`` kabul edilir.

Formül hücreleri güvenilir veri sayılmaz; okunması gereken bir hücrede
formül bulunursa :class:`FormulaCellError` üretilir.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Final

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

__all__ = [
    "DuplicateFieldError",
    "DuplicateHeaderError",
    "ExcelParserError",
    "FileTooLargeError",
    "FormulaCellError",
    "MissingFieldError",
    "MissingHeaderError",
    "MissingSheetError",
    "ParsedWorkbook",
    "RfqData",
    "SupplierQuote",
    "TemplateInfo",
    "UnsupportedFormatError",
    "VatDeclarationError",
    "WorkbookStructure",
    "parse_workbook",
]

# --------------------------------------------------------------------------- #
# Şablon sözleşmesi (docs/EXCEL_TEMPLATE_SPEC.md)
# --------------------------------------------------------------------------- #

SHEET_INSTRUCTIONS: Final = "Instructions"
SHEET_RFQ: Final = "RFQ"
SHEET_QUOTES: Final = "Supplier Quotes"

EXPECTED_SHEETS: Final[tuple[str, ...]] = (
    SHEET_INSTRUCTIONS,
    SHEET_RFQ,
    SHEET_QUOTES,
)

EXPECTED_TEMPLATE_CODE: Final = "RFQ-KUTUPROFIL-40X40X2"

VAT_DECLARATION_YES: Final = "EVET"
VAT_DECLARATION_CELL: Final = "B4"

TEMPLATE_CODE_CELL: Final = "B3"
TEMPLATE_VERSION_CELL: Final = "B2"

FIRST_RFQ_ROW: Final = 3
LAST_RFQ_ROW: Final = 13

HEADER_ROW: Final = 2
FIRST_DATA_ROW: Final = 3
LAST_DATA_ROW: Final = 102

MAX_FILE_SIZE_BYTES: Final = 5 * 1024 * 1024
ALLOWED_SUFFIX: Final = ".xlsx"
REJECTED_SUFFIXES: Final[tuple[str, ...]] = (".xls", ".xlsm", ".csv", ".ods")

#: RFQ sayfası alanları: (beklenen satır, A sütunundaki etiket, veri alanı adı)
RFQ_FIELDS: Final[tuple[tuple[int, str, str], ...]] = (
    (3, "RFQ Number", "rfq_number"),
    (4, "RFQ Date", "rfq_date"),
    (5, "Product Code", "product_code"),
    (6, "Profile Specification", "profile_specification"),
    (7, "Length (m)", "length_m"),
    (8, "Wall Thickness (mm)", "wall_thickness_mm"),
    (9, "Quantity (pcs)", "quantity_pcs"),
    (10, "Delivery Location", "delivery_location"),
    (11, "Requested Delivery Days", "requested_delivery_days"),
    (12, "Currency", "currency"),
    (13, "Description", "description"),
)

#: Supplier Quotes sayfası başlıkları — sıra spesifikasyondaki sıradır.
SUPPLIER_QUOTE_HEADERS: Final[tuple[str, ...]] = (
    "Supplier Name",
    "Supplier Quote Number",
    "Profile Specification",
    "Length (m)",
    "Quantity (pcs)",
    "Weight (kg/m)",
    "Unit Price (TRY/kg, excl. VAT)",
    "Shipping Cost (TRY, excl. VAT)",
    "Shipping Cost Source",
    "Delivery Days",
    "Payment Terms",
    "Quote Date",
    "Quote Valid Until",
    "Notes",
)

_HEADER_TO_ATTR: Final[dict[str, str]] = {
    "Supplier Name": "supplier_name",
    "Supplier Quote Number": "supplier_quote_number",
    "Profile Specification": "profile_specification",
    "Length (m)": "length_m",
    "Quantity (pcs)": "quantity_pcs",
    "Weight (kg/m)": "weight_kg_per_m",
    "Unit Price (TRY/kg, excl. VAT)": "unit_price_try_per_kg",
    "Shipping Cost (TRY, excl. VAT)": "shipping_cost_try",
    "Shipping Cost Source": "shipping_cost_source",
    "Delivery Days": "delivery_days",
    "Payment Terms": "payment_terms",
    "Quote Date": "quote_date",
    "Quote Valid Until": "quote_valid_until",
    "Notes": "notes",
}


# --------------------------------------------------------------------------- #
# Hatalar
# --------------------------------------------------------------------------- #


class ExcelParserError(Exception):
    """Ayrıştırma hatalarının ortak tabanı.

    Mesajlar kullanıcıya gösterilebilir; ``sheet``, ``cell`` ve ``field``
    alanları hata ayıklama ve arayüz etiketlemesi içindir.
    """

    def __init__(
        self,
        message: str,
        *,
        sheet: str | None = None,
        cell: str | None = None,
        field: str | None = None,
    ) -> None:
        super().__init__(message)
        self.sheet = sheet
        self.cell = cell
        self.field = field


class UnsupportedFormatError(ExcelParserError):
    """Dosya uzantısı ``.xlsx`` değil."""


class FileTooLargeError(ExcelParserError):
    """Dosya boyutu izin verilen sınırı aşıyor."""


class MissingSheetError(ExcelParserError):
    """Zorunlu sayfalardan en az biri yok."""


class MissingFieldError(ExcelParserError):
    """RFQ sayfasında beklenen alan etiketi bulunamadı."""


class DuplicateFieldError(ExcelParserError):
    """RFQ sayfasında aynı alan etiketi birden fazla kez geçiyor."""


class MissingHeaderError(ExcelParserError):
    """Supplier Quotes sayfasında zorunlu başlıklardan biri yok."""


class DuplicateHeaderError(ExcelParserError):
    """Supplier Quotes sayfasında aynı başlık birden fazla sütunda var."""


class VatDeclarationError(ExcelParserError):
    """VAT Declaration beyanı ``EVET`` değil."""


class FormulaCellError(ExcelParserError):
    """Okunması gereken hücrede formül var."""


# --------------------------------------------------------------------------- #
# Veri nesneleri
# --------------------------------------------------------------------------- #


@dataclass(frozen=True, slots=True)
class TemplateInfo:
    """``Instructions`` sayfasındaki şablon kimliği ve beyanı."""

    code: str | None
    version: str | None
    vat_declaration: str | None


@dataclass(frozen=True, slots=True)
class RfqData:
    """``RFQ`` sayfasındaki talep başlığı."""

    rfq_number: str | None
    rfq_date: date | datetime | None
    product_code: str | None
    profile_specification: str | None
    length_m: float | int | None
    wall_thickness_mm: float | int | None
    quantity_pcs: float | int | None
    delivery_location: str | None
    requested_delivery_days: float | int | None
    currency: str | None
    description: str | None


@dataclass(frozen=True, slots=True)
class SupplierQuote:
    """``Supplier Quotes`` sayfasındaki tek bir tedarikçi satırı."""

    row_number: int
    supplier_name: str | None
    supplier_quote_number: str | None
    profile_specification: str | None
    length_m: float | int | None
    quantity_pcs: float | int | None
    weight_kg_per_m: float | None
    unit_price_try_per_kg: float | None
    shipping_cost_try: float | None
    shipping_cost_source: str | None
    delivery_days: float | int | None
    payment_terms: str | None
    quote_date: date | datetime | None
    quote_valid_until: date | datetime | None
    notes: str | None


@dataclass(frozen=True, slots=True)
class WorkbookStructure:
    """Dosyanın yapısal tespit sonucu.

    Parser yalnızca *tespit* eder; bu bilgilerin hangi doğrulama kodunu
    üreteceğine Validator karar verir. Böylece eksik sayfa/başlık gibi
    durumlar da bulgu olarak raporlanabilir.
    """

    sheet_names: tuple[str, ...]
    missing_sheets: tuple[str, ...]
    extra_sheets: tuple[str, ...]
    #: Başlık satırında bulunan tanınmış başlıklar ve sütun indeksleri.
    header_order: tuple[tuple[str, int], ...]
    #: Beklenen ama bulunamayan başlıklar.
    missing_headers: tuple[str, ...]
    #: Tanınmayan, veri olarak kullanılmayan başlıklar (başlık, sütun).
    unknown_headers: tuple[tuple[str, int], ...]
    #: Beklenen başlığa benzeyen ama birebir eşleşmeyen başlıklar
    #: (okunan başlık, beklenen başlık, sütun).
    misspelled_headers: tuple[tuple[str, str, int], ...]
    #: Birden fazla sütunda geçen başlıklar ve mükerrer (fazladan) sütun.
    duplicate_headers: tuple[tuple[str, int], ...]

    @property
    def required_columns_present(self) -> bool:
        """14 zorunlu başlığın tamamı birebir mevcut mu."""
        return not self.missing_headers and not self.misspelled_headers

    @property
    def header_columns(self) -> dict[str, int]:
        """Başlık adı -> sütun indeksi eşlemesi."""
        return dict(self.header_order)


@dataclass(frozen=True, slots=True)
class ParsedWorkbook:
    """Ayrıştırma sonucu: şablon bilgisi, RFQ ve tedarikçi satırları."""

    source_path: Path
    template: TemplateInfo
    rfq: RfqData
    suppliers: tuple[SupplierQuote, ...]
    structure: WorkbookStructure


# --------------------------------------------------------------------------- #
# Yardımcılar
# --------------------------------------------------------------------------- #


def _normalize_text(value: Any) -> str:
    """Başlık/etiket karşılaştırması için boşlukları sadeleştirir.

    Spesifikasyon §11.5: baştaki/sondaki boşluklar kırpılır, çoklu boşluk
    tek boşluğa indirilir. Büyük/küçük harf değiştirilmez.
    """
    if not isinstance(value, str):
        return ""
    return " ".join(value.split())


def _is_blank(value: Any) -> bool:
    """``None`` veya yalnızca boşluktan oluşan metin -> boş."""
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    return False


def _clean(value: Any) -> Any:
    """Boş hücreleri ``None`` yapar; dolu değerleri tipini bozmadan döndürür."""
    return None if _is_blank(value) else value


def _read_cell(
    sheet: Worksheet,
    row: int,
    column: int,
    field: str,
) -> Any:
    """Hücreyi okur; formül hücresi ise hata üretir.

    ``data_only=False`` ile açıldığı için formül hücresi hesaplanmış değer
    değil formül metni taşır. Bu değer veri olarak kullanılmaz.
    """
    cell = sheet.cell(row=row, column=column)
    if cell.data_type == "f":
        raise FormulaCellError(
            "Formül hücresi güvenilir veri olarak kabul edilmez: "
            f"{sheet.title}!{cell.coordinate} ({field}). "
            "Hesaplanmış değerler dosyadan okunmaz; lütfen ham değeri girin.",
            sheet=sheet.title,
            cell=cell.coordinate,
            field=field,
        )
    return _clean(cell.value)


def _validate_file(path: Path) -> None:
    if not path.exists() or not path.is_file():
        raise ExcelParserError(
            "Dosya bulunamadı. Dosyanın silinmediğinden veya taşınmadığından "
            "emin olup yeniden yükleyin.",
            field="file",
        )

    suffix = path.suffix.lower()
    if suffix != ALLOWED_SUFFIX:
        raise UnsupportedFormatError(
            f"Desteklenmeyen dosya biçimi: '{suffix or '(uzantısız)'}'. "
            f"Yalnızca {ALLOWED_SUFFIX} kabul edilir; "
            f"{', '.join(REJECTED_SUFFIXES)} reddedilir.",
            field="file",
        )

    size = path.stat().st_size
    if size > MAX_FILE_SIZE_BYTES:
        raise FileTooLargeError(
            f"Dosya boyutu sınırı aşıyor: {size} bayt "
            f"(en fazla {MAX_FILE_SIZE_BYTES} bayt).",
            field="file",
        )


def _open_workbook(path: Path) -> Workbook:
    try:
        return load_workbook(
            filename=path,
            data_only=False,
            keep_links=False,
            read_only=False,
        )
    except (InvalidFileException, zipfile.BadZipFile, OSError, KeyError, ValueError) as exc:
        # Yol ve teknik ayrıntı kullanıcıya gösterilmez; özgün hata
        # ``__cause__`` üzerinde korunur (hata ayıklama/loglama için).
        raise ExcelParserError(
            "Dosya açılamadı veya geçerli bir .xlsx değil. Dosyanın bozuk "
            "olmadığından ve Excel biçiminde kaydedildiğinden emin olup "
            "yeniden yükleyin.",
            field="file",
        ) from exc


def _require_sheets(workbook: Workbook) -> None:
    found = tuple(workbook.sheetnames)
    missing = [name for name in EXPECTED_SHEETS if name not in found]
    if missing:
        raise MissingSheetError(
            f"Zorunlu sayfa(lar) bulunamadı: {', '.join(missing)}. "
            f"Dosyadaki sayfalar: {', '.join(found) or '(yok)'}.",
            sheet=None,
            field="sheets",
        )


# --------------------------------------------------------------------------- #
# Instructions
# --------------------------------------------------------------------------- #


def _read_template(workbook: Workbook) -> TemplateInfo:
    sheet = workbook[SHEET_INSTRUCTIONS]
    return TemplateInfo(
        code=_read_cell(sheet, 3, 2, "Template Code"),
        version=_read_cell(sheet, 2, 2, "Template Version"),
        vat_declaration=_read_cell(sheet, 4, 2, "VAT Declaration"),
    )


def _require_template(template: TemplateInfo | None) -> TemplateInfo:
    assert template is not None
    return template


def _empty_template() -> TemplateInfo:
    """Sayfa hiç yoksa kullanılan boş şablon bilgisi."""
    return TemplateInfo(code=None, version=None, vat_declaration=None)


def _empty_rfq() -> RfqData:
    """RFQ sayfası hiç yoksa kullanılan boş talep bilgisi."""
    return RfqData(
        rfq_number=None,
        rfq_date=None,
        product_code=None,
        profile_specification=None,
        length_m=None,
        wall_thickness_mm=None,
        quantity_pcs=None,
        delivery_location=None,
        requested_delivery_days=None,
        currency=None,
        description=None,
    )


def _require_vat_declaration(template: TemplateInfo) -> None:
    declared = template.vat_declaration
    value = declared.strip().upper() if isinstance(declared, str) else ""
    if value != VAT_DECLARATION_YES:
        shown = "(boş)" if _is_blank(declared) else str(declared)
        raise VatDeclarationError(
            "KDV beyanı onaylanmadığı için dosya işleme alınamaz. "
            f"{SHEET_INSTRUCTIONS}!{VAT_DECLARATION_CELL} değeri '{VAT_DECLARATION_YES}' olmalıdır "
            f"(okunan değer: {shown}).",
            sheet=SHEET_INSTRUCTIONS,
            cell=VAT_DECLARATION_CELL,
            field="VAT Declaration",
        )


# --------------------------------------------------------------------------- #
# RFQ
# --------------------------------------------------------------------------- #


def _read_rfq(workbook: Workbook, *, strict: bool = True) -> RfqData:
    sheet = workbook[SHEET_RFQ]

    positions: dict[str, int] = {}
    for row in range(FIRST_RFQ_ROW, LAST_RFQ_ROW + 1):
        label = _normalize_text(_read_cell(sheet, row, 1, "Field"))
        if not label:
            continue
        if label in positions:
            raise DuplicateFieldError(
                f"RFQ alanı birden fazla kez tanımlı: '{label}' "
                f"(satır {positions[label]} ve {row}).",
                sheet=SHEET_RFQ,
                cell=f"A{row}",
                field=label,
            )
        positions[label] = row

    values: dict[str, Any] = {}
    for expected_row, label, attr in RFQ_FIELDS:
        row = positions.get(label)
        if row is None:
            if not strict:
                values[attr] = None
                continue
            raise MissingFieldError(
                f"RFQ alanı bulunamadı: '{label}' "
                f"({SHEET_RFQ}!A{expected_row} bekleniyordu).",
                sheet=SHEET_RFQ,
                cell=f"A{expected_row}",
                field=label,
            )
        values[attr] = _read_cell(sheet, row, 2, label)

    return RfqData(**values)


# --------------------------------------------------------------------------- #
# Supplier Quotes
# --------------------------------------------------------------------------- #


def _read_header_label(sheet: Worksheet, column: int) -> str:
    """Başlık hücresini formül denetimi olmadan okur.

    Spesifikasyon §10: kullanıcının eklediği hesaplanmış sütunlar okunmaz ve
    yok sayılır. Bu yüzden tanınmayan bir sütunun başlığında formül bulunması
    ayrıştırmayı durdurmaz; yalnızca başlık eşleşmediği için sütun atlanır.
    """
    return _normalize_text(sheet.cell(row=HEADER_ROW, column=column).value)


def _header_key(value: str) -> str:
    """Başlık karşılaştırması için büyük/küçük harf ve boşluk duyarsız anahtar."""
    return " ".join(value.lower().split())


def _inspect_headers(
    sheet: Worksheet,
) -> tuple[tuple[str, int], ...]:
    """Başlık satırını tarar ve (başlık, sütun) çiftlerini döndürür.

    Yalnızca dolu hücreler döner; tanınan/tanınmayan ayrımı çağırana aittir.
    """
    found: list[tuple[str, int]] = []
    for column in range(1, sheet.max_column + 1):
        label = _read_header_label(sheet, column)
        if label:
            found.append((label, column))
    return tuple(found)


def _classify_headers(
    found: tuple[tuple[str, int], ...],
) -> tuple[
    tuple[tuple[str, int], ...],
    tuple[str, ...],
    tuple[tuple[str, int], ...],
    tuple[tuple[str, str, int], ...],
    tuple[tuple[str, int], ...],
]:
    """Başlıkları tanınan / eksik / bilinmeyen / yazımı bozuk / mükerrer olarak ayırır."""
    expected = SUPPLIER_QUOTE_HEADERS
    expected_keys = {_header_key(name): name for name in expected}

    recognized: list[tuple[str, int]] = []
    unknown: list[tuple[str, int]] = []
    misspelled: list[tuple[str, str, int]] = []
    duplicates: list[tuple[str, int]] = []
    seen: set[str] = set()

    for label, column in found:
        if label in expected:
            if label in seen:
                duplicates.append((label, column))
                continue
            seen.add(label)
            recognized.append((label, column))
            continue

        # Birebir değil ama normalize edilmiş hali beklenen bir başlığa
        # eşitse bu bir yazım/boşluk/harf farkıdır (E-HDR-01).
        candidate = expected_keys.get(_header_key(label))
        if candidate is not None:
            misspelled.append((label, candidate, column))
        else:
            unknown.append((label, column))

    missing = tuple(name for name in expected if name not in seen)
    return (
        tuple(recognized),
        missing,
        tuple(unknown),
        tuple(misspelled),
        tuple(duplicates),
    )


def _inspect_structure(workbook: Workbook) -> WorkbookStructure:
    """Dosyanın yapısal durumunu istisna fırlatmadan tespit eder."""
    present = tuple(workbook.sheetnames)
    missing_sheets = tuple(name for name in EXPECTED_SHEETS if name not in present)
    extra_sheets = tuple(name for name in present if name not in EXPECTED_SHEETS)

    if SHEET_QUOTES in present:
        found = _inspect_headers(workbook[SHEET_QUOTES])
        (
            header_order,
            missing_headers,
            unknown_headers,
            misspelled_headers,
            duplicate_headers,
        ) = _classify_headers(found)
    else:
        header_order = ()
        missing_headers = SUPPLIER_QUOTE_HEADERS
        unknown_headers = ()
        misspelled_headers = ()
        duplicate_headers = ()

    return WorkbookStructure(
        sheet_names=present,
        missing_sheets=missing_sheets,
        extra_sheets=extra_sheets,
        header_order=header_order,
        missing_headers=missing_headers,
        unknown_headers=unknown_headers,
        misspelled_headers=misspelled_headers,
        duplicate_headers=duplicate_headers,
    )


def _map_headers(sheet: Worksheet, *, strict: bool = True) -> dict[str, int]:
    """Başlık adı -> sütun indeksi eşlemesi üretir.

    ``strict=True`` (varsayılan) eksik başlıkta istisna fırlatır; mevcut
    davranış korunur. ``strict=False`` ise eşleşen başlıklarla devam eder ve
    eksikleri Validator'ın raporlamasına bırakır.
    """
    found = _inspect_headers(sheet)
    (
        recognized,
        missing,
        _unknown,
        misspelled,
        duplicates,
    ) = _classify_headers(found)

    if not strict:
        return dict(recognized)

    if duplicates:
        duplicate, duplicate_column = duplicates[0]
        columns = [column for label, column in found if label == duplicate]
        raise DuplicateHeaderError(
            f"Başlık birden fazla sütunda tanımlı: '{duplicate}' "
            f"({', '.join(sheet.cell(row=HEADER_ROW, column=c).coordinate for c in columns)}).",
            sheet=sheet.title,
            cell=sheet.cell(row=HEADER_ROW, column=duplicate_column).coordinate,
            field=duplicate,
        )

    if misspelled:
        label, expected_name, column = misspelled[0]
        raise MissingHeaderError(
            f"Başlık yazımı beklenenle eşleşmiyor: {sheet.title}!"
            f"{sheet.cell(row=HEADER_ROW, column=column).coordinate} = '{label}', "
            f"beklenen '{expected_name}'.",
            sheet=sheet.title,
            cell=sheet.cell(row=HEADER_ROW, column=column).coordinate,
            field=expected_name,
        )

    if missing:
        raise MissingHeaderError(
            f"Zorunlu sütun başlığı/baslıkları eksik: {', '.join(missing)}. "
            f"Başlıklar {sheet.title} sayfasının {HEADER_ROW}. satırında olmalıdır.",
            sheet=sheet.title,
            cell=f"A{HEADER_ROW}",
            field=", ".join(missing),
        )

    return dict(recognized)


def _read_suppliers(
    workbook: Workbook, *, strict: bool = True
) -> tuple[SupplierQuote, ...]:
    sheet = workbook[SHEET_QUOTES]
    mapping = _map_headers(sheet, strict=strict)

    last_row = min(sheet.max_row, LAST_DATA_ROW)
    quotes: list[SupplierQuote] = []

    for row in range(FIRST_DATA_ROW, last_row + 1):
        raw = {
            header: _read_cell(sheet, row, column, header)
            for header, column in mapping.items()
        }
        if all(value is None for value in raw.values()):
            continue

        # ``strict=False`` modunda eksik başlıklar olabilir; okunamayan alan
        # ``None`` kalır ve eksikliği Validator raporlar.
        quotes.append(
            SupplierQuote(
                row_number=row,
                **{
                    attr: raw.get(header)
                    for header, attr in _HEADER_TO_ATTR.items()
                },
            )
        )

    return tuple(quotes)


# --------------------------------------------------------------------------- #
# Genel giriş noktası
# --------------------------------------------------------------------------- #


def parse_workbook(source: str | Path, *, strict: bool = True) -> ParsedWorkbook:
    """Şablon dosyasını okur ve yapısal sonucu döndürür.

    Args:
        source: Okunacak ``.xlsx`` dosyasının yolu. Dosya değiştirilmez.
        strict: ``True`` (varsayılan) ise yapısal eksiklikler istisna olarak
            fırlatılır; mevcut davranış budur. ``False`` ise eksik sayfa ve
            başlıklar istisna yerine :attr:`ParsedWorkbook.structure` içinde
            raporlanır ve okunabilen alanlarla devam edilir. Validator
            yapısal kodları üretebilmek için bu modu kullanır.

    Raises:
        ExcelParserError: Dosya açılamıyor, yapı beklentileri karşılanmıyor
            veya KDV beyanı onaylanmamışsa. Alt sınıflar daha özel neden
            bildirir.
    """
    path = Path(source)
    _validate_file(path)

    workbook = _open_workbook(path)
    try:
        structure = _inspect_structure(workbook)
        if strict:
            _require_sheets(workbook)

        template = _read_template(workbook) if SHEET_INSTRUCTIONS in workbook.sheetnames else None
        if strict:
            _require_vat_declaration(_require_template(template))

        rfq = _read_rfq(workbook, strict=strict) if SHEET_RFQ in workbook.sheetnames else None
        suppliers = (
            _read_suppliers(workbook, strict=strict)
            if SHEET_QUOTES in workbook.sheetnames
            else ()
        )

        return ParsedWorkbook(
            source_path=path,
            template=template if template is not None else _empty_template(),
            rfq=rfq if rfq is not None else _empty_rfq(),
            suppliers=suppliers,
            structure=structure,
        )
    finally:
        workbook.close()

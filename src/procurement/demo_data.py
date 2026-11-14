"""Boş şablon ve dolu örnek ("demo") çalışma kitabı üretimi.

Neden dinamik?
--------------
Doğrulama katmanı tarihleri **bugüne göre** denetler (``W-QD-01`` ileri teklif
tarihi, ``W-VAL-01`` süresi geçmiş geçerlilik, ``W-RFQ-02`` ileri RFQ tarihi).
Sabit bir örnek dosya birkaç ay sonra kendiliğinden uyarı üretmeye başlardı.
Bu modül bu yüzden örnek dosyayı her istekte bugünün tarihine göre üretir;
böylece demo her zaman aynı, tutarlı senaryoyu gösterir.

Tasarım sınırları
-----------------
* Üretim, paketlenmiş resmî şablonun (``templates/supplier_quote_template.xlsx``)
  **kopyası** üzerinden yapılır. Böylece stiller, açılır liste doğrulamaları,
  başlıklar, sayfa adları ve ``Instructions`` metinleri şablonla birebir aynı
  kalır; yeni bir sözleşme tanımlanmaz.
* Hiçbir maliyet burada hesaplanmaz. Dosya yalnızca ham teklif alanlarını taşır;
  toplam ağırlık, ürün bedeli ve toplam maliyet uygulama tarafından üretilir.
* Üretilen veriler kurgusaldır. Tedarikçi adları, teklif numaraları ve tutarlar
  gerçek teklif değildir.
* Dosya sistemi kullanılmaz: çıktı yalnızca bellekte ``bytes`` olarak döner.

Örnek veri setinin senaryoları
------------------------------
Veri seti, AI analizinin ve doğrulama bulgularının tüm yüzeylerini gösterecek
şekilde kasıtlı olarak "kirli" ve dağılımlıdır:

* tedarikçiler arasında belirgin **birim fiyat uçurumları** (26,40–35,40 TL/kg),
* **termin darboğazları** (9–420 gün; istenen termin 20 gün),
* **nakliye/maliyet sapmaları** (0 TL fabrika teslim, 16.800 TL'ye varan navlun),
* **konsolidasyon fırsatları** (düşük navlunlu yakın tedarikçiler),
* **ödeme koşulu riski** (Peşin, Net 30–90, karma yüzdeli koşullar),
* **veri kalitesi tuzakları** (referans dışı kg/m, süresi geçmiş teklif,
  eksik nakliye, gerekçesiz nakliye varsayımı).
"""

from __future__ import annotations

from copy import copy

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from io import BytesIO
from pathlib import Path
from typing import Final

from openpyxl import load_workbook

from .excel_parser import (
    EXPECTED_TEMPLATE_CODE,
    SHEET_INSTRUCTIONS,
    SHEET_QUOTES,
    SHEET_RFQ,
    SUPPLIER_QUOTE_HEADERS,
)
from .validator import (
    EXPECTED_CURRENCY,
    EXPECTED_LENGTH_M,
    EXPECTED_PROFILE_SPEC,
    EXPECTED_PRODUCT_CODE,
    EXPECTED_TEMPLATE_VERSION,
    EXPECTED_WALL_THICKNESS_MM,
    MAX_DELIVERY_DAYS,
    WEIGHT_REFERENCE_MAX,
    WEIGHT_REFERENCE_MIN,
)

__all__ = [
    "BLANK_FILENAME",
    "COLUMN_GUIDE",
    "DemoDataError",
    "DemoQuote",
    "SAMPLE_DELIVERY_LOCATION",
    "SAMPLE_FILENAME",
    "SAMPLE_QUANTITY",
    "SAMPLE_QUOTES",
    "SAMPLE_REQUESTED_DAYS",
    "build_blank_template",
    "build_sample_workbook",
    "sample_supplier_count",
]

#: İndirilen dosyaların adları.
BLANK_FILENAME: str = "tedarikci-teklif-sablonu-bos.xlsx"
SAMPLE_FILENAME: str = "ornek-tedarikci-teklifleri.xlsx"

#: Örnek RFQ başlığı. Yıl, dosya üretildiği güne göre yazılır.
SAMPLE_RFQ_NUMBER: Final[str] = "RFQ-{year}-0517"
SAMPLE_QUANTITY: Final[int] = 500
SAMPLE_REQUESTED_DAYS: Final[int] = 20
SAMPLE_DELIVERY_LOCATION: Final[str] = "Bursa / Nilüfer OSB"

#: ``Instructions`` sayfasındaki KDV beyanı hücresi (parser sözleşmesi).
_VAT_CELL: Final[str] = "B4"
_VAT_YES: Final[str] = "EVET"
_VAT_NO: Final[str] = "HAYIR"

_HEADER_ROW: Final[int] = 2
_FIRST_DATA_ROW: Final[int] = 3
_LAST_DATA_ROW: Final[int] = 102

_SHIPPING_QUOTE: Final[str] = "Supplier Quote"
_SHIPPING_ASSUMPTION: Final[str] = "Customer Assumption"


class DemoDataError(Exception):
    """Örnek/boş dosya üretilemediğinde fırlatılır.

    Mesaj kullanıcıya gösterilebilir; dosya yolu veya teknik ayrıntı
    taşımaz. Özgün hata ``__cause__`` üzerinde korunur.
    """


# --------------------------------------------------------------------------- #
# Veri modeli
# --------------------------------------------------------------------------- #


@dataclass(frozen=True, slots=True)
class DemoQuote:
    """Örnek veri setindeki tek bir tedarikçi teklifi.

    Attributes:
        supplier: Tedarikçi adı (kurgusal).
        quote_number: Teklif numarası; her satırda tekil.
        weight_kg_per_m: Tedarikçinin beyan ettiği kg/m. Referans aralığın
            dışındaki değerler kasıtlıdır ve yalnızca ``W-WT-01`` üretir.
        unit_price: KDV hariç TL/kg.
        shipping: KDV hariç TL. ``None`` ise nakliye **bilinmiyor** demektir;
            sıfır değildir ve ``I-SHIP-01`` üretir.
        shipping_source: ``Supplier Quote`` veya ``Customer Assumption``.
        delivery_days: Termin (gün).
        payment_terms: Ödeme koşulu metni (vade günü yazılmaz).
        quote_days_ago: Teklif tarihi bugünden kaç gün önce.
        valid_days: Geçerlilik tarihi bugünden kaç gün sonra (negatif = süresi
            geçmiş).
        notes: Serbest not. ``Customer Assumption`` satırlarında gerekçe
            bulunur; biri bilinçli olarak boştur (``W-SRC-01``).
    """

    supplier: str
    quote_number: str
    weight_kg_per_m: float
    unit_price: float
    shipping: float | None
    shipping_source: str
    delivery_days: int
    payment_terms: str
    quote_days_ago: int
    valid_days: int
    notes: str


#: 35 satırlık örnek veri seti.
#:
#: Dağılım bilinçli kurulmuştur: birkaç "rekabetçi" tedarikçi, geniş bir orta
#: grup, yüksek fiyatlı/kıt tedarikçiler, uç termin (420 gün) ve veri kalitesi
#: tuzakları. En düşük ürün bedeli ile en düşük toplam maliyet **ayrı**
#: satırlara düşer; böylece iki işaretin ayrı hesaplandığı görülür.
SAMPLE_QUOTES: Final[tuple[DemoQuote, ...]] = (
    # --- Rekabetçi çekirdek: düşük fiyat, makul termin ---------------------- #
    DemoQuote(
        "Bursa Metal A.Ş.", "BM-2026-1187", 2.39, 28.50, 7_500.0,
        _SHIPPING_QUOTE, 15, "Net 30", 3, 45,
        "Nakliye dahil; birim fiyat KDV hariç yazılı teyitli.",
    ),
    DemoQuote(
        "Gemlik Profil Ltd.", "GP-8841", 2.31, 30.10, 5_200.0,
        _SHIPPING_QUOTE, 25, "%50 peşin, %50 teslimatta", 5, 30,
        "KDV hariç fiyat teyit edildi; karma ödeme koşulu.",
    ),
    DemoQuote(
        "Karabük Haddecilik A.Ş.", "KH-4471", 2.41, 27.35, 6_800.0,
        _SHIPPING_QUOTE, 12, "Net 30", 2, 40,
        "Hadde çıkışlı; sevkiyat 5 iş günü içinde.",
    ),
    DemoQuote(
        "Gebze Metal Profil", "GM-2026-339", 2.37, 26.95, 4_300.0,
        _SHIPPING_QUOTE, 10, "Net 45", 1, 35,
        "En kısa termin; stoktan sevkiyat. MOQ 500 adet.",
    ),
    DemoQuote(
        "Bilecik Çelik A.Ş.", "BC-6604", 2.35, 26.80, 5_200.0,
        _SHIPPING_QUOTE, 13, "Net 30", 4, 50,
        "Toplam maliyette en rekabetçi kombinasyon.",
    ),
    DemoQuote(
        "İnegöl Çelik San.", "IC-2026-77", 2.58, 27.90, 9_100.0,
        _SHIPPING_ASSUMPTION, 18, "Peşin", 4, 15,
        "Nakliye tahmini (gidiş-dönüş 420 km); fiyat KDV hariç. "
        f"Beyan edilen kg/m referans üstü ({WEIGHT_REFERENCE_MAX:.2f}).",
    ),
    # --- Orta grup: fiyat/termin dengesi ----------------------------------- #
    DemoQuote(
        "Kocaeli Çelik Tic.", "KC-9924", 2.44, 29.40, 3_900.0,
        _SHIPPING_QUOTE, 14, "Net 30", 3, 45,
        "Düşük navlun; bölgesel konsolidasyona uygun.",
    ),
    DemoQuote(
        "Ege Profil A.Ş.", "EP-5518", 2.36, 30.85, 8_700.0,
        _SHIPPING_QUOTE, 22, "Net 60", 6, 60,
        "İzmir çıkışlı; 60 gün vade.",
    ),
    DemoQuote(
        "Akdeniz Metal San.", "AM-2207", 2.42, 31.40, 11_200.0,
        _SHIPPING_QUOTE, 28, "Net 45", 7, 55,
        "Uzak mesafe navlunu; kısmi sevkiyat mümkün.",
    ),
    DemoQuote(
        "Anadolu Boru Profil", "AB-7781", 2.39, 28.15, 6_100.0,
        _SHIPPING_QUOTE, 17, "Net 30", 2, 42,
        "Standart galvaniz; KDV hariç fiyat teyitli.",
    ),
    DemoQuote(
        "Trakya Çelik Ltd.", "TC-3345", 2.47, 29.90, 5_400.0,
        _SHIPPING_QUOTE, 20, "30 gün", 5, 45,
        "İstenen terminle aynı gün sayısı; sınırda.",
    ),
    DemoQuote(
        "Samsun Metal A.Ş.", "SM-6620", 2.40, 27.60, 9_800.0,
        _SHIPPING_QUOTE, 26, "Peşin", 6, 30,
        "Peşin ödeme; nakit akışı üzerinde baskı oluşturur.",
    ),
    DemoQuote(
        "Konya Hadde San.", "KO-1156", 2.43, 27.95, 7_600.0,
        _SHIPPING_QUOTE, 24, "Net 30", 8, 50,
        "İç Anadolu çıkışlı; demiryolu ile konsolidasyon seçeneği.",
    ),
    DemoQuote(
        "İzmir Metal Tic.", "IM-7733", 2.40, 28.80, 6_900.0,
        _SHIPPING_QUOTE, 16, "Net 45", 4, 40,
        "45 gün vade; sevkiyat haftada iki kez.",
    ),
    DemoQuote(
        "Çorum Haddecilik", "CH-2205", 2.37, 29.10, 5_800.0,
        _SHIPPING_QUOTE, 19, "Net 30", 3, 38,
        "Orta ölçekli üretici; ilave kapasite teyidi alınmadı.",
    ),
    DemoQuote(
        "Tekirdağ Metal Ltd.", "TM-1184", 2.41, 28.05, 7_200.0,
        _SHIPPING_QUOTE, 15, "Net 30", 1, 36,
        "KDV hariç teyitli; MOQ 1.000 adet, bu sipariş için onaylandı.",
    ),
    DemoQuote(
        "Sakarya Metal", "SK-9917", 2.46, 31.90, 3_600.0,
        _SHIPPING_QUOTE, 21, "Net 45", 5, 44,
        "En düşük navlunlardan biri; birim fiyat yüksek.",
    ),
    DemoQuote(
        "Ankara Çelik Merkezi", "AC-4410", 2.45, 33.10, 4_800.0,
        _SHIPPING_QUOTE, 18, "Net 30", 4, 40,
        "Stoklu satıcı; fiyat primi belirgin.",
    ),
    # --- Yüksek fiyat / uzun termin --------------------------------------- #
    DemoQuote(
        "Mersin Profil Ltd.", "MP-8892", 2.38, 32.20, 12_400.0,
        _SHIPPING_QUOTE, 34, "Net 60", 7, 60,
        "İthal girdiye bağlı; kur riski tedarikçide değil.",
    ),
    DemoQuote(
        "Adana Profil San.", "AP-2288", 2.49, 34.60, 13_800.0,
        _SHIPPING_QUOTE, 38, "Net 60", 9, 65,
        "Yüksek fiyat ve navlun; acil ihtiyaç için uygun değil.",
    ),
    DemoQuote(
        "Denizli Profil Ltd.", "DP-3312", 2.42, 29.25, 8_300.0,
        _SHIPPING_ASSUMPTION, 27, "Net 30", 6, 35,
        "Nakliye tutarı tahminidir; iki hafta içinde teyit edilecek.",
    ),
    DemoQuote(
        "Manisa Çelik San.", "MC-5561", 2.39, 30.00, 0.0,
        _SHIPPING_QUOTE, 11, "Peşin", 2, 30,
        "Fabrika teslim (EXW); nakliye alıcıya aittir. Sıfır nakliye "
        "gerçek bir maliyet değildir.",
    ),
    DemoQuote(
        "Erzurum Metal San.", "ER-6602", 2.28, 27.35, 12_600.0,
        _SHIPPING_QUOTE, 44, "Net 60", 1, 28,
        "Ürün bedelinde en düşük; uzak mesafe navlunu toplam maliyeti "
        "yükseltiyor.",
    ),
    DemoQuote(
        "Yalova Çelik Tic.", "YC-2280", 2.16, 30.90, 5_700.0,
        _SHIPPING_QUOTE, 9, "Net 30", 1, 28,
        "Beyan edilen kg/m referans altında; ölçüm teyidi istenmeli.",
    ),
    DemoQuote(
        "Kocaeli Hadde A.Ş.", "KA-5573", 2.43, 30.25, 6_400.0,
        _SHIPPING_QUOTE, 23, "Net 30", 5, 46,
        "Termin taahhüdü yazılı değil.",
    ),
    DemoQuote(
        "Eskişehir Metal", "EM-8120", 2.44, 29.65, 6_000.0,
        _SHIPPING_QUOTE, 18, "Net 45", 3, 41,
        "Bölgesel tedarikçi; konsolidasyona uygun.",
    ),
    DemoQuote(
        "Zonguldak Profil", "ZP-3391", 2.71, 26.40, 10_400.0,
        _SHIPPING_QUOTE, 31, "Net 45", 8, 55,
        "En düşük birim fiyat; ancak kg/m beyanı referans üstünde ve "
        "navlun yüksek.",
    ),
    # --- Uç değerler ve veri kalitesi tuzakları --------------------------- #
    DemoQuote(
        "Sivas Demir Çelik", "SD-2219", 2.36, 27.10, 16_800.0,
        _SHIPPING_QUOTE, 420, "Net 30", 9, 60,
        "Fiyat cazip ancak termin 420 gün; planlamada kullanılamaz.",
    ),
    DemoQuote(
        "Van Metal Tic.", "VM-6628", 2.45, 34.90, None,
        _SHIPPING_QUOTE, 52, "Net 60", 7, 50,
        "Nakliye tutarı henüz bildirilmedi; toplam maliyet hesaplanamaz.",
    ),
    DemoQuote(
        "Rize Profil Ltd.", "RP-1188", 2.40, 31.25, 6_300.0,
        _SHIPPING_ASSUMPTION, 29, "Net 45", 4, 32,
        "",
    ),
    DemoQuote(
        "Kars Çelik San.", "KS-7741", 2.39, 30.40, 7_400.0,
        _SHIPPING_QUOTE, 20, "Net 30", 10, -2,
        "Teklif geçerlilik süresi doldu; fiyat teyidi yenilenmeli.",
    ),
    DemoQuote(
        "Hatay Metal A.Ş.", "HM-8840", 2.52, 35.40, 15_600.0,
        _SHIPPING_QUOTE, 42, "Net 60", 8, 58,
        "En yüksek birim fiyat; deprem sonrası lojistik kısıtı.",
    ),
    DemoQuote(
        "Gaziantep Çelik", "GC-4477", 2.48, 33.80, 14_200.0,
        _SHIPPING_QUOTE, 48, "Net 60", 9, 62,
        "Uzun termin ve yüksek navlun; toplam maliyeti belirgin yüksek.",
    ),
    DemoQuote(
        "Kayseri Metal A.Ş.", "KM-9985", 2.43, 32.75, 11_900.0,
        _SHIPPING_QUOTE, 36, "Net 60", 6, 48,
        "60 gün vade; fiyat primi yüksek.",
    ),
    DemoQuote(
        "Bolu Profil San.", "BP-7739", 2.44, 30.60, 6_200.0,
        _SHIPPING_QUOTE, 23, "%40 peşin, %60 teslimatta", 5, 37,
        "Karma ödeme; vade günü tek değere indirgenemez.",
    ),
)


def sample_supplier_count() -> int:
    """Örnek veri setindeki tedarikçi satırı sayısı."""
    return len(SAMPLE_QUOTES)


#: Boş şablonun ``Instructions`` sayfasına yazılan sütun kılavuzu.
#: Başlıklar :data:`procurement.excel_parser.SUPPLIER_QUOTE_HEADERS` ile
#: birebir aynıdır; kılavuz yalnızca insan içindir, parser bu satırları okumaz.
COLUMN_GUIDE: Final[tuple[tuple[str, str], ...]] = (
    (
        "Supplier Name",
        "Metin · zorunlu · tekil. Tedarikçinin tam adı (ör. Bursa Metal A.Ş.).",
    ),
    (
        "Supplier Quote Number",
        "Metin · zorunlu. Tedarikçinin teklif numarası (ör. BM-2026-1187).",
    ),
    (
        "Profile Specification",
        "Sabit: 40x40x2. Değiştirilmez.",
    ),
    (
        "Length (m)",
        "Sabit: 6. Sayı olarak yazılır.",
    ),
    (
        "Quantity (pcs)",
        "Pozitif tam sayı · zorunlu. RFQ!B9 ile aynı olmalı (ör. 500).",
    ),
    (
        "Weight (kg/m)",
        "Ondalık sayı · zorunlu. Referans aralık 2,27–2,51; dışı yalnızca "
        "uyarı üretir (ör. 2,39).",
    ),
    (
        "Unit Price (TRY/kg, excl. VAT)",
        "Pozitif sayı · zorunlu. KDV HARİÇ TL/kg (ör. 28,50).",
    ),
    (
        "Shipping Cost (TRY, excl. VAT)",
        "KDV hariç TL. Bilinmiyorsa BOŞ bırakın; 0 yazmayın (0 gerçek bir "
        "maliyettir).",
    ),
    (
        "Shipping Cost Source",
        "Supplier Quote veya Customer Assumption. Nakliye doluysa zorunlu.",
    ),
    (
        "Delivery Days",
        "Tam sayı (gün). Boş bırakılabilir; 365'in üzeri uyarı üretir.",
    ),
    (
        "Payment Terms",
        "Metin · zorunlu. Örn. Peşin, Net 30, %50 peşin, %50 teslimatta. "
        "Vade günü ayrıca yazılmaz.",
    ),
    (
        "Quote Date",
        "Tarih · zorunlu · yyyy-mm-dd. Gelecek tarih uyarı üretir.",
    ),
    (
        "Quote Valid Until",
        "Tarih · isteğe bağlı · yyyy-mm-dd. Teklif tarihinden önce olamaz.",
    ),
    (
        "Notes",
        "Serbest metin (en fazla 500). KDV teyidi veya nakliye varsayımının "
        "gerekçesi buraya yazılır.",
    ),
)


# --------------------------------------------------------------------------- #
# Yardımcılar
# --------------------------------------------------------------------------- #


def _load_template(template_path: Path):
    """Resmî şablonu tek okumayla belleğe alır ve düzenlenebilir kopyasını açar.

    Şablon baytları üzerinden okunur (yol üzerinden değil): böylece arayüzün
    şablon indirme yolu ile örnek/boş dosya üretimi **aynı** okuma hatasını
    paylaşır ve iki yerde farklı davranış oluşmaz.
    """
    try:
        payload = template_path.read_bytes()
    except OSError as error:
        raise DemoDataError(
            "Şablon dosyası okunamadığı için örnek/boş dosya üretilemedi."
        ) from error
    try:
        return load_workbook(filename=BytesIO(payload))
    except Exception as error:  # noqa: BLE001 - kullanıcıya tek mesaj gösterilir
        raise DemoDataError(
            "Şablon dosyası açılamadığı için örnek/boş dosya üretilemedi."
        ) from error


def _serialize(workbook) -> bytes:
    """Çalışma kitabını bellekte ``.xlsx`` baytlarına çevirir."""
    buffer = BytesIO()
    try:
        workbook.save(buffer)
    except Exception as error:  # noqa: BLE001
        raise DemoDataError(
            "Dosya oluşturulamadı. Lütfen yeniden deneyin."
        ) from error
    finally:
        workbook.close()
    return buffer.getvalue()


def _write_rfq(
    sheet,
    *,
    rfq_number: str | None,
    rfq_date: date | None,
    quantity: int | None,
    delivery_location: str | None,
    requested_days: int | None,
    description: str | None,
) -> None:
    """RFQ sayfasındaki kullanıcı alanlarını yazar; sabitleri teyit eder."""
    sheet["B3"] = rfq_number
    sheet["B4"] = (
        datetime.combine(rfq_date, time.min) if rfq_date is not None else None
    )
    # Sabit ürün kimliği alanları her iki dosyada da aynıdır.
    sheet["B5"] = EXPECTED_PRODUCT_CODE
    sheet["B6"] = EXPECTED_PROFILE_SPEC
    sheet["B7"] = EXPECTED_LENGTH_M
    sheet["B8"] = EXPECTED_WALL_THICKNESS_MM
    sheet["B9"] = quantity
    sheet["B10"] = delivery_location
    sheet["B11"] = requested_days
    sheet["B12"] = EXPECTED_CURRENCY
    sheet["B13"] = description


def _clear_data_rows(sheet) -> None:
    """``Supplier Quotes`` veri satırlarını boşaltır; biçimleri korur."""
    for row in sheet.iter_rows(
        min_row=_FIRST_DATA_ROW,
        max_row=_LAST_DATA_ROW,
        min_col=1,
        max_col=len(SUPPLIER_QUOTE_HEADERS),
    ):
        for cell in row:
            cell.value = None


def _write_quote_row(sheet, row: int, quote: DemoQuote, today: date) -> None:
    """Tek tedarikçi satırını şablonun sütun sırasına göre yazar."""
    quote_date = today - timedelta(days=quote.quote_days_ago)
    valid_until = today + timedelta(days=quote.valid_days)

    values = (
        quote.supplier,
        quote.quote_number,
        EXPECTED_PROFILE_SPEC,
        EXPECTED_LENGTH_M,
        SAMPLE_QUANTITY,
        quote.weight_kg_per_m,
        quote.unit_price,
        quote.shipping,
        quote.shipping_source,
        quote.delivery_days,
        quote.payment_terms,
        datetime.combine(quote_date, time.min),
        datetime.combine(valid_until, time.min),
        quote.notes or None,
    )
    for column, value in enumerate(values, start=1):
        sheet.cell(row=row, column=column, value=value)


def _bold_font(sheet, reference: str):
    """Verilen hücrenin yazı tipinin kalın kopyasını üretir.

    ``openpyxl`` yazı tipi nesneleri paylaşılabilir (immutable) olduğundan
    kopya alınır; aksi hâlde kalınlaştırma aynı stili kullanan diğer
    hücrelere de yansırdı.
    """
    style = copy(sheet[reference].font)
    style.bold = True
    return style


def _write_blank_instructions(sheet) -> None:
    """Boş şablonun ``Instructions`` sayfasına kılavuz bloğu ekler."""
    sheet["A18"] = (
        "Bu dosya BOŞ şablondur: Supplier Quotes sayfasında örnek satır yoktur. "
        "Doldurmaya 3. satırdan başlayın."
    )
    for row in range(19, 60):
        sheet.cell(row=row, column=1, value=None)
        sheet.cell(row=row, column=2, value=None)

    sheet["A20"] = "HIZLI BAŞLANGIÇ — SÜTUN KILAVUZU"
    sheet["A20"].font = _bold_font(sheet, "A1")
    sheet["A21"] = (
        "Supplier Quotes sayfasının 2. satırındaki 14 başlığı değiştirmeyin. "
        "Her tedarikçi için bir satır girin; veri 3. satırdan başlar. "
        "Tamamen boş satırlar yok sayılır."
    )
    sheet["A22"] = (
        "İşleme alınmak için Instructions!B4 (VAT Declaration) hücresini EVET "
        "yapın; KDV beyanı onaylanmadan dosya reddedilir (E-DECL-01)."
    )

    header_row = 24
    for column, label in ((1, "Sütun (birebir başlık)"), (2, "Tür / kural / örnek")):
        cell = sheet.cell(row=header_row, column=column, value=label)
        cell.font = _bold_font(sheet, "A2")

    for offset, (header, guidance) in enumerate(COLUMN_GUIDE, start=1):
        sheet.cell(row=header_row + offset, column=1, value=header)
        sheet.cell(row=header_row + offset, column=2, value=guidance)

    footer_row = header_row + len(COLUMN_GUIDE) + 2
    sheet.cell(
        row=footer_row,
        column=1,
        value=(
            "Dolu bir örnek görmek için ana sayfadaki "
            "'Örnek Dolu Dosyayı İndir' düğmesini kullanın."
        ),
    )


def _write_sample_instructions(sheet, today: date) -> None:
    """Örnek dosyanın ``Instructions`` sayfasına senaryo notlarını yazar."""
    sheet["A18"] = (
        f"Bu dosya ÖRNEK veri içerir ({len(SAMPLE_QUOTES)} tedarikçi teklifi). "
        "Tedarikçi adları, teklif numaraları ve tutarlar kurgusaldır; gerçek "
        "teklif yerine kullanılmaz. Tarihler dosyanın üretildiği güne göre "
        f"({today.isoformat()}) ayarlanmıştır."
    )
    for row in range(19, 60):
        sheet.cell(row=row, column=1, value=None)
        sheet.cell(row=row, column=2, value=None)

    notes = (
        ("Demo senaryoları",
         "Fiyat uçurumları (26,40–35,40 TL/kg), 9–420 gün arası termin "
         "dağılımı, 0–16.800 TL navlun, Peşin/Net 30–90 ve karma ödeme "
         "koşulları, süresi geçmiş teklif, eksik nakliye, gerekçesiz nakliye "
         "varsayımı ve referans dışı kg/m beyanları."),
        ("AI analizi",
         "Örnek dosya yüklendikten sonra 'AI analizi' görünümünde "
         "'AI analizini oluştur' düğmesine basın. Tedarikçi adları AI'a "
         "gönderilmez; yalnızca hesaplanmış değerler ve doğrulama bulguları "
         "yorumlanır."),
        ("Veri kuralı",
         "Tüm tutarlar KDV hariç TRY'dir. Nakliye bilinmiyorsa boş bırakılır; "
         "sıfır yazılmaz. Ödeme koşulları maliyet hesabına dahil edilmez."),
    )
    for offset, (label, text) in enumerate(notes):
        row = 20 + offset
        sheet.cell(row=row, column=1, value=label).font = _bold_font(sheet, "A1")
        sheet.cell(row=row, column=2, value=text)


# --------------------------------------------------------------------------- #
# Genel giriş noktaları
# --------------------------------------------------------------------------- #


def build_blank_template(template_path: str | Path, *, today: date | None = None) -> bytes:
    """Boş (veri satırı olmayan) teklif şablonunu üretir.

    Resmî şablonun kopyası kullanılır: sayfa adları, 14 sütun başlığı, stiller
    ve açılır liste doğrulamaları korunur. Örnek veri satırları silinir; RFQ
    sayfasında yalnızca kullanıcıya ait alanlar boş bırakılır, sabit ürün
    kimliği alanları korunur.

    Args:
        template_path: Paketlenmiş resmî ``.xlsx`` şablonunun yolu.
        today: Kullanılmayan, imza uyumu için tarih. (Kılavuz tarih taşımaz.)

    Returns:
        Bellekte üretilmiş ``.xlsx`` baytları.

    Raises:
        DemoDataError: Şablon açılamıyor veya dosya üretilemiyorsa.
    """
    workbook = _load_template(Path(template_path))
    try:
        instructions = workbook[SHEET_INSTRUCTIONS]
        quotes = workbook[SHEET_QUOTES]

        instructions[_VAT_CELL] = _VAT_NO
        instructions["B2"] = EXPECTED_TEMPLATE_VERSION
        instructions["B3"] = EXPECTED_TEMPLATE_CODE

        _write_rfq(
            workbook[SHEET_RFQ],
            rfq_number=None,
            rfq_date=None,
            quantity=None,
            delivery_location=None,
            requested_days=None,
            description=None,
        )
        _clear_data_rows(quotes)
        _write_blank_instructions(instructions)
        return _serialize(workbook)
    except DemoDataError:
        workbook.close()
        raise
    except Exception as error:  # noqa: BLE001
        workbook.close()
        raise DemoDataError(
            "Boş şablon hazırlanamadı. Lütfen yeniden deneyin."
        ) from error


def build_sample_workbook(template_path: str | Path, *, today: date | None = None) -> bytes:
    """AI demo için dolu örnek teklif dosyasını üretir.

    Tarihler ``today`` referansına göre üretilir; böylece örnek dosya
    doğrulamanın tarih kurallarına (``W-QD-01``, ``W-VAL-01``, ``W-RFQ-02``)
    takılmaz. Yalnızca kasıtlı senaryolar uyarı üretir; satır düzeyinde
    **hiçbir bloklayıcı hata (E- kodu) yoktur**, böylece 35 satırın tamamı
    hesaplanır.

    Args:
        template_path: Paketlenmiş resmî ``.xlsx`` şablonunun yolu.
        today: Örnek tarihlerin çapalandığı gün. Verilmezse sistem tarihi.

    Returns:
        Bellekte üretilmiş ``.xlsx`` baytları.

    Raises:
        DemoDataError: Şablon açılamıyor veya dosya üretilemiyorsa.
    """
    reference_day = today or date.today()
    workbook = _load_template(Path(template_path))
    try:
        instructions = workbook[SHEET_INSTRUCTIONS]
        quotes = workbook[SHEET_QUOTES]

        # Örnek dosya KDV hariç fiyatlarla üretildiği için beyan açıkça
        # işaretlenir; boş şablonda beyan bilinçli olarak HAYIR bırakılır.
        instructions[_VAT_CELL] = _VAT_YES
        instructions["B2"] = EXPECTED_TEMPLATE_VERSION
        instructions["B3"] = EXPECTED_TEMPLATE_CODE

        _write_rfq(
            workbook[SHEET_RFQ],
            rfq_number=SAMPLE_RFQ_NUMBER.format(year=reference_day.year),
            rfq_date=reference_day - timedelta(days=2),
            quantity=SAMPLE_QUANTITY,
            delivery_location=SAMPLE_DELIVERY_LOCATION,
            requested_days=SAMPLE_REQUESTED_DAYS,
            description=(
                "6 m boy, sıcak daldırma galvaniz; tüm fiyatlar KDV hariç. "
                "Örnek veri seti — gerçek teklif değildir."
            ),
        )

        _clear_data_rows(quotes)
        if len(SAMPLE_QUOTES) > _LAST_DATA_ROW - _FIRST_DATA_ROW + 1:
            raise DemoDataError(
                "Örnek veri seti şablonun satır sınırını aşıyor."
            )
        for offset, quote in enumerate(SAMPLE_QUOTES):
            _write_quote_row(quotes, _FIRST_DATA_ROW + offset, quote, reference_day)

        _write_sample_instructions(instructions, reference_day)
        return _serialize(workbook)
    except DemoDataError:
        workbook.close()
        raise
    except Exception as error:  # noqa: BLE001
        workbook.close()
        raise DemoDataError(
            "Örnek dosya hazırlanamadı. Lütfen yeniden deneyin."
        ) from error

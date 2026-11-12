"""Arayüz (``app.py``) testleri — Streamlit ``AppTest`` ile uçtan uca.

Kapsam
------
Arayüz gerçek akışı çalıştırır: yüklenen dosya → ``parse_workbook`` →
``validate`` → ``calculate`` → görünüm. Testler bu zincirin ekranda doğru
sonucu ürettiğini doğrular; hesaplama mantığını yeniden test etmez.

Testler hiçbir demo verisi üretmez: dosya olarak her zaman gerçek şablon
(``templates/supplier_quote_template.xlsx``) kullanılır ve yalnızca test
senaryosu için gereken hücre değiştirilir.
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest
from openpyxl import load_workbook
from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]
APP = ROOT / "app.py"
TEMPLATE = ROOT / "templates" / "supplier_quote_template.xlsx"

XLSX_MIME = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

pytestmark = pytest.mark.integration


# --------------------------------------------------------------------------- #
# Yardımcılar
# --------------------------------------------------------------------------- #


def _workbook_bytes(tmp_path: Path, mutate=None, *, vat: str = "EVET") -> bytes:
    """Şablonun kopyasını üretir ve bayt olarak döndürür."""
    path = tmp_path / "upload.xlsx"
    shutil.copyfile(TEMPLATE, path)

    workbook = load_workbook(path)
    try:
        workbook["Instructions"]["B4"] = vat
        if mutate is not None:
            mutate(workbook)
    finally:
        workbook.save(path)
        workbook.close()

    return path.read_bytes()


def _app() -> AppTest:
    test = AppTest.from_file(str(APP), default_timeout=90)
    test.run()
    assert not test.exception, [str(e) for e in test.exception]
    return test


def _upload(test: AppTest, payload: bytes, *, name: str = "teklif.xlsx") -> AppTest:
    test.file_uploader[0].set_value((name, payload, XLSX_MIME))
    test.run()
    assert not test.exception, [str(e) for e in test.exception]
    return test


def _markdown(test: AppTest) -> str:
    """Ekrandaki tüm markdown/HTML bloklarını tek metinde birleştirir."""
    return "\n".join(block.value for block in test.markdown)


def _table_html(test: AppTest) -> str:
    """Karşılaştırma tablosunun HTML'ini döndürür; yoksa boş metin."""
    for block in test.markdown:
        if "<table" in block.value:
            return block.value
    return ""


def _set(sheet: str, cell: str, value):
    def mutate(workbook) -> None:
        workbook[sheet][cell] = value

    return mutate


# --------------------------------------------------------------------------- #
# Başlangıç durumu
# --------------------------------------------------------------------------- #


def test_app_starts_without_file():
    test = _app()

    assert len(test.file_uploader) == 1
    assert test.file_uploader[0].allowed_type == [".xlsx"]
    text = _markdown(test)
    assert "Henüz dosya yüklenmedi" in text
    assert _table_html(test) == ""


def test_initial_screen_does_not_show_costs():
    """Dosya yokken hiçbir tablo, grafik veya tedarikçi bilgisi gösterilmez."""
    test = _app()

    assert len(test.get("plotly_chart")) == 0
    assert "Genel özet" not in _markdown(test)
    assert "Bursa Metal" not in _markdown(test)


# --------------------------------------------------------------------------- #
# Mutlu yol — gerçek şablon
# --------------------------------------------------------------------------- #


def test_upload_renders_kpi_cards(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path))
    text = _markdown(test)

    assert "Genel özet" in text
    assert "Tedarikçi sayısı" in text
    assert "Hesaplanabilen teklif" in text
    assert "204.345,00 TL" in text
    assert "211.845,00 TL" in text


def test_upload_renders_comparison_table(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path))
    table = _table_html(test)

    assert table
    for name in ("Bursa Metal A.Ş.", "Gemlik Profil Ltd.", "İnegöl Çelik San."):
        assert name in table


def test_table_has_every_required_column(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path))
    table = _table_html(test)

    for header in (
        "Tedarikçi",
        "Miktar (adet)",
        "Toplam ağırlık (kg)",
        "Birim fiyat (TL/kg)",
        "Ürün bedeli (TL)",
        "Nakliye (TL)",
        "Toplam tedarik maliyeti (TL)",
        "Nakliye (TL/kg)",
        "Termin (gün)",
        "Ödeme koşulları",
    ):
        assert header in table


def test_table_shows_calculated_values(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path))
    table = _table_html(test)

    # 500 adet × 6 m × 2,39 kg/m = 7170 kg; × 28,50 = 204.345,00
    assert "7.170,000" in table
    assert "204.345,00" in table
    assert "7.500,00" in table
    assert "211.845,00" in table
    assert "1,0460" in table


def test_table_shows_payment_terms_verbatim(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path))

    assert "%50 peşin, %50 teslimatta" in _table_html(test)


def test_lowest_values_are_marked(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path))
    table = _table_html(test)

    assert "En düşük" in table
    assert "is-lowest" in table


def test_upload_renders_both_charts(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path))
    charts = test.get("plotly_chart")

    assert len(charts) == 2


def test_template_warnings_are_shown_as_warnings(tmp_path):
    """Şablonun kendi iki uyarısı ayrı bölümde gösterilir; hata paneli yok."""
    test = _upload(_app(), _workbook_bytes(tmp_path))
    text = _markdown(test)

    assert "W-PAY-01" in text
    assert "W-WT-01" in text
    assert "<h3>Uyarılar</h3>" in text
    assert "<h3>Hatalar</h3>" not in text


def test_warning_mentions_supplier_and_cell(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path))
    text = _markdown(test)

    assert "Supplier Quotes!F5" in text
    assert "İnegöl Çelik San." in text


# --------------------------------------------------------------------------- #
# Eksik nakliye
# --------------------------------------------------------------------------- #


def test_missing_shipping_is_shown_as_missing_data(tmp_path):
    test = _upload(
        _app(), _workbook_bytes(tmp_path, _set("Supplier Quotes", "H3", None))
    )
    table = _table_html(test)
    row_html = table.split("Bursa Metal A.Ş.")[1].split("</tr>")[0]

    assert "Eksik veri" in row_html
    # Nakliye ve toplam maliyet hücreleri sıfır göstermez.
    assert row_html.count("Eksik veri") == 3


def test_missing_shipping_keeps_product_cost(tmp_path):
    test = _upload(
        _app(), _workbook_bytes(tmp_path, _set("Supplier Quotes", "H3", None))
    )
    table = _table_html(test)

    assert "204.345,00" in table


def test_missing_shipping_row_is_tagged(tmp_path):
    test = _upload(
        _app(), _workbook_bytes(tmp_path, _set("Supplier Quotes", "H3", None))
    )

    assert "Nakliye bilgisi eksik" in _table_html(test)


def test_missing_shipping_is_reported_as_info(tmp_path):
    test = _upload(
        _app(), _workbook_bytes(tmp_path, _set("Supplier Quotes", "H3", None))
    )
    text = _markdown(test)

    assert "Bilgi talepleri" in text
    assert "I-SHIP-01" in text


def test_missing_shipping_excludes_row_from_total_chart(tmp_path):
    """Toplam maliyeti hesaplanamayan satır toplam grafiğine girmez."""
    test = _upload(
        _app(), _workbook_bytes(tmp_path, _set("Supplier Quotes", "H3", None))
    )
    text = _markdown(test)

    assert "1 satır eksik nakliye bilgisi nedeniyle grafiğe alınmadı" in text


def test_missing_shipping_moves_lowest_total(tmp_path):
    test = _upload(
        _app(), _workbook_bytes(tmp_path, _set("Supplier Quotes", "H3", None))
    )
    text = _markdown(test)

    # Satır 3 toplam maliyette yarışamaz; en düşük toplam satır 4'tür.
    assert "213.793,00 TL" in text


def test_explicit_zero_shipping_is_complete(tmp_path):
    test = _upload(
        _app(), _workbook_bytes(tmp_path, _set("Supplier Quotes", "H3", 0))
    )
    text = _markdown(test)

    assert "Eksik veri" not in _table_html(test)
    assert "Bilgi talepleri" not in text


# --------------------------------------------------------------------------- #
# Dosya reddi
# --------------------------------------------------------------------------- #


def test_missing_vat_declaration_rejects_file(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path, vat="HAYIR"))
    text = _markdown(test)

    assert "Dosya reddedildi" in text
    assert "E-DECL-01" in text


def test_rejection_explains_how_to_fix(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path, vat="HAYIR"))
    text = _markdown(test)

    assert "Nasıl düzeltilir" in text
    assert "EVET" in text


def test_rejection_shows_no_comparison(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path, vat="HAYIR"))

    assert _table_html(test) == ""
    assert len(test.get("plotly_chart")) == 0
    assert "Genel özet" not in _markdown(test)


def test_unsupported_currency_rejects_file(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path, _set("RFQ", "B12", "USD")))
    text = _markdown(test)

    assert "Dosya reddedildi" in text
    assert "E-CUR-01" in text
    assert "TRY" in text


def test_corrupt_file_shows_readable_error(tmp_path):
    test = _upload(_app(), b"bu bir excel dosyasi degil", name="bozuk.xlsx")
    text = _markdown(test)

    assert "Dosya okunamadı" in text
    assert _table_html(test) == ""


# --------------------------------------------------------------------------- #
# Satır düzeyi hatalar
# --------------------------------------------------------------------------- #


def test_bad_row_is_excluded_and_others_computed(tmp_path):
    test = _upload(
        _app(), _workbook_bytes(tmp_path, _set("Supplier Quotes", "G3", 0))
    )
    text = _markdown(test)
    table = _table_html(test)

    assert "E-PRICE-01" in text
    assert "Karşılaştırmaya alınmayan satırlar" in text
    assert "Bursa Metal A.Ş." not in table
    assert "Gemlik Profil Ltd." in table


def test_bad_row_keeps_kpi_count_consistent(tmp_path):
    test = _upload(
        _app(), _workbook_bytes(tmp_path, _set("Supplier Quotes", "G3", 0))
    )
    text = _markdown(test)

    assert "1 satır karşılaştırmaya alınmadı" in text


def test_all_rows_bad_shows_empty_state(tmp_path):
    def mutate(workbook) -> None:
        sheet = workbook["Supplier Quotes"]
        for row in (3, 4, 5):
            sheet[f"G{row}"] = 0

    test = _upload(_app(), _workbook_bytes(tmp_path, mutate))
    text = _markdown(test)

    assert "Karşılaştırılabilir teklif yok" in text
    assert _table_html(test) == ""
    assert len(test.get("plotly_chart")) == 0


def test_all_rows_empty_shows_clear_empty_state(tmp_path):
    """Teklif satırı yoksa 'karşılaştırılabilir teklif yok' değil, satır yok denir."""
    def mutate(workbook) -> None:
        sheet = workbook["Supplier Quotes"]
        for row in range(3, 6):
            for column in "ABCDEFGHIJKLMN":
                sheet[f"{column}{row}"] = None

    test = _upload(_app(), _workbook_bytes(tmp_path, mutate))
    text = _markdown(test)

    assert "Teklif satırı bulunamadı" in text
    assert _table_html(test) == ""
    assert len(test.get("plotly_chart")) == 0


# --------------------------------------------------------------------------- #
# Arayüz kuralları
# --------------------------------------------------------------------------- #


def test_app_does_not_offer_ranking_or_winner(tmp_path):
    """Arayüz kazanan ilan etmez veya puan göstermez."""
    test = _upload(_app(), _workbook_bytes(tmp_path))
    text = _markdown(test).lower()

    for forbidden in ("kazanan:", "puan", "skor", "sıralama:", "önerilen tedarikçi"):
        assert forbidden not in text


def test_app_states_no_vat_and_no_automatic_winner(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path))
    text = _markdown(test)

    assert "KDV hesaplanmaz" in text
    assert "Kazanan otomatik belirlenmez" in text


def test_uploaded_filename_is_displayed(tmp_path):
    test = _upload(
        _app(), _workbook_bytes(tmp_path), name="kasim-teklifleri.xlsx"
    )

    assert "kasim-teklifleri.xlsx" in _markdown(test)


def test_file_text_is_html_escaped(tmp_path):
    """Dosyadan gelen metin HTML olarak yorumlanmaz."""
    def mutate(workbook) -> None:
        sheet = workbook["Supplier Quotes"]
        sheet["A3"] = "<img src=x onerror=alert(1)>"
        sheet["K3"] = "<b>Peşin</b>"

    test = _upload(_app(), _workbook_bytes(tmp_path, mutate))
    text = _markdown(test)

    assert "<img src=x" not in text
    assert "&lt;img src=x" in text
    assert "<b>Peşin</b>" not in text


def test_uploaded_filename_is_escaped(tmp_path):
    test = _upload(
        _app(), _workbook_bytes(tmp_path), name="<b>teklif</b>.xlsx"
    )
    text = _markdown(test)

    assert "<b>teklif</b>" not in text
    assert "&lt;b&gt;teklif&lt;/b&gt;.xlsx" in text


def test_upload_does_not_write_files(tmp_path):
    """Yükleme sonrası çalışma dizininde yeni dosya oluşmaz."""
    before = {path.name for path in ROOT.iterdir()}
    _upload(_app(), _workbook_bytes(tmp_path))
    after = {path.name for path in ROOT.iterdir()}

    assert after == before


# --------------------------------------------------------------------------- #
# Biçimsel sınır durumları
# --------------------------------------------------------------------------- #


def test_text_quantity_in_rfq_does_not_break_layout(tmp_path):
    """Validator'ın kabul ettiği ``500,0`` gibi metin miktar arayüzü düşürmez."""
    test = _upload(
        _app(), _workbook_bytes(tmp_path, _set("RFQ", "B9", "500,0"))
    )
    text = _markdown(test)

    assert "500" in text
    assert _table_html(test)
    assert len(test.get("plotly_chart")) == 2

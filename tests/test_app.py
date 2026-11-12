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

import html
import json
import shutil
from pathlib import Path

import pytest
from openpyxl import load_workbook
from streamlit.testing.v1 import AppTest

from procurement import ai_analysis

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


def _visible_markdown(test: AppTest) -> str:
    """``_markdown`` gibi çalışır; global ``<style>`` bloğunu dışarıda bırakır.

    CSS sınıf adları da metinde aranabildiği için görünüm içeriği testleri
    stil tanımlarından arındırılmış metin üzerinde çalışır.
    """
    return "\n".join(
        block.value
        for block in test.markdown
        if not block.value.lstrip().startswith("<style>")
    )


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


DASHBOARD = "Dashboard"
COMPARISON = "Tedarikçi karşılaştırması"
AI_VIEW = "AI analizi"
ROADMAP = "Product Roadmap"


def _view(test: AppTest, label: str) -> AppTest:
    """Görünümü değiştirir; dashboard varsayılan görünümdür."""
    test.segmented_control[0].set_value(label)
    test.run()
    assert not test.exception, [str(e) for e in test.exception]
    return test


def _comparison(test: AppTest) -> AppTest:
    """Karşılaştırma görünümüne geçer."""
    return _view(test, COMPARISON)


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
# Şablon indirme
# --------------------------------------------------------------------------- #


def test_template_download_is_available_on_start_screen():
    """Şablon, dosya yüklenmeden indirilebilir; rapor butonlarıyla karışmaz."""
    test = _app()
    labels = [button.label for button in test.download_button]

    assert labels == ["Excel şablonunu indir"]
    assert "Excel şablonunu indir" in _markdown(test)
    assert "Nasıl Çalışır?" in _markdown(test)


def test_template_download_points_to_xlsx_file():
    test = _app()
    button = test.download_button[0]

    assert button.proto.url.endswith(".xlsx")
    assert button.proto.id.endswith("download-template")


def test_template_download_serves_the_real_template_file(monkeypatch):
    """İndirilen içerik repodaki şablon dosyasının birebir kendisidir."""
    import streamlit as st_module

    captured: dict[str, object] = {}
    original = st_module.download_button

    def spy(label, data=None, **kwargs):
        if kwargs.get("key") == "download-template":
            captured["label"] = label
            captured["data"] = data
            captured["file_name"] = kwargs.get("file_name")
            captured["mime"] = kwargs.get("mime")
        return original(label, data=data, **kwargs)

    monkeypatch.setattr(st_module, "download_button", spy)

    test = _app()

    assert not test.exception, [str(e) for e in test.exception]
    assert captured["data"] == TEMPLATE.read_bytes()
    assert str(captured["file_name"]).endswith(".xlsx")
    assert captured["mime"] == XLSX_MIME


def test_missing_template_shows_readable_error(monkeypatch):
    """Şablon dosyası okunamazsa kullanıcıya anlaşılır mesaj gösterilir."""
    from pathlib import Path as _Path

    original = _Path.read_bytes

    def failing_read(self):
        if self.name == "supplier_quote_template.xlsx":
            raise OSError("disk hatası: C:/gizli/yol")
        return original(self)

    monkeypatch.setattr(_Path, "read_bytes", failing_read)

    test = _app()
    text = _markdown(test)

    assert "Şablon indirilemiyor" in text
    assert "C:/gizli/yol" not in text
    assert len(test.download_button) == 0


def test_uploaded_file_replaces_template_download(tmp_path):
    """Dosya yüklendikten sonra yalnızca rapor indirmeleri görünür."""
    test = _upload(_app(), _workbook_bytes(tmp_path))
    labels = [button.label for button in test.download_button]

    assert labels == ["Excel raporunu indir", "PDF raporunu indir"]


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
    test = _comparison(_upload(_app(), _workbook_bytes(tmp_path)))
    table = _table_html(test)

    assert table
    for name in ("Bursa Metal A.Ş.", "Gemlik Profil Ltd.", "İnegöl Çelik San."):
        assert name in table


def test_table_has_every_required_column(tmp_path):
    test = _comparison(_upload(_app(), _workbook_bytes(tmp_path)))
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
    test = _comparison(_upload(_app(), _workbook_bytes(tmp_path)))
    table = _table_html(test)

    # 500 adet × 6 m × 2,39 kg/m = 7170 kg; × 28,50 = 204.345,00
    assert "7.170,000" in table
    assert "204.345,00" in table
    assert "7.500,00" in table
    assert "211.845,00" in table
    assert "1,0460" in table


def test_table_shows_payment_terms_verbatim(tmp_path):
    test = _comparison(_upload(_app(), _workbook_bytes(tmp_path)))

    assert "%50 peşin, %50 teslimatta" in _table_html(test)


def test_lowest_values_are_marked(tmp_path):
    test = _comparison(_upload(_app(), _workbook_bytes(tmp_path)))
    table = _table_html(test)

    assert "En düşük" in table
    assert "is-lowest" in table


def test_upload_renders_both_charts(tmp_path):
    test = _comparison(_upload(_app(), _workbook_bytes(tmp_path)))
    charts = test.get("plotly_chart")

    assert len(charts) == 2


def test_template_warnings_are_shown_as_warnings(tmp_path):
    """Şablonun kendi iki uyarısı ayrı bölümde gösterilir; hata paneli yok."""
    test = _comparison(_upload(_app(), _workbook_bytes(tmp_path)))
    text = _markdown(test)

    assert "W-PAY-01" in text
    assert "W-WT-01" in text
    assert "<h3>Uyarılar</h3>" in text
    assert "<h3>Hatalar</h3>" not in text


def test_warning_mentions_supplier_and_cell(tmp_path):
    test = _comparison(_upload(_app(), _workbook_bytes(tmp_path)))
    text = _markdown(test)

    assert "Supplier Quotes!F5" in text
    assert "İnegöl Çelik San." in text


# --------------------------------------------------------------------------- #
# Eksik nakliye
# --------------------------------------------------------------------------- #


def test_missing_shipping_is_shown_as_missing_data(tmp_path):
    test = _comparison(
        _upload(
            _app(), _workbook_bytes(tmp_path, _set("Supplier Quotes", "H3", None))
        )
    )
    table = _table_html(test)
    row_html = table.split("Bursa Metal A.Ş.")[1].split("</tr>")[0]

    assert "Eksik veri" in row_html
    # Nakliye ve toplam maliyet hücreleri sıfır göstermez.
    assert row_html.count("Eksik veri") == 3


def test_missing_shipping_keeps_product_cost(tmp_path):
    test = _comparison(
        _upload(
            _app(), _workbook_bytes(tmp_path, _set("Supplier Quotes", "H3", None))
        )
    )
    table = _table_html(test)

    assert "204.345,00" in table


def test_missing_shipping_row_is_tagged(tmp_path):
    test = _comparison(
        _upload(
            _app(), _workbook_bytes(tmp_path, _set("Supplier Quotes", "H3", None))
        )
    )

    assert "Nakliye bilgisi eksik" in _table_html(test)


def test_missing_shipping_is_reported_as_info(tmp_path):
    test = _comparison(
        _upload(
            _app(), _workbook_bytes(tmp_path, _set("Supplier Quotes", "H3", None))
        )
    )
    text = _markdown(test)

    assert "Bilgi talepleri" in text
    assert "I-SHIP-01" in text


def test_missing_shipping_excludes_row_from_total_chart(tmp_path):
    """Toplam maliyeti hesaplanamayan satır toplam grafiğine girmez."""
    test = _comparison(
        _upload(
            _app(), _workbook_bytes(tmp_path, _set("Supplier Quotes", "H3", None))
        )
    )
    text = _markdown(test)

    assert "1 satır eksik nakliye bilgisi nedeniyle grafiğe alınmadı" in text


def test_missing_shipping_moves_lowest_total(tmp_path):
    test = _comparison(
        _upload(
            _app(), _workbook_bytes(tmp_path, _set("Supplier Quotes", "H3", None))
        )
    )
    text = _markdown(test)

    # Satır 3 toplam maliyette yarışamaz; en düşük toplam satır 4'tür.
    assert "213.793,00 TL" in text


def test_explicit_zero_shipping_is_complete(tmp_path):
    test = _comparison(
        _upload(_app(), _workbook_bytes(tmp_path, _set("Supplier Quotes", "H3", 0)))
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
    test = _comparison(
        _upload(_app(), _workbook_bytes(tmp_path, _set("Supplier Quotes", "G3", 0)))
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
# Dashboard
# --------------------------------------------------------------------------- #


def test_dashboard_is_default_view(tmp_path):
    """Dosya yüklenince varsayılan görünüm Dashboard'dır."""
    test = _upload(_app(), _workbook_bytes(tmp_path))

    assert test.segmented_control[0].value == DASHBOARD
    assert _table_html(test) == ""  # karşılaştırma tablosu dashboard'da yok


def test_dashboard_offers_navigation_between_views(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path))
    options = test.segmented_control[0].options

    assert list(options) == [DASHBOARD, COMPARISON, AI_VIEW, ROADMAP]


def test_dashboard_shows_four_kpi_cards(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path))
    text = _markdown(test)

    assert "Genel özet" in text
    for label in (
        "Tedarikçi sayısı",
        "Hesaplanabilen teklif",
        "En düşük ürün bedeli",
        "En düşük toplam tedarik maliyeti",
    ):
        assert label in text


def test_dashboard_kpi_values_come_from_calculator(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path))
    text = _markdown(test)

    assert "204.345,00 TL" in text
    assert "211.845,00 TL" in text
    assert "Bursa Metal A.Ş." in text


def test_dashboard_renders_one_cost_chart(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path))
    charts = test.get("plotly_chart")

    assert len(charts) == 1
    assert "Tedarikçi bazında maliyet" in _markdown(test)


def test_dashboard_chart_is_total_cost_when_available(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path))
    spec = json.loads(test.get("plotly_chart")[0].proto.spec)

    assert spec["data"][0]["x"] == [211845.0, 213793.0, 225046.0]
    assert "Toplam tedarik maliyeti (TL)" in _markdown(test)


def test_dashboard_chart_falls_back_to_product_cost(tmp_path):
    """Hiçbir satırda toplam hesaplanamıyorsa ürün bedeli grafiği çizilir."""

    def mutate(workbook) -> None:
        sheet = workbook["Supplier Quotes"]
        for row in (3, 4, 5):
            sheet[f"H{row}"] = None

    test = _upload(_app(), _workbook_bytes(tmp_path, mutate))
    spec = json.loads(test.get("plotly_chart")[0].proto.spec)

    assert len(spec["data"][0]["x"]) == 3
    assert "Ürün bedeli (TL)" in _markdown(test)


def test_dashboard_shows_analysis_summary(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path))
    text = _markdown(test)

    assert "Analiz özeti" in text
    assert "Son yüklenen teklif dosyası" in text
    assert "teklif.xlsx" in text


def test_dashboard_summary_reports_file_and_validation_counts(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path))
    text = _markdown(test)

    assert "RFQ-2026-0042" in text
    assert "40x40x2" in text
    assert "500 adet" in text
    assert "Tedarikçi satırı" in text
    assert "Bilgi talepleri" in text


def test_dashboard_summary_status_for_clean_template(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path))
    text = _markdown(test)

    # Şablonun iki uyarısı vardır; durum "uyarılarla tamamlandı" olur.
    assert "Uyarılarla tamamlandı" in text
    assert "status-warn" in text


def test_dashboard_chart_note_reports_omitted_rows(tmp_path):
    test = _upload(
        _app(), _workbook_bytes(tmp_path, _set("Supplier Quotes", "H3", None))
    )
    text = _markdown(test)

    assert "1 satır eksik nakliye bilgisi nedeniyle grafiğe alınmadı" in text


def test_dashboard_summary_flags_missing_shipping(tmp_path):
    test = _upload(
        _app(), _workbook_bytes(tmp_path, _set("Supplier Quotes", "H3", None))
    )
    text = _markdown(test)

    assert "Eksik veri var" in text
    assert "Eksik nakliye" in text


def test_dashboard_summary_flags_partial_comparison(tmp_path):
    test = _upload(
        _app(), _workbook_bytes(tmp_path, _set("Supplier Quotes", "G3", 0))
    )
    text = _markdown(test)

    assert "Kısmi karşılaştırma" in text
    assert "status-error" in text


def test_dashboard_empty_file_shows_no_chart(tmp_path):
    def mutate(workbook) -> None:
        sheet = workbook["Supplier Quotes"]
        for row in range(3, 6):
            for column in "ABCDEFGHIJKLMN":
                sheet[f"{column}{row}"] = None

    test = _upload(_app(), _workbook_bytes(tmp_path, mutate))
    text = _markdown(test)

    assert "Teklif satırı bulunamadı" in text
    assert len(test.get("plotly_chart")) == 0
    assert "Teklif satırı yok" in text


def test_rejected_file_shows_no_dashboard(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path, vat="HAYIR"))
    text = _markdown(test)

    assert "Dosya reddedildi" in text
    assert "Genel özet" not in text
    assert "Analiz özeti" not in text
    assert len(test.segmented_control) == 0


def test_dashboard_does_not_repeat_comparison_table(tmp_path):
    """Dashboard'da tablo ve doğrulama bulgu panelleri gösterilmez."""
    test = _upload(_app(), _workbook_bytes(tmp_path))
    text = _markdown(test)

    assert _table_html(test) == ""
    assert "<h3>Uyarılar</h3>" not in text
    assert "<h3>Bilgi talepleri</h3>" not in text


def test_switching_back_to_dashboard_keeps_kpis(tmp_path):
    test = _comparison(_upload(_app(), _workbook_bytes(tmp_path)))
    assert _table_html(test)

    test = _view(test, DASHBOARD)
    assert _table_html(test) == ""
    assert "Genel özet" in _markdown(test)


# --------------------------------------------------------------------------- #
# Product Roadmap görünümü
# --------------------------------------------------------------------------- #


def _roadmap(test: AppTest) -> AppTest:
    """Product Roadmap görünümüne geçer."""
    return _view(test, ROADMAP)


def test_roadmap_view_is_reachable(tmp_path):
    test = _roadmap(_upload(_app(), _workbook_bytes(tmp_path)))
    text = _visible_markdown(test)

    assert test.segmented_control[0].value == ROADMAP
    assert "Product Roadmap" in text
    assert "roadmap-track" in text


def test_roadmap_shows_three_stages_with_titles_and_status(tmp_path):
    test = _roadmap(_upload(_app(), _workbook_bytes(tmp_path)))
    text = _visible_markdown(test)

    for phase, title, status in (
        ("V1", "Core Procurement Platform", "LIVE"),
        ("V2", "Procurement Intelligence", "PLANNED"),
        ("V3", "Enterprise Procurement", "FUTURE"),
    ):
        assert phase in text
        assert title in text
        assert status in text

    assert text.count('class="roadmap-stage') == 3


def test_roadmap_v1_lists_all_shipped_capabilities(tmp_path):
    test = _roadmap(_upload(_app(), _workbook_bytes(tmp_path)))
    text = _visible_markdown(test)

    for feature in (
        "Excel Supplier Quote Upload",
        "Automated Data Validation",
        "Procurement Cost Calculation",
        "Supplier Comparison Dashboard",
        "AI-Powered Procurement Analysis",
        "Excel &amp; PDF Reporting",
    ):
        assert feature in text


def test_roadmap_v2_and_v3_list_planned_capabilities(tmp_path):
    test = _roadmap(_upload(_app(), _workbook_bytes(tmp_path)))
    text = _visible_markdown(test)

    for feature in (
        "Advanced Product Specifications",
        "AI Price Forecasting",
        "Logistics Optimization",
        "Advanced Cost Analysis",
        "Supplier Portal",
        "ERP / SAP Integrations",
        "Persistent Analysis History",
        "Multi-Currency Support",
    ):
        assert feature in text


def test_roadmap_marks_only_v1_as_live(tmp_path):
    test = _roadmap(_upload(_app(), _workbook_bytes(tmp_path)))
    text = _visible_markdown(test)

    assert 'class="roadmap-stage is-live"' in text
    assert 'class="roadmap-stage is-planned"' in text
    assert 'class="roadmap-stage is-future"' in text
    assert text.count("is-live") == 2  # kart + durum rozeti
    assert "henüz kullanıma açık değildir" in text


def test_roadmap_does_not_render_analysis_or_comparison(tmp_path):
    """Roadmap sunumdur: tablo, KPI veya grafik üretmez."""
    test = _roadmap(_upload(_app(), _workbook_bytes(tmp_path)))
    text = _visible_markdown(test)

    assert _table_html(test) == ""
    assert len(test.get("plotly_chart")) == 0
    assert "Genel özet" not in text


def test_roadmap_reports_remain_available(tmp_path):
    """Mevcut akış korunur: rapor indirmeleri roadmap görünümünde de durur."""
    test = _roadmap(_upload(_app(), _workbook_bytes(tmp_path)))
    labels = [button.label for button in test.download_button]

    assert labels == ["Excel raporunu indir", "PDF raporunu indir"]


def test_switching_from_roadmap_back_to_dashboard(tmp_path):
    test = _roadmap(_upload(_app(), _workbook_bytes(tmp_path)))
    assert "roadmap-track" in _visible_markdown(test)

    test = _view(test, DASHBOARD)
    text = _visible_markdown(test)

    assert "roadmap-track" not in text
    assert "Genel özet" in text


# --------------------------------------------------------------------------- #
# AI analizi görünümü
# --------------------------------------------------------------------------- #


def _ai(test: AppTest) -> AppTest:
    """AI analizi görünümüne geçer."""
    return _view(test, AI_VIEW)


def _without_api_key(monkeypatch) -> None:
    """Ortamdan API anahtarını kaldırır; .env okumasını da kapatır."""
    monkeypatch.delenv(ai_analysis.API_KEY_ENV, raising=False)
    monkeypatch.setattr(
        "procurement.ai_analysis._load_dotenv", lambda: None
    )


def test_ai_view_is_reachable_and_asks_for_explicit_action(tmp_path):
    test = _ai(_upload(_app(), _workbook_bytes(tmp_path)))
    text = _markdown(test)

    assert test.segmented_control[0].value == AI_VIEW
    assert "AI analizi" in text
    assert "Executive Summary" not in text  # kullanıcı butona basmadan çağrı yok


def test_ai_view_has_run_button(tmp_path, monkeypatch):
    monkeypatch.setenv(ai_analysis.API_KEY_ENV, "test-anahtar")
    test = _ai(_upload(_app(), _workbook_bytes(tmp_path)))

    assert len(test.button) == 1
    assert test.button[0].label == "AI analizini oluştur"


def test_ai_view_without_api_key_shows_config_notice(tmp_path, monkeypatch):
    _without_api_key(monkeypatch)
    test = _ai(_upload(_app(), _workbook_bytes(tmp_path)))
    text = _markdown(test)

    assert "API anahtarı bulunamadı" in text
    assert ai_analysis.API_KEY_ENV in text
    assert len(test.button) == 0


def test_ai_view_does_not_call_api_before_button(tmp_path, monkeypatch):
    _without_api_key(monkeypatch)
    calls: list[int] = []
    monkeypatch.setattr(
        ai_analysis, "analyze", lambda *a, **k: calls.append(1)
    )
    _ai(_upload(_app(), _workbook_bytes(tmp_path)))

    assert calls == []


def test_ai_view_rejected_file_has_no_ai_section(tmp_path, monkeypatch):
    _without_api_key(monkeypatch)
    test = _upload(_app(), _workbook_bytes(tmp_path, vat="HAYIR"))

    assert len(test.segmented_control) == 0
    assert "AI analizi" not in _markdown(test)
    assert "Satın alma değerlendirmesi" not in _markdown(test)


def test_ai_view_reports_service_error(tmp_path, monkeypatch):
    monkeypatch.setenv(ai_analysis.API_KEY_ENV, "test-anahtar")

    def failing(*args, **kwargs):
        raise ai_analysis.AIAnalysisError(
            "AI servisi zaman aşımına uğradı. Lütfen tekrar deneyin.",
            kind="timeout",
        )

    monkeypatch.setattr(ai_analysis, "analyze", failing)
    test = _ai(_upload(_app(), _workbook_bytes(tmp_path)))
    test.button[0].click()
    test.run()
    assert not test.exception, [str(e) for e in test.exception]
    text = _markdown(test)

    assert "Analiz üretilemedi" in text
    assert "zaman aşımı" in text
    assert "is-error" in text


def test_ai_view_renders_sections_after_run(tmp_path, monkeypatch):
    monkeypatch.setenv(ai_analysis.API_KEY_ENV, "test-anahtar")

    analysis = ai_analysis.AIAnalysis(
        sections=tuple(
            ai_analysis.AiSection(key=key, title=title, body=f"{title} metni.")
            for key, title in ai_analysis.SECTION_TITLES
        ),
        model="deepseek-ai/deepseek-v4.1-flash-1017",
        legend=(("Tedarikçi 1", "Bursa Metal A.Ş."),),
    )
    monkeypatch.setattr(ai_analysis, "analyze", lambda *a, **k: analysis)

    test = _ai(_upload(_app(), _workbook_bytes(tmp_path)))
    test.button[0].click()
    test.run()
    assert not test.exception, [str(e) for e in test.exception]
    text = _markdown(test)

    for _key, title in ai_analysis.SECTION_TITLES:
        assert html.escape(title) in text
        assert html.escape(f"{title} metni.") in text
    assert "Tedarikçi 1" in text
    assert "Bursa Metal A.Ş." in text


def test_ai_view_shows_unverified_figure_details(tmp_path, monkeypatch):
    monkeypatch.setenv(ai_analysis.API_KEY_ENV, "test-anahtar")

    def failing(*args, **kwargs):
        raise ai_analysis.AIAnalysisError(
            "Kaynakla eşleşmeyen ifadeler var.",
            kind="unverified",
            details=("999.999,00 TL",),
        )

    monkeypatch.setattr(ai_analysis, "analyze", failing)
    test = _ai(_upload(_app(), _workbook_bytes(tmp_path)))
    test.button[0].click()
    test.run()
    text = _markdown(test)

    assert "Kaynakla eşleşmeyen ifadeler" in text
    assert "999.999,00 TL" in text


def test_ai_view_available_with_missing_shipping(tmp_path, monkeypatch):
    """Eksik nakliye analizi engellemez; buton görünür kalır."""
    monkeypatch.setenv(ai_analysis.API_KEY_ENV, "test-anahtar")
    test = _ai(
        _upload(
            _app(),
            _workbook_bytes(tmp_path, _set("Supplier Quotes", "H3", None)),
        )
    )
    text = _markdown(test)

    assert len(test.button) == 1
    assert "AI analizi kullanılamıyor" not in text


def test_ai_view_clears_previous_analysis_on_new_file(tmp_path, monkeypatch):
    monkeypatch.setenv(ai_analysis.API_KEY_ENV, "test-anahtar")

    analysis = ai_analysis.AIAnalysis(
        sections=(
            ai_analysis.AiSection(
                key="executive_summary",
                title="Executive Summary",
                body="Önceki dosyanın yorumu.",
            ),
        ),
        model="test-model",
        legend=(),
    )
    monkeypatch.setattr(ai_analysis, "analyze", lambda *a, **k: analysis)

    test = _ai(_upload(_app(), _workbook_bytes(tmp_path)))
    test.button[0].click()
    test.run()
    assert "Önceki dosyanın yorumu." in _markdown(test)

    # Yeni (farklı) dosya yüklenince eski analiz gösterilmez.
    other = _workbook_bytes(tmp_path, _set("Supplier Quotes", "H3", None))
    test.file_uploader[0].set_value(("yeni.xlsx", other, XLSX_MIME))
    test.run()
    assert not test.exception, [str(e) for e in test.exception]

    assert "Önceki dosyanın yorumu." not in _markdown(test)


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
    test = _comparison(_upload(_app(), _workbook_bytes(tmp_path)))
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
# Rapor indirme
# --------------------------------------------------------------------------- #


def test_report_download_buttons_are_available(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path))
    labels = [button.label for button in test.download_button]

    assert labels == ["Excel raporunu indir", "PDF raporunu indir"]
    assert "Raporlar" in _markdown(test)


def test_report_download_buttons_show_expected_file_names(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path))
    urls = [button.proto.url for button in test.download_button]

    assert urls[0].endswith(".xlsx")
    assert urls[1].endswith(".pdf")


def test_reports_are_available_on_every_view(tmp_path):
    payload = _workbook_bytes(tmp_path)
    test = _upload(_app(), payload)

    for view in (DASHBOARD, COMPARISON, AI_VIEW):
        test = _view(test, view)
        assert len(test.download_button) == 2


def test_report_note_states_ai_section_is_optional(tmp_path):
    test = _upload(_app(), _workbook_bytes(tmp_path))
    text = _markdown(test)

    assert "AI bölümü olmadan" in text
    assert "KDV hariç" in text


def test_reports_include_existing_ai_analysis(tmp_path, monkeypatch):
    monkeypatch.setenv(ai_analysis.API_KEY_ENV, "test-anahtar")
    analysis = ai_analysis.AIAnalysis(
        sections=(
            ai_analysis.AiSection(
                key="executive_summary",
                title="Executive Summary",
                body="Yorum metni.",
            ),
        ),
        model="test-model",
        legend=(),
    )
    monkeypatch.setattr(ai_analysis, "analyze", lambda *a, **k: analysis)

    test = _ai(_upload(_app(), _workbook_bytes(tmp_path)))
    test.button[0].click()
    test.run()
    assert not test.exception, [str(e) for e in test.exception]

    text = _markdown(test)
    assert "Raporlar" in text
    assert "ayrı ve etiketli bölüm olarak eklenir" in text


def test_report_error_is_shown_to_user(tmp_path, monkeypatch):
    """Rapor üretimi patlarsa kullanıcıya anlaşılır hata gösterilir."""
    def broken(*args, **kwargs):
        raise RuntimeError("iç ayrıntı: C:/gizli/yol")

    monkeypatch.setattr("procurement.report._summary_sheet", broken)

    test = _upload(_app(), _workbook_bytes(tmp_path))
    text = _markdown(test)

    assert "Rapor oluşturulamadı" in text
    assert "Excel raporu oluşturulamadı" in text
    assert "C:/gizli/yol" not in text
    assert len(test.download_button) == 0


def test_report_does_not_write_files(tmp_path):
    """Raporlar yalnızca bellekte üretilir; diskte dosya bırakmaz."""
    before = {path.name for path in ROOT.iterdir()}
    test = _upload(_app(), _workbook_bytes(tmp_path))
    assert len(test.download_button) == 2
    after = {path.name for path in ROOT.iterdir()}

    assert after == before


# --------------------------------------------------------------------------- #
# Biçimsel sınır durumları
# --------------------------------------------------------------------------- #


def test_text_quantity_in_rfq_does_not_break_layout(tmp_path):
    """Validator'ın kabul ettiği ``500,0`` gibi metin miktar arayüzü düşürmez."""
    test = _comparison(
        _upload(_app(), _workbook_bytes(tmp_path, _set("RFQ", "B9", "500,0")))
    )
    text = _markdown(test)

    assert "500" in text
    assert _table_html(test)
    assert len(test.get("plotly_chart")) == 2


def test_text_quantity_in_rfq_shows_in_dashboard_summary(tmp_path):
    """Dashboard özeti de virgüllü metin miktarı okunur biçimde gösterir."""
    test = _upload(
        _app(), _workbook_bytes(tmp_path, _set("RFQ", "B9", "500,0"))
    )
    text = _markdown(test)

    assert "500 adet" in text
    assert "summary-value" in text

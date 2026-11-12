"""AI analizi testleri — DeepSeek/Deploy istemcisi taklit edilerek.

Kapsam
------
* İstek gövdesinin yalnızca hesaplanmış veriyi taşıması ve maskelenmesi.
* Başarılı yanıtın beş bölüme ayrılması.
* API hatası, zaman aşımı, boş yanıt ve bozuk JSON yönetimi.
* Eksik nakliye bilgisinin analizi engellememesi, sıfır sayılmaması.
* Reddedilen dosyada analiz yapılmaması.
* Yanıttaki kaynakla çelişen sayısal iddiaların reddedilmesi.

Hiçbir test gerçek API çağrısı yapmaz; ağa çıkılmaz.
"""

from __future__ import annotations

import json
import shutil
from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import httpx2
import openai
import pytest
from openpyxl import load_workbook

from procurement.ai_analysis import (
    API_KEY_ENV,
    SECTION_TITLES,
    AIAnalysisError,
    analysis_block_reason,
    analyze,
    build_request,
    unverified_figures,
)
from procurement.calculator import CalculationResult, calculate
from procurement.excel_parser import ParsedWorkbook, parse_workbook
from procurement.validator import ValidationResult, validate

TEMPLATE = (
    Path(__file__).resolve().parents[1]
    / "templates"
    / "supplier_quote_template.xlsx"
)

pytestmark = pytest.mark.unit

BODIES: dict[str, str] = {
    "executive_summary": "Üç geçerli teklif var; toplam maliyetler birbirine yakın.",
    "cost_analysis": (
        "En düşük toplam tedarik maliyeti 211.845,00 TL. En yüksek toplam "
        "maliyet 225.046,00 TL."
    ),
    "delivery_payment": "Terminler 15 ile 25 gün arasında; ödeme koşulları farklı.",
    "risks_missing": "Bir teklifte karma ödeme koşulu var; vade günü belirsiz.",
    "recommendations": (
        "Nakliye varsayımı olan teklifi teyit edin; seçim satın alma ekibine aittir."
    ),
}


# --------------------------------------------------------------------------- #
# Yardımcılar
# --------------------------------------------------------------------------- #


def _workbook(tmp_path: Path, mutate=None) -> Path:
    path = tmp_path / "ai.xlsx"
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


def _set(sheet: str, cell: str, value):
    def mutate(workbook) -> None:
        workbook[sheet][cell] = value

    return mutate


def _analyse(
    tmp_path: Path, mutate=None
) -> tuple[ParsedWorkbook, ValidationResult, CalculationResult]:
    parsed = parse_workbook(_workbook(tmp_path, mutate), strict=False)
    validation = validate(parsed)
    return parsed, validation, calculate(parsed, validation)


class _FakeClient:
    """OpenAI uyumlu istemcinin yalnızca kullanılan yüzeyi."""

    def __init__(self, content: str | None = None, error: Exception | None = None):
        self.content = content
        self.error = error
        self.calls: list[dict[str, Any]] = []
        self.chat = SimpleNamespace(
            completions=SimpleNamespace(create=self._create)
        )

    def _create(self, **kwargs):
        self.calls.append(kwargs)
        if self.error is not None:
            raise self.error
        return SimpleNamespace(
            choices=[
                SimpleNamespace(
                    message=SimpleNamespace(content=self.content)
                )
            ]
        )


def _json_content(**overrides: str) -> str:
    body = dict(BODIES)
    body.update(overrides)
    return json.dumps(body, ensure_ascii=False)


# --------------------------------------------------------------------------- #
# İstek gövdesi — yalnızca hesaplanmış veri, maskelenmiş kimlik
# --------------------------------------------------------------------------- #


def test_request_contains_only_verified_costs(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    payload = build_request(parsed, validation, result).payload

    first = payload["tedarikciler"][0]
    assert first["urun_bedeli"] == "204.345,00 TL"
    assert first["toplam_tedarik_maliyeti"] == "211.845,00 TL"
    assert first["birim_fiyat"] == "28,50 TL/kg"
    assert payload["en_dusuk"]["toplam_tedarik_maliyeti"]["tutar"] == (
        "211.845,00 TL"
    )


def test_request_masks_supplier_names_and_notes(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    payload = build_request(parsed, validation, result).payload
    serialized = json.dumps(payload, ensure_ascii=False)

    for real in ("Bursa Metal", "Gemlik Profil", "İnegöl Çelik", "BM-2026-1187"):
        assert real not in serialized
    assert "Tedarikçi 1" in serialized
    assert "Tedarikçi 3" in serialized


def test_request_does_not_carry_excel_notes(tmp_path):
    """Ham Excel notları (ör. KDV teyidi) AI'a gönderilmez."""
    parsed, validation, result = _analyse(tmp_path)
    serialized = json.dumps(
        build_request(parsed, validation, result).payload, ensure_ascii=False
    )

    assert "KDV hariç teyitli" not in serialized
    assert "36,12" not in serialized


def test_request_keeps_missing_shipping_as_null(tmp_path):
    parsed, validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "H3", None)
    )
    payload = build_request(parsed, validation, result).payload
    first = payload["tedarikciler"][0]

    assert first["nakliye"] is None
    assert first["toplam_tedarik_maliyeti"] is None
    assert first["eksik_veri"] == ["Nakliye bilgisi eksik"]


def test_request_includes_validation_findings_masked(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    payload = build_request(parsed, validation, result).payload
    codes = [item["kod"] for item in payload["dogrulama_bulgulari"]]
    serialized = json.dumps(payload, ensure_ascii=False)

    assert "W-PAY-01" in codes
    assert "W-WT-01" in codes
    assert "Gemlik Profil" not in serialized


def test_request_legend_maps_labels_for_ui(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    legend = build_request(parsed, validation, result).legend

    assert legend[0] == ("Tedarikçi 1", "Bursa Metal A.Ş.")
    assert legend[1][0] == "Tedarikçi 2"


def test_request_excludes_rows_are_reported(tmp_path):
    parsed, validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "G3", 0)
    )
    payload = build_request(parsed, validation, result).payload

    assert len(payload["tedarikciler"]) == 2
    assert payload["karsilastirma_disi"]
    assert payload["karsilastirma_disi"][0]["kodlar"] == ["E-PRICE-01"]


def test_excluded_row_name_is_masked_in_findings(tmp_path):
    """Karşılaştırma dışı satırın adı bulgularda ve dışlama listesinde sızmaz."""
    parsed, validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "G3", 0)
    )
    payload = build_request(parsed, validation, result).payload
    serialized = json.dumps(payload, ensure_ascii=False)

    assert "Bursa Metal" not in serialized
    assert payload["karsilastirma_disi"][0]["tedarikci"] == "Karşılaştırma dışı 1"
    excluded_finding = next(
        item
        for item in payload["dogrulama_bulgulari"]
        if item["kod"] == "E-PRICE-01"
    )
    assert excluded_finding["tedarikci"] == "Karşılaştırma dışı 1"


# --------------------------------------------------------------------------- #
# Maskeleme dayanıklılığı (inceleme R8)
# --------------------------------------------------------------------------- #


def test_prefix_names_are_fully_masked(tmp_path):
    """Biri diğerinin öneki olan adlar kısmi sızıntı bırakmaz."""
    def mutate(workbook) -> None:
        sheet = workbook["Supplier Quotes"]
        sheet["B3"] = "Bursa Metal"
        sheet["B4"] = "Bursa Metal A.Ş."

    parsed, validation, result = _analyse(tmp_path, mutate)
    payload = build_request(parsed, validation, result).payload
    serialized = json.dumps(payload, ensure_ascii=False)

    assert "Bursa Metal" not in serialized
    assert "A.Ş." not in serialized
    assert payload["tedarikciler"][0]["tedarikci"] == "Tedarikçi 1"
    assert payload["tedarikciler"][1]["tedarikci"] == "Tedarikçi 2"


def test_masking_is_case_insensitive(tmp_path):
    """Küçük harfle yazılmış ad da maskelenir."""
    def mutate(workbook) -> None:
        workbook["Supplier Quotes"]["B3"] = "BURSA METAL A.Ş."

    parsed, validation, result = _analyse(tmp_path, mutate)
    payload = build_request(parsed, validation, result).payload
    serialized = json.dumps(payload, ensure_ascii=False)

    assert "BURSA METAL" not in serialized.upper()
    assert "Tedarikçi 1" in serialized


def test_supplier_name_in_exclusion_reason_is_masked(tmp_path):
    """Dışlama gerekçesindeki ad büyük/küçük harf farkıyla da sızmaz."""
    def mutate(workbook) -> None:
        sheet = workbook["Supplier Quotes"]
        sheet["A3"] = "gemlik profil ltd."  # farklı yazım, aynı firma değil
        sheet["G3"] = 0

    parsed, validation, result = _analyse(tmp_path, mutate)
    payload = build_request(parsed, validation, result).payload
    serialized = json.dumps(payload, ensure_ascii=False).casefold()

    assert "gemlik profil" not in serialized


# --------------------------------------------------------------------------- #
# Ön koşullar
# --------------------------------------------------------------------------- #


def test_block_reason_is_none_for_valid_file(tmp_path):
    parsed, validation, result = _analyse(tmp_path)

    assert analysis_block_reason(validation, result) is None


def test_block_reason_for_rejected_file(tmp_path):
    parsed, validation, result = _analyse(
        tmp_path, _set("Instructions", "B4", "HAYIR")
    )

    assert validation.file_rejected
    assert analysis_block_reason(validation, result) is not None


def test_analyze_does_not_call_api_for_rejected_file(tmp_path):
    parsed, validation, result = _analyse(
        tmp_path, _set("Instructions", "B4", "HAYIR")
    )
    client = _FakeClient(_json_content())

    with pytest.raises(AIAnalysisError) as caught:
        analyze(parsed, validation, result, client=client)

    assert caught.value.kind == "unavailable"
    assert client.calls == []


def test_analyze_does_not_call_api_without_rows(tmp_path):
    def clear(workbook) -> None:
        sheet = workbook["Supplier Quotes"]
        for row in range(3, 6):
            for column in "ABCDEFGHIJKLMN":
                sheet[f"{column}{row}"] = None

    parsed, validation, result = _analyse(tmp_path, clear)
    client = _FakeClient(_json_content())

    with pytest.raises(AIAnalysisError) as caught:
        analyze(parsed, validation, result, client=client)

    assert caught.value.kind == "unavailable"
    assert client.calls == []


# --------------------------------------------------------------------------- #
# Başarılı yanıt
# --------------------------------------------------------------------------- #


def test_analyze_returns_five_sections_in_order(tmp_path):
    parsed, validation, result = _analyse(tmp_path)

    analysis = analyze(
        parsed, validation, result, client=_FakeClient(_json_content())
    )

    assert tuple(s.key for s in analysis.sections) == tuple(
        key for key, _title in SECTION_TITLES
    )
    assert tuple(s.title for s in analysis.sections) == (
        "Executive Summary",
        "Cost Analysis",
        "Delivery & Payment Evaluation",
        "Risks & Missing Information",
        "Procurement Recommendations",
    )


def test_analyze_sends_masked_payload_and_prompt(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    client = _FakeClient(_json_content())
    analyze(parsed, validation, result, client=client)

    call = client.calls[0]
    assert call["messages"][0]["role"] == "system"
    assert "Türkçe" in call["messages"][0]["content"]
    assert "kazanan" in call["messages"][0]["content"].lower()
    user = call["messages"][1]["content"]
    assert "Tedarikçi 1" in user
    assert "Bursa Metal" not in user


def test_analyze_accepts_fenced_json(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    fenced = f"```json\n{_json_content()}\n```"

    analysis = analyze(
        parsed, validation, result, client=_FakeClient(fenced)
    )

    assert len(analysis.sections) == 5


def test_analyze_keeps_section_text_verbatim(tmp_path):
    parsed, validation, result = _analyse(tmp_path)

    analysis = analyze(
        parsed, validation, result, client=_FakeClient(_json_content())
    )

    assert analysis.sections[0].body == BODIES["executive_summary"]


# --------------------------------------------------------------------------- #
# Hata durumları
# --------------------------------------------------------------------------- #


def _timeout_error() -> openai.APITimeoutError:
    request = httpx2.Request("POST", "https://example.invalid/v1/chat")
    return openai.APITimeoutError(request=request)


def _connection_error() -> openai.APIConnectionError:
    request = httpx2.Request("POST", "https://example.invalid/v1/chat")
    return openai.APIConnectionError(request=request)


def _status_error(status: int = 500) -> openai.APIStatusError:
    request = httpx2.Request("POST", "https://example.invalid/v1/chat")
    response = httpx2.Response(status, request=request)
    return openai.APIStatusError("boom", response=response, body=None)


def test_analyze_maps_timeout(tmp_path):
    parsed, validation, result = _analyse(tmp_path)

    with pytest.raises(AIAnalysisError) as caught:
        analyze(
            parsed,
            validation,
            result,
            client=_FakeClient(error=_timeout_error()),
        )

    assert caught.value.kind == "timeout"
    assert "zaman aşımı" in str(caught.value)


def test_analyze_maps_connection_error(tmp_path):
    parsed, validation, result = _analyse(tmp_path)

    with pytest.raises(AIAnalysisError) as caught:
        analyze(
            parsed,
            validation,
            result,
            client=_FakeClient(error=_connection_error()),
        )

    assert caught.value.kind == "connection"


def test_analyze_maps_service_error(tmp_path):
    parsed, validation, result = _analyse(tmp_path)

    with pytest.raises(AIAnalysisError) as caught:
        analyze(
            parsed,
            validation,
            result,
            client=_FakeClient(error=_status_error()),
        )

    assert caught.value.kind == "service"


def test_analyze_maps_empty_answer(tmp_path):
    parsed, validation, result = _analyse(tmp_path)

    with pytest.raises(AIAnalysisError) as caught:
        analyze(parsed, validation, result, client=_FakeClient("   "))

    assert caught.value.kind == "empty"


def test_analyze_maps_answer_without_choices(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    client = _FakeClient()
    client.chat.completions.create = lambda **kwargs: SimpleNamespace(
        choices=[]
    )

    with pytest.raises(AIAnalysisError) as caught:
        analyze(parsed, validation, result, client=client)

    assert caught.value.kind == "empty"


def test_analyze_maps_invalid_json(tmp_path):
    parsed, validation, result = _analyse(tmp_path)

    with pytest.raises(AIAnalysisError) as caught:
        analyze(
            parsed,
            validation,
            result,
            client=_FakeClient("Bu bir JSON değil."),
        )

    assert caught.value.kind == "format"


def test_analyze_rejects_missing_section(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    body = dict(BODIES)
    body.pop("recommendations")

    with pytest.raises(AIAnalysisError) as caught:
        analyze(
            parsed,
            validation,
            result,
            client=_FakeClient(json.dumps(body, ensure_ascii=False)),
        )

    assert caught.value.kind == "format"
    assert "Procurement Recommendations" in str(caught.value)


def test_analyze_requires_api_key_when_no_client(tmp_path, monkeypatch):
    parsed, validation, result = _analyse(tmp_path)
    monkeypatch.delenv(API_KEY_ENV, raising=False)
    monkeypatch.setattr(
        "procurement.ai_analysis._load_dotenv", lambda: None
    )

    with pytest.raises(AIAnalysisError) as caught:
        analyze(parsed, validation, result)

    assert caught.value.kind == "config"
    assert API_KEY_ENV in str(caught.value)


# --------------------------------------------------------------------------- #
# Eksik nakliye — analiz sürer, sıfır sayılmaz
# --------------------------------------------------------------------------- #


def test_analyze_runs_with_missing_shipping(tmp_path):
    parsed, validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "H3", None)
    )
    # Nakliyesi eksik satırın toplamı yok; yanıt yalnızca mevcut tutarları anar.
    content = _json_content(
        cost_analysis="Toplam maliyetler 213.793,00 TL ve 225.046,00 TL."
    )

    analysis = analyze(
        parsed, validation, result, client=_FakeClient(content)
    )

    assert len(analysis.sections) == 5
    assert analysis_block_reason(validation, result) is None


def test_missing_shipping_is_null_not_zero_in_request(tmp_path):
    parsed, validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "H3", None)
    )
    payload = build_request(parsed, validation, result).payload
    missing = [
        item for item in payload["tedarikciler"] if item["eksik_veri"]
    ]

    assert len(missing) == 1
    assert missing[0]["eksik_veri"] == ["Nakliye bilgisi eksik"]
    assert missing[0]["nakliye"] is None
    assert missing[0]["toplam_tedarik_maliyeti"] is None


# --------------------------------------------------------------------------- #
# Sayısal tutarlılık denetimi
# --------------------------------------------------------------------------- #


def test_unverified_figures_accepts_source_values(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    figures = build_request(parsed, validation, result).figures
    text = (
        "Toplam maliyet 211.845,00 TL; ürün bedeli 204.345,00 TL. "
        "Birim fiyat 28,50 TL/kg, ağırlık 2,390 kg/m, termin 15 gün."
    )

    assert unverified_figures(text, figures) == ()


def test_unverified_figures_detects_invented_numbers(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    figures = build_request(parsed, validation, result).figures
    text = "Maliyet yaklaşık 999.999,00 TL; birim fiyat 33,75 TL/kg."

    assert set(unverified_figures(text, figures)) == {
        "999.999,00 TL",
        "33,75 TL/kg",
    }


def test_unverified_figures_ignores_plain_counts(tmp_path):
    """Birimi olmayan serbest sayılar (tedarikçi sayısı) iddia sayılmaz."""
    parsed, validation, result = _analyse(tmp_path)
    figures = build_request(parsed, validation, result).figures

    assert unverified_figures("3 tedarikçi incelendi, 2 uyarı var.", figures) == ()


def test_unverified_figures_accepts_total_weight_and_payment_percent(tmp_path):
    """Kaynak ağırlık ve ödeme yüzdeleri iddia olarak kabul edilir."""
    parsed, validation, result = _analyse(tmp_path)
    figures = build_request(parsed, validation, result).figures
    text = (
        "Toplam ağırlık 7.170,000 kg. Ödeme %50 peşin, %50 teslimatta; "
        "vade 30 gün."
    )

    assert unverified_figures(text, figures) == ()


def test_unverified_figures_detects_invented_weight_and_percent(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    figures = build_request(parsed, validation, result).figures
    text = "Toplam ağırlık 8.000,000 kg; ödeme %12 peşin."

    assert set(unverified_figures(text, figures)) == {
        "8.000,000 kg",
        "%12",
    }


# --------------------------------------------------------------------------- #
# Tedarikçi ve metrik eşleşmesi (inceleme R1/R2/R4)
# --------------------------------------------------------------------------- #


def _figures(tmp_path, mutate=None):
    parsed, validation, result = _analyse(tmp_path, mutate)
    return build_request(parsed, validation, result).figures


def test_figures_are_scoped_per_supplier(tmp_path):
    """Her metrik kendi tedarikçisinin kaydında tutulur."""
    figures = _figures(tmp_path)
    first = figures.by_label("Tedarikçi 1")
    second = figures.by_label("Tedarikçi 2")

    assert first.product_cost == Decimal("204345.00")
    assert first.shipping == Decimal("7500.00")
    assert first.total_cost == Decimal("211845.00")
    assert first.delivery_days == Decimal("15")
    assert first.payment_days == Decimal("30")
    assert second.delivery_days == Decimal("25")
    assert second.payment_days is None
    assert second.payment_percents == (Decimal("50"),)


def test_lowest_marks_come_from_cost_mark(tmp_path):
    figures = _figures(tmp_path)

    assert figures.lowest_total_labels == ("Tedarikçi 1",)
    assert figures.lowest_total_amount == Decimal("211845.00")
    assert figures.lowest_product_labels == ("Tedarikçi 1",)


@pytest.mark.parametrize(
    "text",
    [
        "Tedarikçi 2'nin termi 15 gündür.",
        "Tedarikçi 3'ün ürün bedeli 211.845,00 TL'dir.",
        "Tedarikçi 1'in ağırlığı 2,580 kg/m'dir.",
        "Tedarikçi 2'nin birim fiyatı 28,50 TL/kg'dır.",
        "Tedarikçi 2'nin nakliyesi 7.500,00 TL'dir.",
        "Tedarikçi 1'in toplam ağırlığı 6.930,000 kg'dır.",
    ],
)
def test_value_assigned_to_wrong_supplier_is_rejected(tmp_path, text):
    """Bir tedarikçinin değeri başkasına atfedilirse analiz gösterilmez."""
    assert unverified_figures(text, _figures(tmp_path)) != ()


@pytest.mark.parametrize(
    "text",
    [
        "Tedarikçi 1'in termi 15 gündür.",
        "Tedarikçi 1'in ürün bedeli 204.345,00 TL'dir.",
        "Tedarikçi 1'in ağırlığı 2,390 kg/m'dir.",
        "Tedarikçi 2'nin birim fiyatı 30,10 TL/kg'dır.",
        "Tedarikçi 2'nin nakliyesi 5.200,00 TL'dir.",
        "Tedarikçi 1'in toplam ağırlığı 7.170,000 kg'dır.",
    ],
)
def test_value_with_correct_supplier_is_accepted(tmp_path, text):
    assert unverified_figures(text, _figures(tmp_path)) == ()


@pytest.mark.parametrize(
    "text",
    [
        "Tedarikçi 2'nin termi 30 gündür.",  # 30 = T1'in vadesi
        "Tedarikçi 1'in vadesi 15 gündür.",  # 15 = T1'in termi
        "İstenen termin 25 gündür.",  # 25 = T2'nin termi
    ],
)
def test_delivery_payment_and_requested_days_are_distinct(tmp_path, text):
    assert unverified_figures(text, _figures(tmp_path)) != ()


@pytest.mark.parametrize(
    "text",
    [
        "Tedarikçi 2'nin termi 25 gündür.",
        "Tedarikçi 1'in vadesi 30 gündür.",
        "İstenen termin 20 gündür.",
        "Terminler Tedarikçi 1 15 gün, Tedarikçi 2 25 gün; istenen termin 20 gün.",
        "T1 15 gün, T2 25 gün.",
        "Tedarikçi 3'ün ürün bedeli ve toplam maliyeti daha yüksek.",
        "Tedarikçi 2'de en düşük maliyet değil.",
        # Aynı cümlede RFQ ve tedarikçi termini: 25 sayısı RFQ sanılmamalı.
        "İstenen 20 gün termin açısından 25 günlük teklif ayrıca değerlendirilmeli.",
        # Ödeme vadesi ile teslim termini aynı cümlede.
        "Tedarikçi 1 termini 15 gün, ödeme Net 30 (30 gün).",
    ],
)
def test_correct_day_metric_is_accepted(tmp_path, text):
    assert unverified_figures(text, _figures(tmp_path)) == ()


@pytest.mark.parametrize("text", ["T2 %50 peşin ödüyor.", "%50 peşin."])
def test_supplier_index_is_not_read_as_percent(tmp_path, text):
    """'Tedarikçi 2 %50' ifadesindeki 2 yüzde değildir."""
    assert unverified_figures(text, _figures(tmp_path)) == ()


@pytest.mark.parametrize(
    "text",
    [
        "Tedarikçi 3 kg/m değeri W-WT-01 kg/m uyarısı aldı.",
        "Tedarikçi 2 ödeme koşulu W-PAY-01 uyarısı taşıyor.",
    ],
)
def test_finding_codes_are_not_read_as_claims(tmp_path, text):
    """Bulgu kodlarının sonundaki sayı (01) sayısal iddia sayılmaz."""
    assert unverified_figures(text, _figures(tmp_path)) == ()


@pytest.mark.parametrize(
    "text",
    [
        "En düşük toplam tedarik maliyeti Tedarikçi 3'te.",
        "En düşük ürün bedeli Tedarikçi 2'de.",
    ],
)
def test_lowest_cost_attributed_to_wrong_supplier_is_rejected(tmp_path, text):
    """En düşük işareti Python'un CostMark sonucuyla doğrulanır."""
    assert unverified_figures(text, _figures(tmp_path)) != ()


@pytest.mark.parametrize(
    "text",
    [
        "En düşük toplam tedarik maliyeti Tedarikçi 1'de.",
        "En düşük ürün bedeli Tedarikçi 1'de.",
    ],
)
def test_lowest_cost_attributed_to_marked_supplier_is_accepted(tmp_path, text):
    assert unverified_figures(text, _figures(tmp_path)) == ()


@pytest.mark.parametrize(
    "text",
    [
        # Nakliye bir CostMark metriği değildir; "en düşük nakliye" serbesttir.
        "Nakliye Tedarikçi 2'de 5.200,00 TL ile en düşük, Tedarikçi 1'de "
        "7.500,00 TL, Tedarikçi 3'te 9.100,00 TL.",
        # Atıf "en düşük" ifadesine en yakın tedarikçiye kurulur.
        "Ürün bedeli ve toplam tedarik maliyeti açısından Tedarikçi 1 en düşük "
        "konumda, Tedarikçi 2 toplamda ona yakın, Tedarikçi 3 ise her iki "
        "kalemde de daha yüksek.",
    ],
)
def test_lowest_wording_does_not_flag_unrelated_suppliers(tmp_path, text):
    assert unverified_figures(text, _figures(tmp_path)) == ()


def test_unrelated_mentions_do_not_justify_weight_claims(tmp_path):
    """Bulgu metnindeki 1/2/3 sayıları ağırlık iddiası olarak kabul edilmez."""
    figures = _figures(tmp_path)

    assert unverified_figures("Tedarikçi 1'in ağırlığı 3 kg/m'dir.", figures) != ()
    assert unverified_figures("Tedarikçi 2'nin ağırlığı 1 kg/m'dir.", figures) != ()


def test_reference_weight_range_is_accepted_only_as_weight(tmp_path):
    """Uyarıdaki referans aralık (2.27-2.51) kg/m olarak meşrudur."""
    figures = _figures(tmp_path)

    assert unverified_figures(
        "Tedarikçi 3'ün ağırlığı referans 2,27–2,51 aralığının dışında.",
        figures,
    ) == ()
    assert unverified_figures("Nakliye %2 daha pahalı.", figures) != ()


# --------------------------------------------------------------------------- #
# Eksik nakliye (inceleme R3)
# --------------------------------------------------------------------------- #


def test_missing_shipping_scoped_to_affected_supplier(tmp_path):
    figures = _figures(tmp_path, _set("Supplier Quotes", "H3", None))
    first = figures.by_label("Tedarikçi 1")

    assert first.shipping is None
    assert first.total_cost is None
    assert figures.by_label("Tedarikçi 2").shipping == Decimal("5200.00")


@pytest.mark.parametrize(
    "text",
    [
        "Tedarikçi 1'in nakliyesi sıfırdır.",
        "Tedarikçi 1'de nakliye yoktur.",
        "Tedarikçi 1 için nakliyesiz toplam 204.345,00 TL'dir.",
        "Tedarikçi 1'in nakliyesi 0,00 TL'dir.",
        "Tedarikçi 1'in nakliyesi 5.200,00 TL'dir.",  # T2'nin değeri
        "Tedarikçi 1'in nakliyesi 9.100,00 TL'dir.",  # T3'ün değeri
    ],
)
def test_missing_shipping_cannot_be_zeroed_or_borrowed(tmp_path, text):
    figures = _figures(tmp_path, _set("Supplier Quotes", "H3", None))

    assert unverified_figures(text, figures) != ()


def test_contradicting_answer_with_missing_shipping_is_not_shown(tmp_path):
    parsed, validation, result = _analyse(
        tmp_path, _set("Supplier Quotes", "H3", None)
    )
    content = _json_content(
        cost_analysis="Tedarikçi 1'in nakliyesi sıfırdır."
    )

    with pytest.raises(AIAnalysisError) as caught:
        analyze(parsed, validation, result, client=_FakeClient(content))

    assert caught.value.kind == "unverified"


# --------------------------------------------------------------------------- #
# Tolerans ve sayı biçimleri (inceleme R5/R6/R7)
# --------------------------------------------------------------------------- #


@pytest.mark.parametrize(
    "text",
    [
        "Tedarikçi 1'in toplam maliyeti 211.845 TL.",
        "Tedarikçi 1'in toplam maliyeti 211.845,00 TL.",
        "Tedarikçi 1'in toplam maliyeti 211845,00 TL.",
        "Tedarikçi 1'in toplam maliyeti yaklaşık 212 bin TL.",
        "Tedarikçi 1'in toplam ağırlığı 7.170 kg.",
    ],
)
def test_correct_value_in_short_format_is_accepted(tmp_path, text):
    """Doğru değerin kuruşsuz/makul yuvarlanmış yazımı reddedilmez."""
    assert unverified_figures(text, _figures(tmp_path)) == ()


@pytest.mark.parametrize(
    "text",
    [
        "Tedarikçi 1'in toplam maliyeti 226.100,00 TL.",
        "Tedarikçi 1'in toplam maliyeti 212.000,00 TL.",
        "Tedarikçi 1'in toplam maliyeti 211,845.00 TL.",
        "Tedarikçi 1'in toplam maliyeti 999,999.00 TL.",
        "Tedarikçi 1'in toplam maliyeti 213 bin TL.",
    ],
)
def test_incorrect_or_foreign_format_is_rejected(tmp_path, text):
    """Yanlış iddia kabul edilmez; İngilizce biçim güvenle reddedilir."""
    assert unverified_figures(text, _figures(tmp_path)) != ()


def test_analyze_rejects_contradicting_answer(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    content = _json_content(
        cost_analysis="En düşük toplam maliyet 199.999,00 TL."
    )

    with pytest.raises(AIAnalysisError) as caught:
        analyze(parsed, validation, result, client=_FakeClient(content))

    assert caught.value.kind == "unverified"
    assert "199.999,00 TL" in caught.value.details


def test_analyze_rejects_contradicting_delivery_days(tmp_path):
    parsed, validation, result = _analyse(tmp_path)
    content = _json_content(delivery_payment="Tüm teklifler 45 gün terminli.")

    with pytest.raises(AIAnalysisError) as caught:
        analyze(parsed, validation, result, client=_FakeClient(content))

    assert caught.value.kind == "unverified"
    assert "45 gün" in caught.value.details


def test_error_details_do_not_leak_api_key(tmp_path, monkeypatch):
    parsed, validation, result = _analyse(tmp_path)
    monkeypatch.setenv(API_KEY_ENV, "sk-test-gizli-anahtar")

    with pytest.raises(AIAnalysisError) as caught:
        analyze(
            parsed,
            validation,
            result,
            client=_FakeClient(error=_status_error(401)),
        )

    assert "sk-test-gizli-anahtar" not in str(caught.value)

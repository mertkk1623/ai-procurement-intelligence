"""AI Procurement Intelligence — Streamlit arayüzü.

Çalıştırma::

    streamlit run app.py

Katmanlar
---------
* :mod:`procurement.excel_parser` — dosyayı **okur** (salt okunur).
* :mod:`procurement.validator` — doğrular, bulgu üretir.
* :mod:`procurement.calculator` — yalnızca kabul edilen satırları hesaplar.
* :mod:`procurement.presentation` — görünüm modeli ve biçimlendirme.
* Bu dosya — yalnızca giriş/çıkış: dosya yükleme, düzen ve çizim.

Hiçbir maliyet burada hesaplanmaz; arayüz calculator çıktısını gösterir.
Otomatik kazanan, puanlama veya sıralama üretilmez.
"""

from __future__ import annotations

import html
import sys
import tempfile
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parent
if str(ROOT / "src") not in sys.path:
    sys.path.insert(0, str(ROOT / "src"))

from procurement.calculator import CalculationResult, calculate  # noqa: E402
from procurement.excel_parser import (  # noqa: E402
    ExcelParserError,
    ParsedWorkbook,
    parse_workbook,
)
from procurement.presentation import (  # noqa: E402
    MISSING_TEXT,
    KpiCard,
    build_comparison_rows,
    build_cost_figure,
    build_kpis,
    comparison_table_html,
    excluded_row_views,
    figure_title,
    finding_context,
    format_days,
    format_quantity,
    group_findings,
    rejection_hints,
)
from procurement.validator import Finding, ValidationResult, validate  # noqa: E402

APP_TITLE = "AI Procurement Intelligence"
PRODUCT_LINE = "Galvanizli kare kutu profil · 40 × 40 × 2 mm · 6 m · TL"

# --------------------------------------------------------------------------- #
# Stil
# --------------------------------------------------------------------------- #

_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

:root {
  --bg: #f5f6f8;
  --surface: #ffffff;
  --surface-muted: #fafbfc;
  --border: #e4e7ec;
  --border-strong: #d0d5dd;
  --ink: #101828;
  --ink-soft: #344054;
  --muted: #667085;
  --brand: #0f2f5c;
  --accent: #155eef;
  --lowest: #0b7a55;
  --lowest-bg: #ecfdf3;
  --lowest-border: #a6f4c5;
  --danger: #b42318;
  --danger-bg: #fef3f2;
  --danger-border: #fecdca;
  --warn: #b54708;
  --warn-bg: #fffaeb;
  --warn-border: #fedf89;
  --info: #175cd3;
  --info-bg: #eff8ff;
  --info-border: #b2ddff;
  --radius: 14px;
  --shadow: 0 1px 2px rgba(16, 24, 40, 0.04), 0 1px 3px rgba(16, 24, 40, 0.06);
}

html, body, [class*="css"], .stApp {
  font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
}

.stApp { background: var(--bg); color: var(--ink); }

header[data-testid="stHeader"] { background: transparent; }

[data-testid="stAppViewContainer"] > .main .block-container,
.block-container {
  max-width: 1240px;
  padding: 2.25rem 2rem 4rem;
}

#MainMenu, footer, [data-testid="stToolbar"] { visibility: hidden; }

h1, h2, h3 { letter-spacing: -0.015em; }

/* ---------------------------------------------------------------- hero --- */
.hero {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  justify-content: space-between;
  gap: 1rem;
  padding-bottom: 1.5rem;
  border-bottom: 1px solid var(--border);
  margin-bottom: 1.75rem;
}
.brand { display: flex; align-items: center; gap: 0.875rem; }
.brand-mark {
  width: 40px; height: 40px; flex: 0 0 40px;
  border-radius: 11px;
  background: var(--brand);
  color: #fff;
  display: inline-flex; align-items: center; justify-content: center;
  font-weight: 700; font-size: 15px; letter-spacing: 0.02em;
}
.brand-text h1 {
  font-size: 1.1875rem; font-weight: 600; line-height: 1.25;
  margin: 0; color: var(--ink);
}
.brand-text p {
  margin: 0.15rem 0 0; font-size: 0.8125rem; color: var(--muted);
}
.hero-pill {
  display: inline-flex; align-items: center; gap: 0.4rem;
  padding: 0.375rem 0.75rem;
  border: 1px solid var(--border);
  border-radius: 999px;
  background: var(--surface);
  font-size: 0.75rem; font-weight: 500; color: var(--ink-soft);
  white-space: nowrap;
}

/* -------------------------------------------------------------- upload --- */
.upload-head h2 { font-size: 1.0625rem; font-weight: 600; margin: 0 0 0.3rem; }
.upload-head p { margin: 0; font-size: 0.8125rem; color: var(--muted); }

[data-testid="stFileUploader"] { margin-top: 0.6rem; }
[data-testid="stFileUploader"] section,
[data-testid="stFileUploaderDropzone"] {
  background: var(--surface-muted);
  border: 1px dashed var(--border-strong);
  border-radius: 12px;
  padding: 1.1rem 1.25rem;
  transition: border-color 0.15s ease, background 0.15s ease;
}
[data-testid="stFileUploader"] section:hover,
[data-testid="stFileUploaderDropzone"]:hover {
  border-color: var(--accent);
  background: #f8faff;
}
[data-testid="stFileUploader"] small { color: var(--muted); }
[data-testid="stFileUploader"] button {
  border-radius: 8px;
  font-weight: 500;
}

/* --------------------------------------------------------------- strip --- */
.meta-strip {
  display: flex; flex-wrap: wrap; gap: 0.5rem 1.75rem;
  padding: 0.875rem 1.125rem;
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 12px;
  margin: 1.25rem 0 1.75rem;
  box-shadow: var(--shadow);
}
.meta-item { display: flex; flex-direction: column; gap: 0.1rem; }
.meta-item .k {
  font-size: 0.6875rem; font-weight: 500; letter-spacing: 0.06em;
  text-transform: uppercase; color: var(--muted);
}
.meta-item .v { font-size: 0.875rem; font-weight: 500; color: var(--ink); }

/* ------------------------------------------------------------ sections --- */
.section-head { margin: 0 0 0.75rem; }
.section-head h2 {
  font-size: 1.0625rem; font-weight: 600; margin: 0 0 0.25rem; color: var(--ink);
}
.section-head p { margin: 0; font-size: 0.8125rem; color: var(--muted); }
.section-gap { height: 1.5rem; }

/* ----------------------------------------------------------------- kpi --- */
.kpi-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
  gap: 0.875rem;
  margin-bottom: 0.25rem;
}
.kpi {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 1.0625rem 1.125rem 1rem;
  box-shadow: var(--shadow);
  display: flex; flex-direction: column; gap: 0.35rem;
  min-height: 116px;
}
.kpi.is-emphasis { border-color: #dbe4f3; }
.kpi-label {
  font-size: 0.6875rem; font-weight: 600; letter-spacing: 0.07em;
  text-transform: uppercase; color: var(--muted);
}
.kpi-value {
  font-size: 1.625rem; font-weight: 600; line-height: 1.15;
  color: var(--ink); font-variant-numeric: tabular-nums;
  letter-spacing: -0.02em; word-break: break-word;
}
.kpi-caption {
  font-size: 0.75rem; color: var(--muted); line-height: 1.4;
  margin-top: auto;
}

/* --------------------------------------------------------------- table --- */
.table-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  overflow: hidden;
}
.table-scroll { overflow-x: auto; }
table.cmp {
  width: 100%;
  border-collapse: collapse;
  font-size: 0.78125rem;
  min-width: 860px;
}
table.cmp thead th {
  background: var(--surface-muted);
  color: var(--muted);
  font-size: 0.65625rem;
  font-weight: 600;
  letter-spacing: 0.045em;
  text-transform: uppercase;
  text-align: left;
  padding: 0.7rem 0.55rem;
  border-bottom: 1px solid var(--border);
  white-space: normal;
  line-height: 1.35;
  vertical-align: bottom;
  position: sticky; top: 0; z-index: 1;
}
table.cmp thead th.num { text-align: right; }
table.cmp tbody td {
  padding: 0.75rem 0.55rem;
  border-bottom: 1px solid #f0f1f4;
  color: var(--ink-soft);
  vertical-align: middle;
}
table.cmp tbody tr:last-child td { border-bottom: none; }
table.cmp tbody tr:hover td { background: #fbfcfe; }
table.cmp td.num {
  text-align: right;
  font-variant-numeric: tabular-nums;
  white-space: nowrap;
  color: var(--ink);
}
table.cmp td.text { white-space: normal; min-width: 150px; }
table.cmp td.is-lowest {
  background: var(--lowest-bg);
  color: #054f31;
  font-weight: 600;
}
table.cmp td.is-lowest:hover { background: #e3fbee; }
table.cmp td.is-missing { color: var(--muted); font-style: italic; }
.supplier-cell { display: flex; flex-direction: column; gap: 0.15rem; }
.supplier-name { font-weight: 600; color: var(--ink); }
.supplier-meta {
  font-size: 0.6875rem; color: var(--muted);
  font-variant-numeric: tabular-nums;
}
.badge {
  display: inline-block;
  margin-left: 0.4rem;
  padding: 0.0625rem 0.4rem;
  border-radius: 999px;
  background: #ffffff;
  border: 1px solid var(--lowest-border);
  color: var(--lowest);
  font-size: 0.625rem;
  font-weight: 600;
  letter-spacing: 0.03em;
  text-transform: uppercase;
  vertical-align: middle;
  white-space: nowrap;
}
.tag {
  display: inline-block;
  margin-top: 0.25rem;
  padding: 0.0625rem 0.375rem;
  border-radius: 5px;
  font-size: 0.625rem;
  font-weight: 500;
  width: fit-content;
}
.tag-warn {
  background: var(--warn-bg);
  border: 1px solid var(--warn-border);
  color: var(--warn);
}
.table-note {
  margin: 0.625rem 0 0;
  font-size: 0.75rem;
  color: var(--muted);
  line-height: 1.5;
}

/* --------------------------------------------------------------- charts --- */
.st-key-chart-product,
.st-key-chart-total {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 1.125rem 1.25rem 0.75rem;
  box-shadow: var(--shadow);
  margin-bottom: 1rem;
}
.st-key-chart-product [data-testid="stPlotlyChart"],
.st-key-chart-total [data-testid="stPlotlyChart"] { margin: 0; }
.chart-head {
  display: flex; align-items: baseline; justify-content: space-between;
  gap: 0.75rem; margin-bottom: 0.25rem;
}
.chart-head h3 {
  font-size: 0.875rem; font-weight: 600; margin: 0; color: var(--ink);
}
.chart-head span { font-size: 0.6875rem; color: var(--muted); }
.legend {
  display: flex; align-items: center; gap: 1rem;
  font-size: 0.6875rem; color: var(--muted); margin-bottom: 0.35rem;
}
.legend i {
  display: inline-block; width: 9px; height: 9px;
  border-radius: 2px; margin-right: 0.3rem;
}

/* ------------------------------------------------------------- findings --- */
.panel {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  overflow: hidden;
  margin-bottom: 0.875rem;
}
.panel-head {
  padding: 0.875rem 1.125rem;
  border-bottom: 1px solid var(--border);
  display: flex; align-items: center; gap: 0.5rem;
  background: var(--surface-muted);
}
.panel-head h3 { font-size: 0.875rem; font-weight: 600; margin: 0; color: var(--ink); }
.panel-head .count {
  font-size: 0.6875rem; font-weight: 600; color: var(--muted);
  background: #eef0f3; border-radius: 999px; padding: 0.05rem 0.45rem;
}
.panel-body { padding: 0.25rem 0; }
.finding {
  display: flex; gap: 0.75rem;
  padding: 0.75rem 1.125rem;
  border-bottom: 1px solid #f0f1f4;
}
.finding:last-child { border-bottom: none; }
.finding-body { display: flex; flex-direction: column; gap: 0.2rem; min-width: 0; }
.finding-body p {
  margin: 0; font-size: 0.8125rem; color: var(--ink-soft); line-height: 1.5;
}
.finding-meta {
  font-size: 0.6875rem; color: var(--muted);
  font-variant-numeric: tabular-nums;
}
.code {
  flex: 0 0 auto;
  align-self: flex-start;
  font-family: 'SFMono-Regular', ui-monospace, 'Segoe UI Mono', Menlo, monospace;
  font-size: 0.6875rem;
  font-weight: 600;
  padding: 0.125rem 0.4rem;
  border-radius: 5px;
  border: 1px solid transparent;
  white-space: nowrap;
}
.code-error {
  background: var(--danger-bg); border-color: var(--danger-border);
  color: var(--danger);
}
.code-warn { background: var(--warn-bg); border-color: var(--warn-border); color: var(--warn); }
.code-info { background: var(--info-bg); border-color: var(--info-border); color: var(--info); }
.finding-kind {
  font-size: 0.625rem; font-weight: 600; text-transform: uppercase;
  letter-spacing: 0.06em; color: var(--muted);
}

/* ------------------------------------------------------------- rejection --- */
.reject {
  background: var(--danger-bg);
  border: 1px solid var(--danger-border);
  border-radius: var(--radius);
  padding: 1.125rem 1.25rem;
  margin-bottom: 1rem;
}
.reject h2 {
  margin: 0 0 0.35rem; font-size: 0.9375rem; font-weight: 600; color: #7a271a;
}
.reject p { margin: 0; font-size: 0.8125rem; color: #912018; line-height: 1.55; }
.reject ul { margin: 0.6rem 0 0; padding-left: 1.1rem; }
.reject li { font-size: 0.8125rem; color: #912018; line-height: 1.55; }

.empty-state {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 2rem 1.5rem;
  text-align: center;
  box-shadow: var(--shadow);
}
.empty-state h3 { margin: 0 0 0.35rem; font-size: 0.9375rem; font-weight: 600; }
.empty-state p { margin: 0; font-size: 0.8125rem; color: var(--muted); }

.info-strip {
  display: flex; flex-wrap: wrap; gap: 0.4rem 1.25rem;
  padding: 0.7rem 1rem;
  background: var(--surface-muted);
  border: 1px solid var(--border);
  border-radius: 10px;
  font-size: 0.75rem; color: var(--muted);
  margin-bottom: 1rem;
}

/* ------------------------------------------------------------ responsive --- */
@media (max-width: 720px) {
  [data-testid="stAppViewContainer"] > .main .block-container,
  .block-container { padding: 1.25rem 1rem 3rem; }
  .hero { padding-bottom: 1.1rem; margin-bottom: 1.25rem; }
  .hero-pill { white-space: normal; }
  .kpi-grid { grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 0.625rem; }
  .kpi { min-height: 0; padding: 0.875rem 0.9375rem; }
  .kpi-value { font-size: 1.375rem; }
  .meta-strip { gap: 0.5rem 1.25rem; padding: 0.75rem 0.875rem; }
  .finding { padding: 0.6875rem 0.875rem; gap: 0.5rem; }
  .panel-head { padding: 0.75rem 0.875rem; }
}
</style>
"""


# --------------------------------------------------------------------------- #
# Yardımcılar
# --------------------------------------------------------------------------- #


def _analyze(payload: bytes) -> tuple[ParsedWorkbook, ValidationResult, CalculationResult]:
    """Yüklenen baytları geçici dosyaya yazıp parse → validate → calculate çalıştırır.

    Parser yalnızca yol kabul ettiği için geçici dosya kullanılır; dosya
    işlem bittiğinde silinir ve hiçbir yere yazılmaz.
    """
    with tempfile.TemporaryDirectory(prefix="procurement-") as directory:
        path = Path(directory) / "teklif.xlsx"
        path.write_bytes(payload)
        parsed = parse_workbook(path, strict=False)
        validation = validate(parsed, today=date.today())
        result = calculate(parsed, validation)
    return parsed, validation, result


def _kpi_html(cards: tuple[KpiCard, ...]) -> str:
    items: list[str] = []
    for card in cards:
        emphasis = " is-emphasis" if card.emphasis else ""
        items.append(
            f'<div class="kpi{emphasis}">'
            f'<span class="kpi-label">{card.label}</span>'
            f'<span class="kpi-value">{card.value}</span>'
            f'<span class="kpi-caption">{card.caption}</span>'
            "</div>"
        )
    return f'<div class="kpi-grid">{"".join(items)}</div>'


def _section_head(title: str, subtitle: str = "") -> str:
    body = f"<h2>{title}</h2>"
    if subtitle:
        body += f"<p>{subtitle}</p>"
    return f'<div class="section-head">{body}</div>'


def _rfq_quantity_text(value: object) -> str | None:
    """RFQ miktarını gösterim metnine çevirir; okunamayan değeri olduğu gibi döndürür.

    Validator ``5,00`` gibi virgüllü metni de geçerli sayabildiği için burada
    ``int()`` çağrısı yapılmaz — arayüz doğrulamanın kabul ettiği bir dosyada
    çökmemelidir.
    """
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return format_quantity(int(value))
    text = str(value).strip()
    if not text:
        return None
    try:
        return format_quantity(int(Decimal(text.replace(",", "."))))
    except (InvalidOperation, ValueError):
        return text


def _meta_strip(parsed: ParsedWorkbook) -> str:
    rfq = parsed.rfq
    items = [
        ("Talep", rfq.rfq_number),
        ("Ürün", rfq.profile_specification),
        ("Miktar", _rfq_quantity_text(rfq.quantity_pcs)),
        ("Teslim yeri", rfq.delivery_location),
        ("İstenen termin", format_days(rfq.requested_delivery_days)),
        ("Para birimi", rfq.currency),
    ]
    rendered = "".join(
        f'<div class="meta-item"><span class="k">{key}</span>'
        f'<span class="v">{html.escape(str(value))}</span></div>'
        for key, value in items
        if value not in (None, "", MISSING_TEXT)
    )
    if not rendered:
        return ""
    return f'<div class="meta-strip">{rendered}</div>'


_CODE_CLASS = {"E": "code-error", "W": "code-warn", "I": "code-info"}


def _findings_panel(
    title: str, findings: tuple[Finding, ...], severity_key: str
) -> str:
    rows: list[str] = []
    for finding in findings:
        context = finding_context(finding)
        meta = (
            f'<span class="finding-meta">{html.escape(context)}</span>'
            if context
            else ""
        )
        rows.append(
            '<div class="finding">'
            f'<span class="code {_CODE_CLASS.get(severity_key, "code-info")}">'
            f"{html.escape(finding.code)}</span>"
            '<div class="finding-body">'
            f"<p>{html.escape(finding.message)}</p>{meta}</div>"
            "</div>"
        )
    return (
        '<div class="panel">'
        f'<div class="panel-head"><h3>{title}</h3>'
        f'<span class="count">{len(findings)}</span></div>'
        f'<div class="panel-body">{"".join(rows)}</div>'
        "</div>"
    )


def _excluded_panel(rows) -> str:
    body: list[str] = []
    for row in rows:
        codes = (
            f'<span class="finding-meta">Kod: {html.escape(row.codes_text)}</span>'
            if row.codes
            else ""
        )
        body.append(
            '<div class="finding">'
            '<span class="code code-error">—</span>'
            '<div class="finding-body">'
            f"<p>{html.escape(row.supplier_name)} · Satır {row.row_number} — "
            f"{html.escape(row.reason)}</p>"
            f"{codes}</div></div>"
        )
    return (
        '<div class="panel">'
        f'<div class="panel-head"><h3>Karşılaştırmaya alınmayan satırlar</h3>'
        f'<span class="count">{len(rows)}</span></div>'
        f'<div class="panel-body">{"".join(body)}</div>'
        "</div>"
    )


def _rejection_panel(validation: ValidationResult) -> str:
    errors = "".join(
        f"<li><strong>{html.escape(finding.code)}</strong> — "
        f"{html.escape(finding.message)}</li>"
        for finding in validation.file_errors
    )
    hints = "".join(
        f"<li>{html.escape(hint)}</li>" for hint in rejection_hints(validation)
    )
    hint_block = f"<p>Nasıl düzeltilir:</p><ul>{hints}</ul>" if hints else ""
    return (
        '<div class="reject">'
        "<h2>Dosya reddedildi — karşılaştırma yapılamadı</h2>"
        "<p>Dosya düzeyinde en az bir hata bulundu. Veri güvenilir okunamadığı "
        "için hiçbir satır hesaplanmadı.</p>"
        f"<ul>{errors}</ul>{hint_block}</div>"
    )


def _chart_block(result: CalculationResult, metric: str) -> None:
    """Başlık, açıklama ve grafiği tek kart içinde çizer."""
    figure = build_cost_figure(result, metric)
    if figure is None:
        return

    plotted = len(figure.data[0].x) if figure.data else 0
    omitted = len(result.rows) - plotted
    note = (
        f"{omitted} satır eksik nakliye bilgisi nedeniyle grafiğe alınmadı"
        if omitted
        else "Tüm hesaplanan satırlar dahil"
    )
    with st.container(key=f"chart-{metric}"):
        st.markdown(
            f'<div class="chart-head"><h3>{figure_title(metric)}</h3>'
            f"<span>{note}</span></div>"
            '<div class="legend">'
            '<span><i style="background:#0b7a55"></i>En düşük</span>'
            '<span><i style="background:#64748b"></i>Diğer tedarikçiler</span>'
            "</div>",
            unsafe_allow_html=True,
        )
        st.plotly_chart(
            figure, config={"displayModeBar": False}, key=f"plot-{metric}"
        )


# --------------------------------------------------------------------------- #
# Sayfa
# --------------------------------------------------------------------------- #

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="◧",
    layout="wide",
    initial_sidebar_state="collapsed",
)
st.markdown(_CSS, unsafe_allow_html=True)

st.markdown(
    '<div class="hero">'
    '<div class="brand">'
    '<span class="brand-mark">AP</span>'
    '<div class="brand-text">'
    f"<h1>{APP_TITLE}</h1>"
    "<p>Tedarikçi teklif karşılaştırma · KDV hariç TL</p>"
    "</div></div>"
    f'<span class="hero-pill">{PRODUCT_LINE}</span>'
    "</div>",
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="upload-head"><h2>Teklif dosyası</h2>'
    "<p>Standart şablonu yükleyin. Dosya yalnızca bu oturumda okunur; "
    "hiçbir yere kaydedilmez.</p></div>",
    unsafe_allow_html=True,
)

uploaded = st.file_uploader(
    "Excel teklif şablonu",
    type=["xlsx"],
    help="Yalnızca .xlsx kabul edilir; en fazla 5 MB.",
    label_visibility="collapsed",
)

if uploaded is None:
    st.markdown(
        '<div class="info-strip">'
        "<span>1 · Şablonu indirin ve tedarikçi tekliflerini girin</span>"
        "<span>2 · Dosyayı buraya yükleyin</span>"
        "<span>3 · Doğrulama ve maliyet karşılaştırmasını inceleyin</span>"
        "</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="empty-state"><h3>Henüz dosya yüklenmedi</h3>'
        "<p>Karşılaştırma, yüklediğiniz dosyadaki doğrulanmış verilerle üretilir. "
        "Örnek veya varsayılan veri kullanılmaz.</p></div>",
        unsafe_allow_html=True,
    )
    st.stop()

try:
    parsed, validation, result = _analyze(uploaded.getvalue())
except ExcelParserError as error:
    st.markdown(
        '<div class="reject"><h2>Dosya okunamadı</h2>'
        f"<p>{html.escape(str(error))}</p></div>",
        unsafe_allow_html=True,
    )
    st.stop()

findings = group_findings(validation)
st.markdown(
    f'<div class="info-strip"><span>Dosya: '
    f"<strong>{html.escape(uploaded.name)}</strong></span>"
    f"<span>Tedarikçi satırı: {len(parsed.suppliers)}</span></div>",
    unsafe_allow_html=True,
)

# --- Dosya reddi ---------------------------------------------------------- #
if validation.file_rejected:
    st.markdown(_rejection_panel(validation), unsafe_allow_html=True)

    # Dosya düzeyi hatalar red panelinde gösterildi; burada satır düzeyi
    # hatalar (varsa) ve uyarı/bilgi talepleri listelenir.
    row_errors = tuple(f for f in findings.errors if not f.is_file_level)
    if row_errors:
        st.markdown(
            _findings_panel("Satır hataları", row_errors, "E"),
            unsafe_allow_html=True,
        )
    if findings.warnings:
        st.markdown(
            _findings_panel("Uyarılar", findings.warnings, "W"),
            unsafe_allow_html=True,
        )
    if findings.infos:
        st.markdown(
            _findings_panel("Bilgi talepleri", findings.infos, "I"),
            unsafe_allow_html=True,
        )
    st.stop()

# --- Talep özeti ---------------------------------------------------------- #
meta = _meta_strip(parsed)
if meta:
    st.markdown(meta, unsafe_allow_html=True)

# --- Özet kartları -------------------------------------------------------- #
st.markdown(_section_head("Genel özet"), unsafe_allow_html=True)
st.markdown(_kpi_html(build_kpis(parsed, result)), unsafe_allow_html=True)

if not result.rows:
    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    if not parsed.suppliers:
        st.markdown(
            '<div class="empty-state"><h3>Teklif satırı bulunamadı</h3>'
            "<p>Dosyanın <em>Supplier Quotes</em> sayfasında doldurulmuş satır yok. "
            "En az bir tedarikçi teklifi girilmelidir.</p></div>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<div class="empty-state"><h3>Karşılaştırılabilir teklif yok</h3>'
            "<p>Satırlar doğrulama hataları nedeniyle hesaplamaya alınmadı. "
            "Aşağıdaki bulguları inceleyip dosyayı düzeltin.</p></div>",
            unsafe_allow_html=True,
        )

# --- Karşılaştırma tablosu ------------------------------------------------ #
if result.rows:
    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    st.markdown(
        _section_head(
            "Tedarikçi karşılaştırması",
            "Satırlar dosyadaki sırada gösterilir. En düşük ürün bedeli ve en "
            "düşük toplam tedarik maliyeti ayrı işaretlenir.",
        ),
        unsafe_allow_html=True,
    )
    comparison_rows = build_comparison_rows(result, parsed)
    st.markdown(comparison_table_html(comparison_rows), unsafe_allow_html=True)
    st.markdown(
        '<p class="table-note">Ödeme koşulları maliyet hesabına dahil edilmez; '
        "yalnızca karşılaştırma bilgisidir. Nakliyesi eksik satırlarda toplam "
        "maliyet gösterilmez ve sıfır kabul edilmez. Kazanan otomatik "
        "belirlenmez; seçim kullanıcıya aittir.</p>",
        unsafe_allow_html=True,
    )

# --- Grafikler ------------------------------------------------------------ #
if result.rows:
    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    st.markdown(
        _section_head(
            "Maliyet karşılaştırması",
            "Ürün bedeli ve toplam tedarik maliyeti ayrı grafiklerde gösterilir.",
        ),
        unsafe_allow_html=True,
    )
    _chart_block(result, "product")
    _chart_block(result, "total")

# --- Doğrulama bulguları -------------------------------------------------- #
st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
st.markdown(
    _section_head(
        "Doğrulama sonuçları",
        "Hatalar satırı karşılaştırmadan çıkarır; uyarılar hesaplamayı "
        "durdurmaz ve karar verilmesini engellemez.",
    ),
    unsafe_allow_html=True,
)

if not findings.has_any and not result.excluded:
    if parsed.suppliers:
        notice = (
            "Dosya doğrulamadan hatasız geçti; hesaplamalar doğrulanmış "
            "verilerle üretildi."
        )
    else:
        notice = (
            "Doğrulanacak teklif satırı yok. Dosyaya tedarikçi teklifleri "
            "girildikten sonra tekrar yükleyin."
        )
    st.markdown(
        '<div class="panel"><div class="panel-head">'
        "<h3>Bulgu yok</h3>"
        '<span class="count">0</span></div>'
        '<div class="panel-body"><div class="finding"><div class="finding-body">'
        f"<p>{notice}</p></div></div></div></div>",
        unsafe_allow_html=True,
    )

if findings.errors:
    st.markdown(
        _findings_panel("Hatalar", findings.errors, "E"), unsafe_allow_html=True
    )
if findings.warnings:
    st.markdown(
        _findings_panel("Uyarılar", findings.warnings, "W"), unsafe_allow_html=True
    )
if findings.infos:
    st.markdown(
        _findings_panel("Bilgi talepleri", findings.infos, "I"),
        unsafe_allow_html=True,
    )

excluded = excluded_row_views(result)
if excluded:
    st.markdown(_excluded_panel(excluded), unsafe_allow_html=True)

# --- Alt bilgi ------------------------------------------------------------ #
st.markdown(
    '<p class="table-note" style="margin-top:1.75rem">'
    "Maliyetler KDV hariç TL'dir; KDV hesaplanmaz. Tek ürün ve tek para birimi "
    "desteklenir. Toplam maliyet ürün bedeli + nakliyedir. Eksik nakliye "
    "sıfır sayılmaz; tabloda <em>Eksik veri</em> olarak gösterilir."
    "</p>",
    unsafe_allow_html=True,
)

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

import hashlib
import html
import sys
import tempfile
from datetime import date
from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parent
if str(ROOT / "src") not in sys.path:
    sys.path.insert(0, str(ROOT / "src"))

from procurement import ai_analysis  # noqa: E402
from procurement.calculator import CalculationResult, calculate  # noqa: E402
from procurement.demo_data import (  # noqa: E402
    BLANK_FILENAME,
    SAMPLE_FILENAME,
    DemoDataError,
    build_blank_template,
    build_sample_workbook,
    sample_supplier_count,
)
from procurement.excel_parser import (  # noqa: E402
    ExcelParserError,
    ParsedWorkbook,
    parse_workbook,
)
from procurement.presentation import (  # noqa: E402
    MISSING_TEXT,
    AnalysisSummary,
    KpiCard,
    build_analysis_summary,
    build_comparison_rows,
    build_cost_figure,
    build_kpis,
    comparison_table_html,
    dashboard_chart_metric,
    excluded_row_views,
    figure_title,
    finding_context,
    format_days,
    format_quantity_value,
    group_findings,
    rejection_hints,
)
from procurement.report import (  # noqa: E402
    EXCEL_FILENAME,
    EXCEL_MIME,
    PDF_FILENAME,
    PDF_MIME,
    ReportError,
    build_excel_report,
    build_pdf_report,
)
from procurement.validator import Finding, ValidationResult, validate  # noqa: E402

APP_TITLE = "AI Procurement Intelligence"
PRODUCT_LINE = "Galvanizli kare kutu profil · 40 × 40 × 2 mm · 6 m · TL"

TEMPLATE_PATH = ROOT / "templates" / "supplier_quote_template.xlsx"
TEMPLATE_FILENAME = "tedarikci-teklif-sablonu.xlsx"

VIEW_DASHBOARD = "Dashboard"
VIEW_COMPARISON = "Tedarikçi karşılaştırması"
VIEW_AI = "AI analizi"
VIEW_ROADMAP = "Product Roadmap"
VIEW_OPTIONS = (VIEW_DASHBOARD, VIEW_COMPARISON, VIEW_AI, VIEW_ROADMAP)

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
.st-key-chart-total,
.st-key-chart-dash {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 1.125rem 1.25rem 0.75rem;
  box-shadow: var(--shadow);
  margin-bottom: 1rem;
}
.st-key-chart-product [data-testid="stPlotlyChart"],
.st-key-chart-total [data-testid="stPlotlyChart"],
.st-key-chart-dash [data-testid="stPlotlyChart"] { margin: 0; }
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

.st-key-quickstart {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  padding: 1.125rem 1.25rem 1.25rem;
  margin: 0.25rem 0 1.25rem;
}
.qs-head {
  display: flex; flex-wrap: wrap; align-items: baseline;
  justify-content: space-between; gap: 0.5rem;
  padding-bottom: 0.75rem;
  border-bottom: 1px solid #f0f1f4;
}
.st-key-quickstart .qs-title {
  display: flex; align-items: center; gap: 0.5rem;
  margin: 0; padding: 0;
  font-size: 0.9375rem; font-weight: 600; line-height: 1.3; color: var(--ink);
}
.st-key-quickstart .qs-title .qs-tag {
  font-size: 0.625rem; font-weight: 600; letter-spacing: 0.06em;
  text-transform: uppercase;
  padding: 0.1rem 0.45rem;
  border-radius: 999px;
  background: var(--lowest-bg);
  border: 1px solid var(--lowest-border);
  color: #054f31;
}
.st-key-quickstart .qs-sub {
  margin: 0.35rem 0 0; font-size: 0.8125rem; color: var(--ink-soft);
  line-height: 1.5;
}
.st-key-quickstart .qs-meta {
  display: flex; flex-wrap: wrap; gap: 0.3rem 0.5rem;
  margin: 0.75rem 0 0.75rem;
}
.st-key-quickstart .qs-meta span {
  font-size: 0.6875rem; color: var(--muted);
  background: var(--surface-muted);
  border: 1px solid var(--border);
  border-radius: 999px;
  padding: 0.15rem 0.55rem;
  line-height: 1.6;
}
.st-key-quickstart [data-testid="stButton"] button,
.st-key-quickstart [data-testid="stDownloadButton"] button {
  border-radius: 9px;
  font-size: 0.8125rem;
  font-weight: 500;
  width: 100%;
}
[data-testid="stButton"] button:disabled {
  background: var(--surface-muted);
  color: var(--muted);
  border-color: var(--border);
}
.qs-note {
  margin: 0.875rem 0 0;
  padding: 0.75rem 0.875rem;
  background: var(--surface-muted);
  border: 1px solid var(--border);
  border-radius: 10px;
  font-size: 0.75rem; color: var(--ink-soft); line-height: 1.6;
}
.qs-note strong { color: var(--ink); font-weight: 600; }
.qs-note ul { margin: 0.4rem 0 0; padding-left: 1.1rem; }
.qs-note li { margin-bottom: 0.2rem; }
.qs-note li:last-child { margin-bottom: 0; }
.qs-note code {
  font-size: 0.6875rem; background: #ffffff;
  border: 1px solid var(--border); border-radius: 5px; padding: 0.05rem 0.25rem;
}
.qs-active {
  display: flex; align-items: center; gap: 0.5rem;
  margin: 0.75rem 0 0;
  padding: 0.75rem 0.875rem;
  background: var(--lowest-bg);
  border: 1px solid var(--lowest-border);
  border-radius: 10px;
  font-size: 0.8125rem; color: #054f31; line-height: 1.5;
}
.qs-active strong { font-weight: 600; }

/* ------------------------------------------------------------- landing --- */
.howto-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
  gap: 0.75rem;
  margin-bottom: 1.25rem;
}
.howto-step {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 12px;
  box-shadow: var(--shadow);
  padding: 1rem 1.125rem 1.0625rem;
  display: flex; flex-direction: column; gap: 0.4rem;
}
.howto-num {
  display: inline-flex; align-items: center; justify-content: center;
  width: 22px; height: 22px; flex: 0 0 22px;
  border-radius: 6px;
  background: var(--info-bg);
  border: 1px solid var(--info-border);
  color: var(--info);
  font-size: 0.6875rem; font-weight: 600;
  font-variant-numeric: tabular-nums;
}
.howto-step h3 {
  margin: 0; font-size: 0.875rem; font-weight: 600; color: var(--ink);
  line-height: 1.35;
}
.howto-step p {
  margin: 0; font-size: 0.75rem; color: var(--muted); line-height: 1.5;
}

.guide {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  padding: 0 1.25rem;
  margin: 1.25rem 0;
}
.guide > summary {
  cursor: pointer;
  list-style: none;
  display: flex; align-items: center; gap: 0.5rem;
  padding: 1rem 0;
  font-size: 0.875rem; font-weight: 600; color: var(--ink);
}
.guide > summary::-webkit-details-marker { display: none; }
.guide > summary::after {
  content: "+"; margin-left: auto;
  font-size: 1rem; font-weight: 500; color: var(--muted);
}
.guide[open] > summary { border-bottom: 1px solid var(--border); }
.guide[open] > summary::after { content: "\2212"; }
.guide-hint {
  margin-left: 0.35rem; font-size: 0.75rem; font-weight: 400; color: var(--muted);
}
.guide-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
  gap: 1.125rem 1.5rem;
  padding: 1.125rem 0 1.25rem;
}
.guide-item h4 {
  margin: 0 0 0.35rem; font-size: 0.8125rem; font-weight: 600; color: var(--ink);
}
.guide-item p {
  margin: 0 0 0.45rem; font-size: 0.75rem; color: var(--ink-soft); line-height: 1.55;
}
.guide-item p:last-child { margin-bottom: 0; }
.guide-item ul { margin: 0; padding-left: 1.1rem; }
.guide-item li {
  font-size: 0.75rem; color: var(--ink-soft); line-height: 1.55;
}
.guide-item code {
  font-size: 0.6875rem; background: var(--surface-muted);
  border: 1px solid var(--border); border-radius: 5px; padding: 0.05rem 0.25rem;
}
.guide-item.is-callout {
  background: var(--warn-bg);
  border: 1px solid var(--warn-border);
  border-radius: 10px;
  padding: 0.875rem 1rem;
}
.guide-item.is-callout h4, .guide-item.is-callout p { color: #93370d; }
.guide-item.is-callout code { background: #ffffff; border-color: var(--warn-border); }

/* ------------------------------------------------------------------ nav --- */
[data-testid="stButtonGroup"] { margin: 0 0 1.5rem; }
[data-testid="stButtonGroup"] button {
  border-radius: 9px;
  font-size: 0.8125rem;
  font-weight: 500;
}

/* ------------------------------------------------------------- dashboard --- */
.summary-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  padding: 1.125rem 1.25rem 1.25rem;
}
.summary-head {
  display: flex; flex-wrap: wrap; align-items: center;
  justify-content: space-between; gap: 0.75rem;
  padding-bottom: 0.75rem;
  border-bottom: 1px solid #f0f1f4;
}
.summary-file { display: flex; flex-direction: column; gap: 0.15rem; min-width: 0; }
.summary-file-label {
  font-size: 0.6875rem; font-weight: 600; letter-spacing: 0.06em;
  text-transform: uppercase; color: var(--muted);
}
.summary-file-name {
  font-size: 0.9375rem; font-weight: 600; color: var(--ink);
  overflow-wrap: anywhere;
}
.summary-detail {
  margin: 0.75rem 0 0; font-size: 0.8125rem; color: var(--ink-soft);
  line-height: 1.5;
}
.status {
  display: inline-flex; align-items: center;
  padding: 0.2rem 0.6rem;
  border-radius: 999px;
  font-size: 0.6875rem; font-weight: 600;
  border: 1px solid transparent;
  white-space: nowrap;
}
.status-ok {
  background: var(--lowest-bg); border-color: var(--lowest-border);
  color: #054f31;
}
.status-warn {
  background: var(--warn-bg); border-color: var(--warn-border); color: var(--warn);
}
.status-error {
  background: var(--danger-bg); border-color: var(--danger-border);
  color: var(--danger);
}
.summary-group { margin-top: 1.125rem; }
.summary-group h4 {
  margin: 0 0 0.5rem;
  font-size: 0.6875rem; font-weight: 600; letter-spacing: 0.06em;
  text-transform: uppercase; color: var(--muted);
}
.summary-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
  gap: 0.75rem 1.25rem;
}
.summary-item {
  display: flex; flex-direction: column; gap: 0.1rem;
  padding: 0.5rem 0.75rem;
  background: var(--surface-muted);
  border: 1px solid var(--border);
  border-radius: 10px;
  min-width: 0;
}
.summary-label { font-size: 0.6875rem; color: var(--muted); font-weight: 500; }
.summary-value {
  font-size: 0.875rem; font-weight: 600; color: var(--ink);
  font-variant-numeric: tabular-nums;
  overflow-wrap: anywhere;
}
.summary-note { font-size: 0.6875rem; color: var(--muted); line-height: 1.4; }

/* ------------------------------------------------------------- ai-panel --- */
.ai-intro {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  padding: 1.125rem 1.25rem;
  margin-bottom: 1rem;
}
.ai-intro h3 {
  margin: 0 0 0.35rem; font-size: 0.9375rem; font-weight: 600; color: var(--ink);
}
.ai-intro p {
  margin: 0; font-size: 0.8125rem; color: var(--ink-soft); line-height: 1.55;
}
.ai-meta {
  display: flex; flex-wrap: wrap; gap: 0.35rem 0.75rem;
  margin-top: 0.75rem; font-size: 0.6875rem; color: var(--muted);
}
.ai-meta span {
  background: var(--surface-muted);
  border: 1px solid var(--border);
  border-radius: 999px;
  padding: 0.15rem 0.55rem;
}
.ai-legend {
  margin: 0.75rem 0 0; padding: 0;
  list-style: none;
  display: flex; flex-wrap: wrap; gap: 0.35rem 1rem;
  font-size: 0.75rem; color: var(--ink-soft);
}
.ai-legend li { display: flex; align-items: center; gap: 0.35rem; }
.ai-legend b {
  font-weight: 600; color: var(--ink);
  font-variant-numeric: tabular-nums;
}
.ai-status {
  display: flex; align-items: flex-start; gap: 0.6rem;
  padding: 0.875rem 1.125rem;
  border: 1px solid var(--border);
  border-radius: 12px;
  background: var(--surface);
  font-size: 0.8125rem; color: var(--ink-soft); line-height: 1.55;
  margin-bottom: 1rem;
}
.ai-status strong { color: var(--ink); font-weight: 600; }
.ai-status.is-error { background: var(--danger-bg); border-color: var(--danger-border); }
.ai-status.is-error, .ai-status.is-error strong { color: #912018; }
.ai-status.is-warn { background: var(--warn-bg); border-color: var(--warn-border); }
.ai-status.is-warn, .ai-status.is-warn strong { color: #93370d; }
.ai-status ul { margin: 0.35rem 0 0; padding-left: 1.1rem; }
.ai-status li { font-size: 0.75rem; }
.ai-section {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  padding: 1.125rem 1.25rem;
  margin-bottom: 0.875rem;
}
.ai-section h3 {
  margin: 0 0 0.5rem;
  font-size: 0.6875rem; font-weight: 600; letter-spacing: 0.06em;
  text-transform: uppercase; color: var(--muted);
}
.ai-section p {
  margin: 0; font-size: 0.875rem; color: var(--ink-soft); line-height: 1.65;
  white-space: pre-line;
}
.ai-foot {
  margin: 0.5rem 0 0; font-size: 0.6875rem; color: var(--muted); line-height: 1.5;
}

/* --------------------------------------------------------------- reports --- */
.st-key-reports {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 1.125rem 1.25rem;
  box-shadow: var(--shadow);
  margin-top: 1.5rem;
}
.st-key-reports [data-testid="stDownloadButton"] button {
  border-radius: 9px;
  font-size: 0.8125rem;
  font-weight: 500;
}
.report-note {
  margin: 0.75rem 0 0; font-size: 0.6875rem; color: var(--muted); line-height: 1.5;
}

/* --------------------------------------------------------------- roadmap --- */
.roadmap-track {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(230px, 1fr));
  gap: 1rem;
  margin-bottom: 0.5rem;
}
.roadmap-stage {
  position: relative;
  display: flex; flex-direction: column;
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  padding: 1.375rem 1.25rem 1.25rem;
  overflow: hidden;
}
.roadmap-stage::before {
  content: ""; position: absolute; top: 0; left: 0; right: 0; height: 3px;
  background: var(--border-strong);
}
.roadmap-stage.is-live::before { background: var(--lowest); }
.roadmap-stage.is-live { border-color: var(--lowest-border); }
.roadmap-stage.is-future {
  background: var(--surface-muted);
  box-shadow: none;
}
.roadmap-head {
  display: flex; align-items: center; justify-content: space-between;
  gap: 0.5rem; margin-bottom: 0.75rem;
}
.roadmap-phase {
  font-size: 0.6875rem; font-weight: 600; letter-spacing: 0.08em;
  text-transform: uppercase; color: var(--muted);
}
.roadmap-stage.is-live .roadmap-phase { color: var(--lowest); }
.roadmap-status {
  padding: 0.125rem 0.5rem;
  border-radius: 999px;
  border: 1px solid var(--border);
  font-size: 0.625rem; font-weight: 600;
  letter-spacing: 0.06em; text-transform: uppercase;
  white-space: nowrap;
}
.roadmap-status.is-live {
  background: var(--lowest-bg); border-color: var(--lowest-border); color: var(--lowest);
}
.roadmap-status.is-planned {
  background: var(--info-bg); border-color: var(--info-border); color: var(--info);
}
.roadmap-status.is-future { background: #ffffff; color: var(--muted); }
.roadmap-stage h3 {
  margin: 0 0 0.25rem; font-size: 1rem; font-weight: 600; color: var(--ink);
  line-height: 1.35;
}
.roadmap-stage.is-future h3 { color: var(--ink-soft); }
.roadmap-caption {
  margin: 0 0 1rem; font-size: 0.75rem; color: var(--muted); line-height: 1.5;
}
.roadmap-list {
  list-style: none; margin: 0; padding: 0;
  display: flex; flex-direction: column; gap: 0.5rem;
}
.roadmap-list li {
  position: relative;
  padding-left: 1.35rem;
  font-size: 0.8125rem; color: var(--ink-soft); line-height: 1.5;
}
.roadmap-list li::before {
  content: ""; position: absolute; left: 0.3rem; top: 0.45rem;
  width: 0.375rem; height: 0.375rem; border-radius: 50%;
  background: var(--border-strong);
}
.roadmap-stage.is-live .roadmap-list li::before {
  content: "✓"; left: 0; top: -0.05rem;
  width: auto; height: auto; border-radius: 0; background: none;
  color: var(--lowest); font-size: 0.75rem; font-weight: 700;
}
.roadmap-stage.is-live .roadmap-list li { color: var(--ink); }
.roadmap-note {
  margin: 0.75rem 0 0; font-size: 0.6875rem; color: var(--muted); line-height: 1.5;
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
  .summary-card { padding: 1rem; }
  .summary-grid { grid-template-columns: 1fr; }
  .roadmap-track { grid-template-columns: 1fr; gap: 0.75rem; }
  .roadmap-stage { padding: 1.125rem 1rem 1rem; }
  .roadmap-head { flex-wrap: wrap; }
  .howto-grid { grid-template-columns: 1fr; gap: 0.625rem; }
  .howto-step { padding: 0.875rem 0.9375rem; }
  .guide { padding: 0 1rem; }
  .guide-grid { grid-template-columns: 1fr; gap: 1rem; padding: 1rem 0 1.125rem; }
  .guide > summary { padding: 0.875rem 0; }
  .st-key-quickstart { padding: 1rem; }
  .qs-head { flex-direction: column; align-items: flex-start; }
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


def _template_bytes() -> bytes | None:
    """Standart Excel şablonunu okur; dosya yoksa ``None`` döndürür.

    Şablon yalnızca indirme anında okunur; oturumda veya diskte kopya
    tutulmaz. Okuma hatası kullanıcıya teknik ayrıntı göstermeden ele alınır.
    """
    try:
        return TEMPLATE_PATH.read_bytes()
    except OSError:
        return None


def _blank_template_bytes() -> bytes | None:
    """Boş (veri satırı olmayan) şablonu bellekte üretir.

    Üretim resmî şablonun kopyası üzerinden yapılır; sayfa adları, sütun
    başlıkları ve açılır liste doğrulamaları korunur. Üretilemezse şablon
    indirme panelindeki uyarı gösterilir.
    """
    try:
        return build_blank_template(TEMPLATE_PATH)
    except DemoDataError:
        return None


def _sample_workbook_bytes() -> bytes | None:
    """AI demo için dolu örnek dosyayı bellekte üretir."""
    try:
        return build_sample_workbook(TEMPLATE_PATH)
    except DemoDataError:
        return None


def _demo_assets() -> tuple[bytes | None, bytes | None]:
    """Boş şablon ve dolu örnek dosyayı oturum başına bir kez üretir.

    Üretim tarihe bağlıdır (teklif/geçerlilik tarihleri bugüne göre yazılır).
    Bu yüzden sonuç gün anahtarıyla önbelleklenir: aynı gün içindeki her
    yeniden çizim aynı baytları kullanır, gün değişince dosyalar tazelenir.
    Böylece indirme düğmeleri her etkileşimde yeni bir dosya üretmez.
    """
    day = date.today().isoformat()
    cached = st.session_state.get("demo_assets")
    if (
        isinstance(cached, dict)
        and cached.get("day") == day
        and isinstance(cached.get("blank"), bytes)
        and isinstance(cached.get("sample"), bytes)
    ):
        return cached["blank"], cached["sample"]

    blank = _blank_template_bytes()
    sample = _sample_workbook_bytes()
    if blank is not None and sample is not None:
        st.session_state["demo_assets"] = {
            "day": day,
            "blank": blank,
            "sample": sample,
        }
    return blank, sample


def _demo_error(message: str) -> None:
    """Hızlı başlangıç panelinde okunur hata mesajı gösterir."""
    st.markdown(
        '<div class="ai-status is-warn"><div>'
        f"<strong>Örnek veri hazırlanamadı.</strong> {html.escape(message)} "
        "Kendi dosyanızı yükleyerek devam edebilirsiniz.</div></div>",
        unsafe_allow_html=True,
    )


def _start_demo() -> None:
    """Örnek dosyayı üretip oturum state'ine koyar.

    Dosya **diske yazılmaz**; yalnızca oturumda bayt olarak tutulur. Aynı
    oturumda yeniden üretilmediği için AI analizi imzası da sabit kalır.
    """
    _blank, sample = _demo_assets()
    if sample is None:
        st.session_state["demo_error"] = "Örnek çalışma kitabı oluşturulamadı."
        return
    st.session_state["demo_file"] = {
        "name": SAMPLE_FILENAME,
        "payload": sample,
        "suppliers": sample_supplier_count(),
    }
    st.session_state.pop("demo_error", None)


def _close_demo() -> None:
    """Demo dosyasını oturumdan kaldırır."""
    st.session_state.pop("demo_file", None)


def _render_demo_banner(demo: dict) -> None:
    """Analiz ekranında örnek verinin yüklü olduğunu belirtir.

    Demo dosyası kullanıcı yüklemesi gibi okunur; bu şerit yalnızca verinin
    kurgusal olduğunu hatırlatır ve tek tıkla kapatma sunar.
    """
    suppliers = demo.get("suppliers")
    count = f"{suppliers} tedarikçi teklifi" if suppliers else "örnek teklifler"
    columns = st.columns([3, 1], vertical_alignment="center")
    with columns[0]:
        st.markdown(
            '<div class="qs-active"><div>'
            "<strong>Örnek veri seti yüklü (demo).</strong> "
            f"{html.escape(str(count))} kurgusal veriyle analiz çalışıyor. "
            "Tedarikçi adları ve tutarlar gerçek teklif değildir."
            "</div></div>",
            unsafe_allow_html=True,
        )
    with columns[1]:
        if st.button("✕ Örnek veriyi kapat", key="close-demo-active", width="stretch"):
            _close_demo()
            st.rerun()


def _render_quickstart() -> None:
    """Hızlı başlangıç paneli: demo, boş şablon ve dolu örnek indirmesi.

    Panel yalnızca başlangıç ekranında gösterilir. Üç eylem de dosya yükleme
    alanından **önce** sunulur; demo düğmesi yükleme adımını atlar: örnek veri
    doğrudan oturuma konur ve analiz ekranı çalışır.
    """
    with st.container(key="quickstart"):
        st.markdown(
            '<div class="qs-head"><div>'
            '<h3 class="qs-title">Hızlı Başlangıç'
            '<span class="qs-tag">Demo</span></h3>'
            '<p class="qs-sub">Dosya indirmeden deneyin veya kendi veriniz '
            "için hazır şablonu kullanın. Örnek veri yalnızca bu düğmeyle "
            "yüklenir; kendiliğinden çalışmaz.</p>"
            "</div></div>"
            '<div class="qs-meta">'
            f"<span>{sample_supplier_count()} tedarikçi teklifi</span>"
            "<span>Fiyat uçurumu · termin darboğazı</span>"
            "<span>Eksik veri · süresi geçmiş teklif</span>"
            "</div>",
            unsafe_allow_html=True,
        )

        columns = st.columns(3)
        with columns[0]:
            if st.button(
                "🚀 Örnek Veri ile Hemen Dene (AI Demo)",
                key="load-demo",
                type="primary",
                width="stretch",
            ):
                _start_demo()
                st.rerun()

        blank, sample = _demo_assets()

        with columns[1]:
            if blank is None:
                st.button(
                    "📥 Boş Veri Şablonunu İndir",
                    key="blank-template-unavailable",
                    disabled=True,
                    width="stretch",
                    help="Şablon dosyası okunamadığı için indirme hazırlanamadı.",
                )
            else:
                st.download_button(
                    "📥 Boş Veri Şablonunu İndir",
                    data=blank,
                    file_name=BLANK_FILENAME,
                    mime=EXCEL_MIME,
                    key="download-blank-template",
                    width="stretch",
                )

        with columns[2]:
            if sample is None:
                st.button(
                    "📊 Örnek Dolu Dosyayı İndir",
                    key="sample-download-unavailable",
                    disabled=True,
                    width="stretch",
                    help="Örnek dosya oluşturulamadığı için indirme hazırlanamadı.",
                )
            else:
                st.download_button(
                    "📊 Örnek Dolu Dosyayı İndir",
                    data=sample,
                    file_name=SAMPLE_FILENAME,
                    mime=EXCEL_MIME,
                    key="download-sample",
                    width="stretch",
                )

        error = st.session_state.get("demo_error")
        if error:
            _demo_error(str(error))

        st.markdown(
            '<div class="qs-note"><strong>Örnek veri seti; satın alma '
            "önerileri, risk analizi ve maliyet karşılaştırması AI "
            "özelliklerini test etmeniz için gerçekçi senaryolar "
            "içerir.</strong>"
            "<ul>"
            "<li><strong>AI satın alma önerileri:</strong> tedarikçiler arasında "
            "belirgin birim fiyat farkları, ürün bedeli ile toplam maliyetin "
            "ayrıştığı satırlar ve nakliye konsolidasyon fırsatları.</li>"
            "<li><strong>Risk ve veri kalitesi analizi:</strong> 420 güne varan "
            "termin darboğazı, süresi geçmiş teklif, eksik nakliye, gerekçesiz "
            "nakliye varsayımı ve referans dışı kg/m beyanları.</li>"
            "<li><strong>Maliyet karşılaştırması:</strong> 26,40–35,40 TL/kg "
            "fiyat aralığı, Peşin/Net 30–90 ve karma ödeme koşulları.</li>"
            "</ul>"
            "Dosyayı incelemek veya değiştirip yeniden yüklemek için yukarıdaki "
            "<em>Örnek Dolu Dosyayı İndir</em> düğmesini kullanın."
            "</div>",
            unsafe_allow_html=True,
        )


def _render_template_download() -> None:
    """Başlangıç ekranında standart Excel şablonunu indirilebilir yapar."""
    payload = _template_bytes()
    if payload is None:
        st.markdown(
            '<div class="ai-status is-warn"><div>'
            "<strong>Şablon indirilemiyor.</strong> "
            "Standart teklif şablonu şu anda kullanılamıyor. "
            "Dosyayı yöneticinizden isteyebilir veya daha önce indirdiğiniz "
            "şablonu kullanabilirsiniz.</div></div>",
            unsafe_allow_html=True,
        )
        return

    st.download_button(
        "Excel şablonunu indir",
        data=payload,
        file_name=TEMPLATE_FILENAME,
        mime=EXCEL_MIME,
        key="download-template",
    )


#: Karşılama ekranındaki adımlar. Sunum amaçlıdır: hiçbir davranış bu
#: metinlerden türetilmez.
_HOW_IT_WORKS = (
    ("Excel şablonunu indir", "Standart teklif şablonunu indirin."),
    ("Tedarikçi tekliflerini doldur", "RFQ bilgilerini ve her tedarikçi için bir satır girin."),
    ("Excel'i yükle", "Doldurduğunuz .xlsx dosyasını yükleyin."),
    ("Teklifleri analiz et", "Doğrulama bulgularını ve maliyet karşılaştırmasını inceleyin."),
    ("Raporları indir", "Excel veya PDF raporunu indirin."),
)


def _render_how_it_works() -> None:
    """Beş adımlı akış özeti. Yalnızca sunumdur."""
    steps: list[str] = []
    for index, (title, caption) in enumerate(_HOW_IT_WORKS, start=1):
        steps.append(
            '<div class="howto-step">'
            f'<span class="howto-num">{index}</span>'
            f"<h3>{html.escape(title)}</h3>"
            f"<p>{html.escape(caption)}</p>"
            "</div>"
        )
    st.markdown(
        _section_head(
            "Nasıl Çalışır?",
            "Teklifler tek bir şablon üzerinden toplanır ve aynı kurallarla "
            "karşılaştırılır.",
        ),
        unsafe_allow_html=True,
    )
    st.markdown(
        f'<div class="howto-grid">{"".join(steps)}</div>',
        unsafe_allow_html=True,
    )


#: Excel kullanım rehberi. İçerik, şablonun ``Instructions`` sayfası ve
#: doğrulama kurallarıyla birebir uyumludur; yeni alan tanımlamaz.
_GUIDE_ITEMS = (
    (
        "Şablonu indirin",
        "<p>Yukarıdaki <strong>Excel şablonunu indir</strong> düğmesi standart "
        "şablonu <code>tedarikci-teklif-sablonu.xlsx</code> adıyla indirir. "
        "Yalnızca <code>.xlsx</code> kabul edilir; dosya en fazla 5 MB olabilir.</p>"
        "<p>Şablon üç sayfa içerir: <code>Instructions</code>, <code>RFQ</code> "
        "ve <code>Supplier Quotes</code>. Sayfa adlarını ve sütun başlıklarını "
        "değiştirmeyin.</p>",
    ),
    (
        "RFQ ve ürün bilgileri",
        "<p><code>RFQ</code> sayfasını bir kez doldurun. Zorunlu alanlar: "
        "<strong>RFQ Number</strong>, <strong>RFQ Date</strong>, "
        "<strong>Quantity (pcs)</strong> ve <strong>Delivery Location</strong>.</p>"
        "<p>Ürün kodu, profil ölçüsü, uzunluk ve para birimi sabittir; "
        "değiştirilemez. RFQ'da eksik zorunlu alan dosyanın tamamını reddeder.</p>",
    ),
    (
        "Tedarikçi başına bir satır",
        "<p><code>Supplier Quotes</code> sayfasında her tedarikçiyi "
        "<strong>ayrı bir satıra</strong> girin. 2. satır başlıktır; veri 3. "
        "satırdan başlar. Tamamen boş satırlar yok sayılır.</p>"
        "<p>Zorunlu alanlar: Supplier Name, Supplier Quote Number, Profile "
        "Specification, Length (m), Quantity (pcs), Weight (kg/m), Unit Price, "
        "Payment Terms ve Quote Date. Bir satırda eksik zorunlu alan yalnızca o "
        "satırı devre dışı bırakır.</p>",
    ),
    (
        "Birim fiyat, kg/m, termin ve ödeme",
        "<p><strong>Unit Price</strong> TL/kg ve <strong>KDV hariç</strong> "
        "girilir. Şablonda KDV sütunu yoktur; uygulama KDV hesaplamaz.</p>"
        "<p><strong>Weight (kg/m)</strong> tedarikçinin beyanıdır. 2,27–2,51 "
        "aralığı dışındaki değerler yalnızca uyarı üretir, karşılaştırmayı "
        "durdurmaz.</p>"
        "<p><strong>Delivery Days</strong> gün olarak yazılır. "
        "<strong>Payment Terms</strong> yalnızca metindir (ör. Peşin, Net 30, "
        "%50 peşin, %50 teslimatta); vade günü yazılmaz.</p>",
    ),
    (
        "Nakliye ve nakliye kaynağı",
        "<p><strong>Shipping Cost</strong> TL ve KDV hariçtir. "
        "<strong>Shipping Cost Source</strong> yalnızca "
        "<code>Supplier Quote</code> veya <code>Customer Assumption</code> "
        "değerini alır.</p>"
        "<p>Nakliye tutarı doluysa kaynak boş bırakılamaz. "
        "<code>Customer Assumption</code> seçildiğinde varsayımın gerekçesi "
        "<strong>Notes</strong> alanına yazılmalıdır.</p>",
    ),
    (
        "KDV beyanı ve tarihler",
        "<p>Dosyanın işleme alınması için "
        "<code>Instructions</code> sayfasındaki <code>B4</code> hücresinde "
        "<strong>VAT Declaration</strong> değeri <code>EVET</code> olmalıdır. "
        "Şablon <code>HAYIR</code> ile gelir; beyan bilinçli olarak "
        "değiştirilmelidir, aksi hâlde dosya reddedilir.</p>"
        "<p>Tarih biçimi <code>yyyy-mm-dd</code>'dir. <strong>Quote Date</strong> "
        "zorunludur; <strong>Quote Valid Until</strong> teklif tarihinden önce "
        "olamaz.</p>",
    ),
)


def _render_excel_guide() -> None:
    """Katlanabilir Excel kullanım rehberi (varsayılan kapalı)."""
    items: list[str] = []
    for title, body in _GUIDE_ITEMS:
        items.append(
            f'<div class="guide-item"><h4>{html.escape(title)}</h4>{body}</div>'
        )
    items.append(
        '<div class="guide-item is-callout">'
        "<h4>Nakliye bilinmiyorsa</h4>"
        "<p>Nakliye tutarı bilinmiyorsa <strong>Shipping Cost</strong> alanını "
        "boş bırakın; <strong>sıfır yazmayın</strong>. Sıfır, gerçek bir maliyet "
        "gibi okunur. Boş bırakılan nakliye sıfır sayılmaz; tabloda "
        "<em>Eksik veri</em> olarak gösterilir ve karşılaştırma durmaz.</p></div>"
    )
    items.append(
        '<div class="guide-item is-callout">'
        "<h4>Örnek veriler demo amaçlıdır</h4>"
        "<p><code>Supplier Quotes</code> sayfasındaki 3–5. satırlar yalnızca "
        "örnekleme içindir. Kendi tekliflerinizle değiştirin; örnek satırları "
        "gerçek veri olarak kullanmayın.</p></div>"
    )
    st.markdown(
        '<details class="guide">'
        "<summary>Excel kullanım rehberi"
        '<span class="guide-hint">Şablonu doldurmadan önce okuyun</span>'
        "</summary>"
        f'<div class="guide-grid">{"".join(items)}</div>'
        "</details>",
        unsafe_allow_html=True,
    )


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
    """RFQ miktarını gösterim metnine çevirir.

    Dönüşüm :func:`procurement.presentation.format_quantity_value` içindedir;
    böylece arayüz ile dashboard özeti aynı kuralı kullanır.
    """
    return format_quantity_value(value)


def _file_signature(name: str, payload: bytes) -> str:
    """Dosya adı ve içerikten kısa bir imza üretir.

    AI analizi oturumda saklandığı için imza değişince eski sonuç
    temizlenir; böylece yeni dosyanın analizi eski dosyanın metniyle
    karışmaz. İçerik loglanmaz, yalnızca özeti tutulur.
    """
    return f"{name}:{hashlib.sha256(payload).hexdigest()[:16]}"


def _analysis_signature(
    file_signature: str,
    validation: ValidationResult,
    result: CalculationResult,
) -> tuple[object, ...]:
    """AI yorumunun hangi sonuçlara ait olduğunu belirleyen imza."""
    return (
        file_signature,
        len(result.rows),
        str(result.lowest_product_cost),
        str(result.lowest_total_cost),
        tuple(f.code for f in validation.findings),
    )


def _current_analysis(
    signature: tuple[object, ...],
) -> ai_analysis.AIAnalysis | None:
    """Oturumdaki AI yorumunu yalnızca imzası eşleşiyorsa döndürür.

    Dosya değiştiğinde eski yorum ve hata temizlenir; böylece raporda başka
    bir dosyanın AI yorumu görünmez.
    """
    if st.session_state.get("ai_signature") != signature:
        st.session_state.pop("ai_analysis", None)
        st.session_state.pop("ai_error", None)
        st.session_state["ai_signature"] = signature
        return None
    analysis: ai_analysis.AIAnalysis | None = st.session_state.get("ai_analysis")
    return analysis


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


def _chart_block(result: CalculationResult, metric: str, *, key: str) -> None:
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
    with st.container(key=f"chart-{key}"):
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
            figure, config={"displayModeBar": False}, key=f"plot-{key}"
        )


def _summary_card(summary: AnalysisSummary) -> str:
    """Analiz özeti kartını üretir.

    Dosyadan gelen tüm metinler kaçışlanır; kart HTML olarak çizilir.
    """
    groups: list[str] = []
    for group in summary.groups:
        items: list[str] = []
        for item in group.items:
            note = (
                f'<span class="summary-note">{html.escape(item.note)}</span>'
                if item.note
                else ""
            )
            items.append(
                '<div class="summary-item">'
                f'<span class="summary-label">{html.escape(item.label)}</span>'
                f'<span class="summary-value">{html.escape(item.value)}</span>'
                f"{note}</div>"
            )
        groups.append(
            '<div class="summary-group">'
            f"<h4>{html.escape(group.title)}</h4>"
            f'<div class="summary-grid">{"".join(items)}</div>'
            "</div>"
        )

    return (
        '<div class="summary-card">'
        '<div class="summary-head">'
        '<div class="summary-file">'
        '<span class="summary-file-label">Son yüklenen teklif dosyası</span>'
        f'<span class="summary-file-name">{html.escape(summary.file_name)}</span>'
        "</div>"
        f'<span class="status status-{summary.status_tone}">'
        f"{html.escape(summary.status_label)}</span>"
        "</div>"
        f'<p class="summary-detail">{html.escape(summary.status_detail)}</p>'
        f'{"".join(groups)}'
        "</div>"
    )


def _empty_rows_state(parsed: ParsedWorkbook) -> None:
    """Hesaplanabilir satır yokken gösterilecek açıklama."""
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


def _render_dashboard(
    parsed: ParsedWorkbook,
    validation: ValidationResult,
    result: CalculationResult,
    file_name: str,
) -> None:
    """Özet görünüm: KPI kartları, tek maliyet grafiği ve analiz özeti."""
    st.markdown(_section_head("Genel özet"), unsafe_allow_html=True)
    st.markdown(_kpi_html(build_kpis(parsed, result)), unsafe_allow_html=True)

    if not result.rows:
        st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
        _empty_rows_state(parsed)
    else:
        metric = dashboard_chart_metric(result)
        if metric is not None:
            st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
            st.markdown(
                _section_head(
                    "Tedarikçi bazında maliyet",
                    "Çubuk sırası dosyadaki sıradır; en düşük değer yeşil "
                    "vurgulanır. Toplam maliyeti eksik satırlar toplam "
                    "grafiğine alınmaz.",
                ),
                unsafe_allow_html=True,
            )
            _chart_block(result, metric, key="dash")

    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
    st.markdown(
        _section_head(
            "Analiz özeti",
            "Yalnızca bu oturumda yüklenen dosyaya aittir; geçmiş saklanmaz.",
        ),
        unsafe_allow_html=True,
    )
    summary = build_analysis_summary(
        parsed, validation, result, file_name=file_name
    )
    st.markdown(_summary_card(summary), unsafe_allow_html=True)

    st.markdown(
        '<p class="table-note">Tedarikçi bazında ayrıntılı tablo, grafikler ve '
        "doğrulama bulguları için üstteki <strong>Tedarikçi karşılaştırması</strong> "
        "görünümüne geçin.</p>",
        unsafe_allow_html=True,
    )


def _render_comparison(
    parsed: ParsedWorkbook,
    validation: ValidationResult,
    result: CalculationResult,
) -> None:
    """Karşılaştırma görünümü: tablo, iki grafik ve doğrulama bulguları."""
    findings = group_findings(validation)

    meta = _meta_strip(parsed)
    if meta:
        st.markdown(meta, unsafe_allow_html=True)

    st.markdown(_section_head("Genel özet"), unsafe_allow_html=True)
    st.markdown(_kpi_html(build_kpis(parsed, result)), unsafe_allow_html=True)

    if not result.rows:
        st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
        _empty_rows_state(parsed)

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

    if result.rows:
        st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)
        st.markdown(
            _section_head(
                "Maliyet karşılaştırması",
                "Ürün bedeli ve toplam tedarik maliyeti ayrı grafiklerde gösterilir.",
            ),
            unsafe_allow_html=True,
        )
        _chart_block(result, "product", key="product")
        _chart_block(result, "total", key="total")

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
            _findings_panel("Hatalar", findings.errors, "E"),
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

    excluded = excluded_row_views(result)
    if excluded:
        st.markdown(_excluded_panel(excluded), unsafe_allow_html=True)


def _render_ai(
    parsed: ParsedWorkbook,
    validation: ValidationResult,
    result: CalculationResult,
    file_signature: str,
) -> None:
    """AI yorumu görünümü: kullanıcı isterse tek seferlik analiz üretir."""
    st.markdown(
        _section_head(
            "AI analizi",
            "Doğrulanmış ve hesaplanmış sonuçların yorumu. Karar vermez, "
            "sıralama üretmez, maliyet hesaplamaz.",
        ),
        unsafe_allow_html=True,
    )

    block_reason = ai_analysis.analysis_block_reason(validation, result)
    has_key = ai_analysis.api_key() is not None

    meta: list[str] = []
    meta.append(f"{len(result.rows)} hesaplanan teklif")
    meta.append("Tedarikçi adları maskelenir")
    if not has_key:
        meta.append(f"{ai_analysis.API_KEY_ENV} tanımlı değil")
    meta_html = "".join(f"<span>{html.escape(item)}</span>" for item in meta)

    st.markdown(
        '<div class="ai-intro">'
        "<h3>Satın alma değerlendirmesi</h3>"
        "<p>Analiz yalnızca aşağıdaki hesaplanmış verilerle üretilir. Ham Excel "
        "içeriği, notlar ve tedarikçi adları gönderilmez. Sonuç yalnızca bu "
        "oturumda tutulur; geçmiş saklanmaz.</p>"
        f'<div class="ai-meta">{meta_html}</div>'
        "</div>",
        unsafe_allow_html=True,
    )

    if block_reason is not None:
        st.markdown(
            f'<div class="ai-status is-warn"><div>'
            f"<strong>AI analizi kullanılamıyor.</strong> "
            f"{html.escape(block_reason)}</div></div>",
            unsafe_allow_html=True,
        )
        return

    if not has_key:
        st.markdown(
            '<div class="ai-status is-warn"><div>'
            "<strong>API anahtarı bulunamadı.</strong> "
            f"AI analizi için <code>{ai_analysis.API_KEY_ENV}</code> ortam değişkenini "
            "veya <code>.env</code> dosyasını tanımlayın. Diğer tüm özellikler "
            "anahtar olmadan çalışır.</div></div>",
            unsafe_allow_html=True,
        )
        return

    signature = _analysis_signature(file_signature, validation, result)
    _current_analysis(signature)  # imza değiştiyse eski yorumu temizler

    if st.button("AI analizini oluştur", key="ai-run", type="primary"):
        with st.spinner("Analiz oluşturuluyor…"):
            try:
                st.session_state["ai_analysis"] = ai_analysis.analyze(
                    parsed, validation, result
                )
                st.session_state.pop("ai_error", None)
            except ai_analysis.AIAnalysisError as error:
                st.session_state.pop("ai_analysis", None)
                st.session_state["ai_error"] = error

    error = st.session_state.get("ai_error")
    if error is not None:
        details = ""
        if error.details:
            items = "".join(
                f"<li>{html.escape(item)}</li>" for item in error.details
            )
            details = f"<p>Kaynakla eşleşmeyen ifadeler:</p><ul>{items}</ul>"
        st.markdown(
            f'<div class="ai-status is-error"><div>'
            f"<strong>Analiz üretilemedi.</strong> {html.escape(str(error))}"
            f"{details}</div></div>",
            unsafe_allow_html=True,
        )

    analysis: ai_analysis.AIAnalysis | None = st.session_state.get("ai_analysis")
    if analysis is None:
        return

    legend = "".join(
        f"<li><b>{html.escape(label)}</b> {html.escape(name)}</li>"
        for label, name in analysis.legend
    )
    if legend:
        st.markdown(
            f'<ul class="ai-legend">{legend}</ul>',
            unsafe_allow_html=True,
        )

    for section in analysis.sections:
        st.markdown(
            f'<div class="ai-section"><h3>{html.escape(section.title)}</h3>'
            f"<p>{html.escape(section.body)}</p></div>",
            unsafe_allow_html=True,
        )

    st.markdown(
        '<p class="ai-foot">AI yorumu yalnızca doğrulanmış sonuçları açıklar; '
        "maliyet hesaplamaz, tedarikçi sıralamaz ve kazanan belirlemez. "
        f"Model: <code>{html.escape(analysis.model)}</code> · "
        "Sayısal ifadeler kaynak verilerle karşılaştırılmıştır.</p>",
        unsafe_allow_html=True,
    )


#: Statik ürün yol haritası içeriği. Sunum amaçlıdır; hiçbir özellik
#: bu sabitlerden türetilmez veya çalışma zamanında değiştirilmez.
_ROADMAP_STAGES = (
    {
        "phase": "V1",
        "title": "Core Procurement Platform",
        "status": "LIVE",
        "tone": "live",
        "caption": "Canlı ortamda kullanılan çekirdek tedarik platformu.",
        "features": (
            "Excel Supplier Quote Upload",
            "Automated Data Validation",
            "Procurement Cost Calculation",
            "Supplier Comparison Dashboard",
            "AI-Powered Procurement Analysis",
            "Excel & PDF Reporting",
        ),
    },
    {
        "phase": "V2",
        "title": "Procurement Intelligence",
        "status": "PLANNED",
        "tone": "planned",
        "caption": "Planlanan tedarik zekâsı yetenekleri.",
        "features": (
            "Advanced Product Specifications",
            "AI Price Forecasting",
            "Logistics Optimization",
            "Advanced Cost Analysis",
        ),
    },
    {
        "phase": "V3",
        "title": "Enterprise Procurement",
        "status": "FUTURE",
        "tone": "future",
        "caption": "Kurumsal ölçek için gelecek aşaması.",
        "features": (
            "Supplier Portal",
            "ERP / SAP Integrations",
            "Persistent Analysis History",
            "Multi-Currency Support",
        ),
    },
)


def _render_roadmap() -> None:
    """Statik ürün yol haritası: V1 canlıda, V2 planlı, V3 gelecek.

    Görünüm yalnızca sunumdur: hiçbir özellik geliştirilmez, veri okunmaz
    ve mevcut hesaplama/doğrulama akışına dokunulmaz.
    """
    st.markdown(
        _section_head(
            "Product Roadmap",
            "Ürünün bugünkü kapsamı ve planlanan aşamaları. V2 ve V3 "
            "özellikleri henüz kullanıma açık değildir.",
        ),
        unsafe_allow_html=True,
    )

    stages: list[str] = []
    for stage in _ROADMAP_STAGES:
        tone = stage["tone"]
        items = "".join(
            f"<li>{html.escape(feature)}</li>" for feature in stage["features"]
        )
        stages.append(
            f'<div class="roadmap-stage is-{tone}">'
            '<div class="roadmap-head">'
            f'<span class="roadmap-phase">{html.escape(stage["phase"])}</span>'
            f'<span class="roadmap-status is-{tone}">'
            f'{html.escape(stage["status"])}</span>'
            "</div>"
            f'<h3>{html.escape(stage["title"])}</h3>'
            f'<p class="roadmap-caption">{html.escape(stage["caption"])}</p>'
            f'<ul class="roadmap-list">{items}</ul>'
            "</div>"
        )

    st.markdown(
        f'<div class="roadmap-track">{"".join(stages)}</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p class="roadmap-note">V1 kapsamı canlı ortamda kullanımdadır. '
        "V2 ve V3 aşamaları planlama niteliğindedir; kapsam ve öncelik "
        "sırası değişebilir.</p>",
        unsafe_allow_html=True,
    )


def _render_reports(
    parsed: ParsedWorkbook,
    validation: ValidationResult,
    result: CalculationResult,
    file_name: str,
    file_signature: str,
) -> None:
    """Excel ve PDF raporlarını indirilebilir olarak sunar.

    Raporlar her çizimde bellekte üretilir; dosya sistemi veya rapor geçmişi
    kullanılmaz. Bu oturumda üretilmiş AI yorumu varsa rapora **ayrı ve
    etiketli** bölüm olarak eklenir; yoksa rapor AI bölümü olmadan üretilir.
    """
    analysis = _current_analysis(
        _analysis_signature(file_signature, validation, result)
    )

    with st.container(key="reports"):
        st.markdown(
            _section_head(
                "Raporlar",
                "Karşılaştırma, doğrulama bulguları ve varsayımlar. Dosyalar "
                "yalnızca indirme anında üretilir; saklanmaz.",
            ),
            unsafe_allow_html=True,
        )

        try:
            excel = build_excel_report(
                parsed, validation, result, file_name=file_name, analysis=analysis
            )
            pdf = build_pdf_report(
                parsed, validation, result, file_name=file_name, analysis=analysis
            )
        except ReportError as error:
            st.markdown(
                '<div class="ai-status is-error"><div>'
                f"<strong>Rapor oluşturulamadı.</strong> {html.escape(str(error))}"
                "</div></div>",
                unsafe_allow_html=True,
            )
            return

        columns = st.columns(2)
        with columns[0]:
            st.download_button(
                "Excel raporunu indir",
                data=excel,
                file_name=EXCEL_FILENAME,
                mime=EXCEL_MIME,
                key="download-excel",
                type="primary",
            )
        with columns[1]:
            st.download_button(
                "PDF raporunu indir",
                data=pdf,
                file_name=PDF_FILENAME,
                mime=PDF_MIME,
                key="download-pdf",
            )

        note = (
            "Bu oturumda oluşturulan AI yorumu rapora ayrı ve etiketli bölüm "
            "olarak eklenir."
            if analysis is not None
            else "AI yorumu henüz oluşturulmadı; raporlar AI bölümü olmadan "
            "üretilir."
        )
        st.markdown(
            f'<p class="report-note">Maliyetler KDV hariç TL\'dir. Eksik veri '
            f"sıfır olarak gösterilmez. {html.escape(note)}</p>",
            unsafe_allow_html=True,
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

# Hızlı başlangıç paneli dosya yükleme alanının hemen üstünde durur ve yalnızca
# başlangıç ekranında gösterilir. Panel bir yer tutucuya çizilir: yükleme
# yapıldıktan (veya demo etkinleştikten) sonra yer tutucu boşaltılır. Böylece
# panel ve indirme düğmeleri analiz ekranına taşınmaz, mevcut akış korunur.
quickstart_slot = st.empty()

uploaded = st.file_uploader(
    "Excel teklif şablonu",
    type=["xlsx"],
    help="Yalnızca .xlsx kabul edilir; en fazla 5 MB.",
    label_visibility="collapsed",
)

# Yükleme kaynağı tek bir sözleşmeye indirgenir: kullanıcı dosyası varsa o,
# yoksa oturumdaki demo dosyası. Böylece parse → validate → calculate zinciri
# ve dosya imzası iki kaynak için de aynı koddan geçer.
demo_file = st.session_state.get("demo_file")
demo_payload = (
    demo_file.get("payload")
    if isinstance(demo_file, dict) and demo_file.get("payload")
    else None
)

if uploaded is None and demo_payload is None:
    with quickstart_slot.container():
        _render_quickstart()

uploaded_name: str | None = None
uploaded_payload: bytes | None = None

if uploaded is not None:
    uploaded_name = uploaded.name
    uploaded_payload = uploaded.getvalue()
elif demo_payload is not None:
    uploaded_name = str(demo_file.get("name") or SAMPLE_FILENAME)
    uploaded_payload = demo_payload

if uploaded_payload is None or uploaded_name is None:
    st.markdown(
        '<div class="empty-state"><h3>Henüz dosya yüklenmedi</h3>'
        "<p>Karşılaştırma, yüklediğiniz dosyadaki doğrulanmış verilerle "
        "üretilir. Örnek veriyi denemek için yukarıdaki "
        "<strong>Örnek Veri ile Hemen Dene</strong> düğmesini kullanın.</p>"
        "</div>",
        unsafe_allow_html=True,
    )
    _render_how_it_works()
    _render_template_download()
    _render_excel_guide()
    _render_roadmap()
    st.stop()

# Demo dosyası yalnızca kullanıcı dosyası yokken etkindir; kullanıcı kendi
# dosyasını yüklediğinde demo oturumdan düşer ve karışıklık oluşmaz.
is_demo = uploaded is None
if not is_demo:
    st.session_state.pop("demo_file", None)

# Demo şeridi ayrıştırmadan **önce** çizilir: dosya okunamazsa bile kullanıcı
# örnek veriyi kapatabilir; aksi hâlde bozuk bir demo oturumda kilitli kalırdı.
if is_demo:
    _render_demo_banner(demo_file if isinstance(demo_file, dict) else {})

try:
    parsed, validation, result = _analyze(uploaded_payload)
except ExcelParserError as error:
    st.markdown(
        '<div class="reject"><h2>Dosya okunamadı</h2>'
        f"<p>{html.escape(str(error))}</p></div>",
        unsafe_allow_html=True,
    )
    st.stop()

file_name = uploaded_name
file_signature = _file_signature(file_name, uploaded_payload)

findings = group_findings(validation)
st.markdown(
    f'<div class="info-strip"><span>Dosya: '
    f"<strong>{html.escape(file_name)}</strong></span>"
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

# --- Görünüm seçimi ------------------------------------------------------- #
view = st.segmented_control(
    "Görünüm",
    options=VIEW_OPTIONS,
    default=VIEW_DASHBOARD,
    key="view",
    label_visibility="collapsed",
)
if view is None:
    view = VIEW_DASHBOARD

if view == VIEW_DASHBOARD:
    _render_dashboard(parsed, validation, result, file_name)
elif view == VIEW_COMPARISON:
    _render_comparison(parsed, validation, result)
elif view == VIEW_ROADMAP:
    _render_roadmap()
else:
    _render_ai(parsed, validation, result, file_signature)

_render_reports(parsed, validation, result, file_name, file_signature)

# --- Alt bilgi ------------------------------------------------------------ #
st.markdown(
    '<p class="table-note" style="margin-top:1.75rem">'
    "Maliyetler KDV hariç TL'dir; KDV hesaplanmaz. Tek ürün ve tek para birimi "
    "desteklenir. Toplam maliyet ürün bedeli + nakliyedir. Eksik nakliye "
    "sıfır sayılmaz; tabloda <em>Eksik veri</em> olarak gösterilir."
    "</p>",
    unsafe_allow_html=True,
)

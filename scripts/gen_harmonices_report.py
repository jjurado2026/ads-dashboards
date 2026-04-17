#!/usr/bin/env python3
"""Generate Harmonices March 2026 HTML report."""
import json, os, base64

# Load data
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
with open(os.path.join(BASE, "data/harmonices/2026-03.json")) as f:
    d = json.load(f)
with open(os.path.join(BASE, "assets/harmonices-logo.png"), "rb") as f:
    logo_b64 = base64.b64encode(f.read()).decode()

g = d["google_ads"]["kpis"]
m = d["meta_ads"]["kpis"]
hs = d["hubspot"]

# ========== CALCULATIONS ==========
total_inv = g["inversion"]["value"] + m["inversion"]["value"]
total_inv_prev = g["inversion"]["previous"] + m["inversion"]["previous"]
total_leads = g["leads"]["value"] + m["leads"]["value"]
total_leads_prev = g["leads"]["previous"] + m["leads"]["previous"]
total_cpl = total_inv / total_leads
total_cpl_prev = total_inv_prev / total_leads_prev

inv_mom = (total_inv / total_inv_prev - 1) * 100
leads_mom = (total_leads / total_leads_prev - 1) * 100
cpl_mom = (total_cpl / total_cpl_prev - 1) * 100

valid_mar = 5
valid_feb = 33
valid_mom = (valid_mar / valid_feb - 1) * 100
contracts_mar = 0
contracts_feb = 2

meta_pct = m["leads"]["value"] / total_leads * 100
google_pct = g["leads"]["value"] / total_leads * 100
meta_inv_pct = m["inversion"]["value"] / total_inv * 100
google_inv_pct = g["inversion"]["value"] / total_inv * 100

def mom(v, p):
    return (v / p - 1) * 100 if p else 0

g_leads_mom = mom(g["leads"]["value"], g["leads"]["previous"])
g_cpl_mom = mom(g["cpl"]["value"], g["cpl"]["previous"])
g_inv_mom = mom(g["inversion"]["value"], g["inversion"]["previous"])
g_imp_mom = mom(g["impresiones"]["value"], g["impresiones"]["previous"])
g_ctr_mom = mom(g["ctr"]["value"], g["ctr"]["previous"])
g_cpc_mom = mom(g["cpc"]["value"], g["cpc"]["previous"])

m_leads_mom = mom(m["leads"]["value"], m["leads"]["previous"])
m_cpl_mom = mom(m["cpl"]["value"], m["cpl"]["previous"])
m_inv_mom = mom(m["inversion"]["value"], m["inversion"]["previous"])
m_imp_mom = mom(m["impresiones"]["value"], m["impresiones"]["previous"])
m_ctr_mom = mom(m["ctr"]["value"], m["ctr"]["previous"])
m_cpc_mom = mom(m["cpc"]["value"], m["cpc"]["previous"])
m_alcance_mom = mom(m["alcance"]["value"], m["alcance"]["previous"])

# Helpers
def mc(val, inv=False):
    if inv:
        return "good" if val < 0 else "bad"
    return "good" if val > 0 else "bad"

def ma(val):
    return "arrow-up" if val > 0 else "arrow-down"

def fp(val):
    return f"{val:+.1f}%"

# Data references
g_daily = d["google_ads"]["daily"]
m_daily = d["meta_ads"]["daily"]
g_campaigns = d["google_ads"]["campaigns"]
m_campaigns = d["meta_ads"]["campaigns"]
camp_plat = d["meta_ads"]["campaigns_by_platform"]
creatives = d["meta_ads"]["creatives"]
kw_sorted = sorted(d["google_ads"]["keywords"], key=lambda x: x["clics"], reverse=True)[:15]
g_hist = d["google_ads"]["monthly_history"]
m_hist = d["meta_ads"]["monthly_history"]
hs_current = hs["deals_current_month"]
hs_stock = hs["deals_stock_total"]
hs_visitas = hs["visitas_monthly"]
hs_valid = hs["leads_valid_monthly"]

month_names = {"01":"Ene","02":"Feb","03":"Mar","04":"Abr","05":"May","06":"Jun","07":"Jul","08":"Ago","09":"Sep","10":"Oct","11":"Nov","12":"Dic"}

total_active = hs_stock["lead_entrante"] + hs_stock["exploratorio"] + hs_stock["interesado_futuro"] + hs_stock["interesado_caliente"] + hs_stock["reservado_pagado"]

# ========== BUILD HTML ==========
parts = []
W = parts.append

# HEAD + CSS
W("""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Harmonices — Informe de Resultados Marzo 2026</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root {
  --dn: #1A2332; --mb: #3B5998; --lb: #5B7FBF; --pb: #E8EEF7;
  --gold: #C9A96E; --w: #FFF; --g50: #F9FAFB; --g100: #F3F4F6;
  --g200: #E5E7EB; --g300: #D1D5DB; --g400: #9CA3AF; --g500: #6B7280;
  --g600: #4B5563; --g700: #374151; --g800: #1F2937; --g900: #111827;
  --amber: #F59E0B; --red: #DC2626;
}
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family:'Inter',-apple-system,BlinkMacSystemFont,sans-serif; background:var(--g100); color:var(--g800); font-size:10pt; line-height:1.5; -webkit-font-smoothing:antialiased; }
.page { width:210mm; min-height:297mm; max-height:297mm; margin:0 auto; background:var(--w); overflow:hidden; position:relative; page-break-after:always; page-break-inside:avoid; }
.page:last-child { page-break-after:auto; }
.page-inner { padding:18mm 20mm 16mm 20mm; height:100%; }
.page-footer { position:absolute; bottom:8mm; left:20mm; right:20mm; display:flex; justify-content:space-between; align-items:center; font-size:7pt; color:var(--g400); border-top:1px solid var(--g200); padding-top:4px; }
.page-footer-dark { position:absolute; bottom:8mm; left:20mm; right:20mm; display:flex; justify-content:space-between; align-items:center; font-size:7pt; color:rgba(255,255,255,0.3); border-top:1px solid rgba(255,255,255,0.1); padding-top:4px; }

/* COVER */
.cover { background:linear-gradient(160deg,#0F1820 0%,#1A2332 35%,#253448 65%,#1A2332 100%); display:flex; flex-direction:column; justify-content:center; align-items:center; text-align:center; color:var(--w); position:relative; }
.cover::before { content:''; position:absolute; inset:0; background:radial-gradient(ellipse 600px 400px at 20% 30%,rgba(201,169,110,0.08) 0%,transparent 70%),radial-gradient(ellipse 500px 500px at 80% 70%,rgba(59,89,152,0.06) 0%,transparent 70%); pointer-events:none; }
.cover-content { position:relative; z-index:1; }
.cover-logo-text { font-size:28pt; font-weight:800; letter-spacing:8px; text-transform:uppercase; margin-bottom:6px; }
.cover-divider { width:60px; height:2px; background:var(--gold); margin:24px auto; opacity:0.6; }
.cover-footer { position:absolute; bottom:20mm; left:0; right:0; text-align:center; font-size:8pt; color:rgba(255,255,255,0.3); z-index:1; }

/* HEADERS */
.page-header { display:flex; justify-content:space-between; align-items:center; margin-bottom:16px; padding-bottom:10px; border-bottom:2px solid var(--pb); }
.page-header-left h2 { font-size:16pt; font-weight:700; color:var(--dn); line-height:1.2; }
.page-header-left .subtitle { font-size:8.5pt; color:var(--g500); font-weight:400; margin-top:2px; }
.page-header-right { text-align:right; font-size:8pt; color:var(--g400); }
.page-header-right .brand { font-weight:600; color:var(--mb); }

/* KPI */
.kpi-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:10px; margin-bottom:14px; }
.kpi-card { background:var(--w); border:1px solid var(--g200); border-radius:10px; padding:12px 14px; position:relative; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.04); }
.kpi-card::before { content:''; position:absolute; top:0; left:0; right:0; height:3px; }
.kpi-card.green::before { background:var(--mb); }
.kpi-card.red::before { background:var(--red); }
.kpi-card.amber::before { background:var(--amber); }
.kpi-label { font-size:7.5pt; font-weight:600; color:var(--g500); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:4px; }
.kpi-value { font-size:18pt; font-weight:800; color:var(--dn); line-height:1.1; }
.kpi-value.smaller { font-size:15pt; }
.kpi-change { display:inline-flex; align-items:center; gap:3px; font-size:8pt; font-weight:600; margin-top:4px; padding:2px 6px; border-radius:4px; }
.kpi-change.bad { color:var(--red); background:#FEF2F2; }
.kpi-change.good { color:#16A34A; background:#F0FDF4; }
.kpi-prev { font-size:7.5pt; color:var(--g400); margin-top:2px; }
.arrow-up::before { content:'\\25B2'; margin-right:2px; font-size:6pt; }
.arrow-down::before { content:'\\25BC'; margin-right:2px; font-size:6pt; }

/* CHANNELS */
.channel-row { display:grid; grid-template-columns:1fr 1fr; gap:10px; margin-bottom:14px; }
.channel-box { border-radius:10px; padding:12px 16px; display:flex; align-items:center; gap:14px; }
.channel-box.meta-box { background:linear-gradient(135deg,#EEF2FF,#E0E7FF); border:1px solid #C7D2FE; }
.channel-box.google-box { background:linear-gradient(135deg,#FEF3C7,#FDE68A40); border:1px solid #FCD34D; }
.channel-icon { width:40px; height:40px; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0; }
.channel-icon.meta-icon { background:#4F46E5; }
.channel-icon.google-icon { background:#F59E0B; }
.channel-data { flex:1; }
.channel-name { font-size:8pt; font-weight:600; text-transform:uppercase; letter-spacing:0.5px; color:var(--g600); margin-bottom:2px; }
.channel-leads { font-size:16pt; font-weight:800; color:var(--g900); line-height:1.1; }
.channel-leads span { font-size:9pt; font-weight:500; color:var(--g500); }
.channel-cpl { font-size:9pt; font-weight:600; color:var(--g600); margin-top:1px; }

/* SEMAFORO */
.traffic-light { background:var(--g50); border:1px solid var(--g200); border-radius:10px; padding:12px 16px; margin-bottom:14px; }
.traffic-light-title { font-size:8pt; font-weight:700; text-transform:uppercase; letter-spacing:0.5px; color:var(--g500); margin-bottom:8px; }
.tl-item { display:flex; align-items:center; gap:10px; padding:5px 0; }
.tl-dot { width:10px; height:10px; border-radius:50%; flex-shrink:0; }
.tl-dot.red-dot { background:var(--red); box-shadow:0 0 6px rgba(220,38,38,0.3); }
.tl-dot.amber-dot { background:var(--amber); box-shadow:0 0 6px rgba(245,158,11,0.3); }
.tl-dot.green-dot { background:var(--mb); box-shadow:0 0 6px rgba(59,89,152,0.3); }
.tl-text { font-size:9pt; color:var(--g700); }
.tl-text strong { font-weight:600; }

/* VERDICT */
.verdict { background:linear-gradient(135deg,var(--dn),#2A3F5F); border-radius:10px; padding:14px 18px; color:var(--w); font-size:10pt; font-weight:500; line-height:1.5; }
.verdict-label { font-size:7pt; font-weight:700; text-transform:uppercase; letter-spacing:1px; color:rgba(255,255,255,0.5); margin-bottom:4px; }

/* TABLE */
.data-table { width:100%; border-collapse:collapse; font-size:8.5pt; margin-bottom:12px; }
.data-table thead th { background:linear-gradient(135deg,var(--dn),#2A3F5F); color:var(--w); font-weight:600; font-size:7.5pt; text-transform:uppercase; letter-spacing:0.5px; padding:7px 10px; text-align:left; }
.data-table thead th:first-child { border-radius:6px 0 0 0; }
.data-table thead th:last-child { border-radius:0 6px 0 0; }
.data-table tbody td { padding:7px 10px; border-bottom:1px solid var(--g100); vertical-align:middle; }
.data-table tbody tr:hover { background:var(--g50); }
.data-table .num { text-align:right; font-variant-numeric:tabular-nums; font-weight:500; }
.data-table .bold { font-weight:700; }
.data-table tfoot td { padding:7px 10px; font-weight:700; background:var(--pb); border-top:2px solid var(--mb); }
.data-table tfoot td:first-child { border-radius:0 0 0 6px; }
.data-table tfoot td:last-child { border-radius:0 0 6px 0; }

/* INSIGHTS */
.insight-box { border-radius:8px; padding:10px 14px; font-size:8.5pt; line-height:1.55; margin-bottom:10px; }
.insight-box.green-insight { background:#F0FDF4; border-left:3px solid var(--mb); color:#166534; }
.insight-box.amber-insight { background:#FFFBEB; border-left:3px solid var(--amber); color:#92400E; }
.insight-box.red-insight { background:#FEF2F2; border-left:3px solid var(--red); color:#991B1B; }
.insight-box.blue-insight { background:#EFF6FF; border-left:3px solid #3B82F6; color:#1E40AF; }
.insight-label { font-size:7pt; font-weight:700; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:3px; }

/* SECTION */
.section-title { font-size:10pt; font-weight:700; color:var(--dn); margin-bottom:8px; margin-top:10px; background:linear-gradient(90deg,var(--pb) 0%,transparent 80%); margin-left:-20mm; margin-right:-20mm; padding:8px 20mm; }
.section-subtitle { font-size:8.5pt; font-weight:600; color:var(--g600); margin-bottom:6px; margin-top:8px; }
.split { display:grid; grid-template-columns:1fr 1fr; gap:12px; }
.mix-block { border-radius:10px; padding:16px; text-align:center; }
.mix-block.meta-mix { background:linear-gradient(150deg,#EEF2FF,#E0E7FF); border:1px solid #C7D2FE; }
.mix-block.google-mix { background:linear-gradient(150deg,#FFFBEB,#FEF3C7); border:1px solid #FCD34D; }
.mix-pct { font-size:28pt; font-weight:800; line-height:1; }
.mix-pct.meta-pct { color:#4F46E5; }
.mix-pct.google-pct { color:#D97706; }
.mix-label { font-size:8pt; font-weight:500; color:var(--g500); margin-top:2px; }
.mix-leads { font-size:13pt; font-weight:700; color:var(--g800); margin-top:6px; }

/* BARS */
.bar-chart { margin-bottom:12px; }
.bar-row { display:flex; align-items:center; margin-bottom:6px; }
.bar-label { width:80px; font-size:8pt; font-weight:600; color:var(--g600); text-align:right; padding-right:10px; flex-shrink:0; }
.bar-track { flex:1; height:24px; background:var(--g100); border-radius:6px; overflow:hidden; display:flex; }
.bar-fill { height:100%; display:flex; align-items:center; padding:0 8px; font-size:7.5pt; font-weight:600; color:var(--w); border-radius:6px; }
.bar-fill.meta-fill { background:#4F46E5; }
.bar-fill.google-fill { background:#D97706; }

/* KPI ROW 6 */
.kpi-row-6 { display:grid; grid-template-columns:repeat(6,1fr); gap:8px; margin-bottom:12px; }
.kpi-mini { background:var(--g50); border:1px solid var(--g200); border-radius:8px; padding:8px 10px; text-align:center; }
.kpi-mini .kpi-label { font-size:6.5pt; margin-bottom:2px; }
.kpi-mini .kpi-value { font-size:13pt; font-weight:700; }
.kpi-mini .kpi-change { font-size:7pt; margin-top:2px; }

/* CHART */
.chart-container { background:var(--g50); border:1px solid var(--g200); border-radius:10px; padding:14px; margin-bottom:12px; }

/* FUNNEL */
.funnel { margin-bottom:14px; }
.funnel-step { display:flex; align-items:center; margin-bottom:4px; }
.funnel-bar { height:32px; border-radius:6px; display:flex; align-items:center; padding:0 12px; font-size:8.5pt; font-weight:600; color:var(--w); }
.funnel-bar.f1 { background:var(--dn); }
.funnel-bar.f2 { background:var(--mb); }
.funnel-bar.f3 { background:var(--lb); }
.funnel-bar.f4 { background:var(--gold); color:var(--dn); }
.funnel-label { font-size:7.5pt; color:var(--g500); margin-left:8px; flex-shrink:0; }

/* FORMAT PILLS */
.format-row { display:flex; gap:10px; margin-bottom:10px; }
.format-pill { flex:1; border-radius:8px; padding:10px 14px; text-align:center; }
.format-pill.success { background:#F0FDF4; border:1px solid #BBF7D0; }
.format-name { font-size:7.5pt; font-weight:700; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:2px; color:#166534; }
.format-value { font-size:16pt; font-weight:800; color:#16A34A; }
.format-detail { font-size:7.5pt; color:var(--g500); margin-top:2px; }

/* YOY */
.yoy-grid { display:grid; grid-template-columns:1fr 1fr; gap:10px; margin-bottom:12px; }
.yoy-card { border:1px solid var(--g200); border-radius:8px; padding:10px 14px; background:var(--w); }
.yoy-metric { font-size:7.5pt; font-weight:600; color:var(--g500); text-transform:uppercase; letter-spacing:0.3px; margin-bottom:6px; }
.yoy-values { display:flex; align-items:baseline; gap:8px; }
.yoy-old { font-size:12pt; font-weight:400; color:var(--g400); text-decoration:line-through; }
.yoy-arrow { color:var(--g400); font-size:10pt; }
.yoy-new { font-size:14pt; font-weight:700; color:var(--g900); }
.yoy-change { display:inline-block; font-size:8pt; font-weight:600; padding:1px 6px; border-radius:4px; margin-top:4px; }
.yoy-change.pos-good { background:#F0FDF4; color:#16A34A; }
.yoy-change.neg-good { background:#F0FDF4; color:#16A34A; }
.yoy-change.pos-bad { background:#FEF2F2; color:var(--red); }

/* STRATEGY */
.strategy-card { border-radius:10px; padding:14px 16px; margin-bottom:10px; border:1px solid var(--g200); background:var(--w); }
.strategy-card.urgent { border-color:#FECACA; background:#FFFBFB; }
.strategy-header { display:flex; align-items:center; gap:8px; margin-bottom:8px; }
.strategy-number { width:24px; height:24px; border-radius:6px; display:flex; align-items:center; justify-content:center; font-size:8pt; font-weight:800; color:var(--w); flex-shrink:0; }
.strategy-number.sn-red { background:var(--red); }
.strategy-number.sn-amber { background:var(--amber); }
.strategy-number.sn-green { background:var(--mb); }
.strategy-number.sn-blue { background:#3B82F6; }
.strategy-title { font-size:10pt; font-weight:700; color:var(--dn); }
.strategy-actions { padding-left:32px; }
.strategy-actions li { font-size:8.5pt; color:var(--g700); margin-bottom:3px; line-height:1.45; }
.strategy-impact { margin-top:6px; padding-left:32px; font-size:8pt; font-weight:600; color:var(--mb); }

/* THANKYOU */
.thankyou { background:linear-gradient(160deg,#0F1820 0%,#1A2332 35%,#253448 65%,#1A2332 100%); display:flex; flex-direction:column; justify-content:center; align-items:center; text-align:center; color:var(--w); position:relative; }
.thankyou::before { content:''; position:absolute; inset:0; background:radial-gradient(ellipse 600px 400px at 80% 30%,rgba(201,169,110,0.06) 0%,transparent 70%),radial-gradient(ellipse 500px 500px at 20% 70%,rgba(59,89,152,0.05) 0%,transparent 70%); pointer-events:none; }
.thankyou-content { position:relative; z-index:1; }
.thankyou-title { font-size:28pt; font-weight:800; letter-spacing:8px; text-transform:uppercase; margin-bottom:8px; }
.thankyou-subtitle { font-size:12pt; font-weight:300; color:rgba(255,255,255,0.6); margin-bottom:8px; }
.thankyou-contact { font-size:9pt; color:rgba(255,255,255,0.4); }
.thankyou-contact a { color:var(--gold); text-decoration:none; }

/* HIGHLIGHTS */
.star-row { background:linear-gradient(90deg,#F0FDF4,#DCFCE7) !important; }
.kw-efficient { background:#F0FDF4 !important; }

/* ACCENTS */
.page:not(.cover):not(.thankyou)::after { content:''; position:absolute; left:0; top:0; bottom:0; width:4px; background:linear-gradient(180deg,var(--dn),var(--mb),var(--gold)); z-index:2; }
.page:not(.cover):not(.thankyou)::before { content:''; position:absolute; bottom:0; left:0; right:0; height:3px; background:linear-gradient(90deg,transparent,var(--pb) 20%,var(--mb) 50%,var(--pb) 80%,transparent); opacity:0.4; z-index:2; }
.page:not(.cover):not(.thankyou) .page-inner { background-image:radial-gradient(circle at 100% 0%,rgba(26,35,50,0.02) 0%,transparent 50%),radial-gradient(circle at 0% 100%,rgba(59,89,152,0.015) 0%,transparent 50%); }

/* UTILITY */
.mb-12 { margin-bottom:12px; }
.mb-14 { margin-bottom:14px; }

/* PRINT */
@media print { body { background:white; } .page { margin:0; box-shadow:none; } }
@media screen { .page { box-shadow:0 4px 24px rgba(0,0,0,0.08); margin-bottom:20px; } }
</style>
</head>
<body>
""")

# ==================== PAGE 1: COVER ====================
W(f"""<div class="page cover">
<div class="cover-content">
<img src="data:image/png;base64,{logo_b64}" alt="Harmonices" style="width:100px;height:auto;margin-bottom:20px;filter:brightness(0) invert(1);">
<div class="cover-logo-text">HARMONICES</div>
<div style="font-size:9pt;font-weight:400;letter-spacing:4px;text-transform:uppercase;color:rgba(255,255,255,0.4);margin-bottom:6px;">Senior Coliving</div>
<div style="font-size:8pt;font-weight:500;letter-spacing:3px;text-transform:uppercase;color:rgba(255,255,255,0.3);margin-bottom:40px;">Torrelodones &middot; Las Rozas &middot; Majadahonda</div>
<div class="cover-divider"></div>
<div style="font-size:18pt;font-weight:600;color:var(--gold);margin-bottom:8px;">MARZO 2026</div>
<div style="font-size:11pt;font-weight:500;letter-spacing:3px;text-transform:uppercase;color:rgba(255,255,255,0.5);margin-bottom:16px;">Informe de Resultados</div>
<div style="font-size:10pt;font-weight:400;color:rgba(255,255,255,0.5);letter-spacing:1px;">Meta Ads &middot; Google Ads &middot; HubSpot CRM</div>
</div>
<div class="cover-footer">harmonices.com &middot; Informe Mensual Marzo 2026</div>
<div class="page-footer-dark"><span>Harmonices &mdash; Informe Marzo 2026</span><span>1 / 11</span></div>
</div>""")

# ==================== PAGE 2: EXECUTIVE DASHBOARD ====================
W(f"""<div class="page"><div class="page-inner">
<div class="page-header"><div class="page-header-left"><h2>Executive Dashboard</h2><div class="subtitle">Resumen ejecutivo &mdash; Marzo 2026 vs Febrero 2026</div></div><div class="page-header-right"><div class="brand">HARMONICES</div><div>Marzo 2026</div></div></div>

<div class="kpi-grid">
<div class="kpi-card green"><div class="kpi-label">Inversi&oacute;n Total</div><div class="kpi-value smaller">{total_inv:,.2f}&euro;</div><div class="kpi-change {mc(inv_mom)} {ma(inv_mom)}">{fp(inv_mom)}</div><div class="kpi-prev">Feb: {total_inv_prev:,.2f}&euro;</div></div>
<div class="kpi-card green"><div class="kpi-label">Leads Totales</div><div class="kpi-value">{total_leads}</div><div class="kpi-change {mc(leads_mom)} {ma(leads_mom)}">{fp(leads_mom)}</div><div class="kpi-prev">Feb: {total_leads_prev}</div></div>
<div class="kpi-card red"><div class="kpi-label">Leads V&aacute;lidos CRM</div><div class="kpi-value">{valid_mar}</div><div class="kpi-change bad arrow-down">{fp(valid_mom)}</div><div class="kpi-prev">Feb: {valid_feb}</div></div>
<div class="kpi-card red"><div class="kpi-label">Contratos Firmados</div><div class="kpi-value">{contracts_mar}</div><div class="kpi-change bad arrow-down">-2 vs Feb</div><div class="kpi-prev">Feb: {contracts_feb}</div></div>
</div>

<div class="channel-row">
<div class="channel-box meta-box"><div class="channel-icon meta-icon"><svg width="20" height="20" viewBox="0 0 24 24" fill="white"><path d="M12 2C6.477 2 2 6.477 2 12c0 4.991 3.657 9.128 8.438 9.878v-6.987h-2.54V12h2.54V9.797c0-2.506 1.492-3.89 3.777-3.89 1.094 0 2.238.195 2.238.195v2.46h-1.26c-1.243 0-1.63.771-1.63 1.562V12h2.773l-.443 2.89h-2.33v6.988C18.343 21.128 22 16.991 22 12c0-5.523-4.477-10-10-10z"/></svg></div><div class="channel-data"><div class="channel-name">Meta Ads</div><div class="channel-leads">{m["leads"]["value"]} <span>leads</span></div><div class="channel-cpl">CPL: {m["cpl"]["value"]:.2f}&euro; &middot; Inv: {m["inversion"]["value"]:.2f}&euro;</div></div></div>
<div class="channel-box google-box"><div class="channel-icon google-icon"><svg width="20" height="20" viewBox="0 0 24 24" fill="white"><path d="M12.48 10.92v3.28h7.84c-.24 1.84-.853 3.187-1.787 4.133-1.147 1.147-2.933 2.4-6.053 2.4-4.827 0-8.6-3.893-8.6-8.72s3.773-8.72 8.6-8.72c2.6 0 4.507 1.027 5.907 2.347l2.307-2.307C18.747 1.44 16.133 0 12.48 0 5.867 0 .307 5.387.307 12s5.56 12 12.173 12c3.573 0 6.267-1.173 8.373-3.36 2.16-2.16 2.84-5.213 2.84-7.667 0-.76-.053-1.467-.173-2.053H12.48z"/></svg></div><div class="channel-data"><div class="channel-name">Google Ads</div><div class="channel-leads">{g["leads"]["value"]} <span>leads</span></div><div class="channel-cpl">CPL: {g["cpl"]["value"]:.2f}&euro; &middot; Inv: {g["inversion"]["value"]:,.2f}&euro;</div></div></div>
</div>

<div class="traffic-light"><div class="traffic-light-title">Sem&aacute;foro del Mes</div>
<div class="tl-item"><div class="tl-dot green-dot"></div><div class="tl-text"><strong>Google Ads:</strong> +11,2% leads, CPL baja a 11,14&euro; (-4,4%). Motor principal del embudo con 82% de los leads.</div></div>
<div class="tl-item"><div class="tl-dot amber-dot"></div><div class="tl-text"><strong>Meta Ads:</strong> Leads caen -30% (28 vs 40). CPL sube a 11,01&euro; (+57,5%). Requiere atenci&oacute;n en creativos.</div></div>
<div class="tl-item"><div class="tl-dot red-dot"></div><div class="tl-text"><strong>CRM/Ventas:</strong> Solo 5 leads v&aacute;lidos vs 33 en febrero (-84,8%). 0 contratos firmados. Ca&iacute;da severa en conversi&oacute;n.</div></div>
</div>

<div class="verdict"><div class="verdict-label">Veredicto</div>Marzo mantiene el volumen de leads (157 vs 156) pero la calidad baja dr&aacute;sticamente: solo 5 leads v&aacute;lidos en CRM frente a 33 en febrero. Google Ads crece (+11,2% leads) y mejora eficiencia. Meta Ads pierde -30% de leads con CPL casi duplicado. Prioridad: investigar la ca&iacute;da de calidad en CRM y optimizar Meta.</div>
</div><div class="page-footer"><span>harmonices.com &middot; Informe Mensual Marzo 2026</span><span>2 / 11</span></div></div>""")

# ==================== PAGE 3: CHANNEL MIX ====================
W(f"""<div class="page"><div class="page-inner">
<div class="page-header"><div class="page-header-left"><h2>Channel Mix</h2><div class="subtitle">Distribuci&oacute;n de leads e inversi&oacute;n por canal</div></div><div class="page-header-right"><div class="brand">HARMONICES</div><div>Marzo 2026</div></div></div>

<div class="section-title">Distribuci&oacute;n de Leads por Canal</div>
<div class="split mb-14">
<div class="mix-block meta-mix"><div class="mix-pct meta-pct">{meta_pct:.0f}%</div><div class="mix-label">Meta Ads</div><div class="mix-leads">{m["leads"]["value"]} leads</div><div style="font-size:8pt;color:var(--g500);margin-top:4px;">CPL: {m["cpl"]["value"]:.2f}&euro; &middot; Inv: {m["inversion"]["value"]:.2f}&euro;</div></div>
<div class="mix-block google-mix"><div class="mix-pct google-pct">{google_pct:.0f}%</div><div class="mix-label">Google Ads</div><div class="mix-leads">{g["leads"]["value"]} leads</div><div style="font-size:8pt;color:var(--g500);margin-top:4px;">CPL: {g["cpl"]["value"]:.2f}&euro; &middot; Inv: {g["inversion"]["value"]:,.2f}&euro;</div></div>
</div>

<div class="section-title">Distribuci&oacute;n de Inversi&oacute;n</div>
<div class="bar-chart">
<div class="bar-row"><div class="bar-label">Meta Ads</div><div class="bar-track"><div class="bar-fill meta-fill" style="width:{meta_inv_pct:.1f}%">{m["inversion"]["value"]:.2f}&euro; ({meta_inv_pct:.0f}%)</div></div></div>
<div class="bar-row"><div class="bar-label">Google Ads</div><div class="bar-track"><div class="bar-fill google-fill" style="width:{google_inv_pct:.1f}%">{g["inversion"]["value"]:,.2f}&euro; ({google_inv_pct:.0f}%)</div></div></div>
</div>

<div class="section-title">Eficiencia por Canal</div>
<table class="data-table"><thead><tr><th>Canal</th><th class="num">Inversi&oacute;n</th><th class="num">Leads</th><th class="num">CPL</th><th class="num">Impresiones</th><th class="num">Clics</th><th class="num">CTR</th><th class="num">CPC</th></tr></thead>
<tbody>
<tr><td class="bold">Meta Ads</td><td class="num">{m["inversion"]["value"]:.2f}&euro;</td><td class="num bold">{m["leads"]["value"]}</td><td class="num">{m["cpl"]["value"]:.2f}&euro;</td><td class="num">{m["impresiones"]["value"]:,}</td><td class="num">{m["clics"]["value"]:,}</td><td class="num">{m["ctr"]["value"]:.2f}%</td><td class="num">{m["cpc"]["value"]:.2f}&euro;</td></tr>
<tr><td class="bold">Google Ads</td><td class="num">{g["inversion"]["value"]:,.2f}&euro;</td><td class="num bold">{g["leads"]["value"]}</td><td class="num">{g["cpl"]["value"]:.2f}&euro;</td><td class="num">{g["impresiones"]["value"]:,}</td><td class="num">{g["clics"]["value"]:,}</td><td class="num">{g["ctr"]["value"]:.2f}%</td><td class="num">{g["cpc"]["value"]:.2f}&euro;</td></tr>
</tbody>
<tfoot><tr><td class="bold">TOTAL</td><td class="num bold">{total_inv:,.2f}&euro;</td><td class="num bold">{total_leads}</td><td class="num bold">{total_cpl:.2f}&euro;</td><td class="num bold">{m["impresiones"]["value"]+g["impresiones"]["value"]:,}</td><td class="num bold">{m["clics"]["value"]+g["clics"]["value"]:,}</td><td class="num">&mdash;</td><td class="num">&mdash;</td></tr></tfoot>
</table>

<div class="insight-box amber-insight"><div class="insight-label">Insight</div>Google Ads genera el {google_pct:.0f}% de los leads con el {google_inv_pct:.0f}% de la inversi&oacute;n. Meta Ads, con solo el {meta_inv_pct:.0f}% del presupuesto, aporta {meta_pct:.0f}% de los leads pero su CPL ha subido significativamente (+57,5%). La eficiencia de Meta se ha deteriorado respecto a febrero.</div>
</div><div class="page-footer"><span>harmonices.com &middot; Informe Mensual Marzo 2026</span><span>3 / 11</span></div></div>""")

# ==================== PAGE 4: META ADS ====================
meta_daily_svg = ""
max_m = max(day["leads"] for day in m_daily) or 1
for i, day in enumerate(m_daily):
    x = 35 + i * 18.7
    h = (day["leads"] / max_m) * 75
    y = 95 - h
    meta_daily_svg += f'<rect x="{x:.1f}" y="{y:.1f}" width="17" height="{h:.1f}" rx="2" fill="#4F46E5" opacity="0.8"/>'
    if i % 5 == 0:
        meta_daily_svg += f'<text x="{x+8.5:.1f}" y="108" font-size="6" fill="#9CA3AF" text-anchor="middle">{i+1}</text>'

platform_rows = ""
for cp in camp_plat:
    star = ' class="star-row"' if cp["leads"] >= 15 else ""
    platform_rows += f'<tr{star}><td class="bold">{cp["platform"]}</td><td class="num">{cp["impresiones"]:,}</td><td class="num">{cp["clics"]:,}</td><td class="num">{cp["ctr"]:.2f}%</td><td class="num">{cp["cpc"]:.2f}&euro;</td><td class="num bold">{cp["leads"]}</td><td class="num">{cp["cpl"]:.2f}&euro;</td><td class="num">{cp["inversion"]:.2f}&euro;</td></tr>'

W(f"""<div class="page"><div class="page-inner">
<div class="page-header"><div class="page-header-left"><h2>Meta Ads</h2><div class="subtitle">Rendimiento Facebook &amp; Instagram &mdash; Marzo 2026</div></div><div class="page-header-right"><div class="brand">HARMONICES</div><div>Marzo 2026</div></div></div>

<div class="kpi-row-6">
<div class="kpi-mini"><div class="kpi-label">Inversi&oacute;n</div><div class="kpi-value">{m["inversion"]["value"]:.0f}&euro;</div><div class="kpi-change {mc(m_inv_mom)} {ma(m_inv_mom)}">{fp(m_inv_mom)}</div></div>
<div class="kpi-mini"><div class="kpi-label">Leads</div><div class="kpi-value">{m["leads"]["value"]}</div><div class="kpi-change {mc(m_leads_mom)} {ma(m_leads_mom)}">{fp(m_leads_mom)}</div></div>
<div class="kpi-mini"><div class="kpi-label">CPL</div><div class="kpi-value">{m["cpl"]["value"]:.2f}&euro;</div><div class="kpi-change {mc(m_cpl_mom, inv=True)} {ma(m_cpl_mom)}">{fp(m_cpl_mom)}</div></div>
<div class="kpi-mini"><div class="kpi-label">CTR</div><div class="kpi-value">{m["ctr"]["value"]:.2f}%</div><div class="kpi-change {mc(m_ctr_mom)} {ma(m_ctr_mom)}">{fp(m_ctr_mom)}</div></div>
<div class="kpi-mini"><div class="kpi-label">CPC</div><div class="kpi-value">{m["cpc"]["value"]:.2f}&euro;</div><div class="kpi-change {mc(m_cpc_mom, inv=True)} {ma(m_cpc_mom)}">{fp(m_cpc_mom)}</div></div>
<div class="kpi-mini"><div class="kpi-label">Alcance</div><div class="kpi-value">{m["alcance"]["value"]:,}</div><div class="kpi-change {mc(m_alcance_mom)} {ma(m_alcance_mom)}">{fp(m_alcance_mom)}</div></div>
</div>

<div class="section-title">Rendimiento por Plataforma</div>
<table class="data-table mb-12"><thead><tr><th>Plataforma</th><th class="num">Impresiones</th><th class="num">Clics</th><th class="num">CTR</th><th class="num">CPC</th><th class="num">Leads</th><th class="num">CPL</th><th class="num">Inversi&oacute;n</th></tr></thead>
<tbody>{platform_rows}</tbody>
<tfoot><tr><td class="bold">TOTAL</td><td class="num bold">{m["impresiones"]["value"]:,}</td><td class="num bold">{m["clics"]["value"]:,}</td><td class="num bold">{m["ctr"]["value"]:.2f}%</td><td class="num bold">{m["cpc"]["value"]:.2f}&euro;</td><td class="num bold">{m["leads"]["value"]}</td><td class="num bold">{m["cpl"]["value"]:.2f}&euro;</td><td class="num bold">{m["inversion"]["value"]:.2f}&euro;</td></tr></tfoot></table>

<div class="section-title">Leads Diarios &mdash; Meta Ads</div>
<div class="chart-container"><svg viewBox="0 0 620 120" style="width:100%;height:auto;">
<line x1="30" y1="10" x2="610" y2="10" stroke="#E5E7EB" stroke-width="0.5"/><line x1="30" y1="37" x2="610" y2="37" stroke="#E5E7EB" stroke-width="0.5"/><line x1="30" y1="64" x2="610" y2="64" stroke="#E5E7EB" stroke-width="0.5"/><line x1="30" y1="91" x2="610" y2="91" stroke="#E5E7EB" stroke-width="0.5"/>
<text x="25" y="14" font-size="7" fill="#9CA3AF" text-anchor="end">4</text><text x="25" y="41" font-size="7" fill="#9CA3AF" text-anchor="end">3</text><text x="25" y="68" font-size="7" fill="#9CA3AF" text-anchor="end">2</text><text x="25" y="95" font-size="7" fill="#9CA3AF" text-anchor="end">1</text>
{meta_daily_svg}
</svg></div>

<div class="insight-box red-insight"><div class="insight-label">Alerta</div>Meta Ads pierde -30% de leads (28 vs 40 en febrero). El CPL sube de 6,99&euro; a 11,01&euro; (+57,5%). Facebook sigue siendo m&aacute;s eficiente (CPL 9,82&euro;) que Instagram (CPL 12,98&euro;). Se recomienda renovar creativos y revisar la segmentaci&oacute;n.</div>
</div><div class="page-footer"><span>harmonices.com &middot; Informe Mensual Marzo 2026</span><span>4 / 11</span></div></div>""")

# ==================== PAGE 5: CREATIVE DEEP DIVE ====================
total_creative_leads = sum(c["leads"] for c in creatives)
creative_rows = ""
for c in creatives:
    pct = (c["leads"] / total_creative_leads * 100) if total_creative_leads else 0
    star = ' class="star-row"' if c["leads"] >= 15 else ""
    creative_rows += f'<tr{star}><td class="bold">{c["name"]}</td><td class="num">{c["impresiones"]:,}</td><td class="num">{c["clics"]:,}</td><td class="num">{c["ctr"]:.2f}%</td><td class="num bold">{c["leads"]}</td><td class="num">{c["cpl"]:.2f}&euro;</td><td class="num">{c["inversion"]:.2f}&euro;</td><td class="num">{pct:.0f}%</td></tr>'

W(f"""<div class="page"><div class="page-inner">
<div class="page-header"><div class="page-header-left"><h2>Creative Deep Dive</h2><div class="subtitle">An&aacute;lisis de creativos Meta Ads &mdash; Marzo 2026</div></div><div class="page-header-right"><div class="brand">HARMONICES</div><div>Marzo 2026</div></div></div>

<div class="section-title">Rendimiento por Creativo</div>
<table class="data-table mb-12"><thead><tr><th>Creativo</th><th class="num">Impresiones</th><th class="num">Clics</th><th class="num">CTR</th><th class="num">Leads</th><th class="num">CPL</th><th class="num">Inversi&oacute;n</th><th class="num">% Leads</th></tr></thead>
<tbody>{creative_rows}</tbody>
<tfoot><tr><td class="bold">TOTAL</td><td class="num bold">{sum(c["impresiones"] for c in creatives):,}</td><td class="num bold">{sum(c["clics"] for c in creatives):,}</td><td class="num">&mdash;</td><td class="num bold">{total_creative_leads}</td><td class="num bold">{m["cpl"]["value"]:.2f}&euro;</td><td class="num bold">{m["inversion"]["value"]:.2f}&euro;</td><td class="num bold">100%</td></tr></tfoot></table>

<div class="section-title">An&aacute;lisis por Formato</div>
<div class="format-row">
<div class="format-pill success"><div class="format-name">Video (2 creativos)</div><div class="format-value">26</div><div class="format-detail">leads &middot; CPL medio: 11,47&euro;</div></div>
<div class="format-pill success"><div class="format-name">Carrusel (1 creativo)</div><div class="format-value">2</div><div class="format-detail">leads &middot; CPL: 4,95&euro;</div></div>
</div>

<div class="insight-box green-insight"><div class="insight-label">Top Performer</div><strong>&quot;Video sept 25&quot;</strong> genera el 71% de los leads de Meta (20 de 28) con CPL de 11,60&euro;. Acumula el 76% de las impresiones, lo que sugiere que el algoritmo lo favorece. Sin embargo, lleva 6 meses activo &mdash; posible fatiga creativa.</div>
<div class="insight-box amber-insight"><div class="insight-label">Oportunidad</div>El <strong>&quot;Carrusel nov 25&quot;</strong> tiene el mejor CPL (4,95&euro;) con solo 592 impresiones. Recibe muy poca distribuci&oacute;n. Aumentar su presupuesto podr&iacute;a mejorar la eficiencia global del canal.</div>
<div class="insight-box blue-insight"><div class="insight-label">Recomendaci&oacute;n</div>Con un CPL global que sube de 6,99&euro; a 11,01&euro;, se recomienda: 1) Producir nuevos creativos de video para combatir la fatiga. 2) Dar m&aacute;s presupuesto al carrusel que tiene CPL de 4,95&euro;. 3) Evaluar si la audiencia de prospecting necesita renovaci&oacute;n.</div>
</div><div class="page-footer"><span>harmonices.com &middot; Informe Mensual Marzo 2026</span><span>5 / 11</span></div></div>""")

# ==================== PAGE 6: GOOGLE ADS ====================
google_daily_svg = ""
max_g = max(day["leads"] for day in g_daily) or 1
for i, day in enumerate(g_daily):
    x = 35 + i * 18.7
    h = (day["leads"] / max_g) * 75
    y = 95 - h
    google_daily_svg += f'<rect x="{x:.1f}" y="{y:.1f}" width="17" height="{h:.1f}" rx="2" fill="#D97706" opacity="0.8"/>'
    if i % 5 == 0:
        google_daily_svg += f'<text x="{x+8.5:.1f}" y="108" font-size="6" fill="#9CA3AF" text-anchor="middle">{i+1}</text>'

gcampaign_rows = ""
for gc in g_campaigns:
    star = ' class="star-row"' if gc["leads"] >= 100 else ""
    gcampaign_rows += f'<tr{star}><td class="bold" style="font-size:7.5pt;">{gc["name"]}</td><td class="num">{gc["impresiones"]:,}</td><td class="num">{gc["clics"]:,}</td><td class="num">{gc["ctr"]:.2f}%</td><td class="num">{gc["cpc"]:.2f}&euro;</td><td class="num bold">{gc["leads"]}</td><td class="num">{gc["cpl"]:.2f}&euro;</td><td class="num">{gc["inversion"]:,.2f}&euro;</td></tr>'

W(f"""<div class="page"><div class="page-inner">
<div class="page-header"><div class="page-header-left"><h2>Google Ads</h2><div class="subtitle">Rendimiento Search &amp; PMax &mdash; Marzo 2026</div></div><div class="page-header-right"><div class="brand">HARMONICES</div><div>Marzo 2026</div></div></div>

<div class="kpi-row-6">
<div class="kpi-mini"><div class="kpi-label">Inversi&oacute;n</div><div class="kpi-value" style="font-size:11pt;">{g["inversion"]["value"]:,.0f}&euro;</div><div class="kpi-change {mc(g_inv_mom)} {ma(g_inv_mom)}">{fp(g_inv_mom)}</div></div>
<div class="kpi-mini"><div class="kpi-label">Leads</div><div class="kpi-value">{g["leads"]["value"]}</div><div class="kpi-change {mc(g_leads_mom)} {ma(g_leads_mom)}">{fp(g_leads_mom)}</div></div>
<div class="kpi-mini"><div class="kpi-label">CPL</div><div class="kpi-value">{g["cpl"]["value"]:.2f}&euro;</div><div class="kpi-change {mc(g_cpl_mom, inv=True)} {ma(g_cpl_mom)}">{fp(g_cpl_mom)}</div></div>
<div class="kpi-mini"><div class="kpi-label">CTR</div><div class="kpi-value">{g["ctr"]["value"]:.2f}%</div><div class="kpi-change {mc(g_ctr_mom)} {ma(g_ctr_mom)}">{fp(g_ctr_mom)}</div></div>
<div class="kpi-mini"><div class="kpi-label">CPC</div><div class="kpi-value">{g["cpc"]["value"]:.2f}&euro;</div><div class="kpi-change {mc(g_cpc_mom, inv=True)} {ma(g_cpc_mom)}">{fp(g_cpc_mom)}</div></div>
<div class="kpi-mini"><div class="kpi-label">Impresiones</div><div class="kpi-value" style="font-size:11pt;">{g["impresiones"]["value"]:,}</div><div class="kpi-change {mc(g_imp_mom)} {ma(g_imp_mom)}">{fp(g_imp_mom)}</div></div>
</div>

<div class="section-title">Campa&ntilde;as Google Ads</div>
<table class="data-table mb-12"><thead><tr><th>Campa&ntilde;a</th><th class="num">Impresiones</th><th class="num">Clics</th><th class="num">CTR</th><th class="num">CPC</th><th class="num">Leads</th><th class="num">CPL</th><th class="num">Inversi&oacute;n</th></tr></thead>
<tbody>{gcampaign_rows}</tbody>
<tfoot><tr><td class="bold">TOTAL</td><td class="num bold">{g["impresiones"]["value"]:,}</td><td class="num bold">{g["clics"]["value"]:,}</td><td class="num bold">{g["ctr"]["value"]:.2f}%</td><td class="num bold">{g["cpc"]["value"]:.2f}&euro;</td><td class="num bold">{g["leads"]["value"]}</td><td class="num bold">{g["cpl"]["value"]:.2f}&euro;</td><td class="num bold">{g["inversion"]["value"]:,.2f}&euro;</td></tr></tfoot></table>

<div class="section-title">Leads Diarios &mdash; Google Ads</div>
<div class="chart-container"><svg viewBox="0 0 620 120" style="width:100%;height:auto;">
<line x1="30" y1="10" x2="610" y2="10" stroke="#E5E7EB" stroke-width="0.5"/><line x1="30" y1="37" x2="610" y2="37" stroke="#E5E7EB" stroke-width="0.5"/><line x1="30" y1="64" x2="610" y2="64" stroke="#E5E7EB" stroke-width="0.5"/><line x1="30" y1="91" x2="610" y2="91" stroke="#E5E7EB" stroke-width="0.5"/>
<text x="25" y="14" font-size="7" fill="#9CA3AF" text-anchor="end">10</text><text x="25" y="41" font-size="7" fill="#9CA3AF" text-anchor="end">7</text><text x="25" y="68" font-size="7" fill="#9CA3AF" text-anchor="end">4</text><text x="25" y="95" font-size="7" fill="#9CA3AF" text-anchor="end">1</text>
{google_daily_svg}
</svg></div>

<div class="insight-box green-insight"><div class="insight-label">Rendimiento</div>Google Ads es el motor principal con 129 leads (+11,2% MoM). La campa&ntilde;a PMax genera el 98% de los leads (127) con un CPL excelente de 10,37&euro;. La campa&ntilde;a de b&uacute;squeda (Search) aporta solo 2 leads con CPL alto de 59,71&euro;.</div>
</div><div class="page-footer"><span>harmonices.com &middot; Informe Mensual Marzo 2026</span><span>6 / 11</span></div></div>""")

# ==================== PAGE 7: KEYWORDS ====================
kw_rows = ""
for kw in kw_sorted:
    ec = ' class="kw-efficient"' if kw["leads"] > 0 else ""
    cpl_str = f'{kw["cpl"]:.2f}&euro;' if kw["cpl"] > 0 else "&mdash;"
    kw_rows += f'<tr{ec}><td class="bold" style="font-size:7.5pt;">{kw["keyword"]}</td><td class="num">{kw["impresiones"]}</td><td class="num">{kw["clics"]}</td><td class="num">{kw["ctr"]:.2f}%</td><td class="num">{kw["cpc"]:.2f}&euro;</td><td class="num bold">{kw["leads"]}</td><td class="num">{cpl_str}</td><td class="num">{kw["inversion"]:.2f}&euro;</td></tr>'

W(f"""<div class="page"><div class="page-inner">
<div class="page-header"><div class="page-header-left"><h2>Keywords &amp; Search Terms</h2><div class="subtitle">Rendimiento de palabras clave (Search) &mdash; Marzo 2026</div></div><div class="page-header-right"><div class="brand">HARMONICES</div><div>Marzo 2026</div></div></div>

<div class="section-title">Top Keywords por Clics</div>
<table class="data-table"><thead><tr><th>Keyword</th><th class="num">Impresiones</th><th class="num">Clics</th><th class="num">CTR</th><th class="num">CPC</th><th class="num">Leads</th><th class="num">CPL</th><th class="num">Inversi&oacute;n</th></tr></thead>
<tbody>{kw_rows}</tbody></table>

<div class="insight-box blue-insight"><div class="insight-label">An&aacute;lisis</div>Las keywords de la campa&ntilde;a Search generan buen tr&aacute;fico con CTRs altos (16-30%) pero la conversi&oacute;n a leads es baja: solo 2 de 40 keywords con leads. &quot;cohousing senior&quot; (CTR 16,5%, 1 lead, CPL 6,91&euro;) y &quot;cohousing jubilados&quot; (CTR 50%, 1 lead, CPL 0,89&euro;) son las &uacute;nicas que convierten. El grueso de conversiones viene de PMax.</div>
<div class="insight-box amber-insight"><div class="insight-label">Oportunidad</div>&quot;cohousing madrid&quot; tiene 45 clics y 24,46% CTR sin convertir a lead. &quot;coliving madrid mayores&quot; destaca con 30,3% CTR. Estas keywords muestran alta intenci&oacute;n pero falta optimizar la landing page para mejorar la conversi&oacute;n.</div>
</div><div class="page-footer"><span>harmonices.com &middot; Informe Mensual Marzo 2026</span><span>7 / 11</span></div></div>""")

# ==================== PAGE 8: HUBSPOT ====================
W(f"""<div class="page"><div class="page-inner">
<div class="page-header"><div class="page-header-left"><h2>HubSpot CRM</h2><div class="subtitle">Embudo de ventas y estado del pipeline &mdash; Marzo 2026</div></div><div class="page-header-right"><div class="brand">HARMONICES</div><div>Marzo 2026</div></div></div>

<div class="section-title">Movimientos del Mes (Marzo 2026)</div>
<div class="funnel">
<div class="funnel-step"><div class="funnel-bar f1" style="width:100%;">97 leads entrantes</div><div class="funnel-label">Nuevos leads</div></div>
<div class="funnel-step"><div class="funnel-bar f2" style="width:{24/97*100:.0f}%;">24 en exploratorio</div><div class="funnel-label">Exploratorio</div></div>
<div class="funnel-step"><div class="funnel-bar f3" style="width:{5/97*100:.0f}%;">5 interesado futuro</div><div class="funnel-label">Interesado futuro</div></div>
<div class="funnel-step"><div class="funnel-bar f4" style="width:3%;">0 reservas</div><div class="funnel-label">Reservado/pagado</div></div>
</div>

<div class="split mb-14">
<div><div class="section-subtitle">Descartados este mes</div><div style="font-size:20pt;font-weight:800;color:var(--red);">25</div><div style="font-size:8pt;color:var(--g500);">leads descartados en marzo</div></div>
<div><div class="section-subtitle">Tasa de descarte</div><div style="font-size:20pt;font-weight:800;color:var(--amber);">{25/97*100:.0f}%</div><div style="font-size:8pt;color:var(--g500);">de los leads entrantes descartados</div></div>
</div>

<div class="section-title">Stock Total del Pipeline</div>
<table class="data-table mb-12"><thead><tr><th>Etapa</th><th class="num">Total acumulado</th><th class="num">% del total</th></tr></thead>
<tbody>
<tr><td class="bold">Lead Entrante</td><td class="num">{hs_stock["lead_entrante"]}</td><td class="num">{hs_stock["lead_entrante"]/total_active*100:.1f}%</td></tr>
<tr><td class="bold">Exploratorio</td><td class="num">{hs_stock["exploratorio"]}</td><td class="num">{hs_stock["exploratorio"]/total_active*100:.1f}%</td></tr>
<tr><td class="bold">Interesado Futuro</td><td class="num">{hs_stock["interesado_futuro"]}</td><td class="num">{hs_stock["interesado_futuro"]/total_active*100:.1f}%</td></tr>
<tr><td class="bold">Interesado Caliente</td><td class="num">{hs_stock["interesado_caliente"]}</td><td class="num">{hs_stock["interesado_caliente"]/total_active*100:.1f}%</td></tr>
<tr class="star-row"><td class="bold">Reservado / Pagado</td><td class="num bold">{hs_stock["reservado_pagado"]}</td><td class="num">{hs_stock["reservado_pagado"]/total_active*100:.1f}%</td></tr>
</tbody>
<tfoot><tr><td class="bold">TOTAL ACTIVOS</td><td class="num bold">{total_active}</td><td class="num bold">100%</td></tr></tfoot></table>

<div class="insight-box red-insight"><div class="insight-label">Alerta CRM</div>Solo 5 leads v&aacute;lidos (Interesado Futuro) en marzo, frente a 26 en febrero y 28 en enero. Cero contratos firmados. El pipeline acumula 2.909 descartados hist&oacute;ricos. Hay 234 deals en exploratorio que necesitan seguimiento activo.</div>
</div><div class="page-footer"><span>harmonices.com &middot; Informe Mensual Marzo 2026</span><span>8 / 11</span></div></div>""")

# ==================== PAGE 9: HISTORICAL DATA ====================
g_hist_rows = ""
for h in g_hist:
    ms = h["month"][:7]; y=ms[:4]; mo=ms[5:7]
    label = f"{month_names.get(mo,mo)} {y[2:]}"
    star = ' class="star-row"' if ms == "2026-03" else ""
    g_hist_rows += f'<tr{star}><td class="bold">{label}</td><td class="num">{h["impresiones"]:,}</td><td class="num">{h["clics"]:,}</td><td class="num">{h["ctr"]:.2f}%</td><td class="num">{h["cpc"]:.2f}&euro;</td><td class="num bold">{h["leads"]}</td><td class="num">{h["cpl"]:.2f}&euro;</td><td class="num">{h["inversion"]:,.2f}&euro;</td></tr>'

m_hist_rows = ""
for h in m_hist:
    ms = h["month"]; y=ms[:4]; mo=ms[5:7]
    label = f"{month_names.get(mo,mo)} {y[2:]}"
    star = ' class="star-row"' if ms == "2026-03" else ""
    m_hist_rows += f'<tr{star}><td class="bold">{label}</td><td class="num">{h["impresiones"]:,}</td><td class="num">{h["clics"]:,}</td><td class="num">{h["ctr"]:.2f}%</td><td class="num">{h["cpc"]:.2f}&euro;</td><td class="num bold">{h["leads"]}</td><td class="num">{h["cpl"]:.2f}&euro;</td><td class="num">{h["inversion"]:.2f}&euro;</td></tr>'

W(f"""<div class="page"><div class="page-inner">
<div class="page-header"><div class="page-header-left"><h2>Datos Hist&oacute;ricos</h2><div class="subtitle">Evoluci&oacute;n mensual &mdash; Google Ads &amp; Meta Ads</div></div><div class="page-header-right"><div class="brand">HARMONICES</div><div>Marzo 2026</div></div></div>

<div class="section-title">Google Ads &mdash; &Uacute;ltimos 15 Meses</div>
<table class="data-table" style="font-size:7.5pt;"><thead><tr><th>Mes</th><th class="num">Impresiones</th><th class="num">Clics</th><th class="num">CTR</th><th class="num">CPC</th><th class="num">Leads</th><th class="num">CPL</th><th class="num">Inversi&oacute;n</th></tr></thead>
<tbody>{g_hist_rows}</tbody></table>

<div class="section-title">Meta Ads &mdash; &Uacute;ltimos 11 Meses</div>
<table class="data-table" style="font-size:7.5pt;"><thead><tr><th>Mes</th><th class="num">Impresiones</th><th class="num">Clics</th><th class="num">CTR</th><th class="num">CPC</th><th class="num">Leads</th><th class="num">CPL</th><th class="num">Inversi&oacute;n</th></tr></thead>
<tbody>{m_hist_rows}</tbody></table>
</div><div class="page-footer"><span>harmonices.com &middot; Informe Mensual Marzo 2026</span><span>9 / 11</span></div></div>""")

# ==================== PAGE 10: STRATEGY ====================
yoy_leads_pct = (129/28-1)*100
yoy_cpl_pct = (11.14/48.87-1)*100
yoy_inv_pct = (1436.56/1368.38-1)*100

W(f"""<div class="page"><div class="page-inner">
<div class="page-header"><div class="page-header-left"><h2>Estrategia &amp; Pr&oacute;ximos Pasos</h2><div class="subtitle">Recomendaciones basadas en datos &mdash; Abril 2026</div></div><div class="page-header-right"><div class="brand">HARMONICES</div><div>Marzo 2026</div></div></div>

<div class="section-title">Interanual Google Ads: Marzo 2025 vs Marzo 2026</div>
<div class="yoy-grid">
<div class="yoy-card"><div class="yoy-metric">Leads</div><div class="yoy-values"><span class="yoy-old">28</span><span class="yoy-arrow">&rarr;</span><span class="yoy-new">129</span></div><div class="yoy-change pos-good">+{yoy_leads_pct:.0f}%</div></div>
<div class="yoy-card"><div class="yoy-metric">CPL</div><div class="yoy-values"><span class="yoy-old">48,87&euro;</span><span class="yoy-arrow">&rarr;</span><span class="yoy-new">11,14&euro;</span></div><div class="yoy-change neg-good">{yoy_cpl_pct:.0f}%</div></div>
<div class="yoy-card"><div class="yoy-metric">Inversi&oacute;n</div><div class="yoy-values"><span class="yoy-old">1.368&euro;</span><span class="yoy-arrow">&rarr;</span><span class="yoy-new">1.437&euro;</span></div><div class="yoy-change pos-bad">+{yoy_inv_pct:.0f}%</div></div>
<div class="yoy-card"><div class="yoy-metric">Impresiones</div><div class="yoy-values"><span class="yoy-old">28.956</span><span class="yoy-arrow">&rarr;</span><span class="yoy-new">331.153</span></div><div class="yoy-change pos-good">+1.044%</div></div>
</div>

<div class="section-title">Plan de Acci&oacute;n &mdash; Abril 2026</div>

<div class="strategy-card urgent">
<div class="strategy-header"><div class="strategy-number sn-red">1</div><div class="strategy-title">Investigar ca&iacute;da de calidad en CRM</div></div>
<ul class="strategy-actions">
<li>Solo 5 leads v&aacute;lidos vs 33 en febrero: ca&iacute;da del 85%. Auditar el origen y la segmentaci&oacute;n de los leads entrantes.</li>
<li>Revisar si el formulario de captaci&oacute;n ha cambiado o si el perfil del lead se ha desviado del target senior.</li>
<li>Implementar scoring de leads para filtrar antes de pasar a ventas.</li>
</ul>
<div class="strategy-impact">Impacto esperado: Recuperar &gt;20 leads v&aacute;lidos/mes y reactivar contratos.</div>
</div>

<div class="strategy-card">
<div class="strategy-header"><div class="strategy-number sn-amber">2</div><div class="strategy-title">Renovar creativos en Meta Ads</div></div>
<ul class="strategy-actions">
<li>El video principal (&quot;Video sept 25&quot;) lleva 6+ meses activo. Producir 2-3 nuevos videos con testimonios reales de residentes.</li>
<li>Escalar el Carrusel (CPL 4,95&euro;) aumentando su porcentaje de presupuesto del 3% al 20%.</li>
<li>Probar nuevos formatos: Reels, Stories, instant experiences con tour virtual.</li>
</ul>
<div class="strategy-impact">Impacto esperado: Reducir CPL Meta de 11,01&euro; a &lt;8&euro; y recuperar volumen de leads.</div>
</div>

<div class="strategy-card">
<div class="strategy-header"><div class="strategy-number sn-green">3</div><div class="strategy-title">Optimizar Google Ads Search</div></div>
<ul class="strategy-actions">
<li>La campa&ntilde;a Search tiene CPL de 59,71&euro; con solo 2 leads. Evaluar si mantenerla o reasignar presupuesto a PMax.</li>
<li>Keywords con alto CTR sin conversi&oacute;n (&quot;cohousing madrid&quot; 24,5% CTR, 0 leads): optimizar landing page espec&iacute;fica.</li>
<li>PMax funciona excelente (CPL 10,37&euro;, 127 leads): considerar incrementar su presupuesto un 10-15%.</li>
</ul>
<div class="strategy-impact">Impacto esperado: Reasignar 119&euro;/mes de Search a PMax para +10-12 leads adicionales.</div>
</div>

<div class="strategy-card">
<div class="strategy-header"><div class="strategy-number sn-blue">4</div><div class="strategy-title">Activar el pipeline de exploratorio</div></div>
<ul class="strategy-actions">
<li>234 deals en exploratorio en el pipeline total. Lanzar secuencia de nurturing por email con contenido educativo sobre senior coliving.</li>
<li>Programar jornadas de puertas abiertas en Torrelodones, Las Rozas y Majadahonda para convertir exploratorios.</li>
<li>Las visitas cayeron a 11 en marzo (vs 20 en febrero). Incentivar visitas con ofertas early bird o descuentos por reserva anticipada.</li>
</ul>
<div class="strategy-impact">Impacto esperado: Convertir 5-10% del pipeline exploratorio en interesados calientes en Q2.</div>
</div>
</div><div class="page-footer"><span>harmonices.com &middot; Informe Mensual Marzo 2026</span><span>10 / 11</span></div></div>""")

# ==================== PAGE 11: THANK YOU ====================
W(f"""<div class="page thankyou">
<div class="thankyou-content">
<img src="data:image/png;base64,{logo_b64}" alt="Harmonices" style="width:90px;height:auto;margin-bottom:20px;filter:brightness(0) invert(1);">
<div class="thankyou-title">Harmonices</div>
<div class="thankyou-subtitle">Senior Coliving</div>
<div style="font-size:9pt;color:rgba(255,255,255,0.35);letter-spacing:3px;text-transform:uppercase;margin-bottom:24px;">Torrelodones &middot; Las Rozas &middot; Majadahonda</div>
<div style="width:60px;height:2px;background:var(--gold);margin:0 auto 24px;opacity:0.5;"></div>
<div class="thankyou-contact">
<a href="https://harmonices.com">harmonices.com</a><br>
<span style="color:rgba(255,255,255,0.3);">Informe Mensual &mdash; Marzo 2026</span>
</div>
</div>
<div class="page-footer-dark"><span>Harmonices &mdash; Informe Marzo 2026</span><span>11 / 11</span></div>
</div>""")

W("</body></html>")

# ========== WRITE FILE ==========
output = os.path.join(BASE, "reports/harmonices/Harmonices_Marzo_2026_v1.html")
with open(output, "w", encoding="utf-8") as out:
    out.write("\n".join(parts))

size = os.path.getsize(output)
print(f"Report written: {output}")
print(f"Size: {size:,} bytes ({size/1024:.1f} KB)")
print(f"Pages: 11")

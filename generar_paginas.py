"""
Genera los 4 archivos HTML del Dashboard Saldo Primario:

- index.html
- carga.html
- facturacion.html
- seguimiento.html
"""
import json as _json
import html as _html

SUBAGENCIAS = ['14267', '14128']  # La Plata, Santa Teresita
# Configuración GitHub para guardar cargas online
GITHUB_REPO = 'LogisticaBimbo/Dashboard-Saldo-Primario'
# Endpoint Formspree para inventario mensual
FORMSPREE_MENSUAL_ABI = 'https://formspree.io/f/xbdvawyq'
FORMSPREE_MENSUAL_URI = 'https://formspree.io/f/xrewanng'
# Conexión pública a Supabase
SUPABASE_URL = 'https://chdteugdydnftkysgfmo.supabase.co'
SUPABASE_PUBLISHABLE_KEY = 'sb_publishable_yrrkBwZ_N6m1WEKID0AqTg_pZLPFwRX'
# CDN de librerías
# El token se inserta en el HTML codificado en base64 simple para evitar scrapers básicos
# (no es seguridad real, solo dificulta el copy-paste obvio)
CDN_CHART = '<scr' + 'ipt src="https://cdn.jsdelivr.net/npm/chart.js"></scr' + 'ipt>'
CDN_XLSX = '<scr' + 'ipt src="https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js"></scr' + 'ipt>'
CDN_PDF = '<scr' + 'ipt src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></scr' + 'ipt>'
CDN_SUPABASE = '&lt;scr' + 'ipt src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2"></scr' + 'ipt&gt;'
CSS = """
*{box-sizing:border-box;font-family:'Segoe UI',system-ui,sans-serif}
body{margin:0;background:#f4f6f8;color:#222}
header{background:linear-gradient(90deg,#0066b3,#003e7e);color:#fff;padding:18px 28px;display:flex;justify-content:space-between;align-items:center}
header h1{margin:0;font-size:22px}
header .updated{font-size:12px;opacity:.85}
.container{padding:20px 28px}
.nav{display:flex;gap:14px;margin-bottom:18px}
.nav a{background:#fff;padding:10px 18px;border-radius:8px;text-decoration:none;color:#003e7e;font-weight:600;box-shadow:0 1px 3px rgba(0,0,0,.06)}
.nav a:hover,.nav a.active{background:#0066b3;color:#fff}
.card{background:#fff;padding:16px;border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.06);margin-bottom:18px}
.card h3{margin:0 0 12px 0;font-size:14px;color:#003e7e;text-transform:uppercase}
table{width:100%;border-collapse:collapse;font-size:12px}
th{background:#f0f4f8;padding:8px;text-align:left;font-weight:600;border-bottom:2px solid #d0d7de;white-space:nowrap}
td{padding:7px 8px;border-bottom:1px solid #eee}
.num{text-align:right;font-variant-numeric:tabular-nums}
.neg{color:#d93025;font-weight:600}
.pos{color:#1e8e3e;font-weight:600}
.sortable{cursor:pointer;user-select:none}
.sortable:hover{background:#dde6ef}
.sortable .arrow{opacity:.4;font-size:10px;margin-left:4px}
.sortable.asc .arrow,.sortable.desc .arrow{opacity:1;color:#003e7e}
.filters{background:#fff;padding:16px;border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.06);display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:20px}
.filters label{font-size:11px;font-weight:600;color:#555;text-transform:uppercase;display:block;margin-bottom:4px}
.filters select,.filters input{width:100%;padding:7px;border:1px solid #d0d7de;border-radius:6px;font-size:13px}
.kpis{display:grid;grid-template-columns:repeat(5,1fr);gap:14px;margin-bottom:20px}
.kpi{background:#fff;padding:16px;border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.06);border-left:4px solid #0066b3}
.kpi .label{font-size:11px;color:#666;text-transform:uppercase;font-weight:600}
.kpi .value{font-size:24px;font-weight:700;margin-top:4px;color:#003e7e}
.kpi.negative{border-left-color:#d93025}.kpi.negative .value{color:#d93025}
.kpi.positive{border-left-color:#1e8e3e}.kpi.positive .value{color:#1e8e3e}
.grid{display:grid;grid-template-columns:2fr 1fr;gap:18px;margin-bottom:20px}
.resumen-row{display:flex;justify-content:space-between;padding:7px 10px;font-size:13px;border-bottom:1px solid #f0f0f0}
.resumen-row.header{background:#003e7e;color:#fff;font-weight:700;border-radius:4px 4px 0 0}
.resumen-row.total{background:#fff3cd;font-weight:700;border-bottom:2px solid #003e7e}
.resumen-row.final{background:#003e7e;color:#fff;font-weight:700;border-radius:0 0 4px 4px}
.resumen-row .v{font-variant-numeric:tabular-nums}
.period-mini{display:inline-flex;align-items:center;gap:8px;background:#f0f4f8;padding:6px 10px;border-radius:6px;font-size:12px}
.period-mini select{padding:4px 8px;border:1px solid #d0d7de;border-radius:4px;font-size:12px}
.download-btn{background:#1e8e3e;color:#fff;border:none;padding:10px 16px;border-radius:6px;cursor:pointer;font-weight:600;margin-left:8px}
.download-btn:hover{background:#176b30}
.download-btn.blue{background:#0066b3}
.download-btn.blue:hover{background:#004f8c}
.summary{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:20px}
.badge{display:inline-block;padding:3px 10px;border-radius:12px;font-size:11px;font-weight:700}
.badge.ok{background:#e6f4ea;color:#1e8e3e}
.badge.warn{background:#fef7e0;color:#b06000}
.badge.err{background:#fce8e6;color:#d93025}
.cal-wrap{overflow-x:auto}
.cal{border-collapse:collapse;font-size:11px}
.cal th,.cal td{border:1px solid #e0e6ec;padding:4px 6px;min-width:38px;text-align:center;white-space:nowrap}
.cal th{background:#003e7e;color:#fff}
.cal td.ceve-name{background:#f0f4f8;font-weight:600;text-align:left;min-width:140px}
.cal .c-ok{background:#1e8e3e;color:#fff}
.cal .c-no{background:#fce8e6}
.cal .c-fut{background:#f8f9fa;color:#bbb}
.tabs{display:flex;gap:0}
.tab-btn{flex:1;background:#e9eef3;border:none;padding:14px;cursor:pointer;font-weight:600;color:#555;border-radius:10px 10px 0 0;font-size:14px}
.tab-btn.active{background:#fff;color:#003e7e;border-bottom:3px solid #0066b3}
.tab-content{background:#fff;padding:24px;border-radius:0 0 10px 10px;box-shadow:0 1px 4px rgba(0,0,0,.06);display:none}
.tab-content.active{display:block}
.form-row{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}
.form-row.three{grid-template-columns:1fr 1fr 1fr}
.form-row.four{grid-template-columns:1fr 1fr 1fr 1fr}
.form-row.full{grid-template-columns:1fr}
label{font-size:11px;font-weight:600;color:#555;text-transform:uppercase;display:block;margin-bottom:5px}
input,select,textarea{width:100%;padding:9px;border:1px solid #d0d7de;border-radius:6px;font-size:14px}
textarea{resize:vertical;min-height:60px}
button[type=submit]{background:#0066b3;color:#fff;border:none;padding:12px 24px;border-radius:6px;cursor:pointer;font-weight:600;font-size:14px;width:100%}
button[type=submit]:hover{background:#004f8c}
.msg{padding:12px;border-radius:6px;margin-bottom:14px;font-size:14px}
.msg.ok{background:#e6f4ea;color:#1e8e3e;border:1px solid #1e8e3e}
.msg.err{background:#fce8e6;color:#d93025;border:1px solid #d93025}
.info{background:#e8f0fe;padding:12px;border-radius:6px;font-size:13px;color:#003e7e;margin-bottom:14px}
.warning{background:#fef7e0;color:#b06000;padding:14px;border-radius:6px;border-left:4px solid #f9ab00;margin-bottom:14px;font-size:13px}
.section-title{background:#003e7e;color:#fff;padding:10px 14px;border-radius:6px;margin:18px 0 12px 0;font-size:13px;font-weight:700;text-transform:uppercase}
.inv-table{width:100%;border-collapse:collapse;margin-bottom:14px;background:#f7f9fb;border-radius:8px;overflow:hidden}
.inv-table th{background:#003e7e;color:#fff;padding:10px;font-size:12px;text-align:center}
.inv-table td{padding:6px;border-bottom:1px solid #e0e6ec}
.inv-table td.equipo{font-weight:600;background:#fff;color:#003e7e;font-size:12px}
.inv-table input{padding:7px;font-size:13px;text-align:right}
.inv-table .total-cell{background:#fff3cd;font-weight:700;text-align:right;font-variant-numeric:tabular-nums;padding:8px}
.accordion{background:#f7f9fb;border:1px solid #e0e6ec;border-radius:8px;margin-bottom:14px}
.acc-header{padding:12px 16px;cursor:pointer;font-weight:600;color:#003e7e;display:flex;justify-content:space-between;align-items:center}
.acc-header:hover{background:#eef2f6}
.acc-body{padding:14px;display:none;border-top:1px solid #e0e6ec}
.accordion.open .acc-body{display:block}
.mov-table{width:100%;border-collapse:collapse;font-size:11px;margin-top:8px}
.mov-table th{background:#003e7e;color:#fff;padding:6px 4px;font-weight:600}
.mov-table td{padding:3px;border-bottom:1px solid #e0e6ec}
.mov-table input{padding:5px;font-size:11px}
.btn-add{background:#1e8e3e;color:#fff;border:none;padding:6px 12px;border-radius:5px;cursor:pointer;font-size:12px;margin-top:8px}
.btn-del{background:#d93025;color:#fff;border:none;padding:3px 8px;border-radius:4px;cursor:pointer;font-size:11px}
.modal{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.6);z-index:1000;justify-content:center;align-items:center}
.modal.show{display:flex}
.modal-box{background:#fff;padding:24px;border-radius:10px;max-width:700px;width:90%;max-height:90vh;overflow-y:auto;box-shadow:0 4px 20px rgba(0,0,0,.3)}
.modal-box h3{margin:0 0 14px 0}
.modal-box.warn h3{color:#d93025}
.modal-box.info h3{color:#0066b3}
.modal-actions{display:flex;gap:10px;margin-top:18px}
.modal-actions button{flex:1;padding:10px;border-radius:6px;border:none;cursor:pointer;font-weight:600}
.btn-confirm{background:#0066b3;color:#fff}
.btn-confirm.warn{background:#d93025}
.btn-cancel{background:#e9eef3;color:#555}
.pdf-summary{font-family:Arial,sans-serif;color:#222;padding:15px;width:760px;max-width:760px;margin:0 auto;background:#fff}
.pdf-summary h2{color:#003e7e;border-bottom:2px solid #003e7e;padding-bottom:6px;font-size:18px;margin:0 0 10px 0}
.pdf-summary h3{color:#003e7e;font-size:13px;margin:12px 0 6px 0;padding-bottom:3px;border-bottom:1px solid #ccc}
.pdf-summary p{margin:4px 0;font-size:12px}
.pdf-summary table{width:100%;border-collapse:collapse;margin:6px 0;table-layout:fixed}
.pdf-summary th,.pdf-summary td{border:1px solid #888;padding:4px;font-size:10px;word-wrap:break-word;text-align:left}
.pdf-summary th{background:#003e7e;color:#fff;font-size:10px}
.pdf-summary tr{page-break-inside:avoid}
.success-screen{text-align:center;padding:40px 20px}
.success-screen .ok-icon{font-size:60px;color:#1e8e3e;margin-bottom:10px}
.success-screen h2{color:#1e8e3e;margin:10px 0}
.operativo-kpis{
  display:grid;
  grid-template-columns:1fr 1fr 1.25fr;
  gap:18px;
  margin-bottom:18px
}
.operativo-kpi{
  background:#fff;
  padding:22px;
  border-radius:12px;
  box-shadow:0 2px 8px rgba(0,0,0,.08);
  border-top:5px solid #0066b3;
  text-align:center
}
.operativo-kpi .label{
  font-size:12px;
  font-weight:700;
  color:#555;
  text-transform:uppercase;
  letter-spacing:.4px
}
.operativo-kpi .value{
  font-size:32px;
  font-weight:750;
  margin-top:8px;
  color:#003e7e;
  font-variant-numeric:tabular-nums
}
.operativo-kpi.principal{
  background:linear-gradient(135deg,#0066b3,#003e7e);
  border-top-color:#003e7e;
  transform:scale(1.02)
}
.operativo-kpi.principal .label,
.operativo-kpi.principal .value{
  color:#fff
}
.operativo-kpi.negativo{
  border-top-color:#d93025
}
.operativo-kpi.negativo .value{
  color:#d93025
}
.operativo-kpi.positivo{
  border-top-color:#1e8e3e
}
.operativo-kpi.positivo .value{
  color:#1e8e3e
}
.operativo-kpi.principal.negativo{
  background:linear-gradient(135deg,#d93025,#a32018)
}
.operativo-kpi.principal.positivo{
  background:linear-gradient(135deg,#1e8e3e,#126b2b)
}
.lectura-flujo{
  background:#fff;
  border-left:4px solid #0066b3;
  padding:12px 16px;
  border-radius:8px;
  box-shadow:0 1px 4px rgba(0,0,0,.06);
  margin-bottom:20px;
  font-size:13px;
  color:#444;
  text-align:center
}
.analisis-toggle{
  width:100%;
  border:none;
  background:#fff;
  color:#003e7e;
  padding:14px 18px;
  border-radius:10px;
  box-shadow:0 1px 4px rgba(0,0,0,.08);
  font-size:14px;
  font-weight:700;
  cursor:pointer;
  margin-bottom:14px
}
.analisis-toggle:hover{
  background:#eef5fb
}
.analisis-panel{
  display:none;
  margin-bottom:18px
}
.analisis-panel.open{
  display:block
}
.analisis-bloques{
  display:grid;
  grid-template-columns:repeat(3,1fr);
  gap:14px;
  margin-bottom:18px
}
.analisis-bloque{
  background:#fff;
  padding:16px;
  border-radius:10px;
  box-shadow:0 1px 4px rgba(0,0,0,.06)
}
.analisis-bloque h4{
  margin:0 0 12px;
  color:#003e7e;
  font-size:13px;
  text-transform:uppercase
}
.analisis-fila{
  display:flex;
  justify-content:space-between;
  gap:10px;
  border-bottom:1px solid #eee;
  padding:7px 0;
  font-size:12px
}
.analisis-fila:last-child{
  border-bottom:none
}
.analisis-fila .valor{
  font-weight:700;
  font-variant-numeric:tabular-nums;
  text-align:right
}
.sin-dato-periodo{
  color:#b06000;
  font-size:11px;
  margin-top:8px
}
@media(max-width:900px){
  .operativo-kpis,
  .analisis-bloques{
    grid-template-columns:1fr
  }
  .operativo-kpi.principal{
    transform:none
  }
}
"""

_A = (
    '&lt;a href="index.html"__C1__&gt;📊 Dashboard&lt;/a&gt;'
    '&lt;a href="carga.html"__C2__&gt;📝 Carga Stock&lt;/a&gt;'
    '&lt;a href="facturacion.html"__C3__&gt;🧾 Carga Facturaciones&lt;/a&gt;'
    '&amp;lt;a href="seguimiento.html"__C4__&amp;gt;🔐 Acceso Control de Equipos&amp;lt;/a&amp;gt;'
)

NAV_DASH = (
    '&lt;div class="nav"&gt;'
    + _A
        .replace('__C1__', ' class="active"')
        .replace('__C2__', '')
        .replace('__C3__', '')
        .replace('__C4__', '')
    + '&lt;/div&gt;'
)

NAV_CARGA = (
    '&lt;div class="nav"&gt;'
    + _A
        .replace('__C1__', '')
        .replace('__C2__', ' class="active"')
        .replace('__C3__', '')
        .replace('__C4__', '')
    + '&lt;/div&gt;'
)

NAV_FACTURACION = (
    '&lt;div class="nav"&gt;'
    + _A
        .replace('__C1__', '')
        .replace('__C2__', '')
        .replace('__C3__', ' class="active"')
        .replace('__C4__', '')
    + '&lt;/div&gt;'
)

NAV_SEG = (
    '&lt;div class="nav"&gt;'
    + _A
        .replace('__C1__', '')
        .replace('__C2__', '')
        .replace('__C3__', '')
        .replace('__C4__', ' class="active"')
    + '&lt;/div&gt;'
)

SUB_JS = "[" + ",".join("'" + s + "'" for s in SUBAGENCIAS) + "]"

EQUIPOS_INV = [
    ("BG", "BANDEJAS (028)"),
    ("DOG", "DOLLYS Grandes (031)"),
    ("PM", "Pallets Madera"),
    ("PP", "Pallets Plástico"),
    ("PASC", "Pascualineros"),
]
INV_ROWS = ""
for cod, lbl in EQUIPOS_INV:
    INV_ROWS += (
        '<tr><td class="equipo">' + lbl + '</td>'
        '<td><input type="number" min="0" value="0" id="m' + cod + 'CP" oninput="recalcTotal(this)" data-eq="' + cod + '"></td>'
        '<td><input type="number" min="0" value="0" id="m' + cod + 'VA" oninput="recalcTotal(this)" data-eq="' + cod + '"></td>'
        '<td><input type="number" min="0" value="0" id="m' + cod + 'OT" oninput="recalcTotal(this)" data-eq="' + cod + '"></td>'
        '<td class="total-cell" id="tot' + cod + '">0</td></tr>'
    )

SECS = [
    ('rd', '📥 Recibidas (DENTRO del conteo)'),
    ('rf', '📥 Recibidas (NO dentro del conteo)'),
    ('dd', '📤 Despachadas (DENTRO del conteo)'),
    ('df', '📤 Despachadas (NO dentro del conteo)'),
]
ACCORDIONS = ""
for code, label in SECS:
    ACCORDIONS += (
        '<div class="accordion" data-mov="' + code + '">'
        '<div class="acc-header"><span>' + label + '</span><span>▶</span></div>'
        '<div class="acc-body">'
        '<table class="mov-table"><thead><tr>'
        '<th>Fecha</th><th>Chofer</th><th>Nº CONEQUIP</th><th>Nº Remito</th>'
        '<th>Destino/Origen</th><th>Hora</th><th>Band</th><th>Pasc</th>'
        '<th>D.Gr</th><th>D.Ch</th><th>Mad</th><th>Plast</th><th></th>'
        '</tr></thead><tbody></tbody></table>'
        '<button type="button" class="btn-add" onclick="agregarFila(' + chr(39) + code + chr(39) + ')">+ Agregar fila</button>'
        '</div></div>'
    )

# ========= INDEX.HTML =========
INDEX = """<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8"><title>Dashboard Saldo Primario</title>
__CDN_CHART__
__CDN_XLSX__
__CDN_SUPABASE__
<style>__CSS__</style></head><body>
<header><h1>📊 Dashboard Saldo Primario - CeVes Argentina</h1><div class="updated" id="lastUpdate">Cargando…</div></header>
<div class="container">__NAV__
<div class="filters">
<div><label>Desde</label><input type="date" id="fDesde"></div>
<div><label>Hasta</label><input type="date" id="fHasta"></div>
<div><label>Región</label><select id="fRegion"><option value="">Todas</option><option>AMBA</option><option>Interior</option></select></div>
<div><label>CeVe</label><select id="fCeve"><option value="">Todos</option></select></div>
<div><label>Equipo</label><select id="fEquipo"><option value="028" selected>BG-Bandeja Grande</option><option value="031">DO-Dolly</option><option value="132">PAL-Pallet</option><option value="">Todos</option></select></div>
</div>
<div class="kpis" id="kpis"></div>
<div class="grid">
<div class="card"><h3>Saldo Neto por CeVe</h3><canvas id="chartCeves" height="120"></canvas></div>
<div class="card"><h3>📋 Resumen Detallado</h3><div id="resumenContenido"></div></div>
</div>

<div class="card">
<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;flex-wrap:wrap;gap:10px">
<h3 style="margin:0">Evolución de Saldo</h3>
<div class="period-mini"><span>Período:</span><select id="fPeriodo"><option value="semana">Semana</option><option value="mes" selected>Mes</option><option value="anio">Año</option></select></div>
</div>
<canvas id="chartEvol" height="60"></canvas>
</div>

<div class="card">
<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;flex-wrap:wrap;gap:10px">
<h3 style="margin:0">📈 Variación Semanal del % Saldo / Flujo</h3>
<div class="period-mini">
<span>Comparar:</span>
<select id="fSemanaTipo">
<option value="lunes" selected>Lunes vs Lunes anterior</option>
<option value="sabado">Sábado vs Sábado anterior</option>
<option value="personalizado">Fecha personalizada</option>
</select>
<input type="date" id="fSemanaCustom" style="display:none">
</div>
<button class="download-btn" onclick="exportarSemanal()">⬇ Excel</button>
</div>
<table id="tablaSemanal"><thead></thead><tbody></tbody></table>
</div>

<div class="card">
<div style="display:flex;align-items:center;margin-bottom:10px;gap:10px;flex-wrap:wrap">
<h3 style="margin:0">Tabla Maestra por CeVe</h3>
<button class="download-btn" onclick="exportarExcel()">⬇ Tabla Excel</button>
<button class="download-btn blue" onclick="abrirModalMov()">📥 Descargar Movimientos</button>
</div>
<table id="tabla">
<thead><tr>
<th class="sortable" data-col="region">Región<span class="arrow">⇅</span></th>
<th class="sortable" data-col="ceve">CeVe<span class="arrow">⇅</span></th>
<th class="sortable num" data-col="inv">Inv. Inicial<span class="arrow">⇅</span></th>
<th class="sortable num" data-col="sa">Conteo Físico<span class="arrow">⇅</span></th>
<th class="sortable num" data-col="vs">Var. Stock<span class="arrow">⇅</span></th>
<th class="sortable num" data-col="ing">Ingresos<span class="arrow">⇅</span></th>
<th class="sortable num" data-col="egr">Egresos<span class="arrow">⇅</span></th>
<th class="sortable num" data-col="sb">Saldo<span class="arrow">⇅</span></th>
<th class="sortable num" data-col="sn">Saldo Neto<span class="arrow">⇅</span></th>
<th class="sortable num" data-col="fl">% Flujo<span class="arrow">⇅</span></th>
</tr></thead><tbody></tbody><tfoot id="tablaFoot"></tfoot></table>
</div>

<div class="card">
<div style="display:flex;align-items:center;margin-bottom:10px;gap:10px;flex-wrap:wrap">
<h3 style="margin:0">📋 Detalle de Movimientos con Planta</h3>
<button class="download-btn" onclick="exportarDetalle()">⬇ Descargar Excel</button>
<span id="detalleContador" style="margin-left:auto;font-size:12px;color:#666"></span>
</div>
<div class="period-mini" style="margin-bottom:10px;flex-wrap:wrap">
<span>CeVe:</span>
<select id="detCeve"><option value="">Todos</option></select>
<span>Desde:</span><input type="date" id="detDesde">
<span>Hasta:</span><input type="date" id="detHasta">
</div>
<div style="max-height:500px;overflow:auto">
<table id="tablaDetalle">
<thead>
<tr>
<th class="sortableDet" data-col="fecha" style="cursor:pointer">Fecha ⇅</th>
<th class="sortableDet" data-col="ceve" style="cursor:pointer">CeVe ⇅</th>
<th class="sortableDet" data-col="motivo" style="cursor:pointer">Motivo ⇅</th>
<th class="sortableDet" data-col="remito" style="cursor:pointer">Nº Remito ⇅</th>
<th class="sortableDet" data-col="origen" style="cursor:pointer">Origen ⇅</th>
<th class="sortableDet" data-col="destino" style="cursor:pointer">Destino ⇅</th>
<th class="sortableDet num" data-col="bg" style="cursor:pointer">
  Bandejas (BG) ⇅
</th>

<th class="sortableDet num" data-col="doll" style="cursor:pointer">
  Dollys (DO) ⇅
</th>

<th class="sortableDet num" data-col="pal" style="cursor:pointer">
  Pallets (PAL) ⇅
</th>
</tr>
<tr style="background:#f7f9fb">
<th style="padding:4px"></th>
<th style="padding:4px"></th>
<th style="padding:4px"><select id="fltMotivo" style="width:100%;padding:4px;font-size:11px"><option value="">Todos</option></select></th>
<th style="padding:4px"><input type="text" id="fltRemito" placeholder="Filtrar..." style="width:100%;padding:4px;font-size:11px"></th>
<th style="padding:4px"><select id="fltOrigen" style="width:100%;padding:4px;font-size:11px"><option value="">Todos</option></select></th>
<th style="padding:4px"><select id="fltDestino" style="width:100%;padding:4px;font-size:11px"><option value="">Todos</option></select></th>
<th style="padding:4px"></th>
<th style="padding:4px"></th>
<th style="padding:4px"></th>
</tr>
</thead>
<tbody></tbody>
</table>
</div>
</div>

<div class="modal" id="modMov"><div class="modal-box info">
<h3>📥 Descargar Movimientos</h3>
<p>Filtrá los movimientos a exportar:</p>
<div class="form-row"><div><label>Desde</label><input type="date" id="movDesde"></div><div><label>Hasta</label><input type="date" id="movHasta"></div></div>
<div class="form-row"><div><label>Región</label><select id="movRegion"><option value="">Todas</option><option>AMBA</option><option>Interior</option></select></div><div><label>CeVe</label><select id="movCeve"><option value="">Todos</option></select></div></div>
<div class="form-row"><div><label>Equipo</label>
<select id="movEquipo">
  <option value="">Todos</option>
  <option value="028" selected>BG-Bandeja Grande</option>
  <option value="031">DO-Dolly</option>
  <option value="132">PAL-Pallet</option>
</select>
</div><div><label>Tipo</label><select id="movTipo"><option value="">Ambos</option><option value="ing">Solo Ingresos</option><option value="egr">Solo Egresos</option></select></div></div>
<div class="modal-actions">
<button type="button" class="btn-cancel" onclick="document.getElementById('modMov').classList.remove('show')">Cancelar</button>
<button type="button" class="btn-confirm" onclick="descargarMov()">⬇ Descargar Excel</button>
</div></div></div>

<script>
const SUBAG = __SUB_JS__;
const SUPABASE_URL = '__SUPABASE_URL__';
const SUPABASE_KEY = '__SUPABASE_KEY__';
const supabaseClient = window.supabase.createClient(
   SUPABASE_URL,
   SUPABASE_KEY
);
let MAESTRO=null;
let INVENTARIO_PENDIENTE = null;
let MOVS=[];
let DETALLE=[];
let FACTURACIONES=[ ];
let STOCK={
      registros_diarios:[ ]
};

let AG_MDP=null;
const AG_ST='14128';
let sortCol='sn', sortDir='asc';
let sortColDet='fecha', sortDirDet='asc';
let sortColSem='variacion', sortDirSem='asc';
async function cargar(){
  const raw = await fetch(
  'data/ceves_maestro.json?ts=' + Date.now(),
  { cache: 'no-store' }
).then(r => r.json());

const cevesBase = raw.ceves.map(
  c => ({...c})
);

const mdp = cevesBase.find(
  c =>
    String(c.nombre)
      .trim()
      .toLowerCase() === 'mar del plata'
);

const st = cevesBase.find(
  c => String(c.agencia) === AG_ST
);

if(!mdp){
  throw new Error(
    'No se encontró Mar del Plata en ceves_maestro.json'
  );
}

AG_MDP = String(mdp.agencia);

if(st){
  mdp.inv_inicial_BG =
    (Number(mdp.inv_inicial_BG) || 0)
    + (Number(st.inv_inicial_BG) || 0);

  mdp.inv_inicial_DO =
    (Number(mdp.inv_inicial_DO) || 0)
    + (Number(st.inv_inicial_DO) || 0);

  mdp.inv_inicial_PAL =
    (Number(mdp.inv_inicial_PAL) || 0)
    + (Number(st.inv_inicial_PAL) || 0);
}

mdp.nombre = 'Mar del Plata + ST';

MAESTRO = {
  ...raw,
  ceves: cevesBase.filter(
    c => !SUBAG.includes(String(c.agencia))
  )
};
  const m = await fetch(
  'data/movimientos_mc1.json?ts=' + Date.now(),
  { cache: 'no-store' }
).then(r => {
  if(!r.ok) {
    throw new Error(
      'No se pudo cargar movimientos_mc1.json: ' + r.status
    );
  }
  return r.json();
});
  MOVS = m.movimientos.filter(x=>!SUBAG.includes(x.agencia));
  try {
    const det = await fetch(
  'data/movimientos_detalle.json?ts=' + Date.now(),
  { cache: 'no-store' }
).then(r => {
  if(!r.ok) {
    throw new Error(
      'No se pudo cargar movimientos_detalle.json: ' + r.status
    );
  }
  return r.json();
});
    DETALLE = det.movimientos.filter(x=>!SUBAG.includes(x.agencia));
  } catch(e){ DETALLE = []; }

/*
Cargar facturaciones manuales desde Supabase.

Cada facturación se convierte en:
- egreso a Planta Pilar;
- movimiento visible en el detalle;
- impacto en el saldo del CeVe y equipo.
*/
try {
  const {
    data: facturacionesSupabase,
    error: errorFacturaciones
  } = await supabaseClient
    .from('facturaciones')
    .select(
      'id, created_at, agencia, ceve, ' +
      'fecha_facturacion, numero_factura, ' +
      'cliente, codigo_facturado, ' +
      'cantidad_facturada'
    )
    .order(
      'fecha_facturacion',
      { ascending: true }
    )
    .order(
      'created_at',
      { ascending: true }
    );

  if(errorFacturaciones){
    throw errorFacturaciones;
  }

  FACTURACIONES =
    facturacionesSupabase || [];

  /*
  Incorporar cada factura a MOVS.

  Así se suma automáticamente a:
  - Egresos Planta;
  - Saldo Bruto;
  - Saldo Neto;
  - gráficos;
  - tabla maestra;
  - exportaciones de movimientos.
  */
  FACTURACIONES.forEach(factura =&gt; {
    const agencia =
      String(factura.agencia);

    const equipo =
      String(factura.codigo_facturado)
        .padStart(3, '0');

    const cantidad =
      Number(
        factura.cantidad_facturada
      ) || 0;

    if(
  SUBAG.includes(agencia) ||
  !['028', '031'].includes(equipo) ||
  cantidad <= 0
){
  return;
}

    MOVS.push({
      agencia: agencia,
      ceve: factura.ceve,
      fecha: factura.fecha_facturacion,
      equipo: equipo,
      ingresos_planta: 0,
      egresos_planta: cantidad,
      saldo_bruto: cantidad,
      fuente: 'facturacion_manual',
      id_facturacion: factura.id
    });

    /*
    Incorporar la misma factura al detalle.

    Destino HIZ se visualiza como Planta Pilar
    mediante plantaLabel().
    */
    DETALLE.push({
      fecha: factura.fecha_facturacion,
      agencia: agencia,
      ceve: factura.ceve,
      equipo: equipo,
      motivo: 'FACTURACIÓN MANUAL',
      tipo: 'egreso_planta',
      origen:
        factura.ceve +
        ' - ' +
        factura.cliente,
      destino: 'HIZ',
      remito: factura.numero_factura,
      cantidad: cantidad,
      fuente: 'facturacion_manual',
      id_facturacion: factura.id,
      timestamp: factura.created_at
    });
  });

  console.log(
    'Facturaciones cargadas desde Supabase:',
    FACTURACIONES.length
  );

} catch(e) {
  console.error(
    'Error cargando facturaciones:',
    e
  );

  FACTURACIONES = [];
}

try {
  const { data: stockSupabase, error } = await supabaseClient
    .from('stock_diario')
    .select(
  'agencia, ceve, fecha_conteo, ' +
  'bg, do_stock, pal, ' +
  'observaciones, usuario, created_at'
)
    .order('fecha_conteo', { ascending: false })
    .order('created_at', { ascending: false });

  if(error) throw error;

  STOCK = {
    registros_diarios: (stockSupabase || []).map(r => ({
      tipo: 'diario',
      agencia: String(r.agencia),
      ceve: r.ceve,
      fecha: r.fecha_conteo,
      BG: Number(r.bg) || 0,
      DO: Number(r.do_stock) || 0,
      PAL:
                     r.pal === null
                         ? null
                         : Number(r.pal),
      observaciones: r.observaciones || '',
      supervisor: r.usuario || '',
      timestamp: r.created_at
    }))
  };

  console.log(
    'Stock cargado desde Supabase:',
    STOCK.registros_diarios.length
  );
} catch(e) {
  console.error('Error cargando stock desde Supabase:', e);
  STOCK = { registros_diarios: [] };
}
  document.getElementById('lastUpdate').textContent='Actualizado: '+new Date(m.actualizado).toLocaleString('es-AR');
  MAESTRO.ceves.forEach(c=>{
    document.getElementById('fCeve').innerHTML += '<option value="'+c.agencia+'">'+c.nombre+'</option>';
    document.getElementById('movCeve').innerHTML += '<option value="'+c.agencia+'">'+c.nombre+'</option>';
    document.getElementById('detCeve').innerHTML += '<option value="'+c.agencia+'">'+c.nombre+'</option>';
  });
  document.getElementById('fDesde').value='2026-01-01';
  document.getElementById('fHasta').value=new Date().toISOString().slice(0,10);
  document.getElementById('movDesde').value='2026-01-01';
  document.getElementById('movHasta').value=new Date().toISOString().slice(0,10);
  document.getElementById('detDesde').value='2026-01-01';
     document.getElementById('detHasta').value=new Date().toISOString().slice(0,10);
     document.querySelectorAll('#detCeve, #detDesde, #detHasta').forEach(el=>el.addEventListener('change', renderDetalle));
  document.addEventListener('click', e => {
    if(e.target.classList.contains('sortableDet')){
      const col = e.target.dataset.col;
      if(sortColDet===col) sortDirDet = sortDirDet==='asc'?'desc':'asc';
      else { sortColDet=col; sortDirDet='asc'; }
      renderDetalle();
    }
  });
  document.querySelectorAll('.filters select,.filters input,.period-mini select,#fSemanaCustom').forEach(el=>el.addEventListener('change',render));
  document.querySelectorAll('.sortable').forEach(th=>th.addEventListener('click',()=>{
    const col=th.dataset.col;
    if(sortCol===col) sortDir = sortDir==='asc'?'desc':'asc';
    else { sortCol=col; sortDir='asc'; }
    render();
  }));
  render();

  // Actualizar automáticamente cuando un supervisor carga stock.
  if(!window._stockRealtime){
    window._stockRealtime = supabaseClient
      .channel('stock-diario-dashboard')
      .on(
        'postgres_changes',
        {
          event: 'INSERT',
          schema: 'public',
          table: 'stock_diario'
        },
        payload => {
          console.log('Nueva carga recibida:', payload.new);
          location.reload();
        }
      )
      .subscribe(status => {
        console.log('Estado Realtime:', status);
      });
  }
/*
Actualizar automáticamente el dashboard
cuando se registra una nueva facturación.
*/
if(!window._facturacionRealtime){
  window._facturacionRealtime =
    supabaseClient
      .channel(
        'facturaciones-dashboard'
      )
      .on(
        'postgres_changes',
        {
          event: 'INSERT',
          schema: 'public',
          table: 'facturaciones'
        },
        payload =&gt; {
          console.log(
            'Nueva facturación recibida:',
            payload.new
          );

          location.reload();
        }
      )
      .subscribe(status =&gt; {
        console.log(
          'Estado Realtime facturaciones:',
          status
        );
      });
  }
}
function filtrar(){
  const r=document.getElementById('fRegion').value, c=document.getElementById('fCeve').value, e=document.getElementById('fEquipo').value;
  const d=document.getElementById('fDesde').value, h=document.getElementById('fHasta').value;
  return MOVS.filter(m=>{
    const ce=MAESTRO.ceves.find(x=>x.agencia===m.agencia);
    if(!ce) return false;
    if(r&&ce.region!==r) return false;
    if(c&&m.agencia!==c) return false;
    if(e&&m.equipo!==e) return false;
    if(d&&m.fecha<d) return false;
    if(h&&m.fecha>h) return false;
    return true;
  });
}
function fmt(n){return(n||0).toLocaleString('es-AR')}
function fmtP(n){return(n*100).toFixed(2)+'%'}
function fmtPSigned(n){const v=(n*100).toFixed(2);return (n>=0?'+':'')+v+'%'}

function stockAct(ag, eq){
  /*
  Mar del Plata + ST registra un único
  stock consolidado bajo la agencia
  técnica de Mar del Plata.
  */

  const registros = (
    STOCK.registros_diarios || []
  )
    .filter(
      s =&gt;
        String(s.agencia) === String(ag)
    )
    .sort((a,b) =&gt; {
      const claveA =
        (a.fecha || '') +
        '|' +
        (a.timestamp || '');

      const claveB =
        (b.fecha || '') +
        '|' +
        (b.timestamp || '');

      return claveB.localeCompare(claveA);
    });

  if(!registros.length){
    return null;
  }

  const ultimo = registros[0];

  if(eq === '028'){
    return Number(ultimo.BG) || 0;
  }

  if(eq === '031'){
    return Number(ultimo.DO) || 0;
  }

  if(eq === '132'){
  return ultimo.PAL === null
    || ultimo.PAL === undefined
      ? null
      : Number(ultimo.PAL);
}

  return null;
}

function calc(ag,eq,d){
  const c=MAESTRO.ceves.find(x=>x.agencia===ag);
  const inv =
  eq === '028'
    ? Number(c.inv_inicial_BG) || 0
    : eq === '031'
      ? Number(c.inv_inicial_DO) || 0
      : Number(c.inv_inicial_PAL) || 0;
  const sa=stockAct(ag,eq);
  const vs=sa!==null?sa-inv:0;
  const ms=d.filter(m=>m.agencia===ag&&m.equipo===eq);
  const ing=ms.reduce((s,m)=>s+m.ingresos_planta,0);
  const egr=ms.reduce((s,m)=>s+m.egresos_planta,0);
  const sb=egr-ing, sn=sb+vs, fl=ing?sn/ing:0;
  return {ceve:c,inv,sa,vs,ing,egr,sb,sn,fl};
}
function isoWeek(fecha){
  const d=new Date(fecha), t=new Date(d.valueOf()), dn=(d.getDay()+6)%7;
  t.setDate(t.getDate()-dn+3);
  const ft=t.valueOf();
  t.setMonth(0,1);
  if(t.getDay()!==4) t.setMonth(0,1+((4-t.getDay())+7)%7);
  const wk=1+Math.ceil((ft-t)/604800000);
  return d.getFullYear()+'-S'+String(wk).padStart(2,'0');
}
function render(){
  const data=filtrar();
  let tI=0,tE=0,tVS=0;
  const gr={};
  data.forEach(m=>{const k=m.agencia+'|'+m.equipo;if(!gr[k])gr[k]={ag:m.agencia,eq:m.equipo}});
  Object.values(gr).forEach(g=>{const r=calc(g.ag,g.eq,data);tI+=r.ing;tE+=r.egr;tVS+=r.vs});
  const tSB=tE-tI, tSN=tSB+tVS, tF=tI?tSN/tI:0;
  document.getElementById('kpis').innerHTML =
    '<div class="kpi"><div class="label">Ingresos Planta</div><div class="value">'+fmt(tI)+'</div></div>'+
    '<div class="kpi"><div class="label">Egresos Planta</div><div class="value">'+fmt(tE)+'</div></div>'+
    '<div class="kpi '+(tSB<0?'negative':'positive')+'"><div class="label">Saldo Bruto</div><div class="value">'+fmt(tSB)+'</div></div>'+
    '<div class="kpi '+(tSN<0?'negative':'positive')+'"><div class="label">Saldo Neto</div><div class="value">'+fmt(tSN)+'</div></div>'+
    '<div class="kpi '+(tF<0?'negative':'positive')+'"><div class="label">% Flujo</div><div class="value">'+fmtP(tF)+'</div></div>';
  renderResumen(data);
  renderSemanal(data);
  renderDetalle();

  const eqF = document.getElementById('fEquipo').value;
  const filas = MAESTRO.ceves.filter(c=>{
    const r=document.getElementById('fRegion').value, cf=document.getElementById('fCeve').value;
    if(r&&c.region!==r) return false;
    if(cf&&c.agencia!==cf) return false;
    return true;
  }).flatMap(c=>{
    const eqs = eqF ? [eqF] : ['028','031','132']
    return eqs.map(eq=>{
      const r=calc(c.agencia,eq,data);
      return {region:c.region, ceve:c.nombre, eq:
  eq === '028'
    ? 'BG'
    : eq === '031'
      ? 'DO'
      : 'PAL', inv:r.inv, sa:r.sa, vs:r.vs, ing:r.ing, egr:r.egr, sb:r.sb, sn:r.sn, fl:r.fl};
    });
  });
  filas.sort((a,b)=>{
    let av=a[sortCol], bv=b[sortCol];
    if(av===null) av=-Infinity; if(bv===null) bv=-Infinity;
    if(typeof av==='string') return sortDir==='asc'?av.localeCompare(bv):bv.localeCompare(av);
    return sortDir==='asc'?av-bv:bv-av;
  });
  document.querySelectorAll('.sortable').forEach(th=>{
    th.classList.remove('asc','desc');
    if(th.dataset.col===sortCol) th.classList.add(sortDir);
  });
  document.querySelector('#tabla tbody').innerHTML = filas.map(f=>
    '<tr><td>'+f.region+'</td><td>'+f.ceve+'</td>'+
    '<td class="num">'+fmt(f.inv)+'</td>'+
    '<td class="num">'+(f.sa!==null?fmt(f.sa):'—')+'</td>'+
    '<td class="num '+(f.vs<0?'neg':'pos')+'">'+fmt(f.vs)+'</td>'+
    '<td class="num">'+fmt(f.ing)+'</td><td class="num">'+fmt(f.egr)+'</td>'+
    '<td class="num '+(f.sb<0?'neg':'pos')+'">'+fmt(f.sb)+'</td>'+
    '<td class="num '+(f.sn<0?'neg':'pos')+'">'+fmt(f.sn)+'</td>'+
    '<td class="num '+(f.fl<0?'neg':'pos')+'">'+fmtP(f.fl)+'</td></tr>'
  ).join('');
  const sm=k=>filas.reduce((s,x)=>s+(x[k]||0),0);
  const sI=sm('ing'),sE=sm('egr'),sN=sm('sn'),sF=sI?sN/sI:0;
  document.getElementById('tablaFoot').innerHTML='<tr style="background:#003e7e;color:#fff;font-weight:700"><td colspan="2">TOTAL</td><td class="num">'+fmt(sm('inv'))+'</td><td class="num">'+fmt(filas.reduce((s,x)=>s+(x.sa||0),0))+'</td><td class="num">'+fmt(sm('vs'))+'</td><td class="num">'+fmt(sI)+'</td><td class="num">'+fmt(sE)+'</td><td class="num">'+fmt(sm('sb'))+'</td><td class="num">'+fmt(sN)+'</td><td class="num">'+fmtP(sF)+'</td></tr>';
  renderCharts(filas,data);
  window._filas=filas;
}
function renderResumen(data){
  const ag=document.getElementById('fCeve').value;
  const eq=document.getElementById('fEquipo').value || '028';
  let inv=0,co=null,vs=0,ing=0,egr=0;
  if(ag){const r=calc(ag,eq,data);inv=r.inv;co=r.sa;vs=r.vs;ing=r.ing;egr=r.egr}
  else MAESTRO.ceves.forEach(c=>{const r=calc(c.agencia,eq,data);inv+=r.inv;vs+=r.vs;ing+=r.ing;egr+=r.egr;if(r.sa!==null)co=(co||0)+r.sa});
  const s=egr-ing, sn=s+vs, fl=ing?sn/ing:0;
  const nom=ag?MAESTRO.ceves.find(c=>c.agencia===ag).nombre:'TODOS LOS CEVES';
  const en =
  eq === '028'
    ? 'BG-Bandeja Grande'
    : eq === '031'
      ? 'DO-Dolly'
      : 'PAL-Pallet';
  document.getElementById('resumenContenido').innerHTML =
    '<div class="resumen-row header"><span>'+nom+' — '+en+'</span><span class="v"></span></div>'+
    '<div class="resumen-row"><span>Inventario Inicial 2026</span><span class="v">'+fmt(inv)+'</span></div>'+
    '<div class="resumen-row"><span>Conteo Físico Diario</span><span class="v">'+(co!==null?fmt(co):'—')+'</span></div>'+
    '<div class="resumen-row"><span>Variación de Stock</span><span class="v '+(vs<0?'neg':'pos')+'">'+fmt(vs)+'</span></div>'+
    '<div class="resumen-row"><span>Ingresos de Planta</span><span class="v">'+fmt(ing)+'</span></div>'+
    '<div class="resumen-row"><span>Egresos a Planta</span><span class="v">'+fmt(egr)+'</span></div>'+
    '<div class="resumen-row total"><span>Saldo (Egr − Ing)</span><span class="v '+(s<0?'neg':'pos')+'">'+fmt(s)+'</span></div>'+
    '<div class="resumen-row total"><span>Saldo Neto</span><span class="v '+(sn<0?'neg':'pos')+'">'+fmt(sn)+'</span></div>'+
    '<div class="resumen-row final"><span>Saldo Neto / Flujo del CV</span><span class="v">'+fmtP(fl)+'</span></div>';
}

function renderSemanal(data){
  // Determinar fechas de comparación según selector
  const tipo = document.getElementById('fSemanaTipo').value;
  const custom = document.getElementById('fSemanaCustom');
  custom.style.display = tipo === 'personalizado' ? 'inline-block' : 'none';

  const hoy = new Date();
  let fechaActual, fechaAnterior;

if(tipo === 'personalizado' && custom.value){
    fechaActual = new Date(custom.value + 'T12:00:00');
    fechaAnterior = new Date(custom.value + 'T12:00:00');
    fechaAnterior.setDate(fechaAnterior.getDate() - 7);
  } else if(tipo === 'sabado'){
    const dow = hoy.getDay();
    fechaActual = new Date(hoy); fechaActual.setDate(hoy.getDate() + (6 - dow));
    fechaAnterior = new Date(fechaActual); fechaAnterior.setDate(fechaActual.getDate() - 7);
  } else {
    // Lunes: el último lunes pasado (si hoy es lunes, usar el de hace 7 días)
    const dow = hoy.getDay();
    const diasAlLunes = dow === 0 ? 6 : dow - 1;
    fechaActual = new Date(hoy); fechaActual.setDate(hoy.getDate() - diasAlLunes);
    fechaAnterior = new Date(fechaActual); fechaAnterior.setDate(fechaActual.getDate() - 7);
  }

  const fmtFecha = d => d.toISOString().slice(0,10);
  const finActual = fmtFecha(fechaActual);
  const finAnterior = fmtFecha(fechaAnterior);
  const lblActual = String(fechaActual.getDate()).padStart(2,'0')+'-'+String(fechaActual.getMonth()+1).padStart(2,'0');
  const lblAnterior = String(fechaAnterior.getDate()).padStart(2,'0')+'-'+String(fechaAnterior.getMonth()+1).padStart(2,'0');

  const eqF = document.getElementById('fEquipo').value || '028';
  const filas = MAESTRO.ceves.filter(c=>{
    const r=document.getElementById('fRegion').value, cf=document.getElementById('fCeve').value;
    if(r&&c.region!==r) return false;
    if(cf&&c.agencia!==cf) return false;
    return true;
  }).map(c=>{
    const ms = data.filter(m=>m.agencia===c.agencia && m.equipo===eqF);
    const msAct = ms.filter(m=>m.fecha <= finActual);
    const ing26 = msAct.reduce((s,m)=>s+m.ingresos_planta,0);
    const egr26 = msAct.reduce((s,m)=>s+m.egresos_planta,0);
    const inv =
  eqF === '028'
    ? Number(c.inv_inicial_BG) || 0
    : eqF === '031'
      ? Number(c.inv_inicial_DO) || 0
      : Number(c.inv_inicial_PAL) || 0;
    const sa = stockAct(c.agencia, eqF);
    const vs = sa!==null ? sa - inv : 0;
    const sn26 = (egr26 - ing26) + vs;
    const fl26 = ing26 ? sn26/ing26 : 0;
    const msAnt = ms.filter(m=>m.fecha <= finAnterior);
    const ingAnt = msAnt.reduce((s,m)=>s+m.ingresos_planta,0);
    const egrAnt = msAnt.reduce((s,m)=>s+m.egresos_planta,0);
    const snAnt = (egrAnt - ingAnt) + vs;
    const flAnt = ingAnt ? snAnt/ingAnt : 0;
    const variacion = fl26 - flAnt;
    return {region:c.region, ceve:c.nombre, fl26, flAnt, variacion};
  });

  // Ordenar según columna activa
  filas.sort((a,b)=>{
    let av=a[sortColSem], bv=b[sortColSem];
    if(typeof av==='string') return sortDirSem==='asc'?av.localeCompare(bv):bv.localeCompare(av);
    return sortDirSem==='asc'?av-bv:bv-av;
  });

  // Totales
  const allMs = data.filter(m=>m.equipo===eqF);
  const allMsAct = allMs.filter(m=>m.fecha <= finActual);
  const allMsAnt = allMs.filter(m=>m.fecha <= finAnterior);
  const tI = allMsAct.reduce((s,m)=>s+m.ingresos_planta,0);
  const tE = allMsAct.reduce((s,m)=>s+m.egresos_planta,0);
  const tVS = MAESTRO.ceves.reduce((s,c)=>{
    const sa=stockAct(c.agencia,eqF);
    const inv=eqF==='028'?c.inv_inicial_BG:c.inv_inicial_DO;
    return s + (sa!==null?sa-inv:0);
  },0);
  const tSN = (tE - tI) + tVS;
  const tFl = tI ? tSN/tI : 0;
  const tIA = allMsAnt.reduce((s,m)=>s+m.ingresos_planta,0);
  const tEA = allMsAnt.reduce((s,m)=>s+m.egresos_planta,0);
  const tSNA = (tEA - tIA) + tVS;
  const tFlA = tIA ? tSNA/tIA : 0;
  const tVar = tFl - tFlA;

  // Helper para flecha en cabecera
  const arr = c => c===sortColSem ? (sortDirSem==='asc' ? '▲' : '▼') : '⇅';

  document.querySelector('#tablaSemanal thead').innerHTML =
    '<tr>'+
    '<th class="sortableSem" data-col="region" style="cursor:pointer">Región <span style="font-size:10px">'+arr('region')+'</span></th>'+
    '<th class="sortableSem" data-col="ceve" style="cursor:pointer">Centro de venta <span style="font-size:10px">'+arr('ceve')+'</span></th>'+
    '<th class="sortableSem num" data-col="fl26" style="cursor:pointer">Saldo/Flujo 2026<br>(al '+lblActual+') <span style="font-size:10px">'+arr('fl26')+'</span></th>'+
    '<th class="sortableSem num" data-col="flAnt" style="cursor:pointer">Saldo/Flujo<br>(al '+lblAnterior+') <span style="font-size:10px">'+arr('flAnt')+'</span></th>'+
    '<th class="sortableSem num" data-col="variacion" style="cursor:pointer">Variación semanal <span style="font-size:10px">'+arr('variacion')+'</span></th>'+
    '</tr>';

  // Vincular click a los nuevos headers
  document.querySelectorAll('.sortableSem').forEach(th=>{
    th.addEventListener('click',()=>{
      const col=th.dataset.col;
      if(sortColSem===col) sortDirSem = sortDirSem==='asc'?'desc':'asc';
      else { sortColSem=col; sortDirSem='asc'; }
      renderSemanal(data);
    });
  });

  document.querySelector('#tablaSemanal tbody').innerHTML = filas.map(f=>
    '<tr><td>'+f.region+'</td><td>'+f.ceve+'</td>'+
    '<td class="num '+(f.fl26<0?'neg':'pos')+'">'+fmtP(f.fl26)+'</td>'+
    '<td class="num '+(f.flAnt<0?'neg':'pos')+'">'+fmtP(f.flAnt)+'</td>'+
    '<td class="num '+(f.variacion<0?'neg':'pos')+'">'+fmtPSigned(f.variacion)+'</td></tr>'
  ).join('') + '<tr style="background:#003e7e;color:#fff;font-weight:700"><td colspan="2">TOTAL</td>'+
    '<td class="num">'+fmtP(tFl)+'</td><td class="num">'+fmtP(tFlA)+'</td><td class="num">'+fmtPSigned(tVar)+'</td></tr>';

  window._semanal = filas;
}

let chC,chE;
function renderCharts(filas,data){
  const pC={};
  filas.forEach(f=>{pC[f.ceve]=(pC[f.ceve]||0)+f.sn});
  if(chC) chC.destroy();
  chC = new Chart(document.getElementById('chartCeves'),{type:'bar',data:{labels:Object.keys(pC),datasets:[{label:'Saldo Neto',data:Object.values(pC),backgroundColor:Object.values(pC).map(v=>v<0?'#d93025':'#1e8e3e')}]},options:{responsive:true,plugins:{legend:{display:false}}}});
  const per=document.getElementById('fPeriodo').value, pP={};
  data.forEach(m=>{const k=bucket(m.fecha,per);if(!pP[k])pP[k]={i:0,e:0};pP[k].i+=m.ingresos_planta;pP[k].e+=m.egresos_planta});
  const ks=Object.keys(pP).sort();
  if(chE) chE.destroy();
  chE = new Chart(document.getElementById('chartEvol'),{type:'line',data:{labels:ks,datasets:[
    {label:'Ingresos',data:ks.map(k=>pP[k].i),borderColor:'#1e8e3e',tension:.3},
    {label:'Egresos',data:ks.map(k=>pP[k].e),borderColor:'#d93025',tension:.3},
    {label:'Saldo Bruto',data:ks.map(k=>pP[k].e-pP[k].i),borderColor:'#0066b3',tension:.3,borderDash:[5,5]}
  ]},options:{responsive:true,plugins:{title:{display:true,text:'Agrupado por: '+per.toUpperCase()}}}});
}
function bucket(f,p){
  if(p==='mes') return f.slice(0,7);
  if(p==='anio') return f.slice(0,4);
  return isoWeek(f);
}
function exportarExcel(){
  const ws=XLSX.utils.json_to_sheet(window._filas);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Saldo Primario');
  XLSX.writeFile(wb,'SaldoPrimario_'+new Date().toISOString().slice(0,10)+'.xlsx');
}
function exportarSemanal(){
  const ws=XLSX.utils.json_to_sheet(window._semanal||[]);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Variacion Semanal');
  XLSX.writeFile(wb,'VariacionSemanal_'+new Date().toISOString().slice(0,10)+'.xlsx');
}
function ceveLabel(codigo){
  const c = (codigo||'').toString().toUpperCase().trim();
  if(!MAESTRO || !MAESTRO.ceves) return codigo;
  // Buscar en el maestro por agencia
  const ceve = MAESTRO.ceves.find(x => x.agencia === c);
  if(ceve) return ceve.nombre;
  // Algunos casos especiales que no están en el maestro (subagencias filtradas u otros)
  const especiales = {
    '14117': 'Tablada',
    '14134': 'Valente',
    '14267': 'La Plata',
    '14128': 'Santa Teresita',
    '14272': 'Villa Tesei'
  };
  return especiales[c] || codigo;
}

function nodoLabel(codigo){
  const c = (codigo||'').toString().toUpperCase().trim();
  if(!c) return '—';
  // Si empieza con "HI" es una planta (letra + código)
  if(/^HI[A-Z]/.test(c)) return plantaLabel(c);
  // Si es todo números, es un CeVe/agencia
  if(/^\\d+$/.test(c)) return ceveLabel(c);
  return codigo;
}
function plantaLabel(codigo){
  const c = (codigo||'').toUpperCase().trim();
  const plantas = {
    'HIZ': 'Planta Pilar',
    'HID': 'Planta Tesei',
    'HIR': 'Planta Córdoba',
    'HIG': 'Cedis Moreno',
    'HIN': 'Valente'
  };
  return plantas[c] || codigo;
}
function motivoLabel(motivo){
  const m = (motivo||'').toUpperCase();
  if(m === 'ENTRADA ESTOQUE - ASN AUTOMATICA' || m === 'ENTRADA ESTOQUE - AJUSTE MANUAL'){
    return 'ENTRADA ESTOQUE';
  }
  return motivo;
}

function renderDetalle(){
  // Poblar dropdowns de los filtros por columna (una sola vez)
  if(!window._dropdownsPobladosDet){
    const motivosSet = new Set();
    const origenesSet = new Set();
    const destinosSet = new Set();
    DETALLE.forEach(m=>{
      motivosSet.add(motivoLabel(m.motivo));
      if(m.origen) origenesSet.add(nodoLabel(m.origen));
      if(m.destino) destinosSet.add(nodoLabel(m.destino));
    });
    const orderStr = (a,b) => a.localeCompare(b);
    const opts = (set, id) => {
      const sel = document.getElementById(id);
      Array.from(set).sort(orderStr).forEach(v => sel.innerHTML += '<option value="'+v+'">'+v+'</option>');
    };
    opts(motivosSet, 'fltMotivo');
    opts(origenesSet, 'fltOrigen');
    opts(destinosSet, 'fltDestino');
    document.querySelectorAll('#fltMotivo, #fltRemito, #fltOrigen, #fltDestino').forEach(el => {
      el.addEventListener('input', renderDetalle);
      el.addEventListener('change', renderDetalle);
    });
    window._dropdownsPobladosDet = true;
  }

  const ag = document.getElementById('detCeve').value;
  const d = document.getElementById('detDesde').value;
  const h = document.getElementById('detHasta').value;
  const fltMotivo = document.getElementById('fltMotivo')?.value || '';
  const fltRemito = (document.getElementById('fltRemito')?.value || '').toUpperCase().trim();
  const fltOrigen = document.getElementById('fltOrigen')?.value || '';
  const fltDestino = document.getElementById('fltDestino')?.value || '';

  const filtradas = DETALLE.filter(m=>{
    if(ag && m.agencia!==ag) return false;
    if(d && m.fecha<d) return false;
    if(h && m.fecha>h) return false;
    return true;
  });

  // Agrupar por (fecha+ceve+remito) o (fecha+ceve+motivo si no hay remito)
  const grupos = {};
  filtradas.forEach(m=>{
    const key = m.remito
      ? m.fecha+'|'+m.agencia+'|'+m.remito
      : m.fecha+'|'+m.agencia+'|SIN_REMITO|'+m.motivo;
    if(!grupos[key]){
      grupos[key] = {
        fecha: m.fecha, ceve: m.ceve, tipo: m.tipo,
        remito: m.remito, origen: m.origen, destino: m.destino,
        motivos:
  new Set([
    motivoLabel(m.motivo)
  ]),

bg: 0,
doll: 0,
pal: 0
      };
    } else {
      grupos[key].motivos.add(motivoLabel(m.motivo));
      if(!grupos[key].origen && m.origen) grupos[key].origen = m.origen;
      if(!grupos[key].destino && m.destino) grupos[key].destino = m.destino;
    }
    if(m.equipo === '028'){
  grupos[key].bg +=
    Number(m.cantidad) || 0;

} else if(m.equipo === '031'){
  grupos[key].doll +=
    Number(m.cantidad) || 0;

} else if(m.equipo === '132'){
  grupos[key].pal +=
    Number(m.cantidad) || 0;
}
  });

  let filas = Object.values(grupos);

  // Aplicar filtros de columnas
  if(fltMotivo){
    filas = filas.filter(f => f.motivos.has(fltMotivo));
  }
  if(fltRemito){
    filas = filas.filter(f => (f.remito||'').toUpperCase().includes(fltRemito));
  }
  if(fltOrigen){
    filas = filas.filter(f => nodoLabel(f.origen) === fltOrigen);
  }
  if(fltDestino){
    filas = filas.filter(f => nodoLabel(f.destino) === fltDestino);
  }

  // Ordenar
  filas.sort((a,b)=>{
    let av, bv;
    if(sortColDet==='motivo'){
      av = Array.from(a.motivos).join('+');
      bv = Array.from(b.motivos).join('+');
    } else {
      av = a[sortColDet];
      bv = b[sortColDet];
    }
    if(av==null) av='';
    if(bv==null) bv='';
    if(typeof av==='string') return sortDirDet==='asc'?av.localeCompare(bv):bv.localeCompare(av);
    return sortDirDet==='asc'?av-bv:bv-av;
  });

  document.getElementById('detalleContador').textContent = filas.length.toLocaleString('es-AR') + ' movimientos';
  const MAX = 500;
  const mostrar = filas.slice(0, MAX);
  document.querySelector('#tablaDetalle tbody').innerHTML = mostrar.map(f=>{
    const tipoColor = f.tipo==='ingreso_planta' ? '#1e8e3e' : '#d93025';
    const motivosArr = Array.from(f.motivos);
    let motivo = motivosArr.join(' + ');
    if(motivo.length>40) motivo = motivo.slice(0,40)+'…';
    const fechaFmt = f.fecha.split('-').reverse().join('/');
    return '<tr>'+
      '<td>'+fechaFmt+'</td>'+
      '<td>'+f.ceve+'</td>'+
      '<td style="font-size:11px;color:'+tipoColor+'">'+motivo+'</td>'+
      '<td>'+(f.remito||'—')+'</td>'+
      '<td>'+nodoLabel(f.origen)+'</td>'+
      '<td>'+nodoLabel(f.destino)+'</td>'+
      '<td class="num">' +
  (f.bg ? fmt(f.bg) : '') +
'</td>' +

'<td class="num">' +
  (f.doll ? fmt(f.doll) : '') +
'</td>' +

'<td class="num">' +
  (f.pal ? fmt(f.pal) : '') +
'</td>' +

'</tr>';
  }).join('');
  if(filas.length > MAX){
    document.querySelector('#tablaDetalle tbody').innerHTML += '<tr><td colspan="9" style="text-align:center;background:#fef7e0;padding:10px">⚠️ Mostrando los primeros '+MAX+' de '+filas.length.toLocaleString('es-AR')+' movimientos. Descargá el Excel para ver todos.</td></tr>';
  }
  window._detalle = filas;
}
function exportarDetalle(){
  const filas = window._detalle || [];
  if(!filas.length){ alert('No hay movimientos para exportar.'); return; }
  const out = filas.map(f=>({
    Fecha: f.fecha.split('-').reverse().join('/'),
    CeVe: f.ceve,
    Tipo: f.tipo==='ingreso_planta' ? 'Ingreso' : 'Egreso',
    Motivo: Array.from(f.motivos).join(' + '),
    Remito: f.remito,
    Origen: nodoLabel(f.origen),
           Destino: nodoLabel(f.destino),
    Bandejas_BG: f.bg,
    Dollys_DO: f.doll,
    Pallets_PAL: f.pal
  }));
  const ws=XLSX.utils.json_to_sheet(out);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Detalle Movimientos');
  XLSX.writeFile(wb,'DetalleMovimientos_'+new Date().toISOString().slice(0,10)+'.xlsx');
}
function abrirModalMov(){ document.getElementById('modMov').classList.add('show'); }
function descargarMov(){
  const d=document.getElementById('movDesde').value, h=document.getElementById('movHasta').value;
  const r=document.getElementById('movRegion').value, c=document.getElementById('movCeve').value;
  const e=document.getElementById('movEquipo').value, t=document.getElementById('movTipo').value;
  const filas = MOVS.filter(m=>{
    const ce=MAESTRO.ceves.find(x=>x.agencia===m.agencia);
    if(!ce) return false;
    if(r && ce.region!==r) return false;
    if(c && m.agencia!==c) return false;
    if(e && m.equipo!==e) return false;
    if(d && m.fecha<d) return false;
    if(h && m.fecha>h) return false;
    return true;
  }).map(m=>{
    const ce=MAESTRO.ceves.find(x=>x.agencia===m.agencia);
    const row = {
  Fecha: m.fecha,
  Region: ce.region,
  CeVe: ce.nombre,
  Equipo:
    m.equipo === '028'
      ? 'BG-Bandeja Grande'
      : m.equipo === '031'
        ? 'DO-Dolly'
        : 'PAL-Pallet'
};
    if(t==='ing') row.Ingresos_Planta=m.ingresos_planta;
    else if(t==='egr') row.Egresos_Planta=m.egresos_planta;
    else { row.Ingresos_Planta=m.ingresos_planta; row.Egresos_Planta=m.egresos_planta; row.Saldo_Bruto=m.saldo_bruto; }
    return row;
  });
  if(!filas.length){ alert('No hay movimientos para los filtros seleccionados.'); return; }
  const ws=XLSX.utils.json_to_sheet(filas);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Movimientos');
  XLSX.writeFile(wb,'Movimientos_'+new Date().toISOString().slice(0,10)+'.xlsx');
  document.getElementById('modMov').classList.remove('show');
}
cargar();
</script></body></html>"""
INDEX = (
    INDEX
    .replace('__CDN_CHART__', CDN_CHART)
    .replace('__CDN_XLSX__', CDN_XLSX)
    .replace('__CDN_SUPABASE__', CDN_SUPABASE)
    .replace('__CSS__', CSS)
    .replace('__NAV__', NAV_DASH)
    .replace('__SUB_JS__', SUB_JS)
    .replace('__SUPABASE_URL__', SUPABASE_URL)
    .replace('__SUPABASE_KEY__', SUPABASE_PUBLISHABLE_KEY)
)

# ========= CARGA.HTML =========
CARGA = """<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8"><title>Carga Stock</title>
__CDN_PDF__
__CDN_SUPABASE__
<style>__CSS__.container{max-width:1100px;margin:0 auto}</style></head><body>
<header><h1>📝 Carga de Stock - CeVes Argentina</h1></header>
<div class="container">__NAV__
<div class="tabs"><button class="tab-btn active" id="btnD">📅 Carga Stock Diario</button><button class="tab-btn" id="btnM">📆 Carga Inventario Mensual</button></div>

<div id="tab-diario" class="tab-content active">
<div class="info">
  📅 Cargá el conteo físico diario de bandejas,
  dollys y pallets en tu CeVe.
</div>
<div class="warning"> ⚠️ <b>Importante:</b> el stock diario debe cargarse una sola vez por día. Antes de confirmar, verificá cuidadosamente el CeVe, la fecha y las cantidades. Si necesitás corregir una carga, realizá un nuevo envío y el dashboard tomará la carga más reciente. </div>
<div id="msgD"></div>
<form id="formD">
<div class="form-row"><div><label>CeVe</label><select id="dCeve" required></select></div><div><label>Fecha del conteo</label><input type="date" id="dFecha" required></div></div>
<div class="form-row"><div><label>Hora del conteo</label><input type="time" id="dHora" required></div><div><label>Supervisor / Responsable del conteo</label><input type="text" id="dSup" required placeholder="Nombre y apellido"></div></div>
<div class="section-title">📦 Cantidad por equipo</div>

<div class="form-row three">

  <div>
    <label>BG - Bandeja Grande (028)</label>
    <input
      type="number"
      id="dBG"
      min="0"
      value="0"
      required
    >
  </div>

  <div>
    <label>DO - Dolly (031)</label>
    <input
      type="number"
      id="dDO"
      min="0"
      value="0"
      required
    >
  </div>

  <div>
    <label>PAL - Pallet (132)</label>
    <input
      type="number"
      id="dPAL"
      min="0"
      value="0"
      required
    >
  </div>

</div>
<div class="form-row full"><div><label>Observaciones</label><textarea id="dObs"></textarea></div></div>
<button type="submit">💾 Guardar conteo diario</button>
</form>
</div>

<div id="tab-mensual" class="tab-content">
<div id="formMensualWrap">
<div class="warning">⚠️ <b>Atención:</b> Este formulario se carga UNA SOLA VEZ por mes y los datos se reportan a Gerencia.Verificá cuidadosamente la información antes de enviarla. Si ya existe un reporte de inventario del CeVe y período seleccionados, podrás registrar una nueva versión indicando obligatoriamente el motivo de la corrección.</div>
<div id="msgM"></div>
<form id="formM">
<div class="section-title">📍 Datos generales</div>
<div class="form-row"><div><label>CeVe</label><select id="mCeve" required></select></div><div><label>Mes / Año del inventario</label><input type="month" id="mFecha" required></div></div>
<div class="form-row three"><div><label>Fecha de toma del inventario</label><input type="date" id="mFC" required></div><div><label>Hora de toma del inventario</label><input type="time" id="mHC" required></div><div><label>Supervisor / Responsable</label><input type="text" id="mSup" required></div></div>

<div class="section-title">📦 Inventario físico</div>
<table class="inv-table">
<thead><tr><th>Equipo</th><th>Con producto (B/E o M/E)</th><th>Vacíos (rotos/sanos/color)</th><th>Otros</th><th>TOTAL</th></tr></thead>
<tbody>__INV_ROWS__</tbody>
</table>

<div class="section-title">🚛 Ingresos y Egresos del día</div>
<div class="info">Cargá acá los remitos / CONEQUIP del día del conteo. Click para abrir cada sección.</div>
__ACCORDIONS__

<div class="section-title">✍️ Firmas y observaciones</div>
<div class="form-row four"><div><label>Firma Op. Logístico</label><input type="text" id="mF1"></div><div><label>Firma Sup. Despacho</label><input type="text" id="mF2"></div><div><label>Firma Contador/Despacho</label><input type="text" id="mF3"></div><div><label>Firma Administrador</label><input type="text" id="mF4"></div></div>
<div class="form-row full"><div><label>Observaciones / Novedades</label><textarea id="mObs"></textarea></div></div>
<button type="submit">💾 Guardar inventario mensual</button>
</form>
</div>

<div id="postSave" style="display:none">
<div class="success-screen">
<div class="ok-icon">✓</div>
<h2>Inventario mensual guardado</h2>
<p>El inventario se registró correctamente y será enviado a Gerencia.</p>
<div style="display:flex;gap:10px;justify-content:center;margin-top:20px">
<button type="button" class="download-btn blue" onclick="descargarPDF()">⬇ Descargar Inventario Mensual en PDF</button>
<button type="button" class="btn-cancel" style="padding:10px 20px;border:none;border-radius:6px;cursor:pointer;font-weight:600" onclick="location.reload()">Volver</button>
</div>
</div>
<div style="display:flex;justify-content:center"><div id="pdfArea" class="pdf-summary"></div></div>
</div>

</div>
</div>

<div class="modal" id="modD"><div class="modal-box info"> <h3>📅 Carga de Stock Diario</h3> <p>Estás por registrar el stock diario de bandejas, dollys y pallets.</p> <p><b>Verificá el CeVe, la fecha y las cantidades antes de continuar.</b></p> <p>El reporte debe cargarse una sola vez por día. Si necesitás corregirlo, podés realizar un nuevo envío y el dashboard utilizará la carga más reciente.</p> <p>¿Estás seguro de que querés continuar?</p><div class="modal-actions"><button type="button" class="btn-cancel" onclick="cerrar('modD')">Cancelar</button><button type="button" class="btn-confirm" onclick="enviarD()">Sí, continuar</button></div></div></div>
<div class="modal" id="modM"><div class="modal-box warn"><h3>⚠️ Carga de Inventario Mensual</h3><p><b>Estás por ingresar a la Carga de Inventario Mensual</b>, ¿estás seguro que deseas completar la tarea?</p><p>Prestá mucha atención antes de guardar. Si el inventario del período ya fue presentado, el sistema solicitará una justificación y conservará ambas versiones en el historial.</b>.</p><p>Este dato se releva de todos los CeVes una vez por mes y se pasa a Gerencia.</p><div class="modal-actions"><button type="button" class="btn-cancel" onclick="cerrar('modM')">Cancelar</button><button type="button" class="btn-confirm warn" onclick="confirmarM()">Sí, continuar</button></div></div></div><div class="modal" id="modInventarioRepetido">
  <div class="modal-box warn">

    <h3>⚠️ Inventario ya registrado</h3>

    <p>
      Ya existe un inventario mensual para el CeVe
      y período seleccionados.
    </p>

    <p>
      Si necesitás corregir o reemplazar la información
      enviada, seleccioná <b>Continuar</b> e indicá el
      motivo de la nueva presentación.
    </p>

    <p>
      La carga anterior permanecerá disponible en el
      historial y la versión más reciente será utilizada
      para el seguimiento y los gráficos.
    </p>

    <p>
      Si querés revisar los datos antes de enviarlos,
      seleccioná <b>Cancelar</b>.
    </p>

    <div id="pasoMotivoInventario" style="display:none">

      <div class="section-title">
        Justificación de la nueva carga
      </div>

      <label>
        Motivo de la nueva presentación
      </label>

      <textarea
        id="motivoNuevaCarga"
        placeholder="Ej.: Corrección de la cantidad informada en bandejas."
        minlength="10"
      ></textarea>

      <div
        id="msgMotivoInventario"
        style="margin-top:10px"
      ></div>

    </div>

    <div class="modal-actions" id="accionesInventarioIniciales">

      <button
        type="button"
        class="btn-cancel"
        onclick="cancelarInventarioRepetido()"
      >
        Cancelar
      </button>

      <button
        type="button"
        class="btn-confirm warn"
        onclick="mostrarMotivoInventario()"
      >
        Continuar
      </button>

    </div>

    <div
      class="modal-actions"
      id="accionesInventarioMotivo"
      style="display:none"
    >

      <button
        type="button"
        class="btn-cancel"
        onclick="volverInventarioRepetido()"
      >
        Volver
      </button>

      <button
        type="button"
        class="btn-confirm warn"
        id="btnConfirmarNuevaCarga"
        onclick="confirmarInventarioRepetido()"
      >
        Confirmar nueva carga
      </button>

    </div>

  </div>
</div>

<script>
const SUBAG = __SUB_JS__;

const SUPABASE_URL = '__SUPABASE_URL__';
const SUPABASE_KEY = '__SUPABASE_KEY__';
const supabaseClient = window.supabase.createClient(
  SUPABASE_URL,
  SUPABASE_KEY
);

const EQUIPOS = [
  ['BG', 'BANDEJAS (028)'],
  ['DOG', 'DOLLYS Grandes (031)'],
  ['PM', 'Pallets Madera'],
  ['PP', 'Pallets Plástico'],
  ['PASC', 'Pascualineros']
];

let MAESTRO = null;
let INVENTARIO_PENDIENTE = null;
let GUARDANDO_INVENTARIO = false;

async function cargar(){
  const raw = await fetch(
    'data/ceves_maestro.json?ts=' + Date.now(),
    { cache: 'no-store' }
  ).then(r =&gt; r.json());

  const cevesBase = raw.ceves.map(
    c =&gt; ({...c})
  );

  const mdp = cevesBase.find(
    c =&gt;
      String(c.nombre)
        .trim()
        .toLowerCase() === 'mar del plata'
  );

  if(mdp){
    mdp.nombre = 'Mar del Plata + ST';
  }

  MAESTRO = {
    ...raw,
    ceves: cevesBase.filter(
      c =&gt; !SUBAG.includes(
        String(c.agencia)
      )
    )
  };

  ['dCeve','mCeve'].forEach(id =&gt; {
    const selector =
      document.getElementById(id);

    selector.innerHTML = '';

    MAESTRO.ceves.forEach(c =&gt; {
      selector.innerHTML +=
        '&lt;option value="' +
        c.agencia +
        '"&gt;' +
        c.nombre +
        '&lt;/option&gt;';
    });
  });

  const hoy = new Date();

  document.getElementById('dFecha').value =
    hoy.toISOString().slice(0,10);

  document.getElementById('dHora').value =
    hoy.toTimeString().slice(0,5);

  document.getElementById('mFecha').value =
    hoy.toISOString().slice(0,7);

  document.getElementById('mFC').value =
    hoy.toISOString().slice(0,10);

  document.getElementById('mHC').value =
    hoy.toTimeString().slice(0,5);
}

document.getElementById('btnD')
  .addEventListener(
    'click',
    () =&gt; activar('D')
  );

document.getElementById('btnM')
  .addEventListener(
    'click',
    () =&gt;
      document
        .getElementById('modM')
        .classList.add('show')
  );

function activar(t){
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(c=>c.classList.remove('active'));
  document.getElementById('tab-'+(t==='D'?'diario':'mensual')).classList.add('active');
  document.getElementById('btn'+t).classList.add('active');
}
function cerrar(id){document.getElementById(id).classList.remove('show')}
function confirmarM(){cerrar('modM');activar('M')}
function abrirInventarioRepetido(payload){
  INVENTARIO_PENDIENTE = payload;

  document.getElementById(
    'motivoNuevaCarga'
  ).value = '';

  document.getElementById(
    'msgMotivoInventario'
  ).innerHTML = '';

  document.getElementById(
    'pasoMotivoInventario'
  ).style.display = 'none';

  document.getElementById(
    'accionesInventarioIniciales'
  ).style.display = 'flex';

  document.getElementById(
    'accionesInventarioMotivo'
  ).style.display = 'none';

  document.getElementById(
    'modInventarioRepetido'
  ).classList.add('show');
}


function cancelarInventarioRepetido(){
  document.getElementById(
    'modInventarioRepetido'
  ).classList.remove('show');

  INVENTARIO_PENDIENTE = null;
}


function mostrarMotivoInventario(){
  document.getElementById(
    'pasoMotivoInventario'
  ).style.display = 'block';

  document.getElementById(
    'accionesInventarioIniciales'
  ).style.display = 'none';

  document.getElementById(
    'accionesInventarioMotivo'
  ).style.display = 'flex';

  document.getElementById(
    'motivoNuevaCarga'
  ).focus();
}


function volverInventarioRepetido(){
  document.getElementById(
    'pasoMotivoInventario'
  ).style.display = 'none';

  document.getElementById(
    'accionesInventarioIniciales'
  ).style.display = 'flex';

  document.getElementById(
    'accionesInventarioMotivo'
  ).style.display = 'none';

  document.getElementById(
    'msgMotivoInventario'
  ).innerHTML = '';
}


async function confirmarInventarioRepetido(){
  if(
    !INVENTARIO_PENDIENTE ||
    GUARDANDO_INVENTARIO
  ){
    return;
  }

  const motivo =
    document.getElementById(
      'motivoNuevaCarga'
    ).value.trim();

  if(motivo.length < 10){
    document.getElementById(
      'msgMotivoInventario'
    ).innerHTML =
      '<div class="msg err">' +
      'Indicá un motivo claro de al menos ' +
      '10 caracteres.' +
      '</div>';

    return;
  }

  const boton =
    document.getElementById(
      'btnConfirmarNuevaCarga'
    );

  boton.disabled = true;
  boton.textContent =
    'Guardando...';

  /*
  Se crea una copia para no modificar
  accidentalmente el objeto pendiente.
  */
  const payload = {
    ...INVENTARIO_PENDIENTE,
    motivo_nueva_carga:
      motivo
  };

  try {
    await guardar(
      payload,
      'msgM',
      'formM',
      mostrarExito
    );

    document.getElementById(
      'modInventarioRepetido'
    ).classList.remove('show');

    INVENTARIO_PENDIENTE = null;

  } catch(error){
    /*
    Si falla, el modal permanece abierto
    para que el supervisor no pierda el motivo.
    */
    document.getElementById(
      'msgMotivoInventario'
    ).innerHTML =
      '<div class="msg err">' +
      'No se pudo guardar la nueva versión: ' +
      (
        error.message ||
        'error no identificado'
      ) +
      '</div>';

  } finally {
    boton.disabled = false;
    boton.textContent =
      'Confirmar nueva carga';
  }
}

document.querySelectorAll('.acc-header').forEach(h=>h.addEventListener('click',()=>h.parentElement.classList.toggle('open')));
document.querySelectorAll('.btn-add').forEach(b=>{
  b.addEventListener('click',()=>{
    const seccion=b.parentElement.closest('.accordion').dataset.mov;
    agregarFila(seccion);
  });
});

function recalcTotal(input){
  const eq=input.dataset.eq;
  const cp=parseInt(document.getElementById('m'+eq+'CP').value)||0;
  const va=parseInt(document.getElementById('m'+eq+'VA').value)||0;
  const ot=parseInt(document.getElementById('m'+eq+'OT').value)||0;
  document.getElementById('tot'+eq).textContent=(cp+va+ot).toLocaleString('es-AR');
}

function agregarFila(s){
  const tb=document.querySelector('.accordion[data-mov="'+s+'"] tbody');
  const tr=document.createElement('tr');
  tr.innerHTML='<td><input type="date"></td><td><input type="text"></td><td><input type="text"></td><td><input type="text"></td><td><input type="text"></td><td><input type="time"></td><td><input type="number" min="0" value="0"></td><td><input type="number" min="0" value="0"></td><td><input type="number" min="0" value="0"></td><td><input type="number" min="0" value="0"></td><td><input type="number" min="0" value="0"></td><td><input type="number" min="0" value="0"></td><td><button type="button" class="btn-del" onclick="this.closest(\\'tr\\').remove()">✕</button></td>';
  tb.appendChild(tr);
}
function leerMov(s){
  const f=[];
  document.querySelectorAll('.accordion[data-mov="'+s+'"] tbody tr').forEach(tr=>{
    const i=tr.querySelectorAll('input');
    f.push({fecha:i[0].value,chofer:i[1].value,conequip:i[2].value,remito:i[3].value,destino:i[4].value,hora:i[5].value,band:+i[6].value||0,pasc:+i[7].value||0,dolly_gr:+i[8].value||0,dolly_ch:+i[9].value||0,madera:+i[10].value||0,plastico:+i[11].value||0});
  });
  return f;
}
function msg(id,txt,ok){document.getElementById(id).innerHTML='<div class="msg '+(ok?'ok':'err')+'">'+txt+'</div>';setTimeout(()=>document.getElementById(id).innerHTML='',6000)}

async function guardar(p, mi, fi, cb){
  try {
    if(p.tipo === 'mensual'){
      if(GUARDANDO_INVENTARIO) return;
      GUARDANDO_INVENTARIO = true;

      const { data, error } = await supabaseClient.rpc(
        'registrar_inventario_mensual',
        {
          p_agencia: String(p.agencia),
          p_ceve: p.ceve,
          p_mes: p.mes,
          p_fecha_conteo: p.fecha_conteo,
          p_hora_conteo: p.hora_conteo,
          p_supervisor: p.supervisor,
          p_inventario: p.inventario,
          p_movimientos: p.movimientos,
          p_firma_op_log: p.firma_op_log || null,
          p_firma_sup_desp: p.firma_sup_desp || null,
          p_firma_contador: p.firma_contador || null,
          p_firma_admin: p.firma_admin || null,
          p_observaciones: p.observaciones || '',
          p_motivo_nueva_carga: p.motivo_nueva_carga || null
        }
      );
      if(error) throw error;

      const guardado = Array.isArray(data) ? data[0] : data;
      if(!guardado) throw new Error('Supabase no devolvió el inventario guardado.');

      p.id = guardado.id;
      p.version = Number(guardado.version) || 1;
      p.vigente = guardado.vigente === true;
      p.reemplaza_id = guardado.reemplaza_id || null;
      p.motivo_nueva_carga = guardado.motivo_nueva_carga || p.motivo_nueva_carga || null;

      const formData = new FormData();
      const tipoPresentacion = p.version &gt; 1 ? 'RECTIFICACIÓN' : 'PRIMERA PRESENTACIÓN';
      formData.append('_subject', '📦 Inventario Mensual - ' + p.ceve + ' - ' + p.mes + ' - Versión ' + p.version);
      formData.append('CeVe', p.ceve);
      formData.append('Agencia', String(p.agencia));
      formData.append('Reportado por', p.supervisor);
      formData.append('Mes', p.mes);
      formData.append('Versión', String(p.version));
      formData.append('Tipo de presentación', tipoPresentacion);
      formData.append('Vigente', p.vigente ? 'Sí' : 'No');
      formData.append('Motivo de la nueva presentación', p.motivo_nueva_carga || 'No aplica. Primera presentación.');
      formData.append('Bandejas (BG-028)', (p.inventario.BG?.total || 0).toLocaleString('es-AR'));
      formData.append('Dollys (DO-031)', (p.inventario.DOG?.total || 0).toLocaleString('es-AR'));
      formData.append('Pallets Madera', (p.inventario.PM?.total || 0).toLocaleString('es-AR'));
      formData.append('Pallets Plástico', (p.inventario.PP?.total || 0).toLocaleString('es-AR'));
      formData.append('Pascualineros', (p.inventario.PASC?.total || 0).toLocaleString('es-AR'));
      formData.append('Ver seguimiento', 'https://logisticabimbo.github.io/Dashboard-Saldo-Primario/seguimiento.html');

      try {
        const respuestas = await Promise.all([
          fetch('__FORMSPREE_ABI__', {method:'POST', body:formData, headers:{Accept:'application/json'}}),
          fetch('__FORMSPREE_URI__', {method:'POST', body:formData, headers:{Accept:'application/json'}})
        ]);
        if(respuestas.some(r =&gt; !r.ok)) console.warn('Algún aviso de inventario no pudo enviarse.');
      } catch(errorCorreo){
        console.warn('Inventario guardado; falló el aviso:', errorCorreo);
      }

      try {
        const local = JSON.parse(localStorage.getItem('mensuales') || '[]');
        local.push(p);
        localStorage.setItem('mensuales', JSON.stringify(local));
      } catch(errorLocal){
        console.warn('No se pudo crear la copia local:', errorLocal);
      }
    } else {
      const { error } = await supabaseClient.from('stock_diario').insert({
        ceve: p.ceve,
        agencia: String(p.agencia),
        fecha_conteo: p.fecha,
        bg: Number(p.BG) || 0,
        do_stock: Number(p.DO) || 0,
        pal: Number(p.PAL) || 0,
        observaciones: p.observaciones || '',
        usuario: p.supervisor || ''
      });
      if(error) throw error;
    }

    msg(mi, p.tipo === 'mensual' ? 'Inventario mensual registrado como versión ' + p.version + '.' : 'Carga registrada correctamente.', true);
    if(cb) cb(p);
    else {
      document.getElementById(fi).reset();
      cargar();
    }
    return p;
  } catch(e){
    console.error('Error registrando la carga:', e);
    msg(mi, 'Error: ' + (e.message || 'No se pudo registrar la carga.'), false);
    throw e;
  } finally {
    if(p.tipo === 'mensual') GUARDANDO_INVENTARIO = false;
  }
}

document.getElementById('formD').addEventListener('submit',e=>{
  e.preventDefault();
  if(!document.getElementById('formD').checkValidity()) return;
  document.getElementById('modD').classList.add('show');
});
function enviarD(){
  cerrar('modD');
  const a=document.getElementById('dCeve').value, c=MAESTRO.ceves.find(x=>x.agencia===a);
  const fecha=document.getElementById('dFecha').value, hora=document.getElementById('dHora').value;
  const ts = fecha+'T'+hora+':00';
    guardar(
    {
      tipo: 'diario',
      agencia: c.agencia,
      ceve: c.nombre,
      fecha: fecha,
      hora: hora,

      supervisor:
        document.getElementById(
          'dSup'
        ).value,

      BG:
        Number(
          document.getElementById(
            'dBG'
          ).value
        ) || 0,

      DO:
        Number(
          document.getElementById(
            'dDO'
          ).value
        ) || 0,

      PAL:
        Number(
          document.getElementById(
            'dPAL'
          ).value
        ) || 0,

      observaciones:
        document.getElementById(
          'dObs'
        ).value,

      timestamp: ts
    },

    'msgD',
    'formD'
  );
}

document
  .getElementById('formM')
  .addEventListener(
    'submit',
    async evento => {
      evento.preventDefault();

      const formulario =
        document.getElementById(
          'formM'
        );

      if(!formulario.checkValidity()){
        formulario.reportValidity();
        return;
      }

      const mesInventario =
        document.getElementById(
          'mFecha'
        ).value;

      const fechaConteo =
        document.getElementById(
          'mFC'
        ).value;

      /*
      La fecha del conteo debe pertenecer
      al período informado.
      */
      if(
        mesInventario &&
        fechaConteo &&
        fechaConteo.slice(0, 7) !==
          mesInventario
      ){
        msg(
          'msgM',
          'La fecha de toma debe corresponder ' +
          'al mes y año del inventario.',
          false
        );

        return;
      }

      const agencia =
        document.getElementById(
          'mCeve'
        ).value;

      const ceve =
        MAESTRO.ceves.find(
          c =>
            String(c.agencia) ===
            String(agencia)
        );

      if(!ceve){
        msg(
          'msgM',
          'No se pudo identificar el CeVe seleccionado.',
          false
        );

        return;
      }

      const payload = {
        tipo:
          'mensual',

        agencia:
          String(ceve.agencia),

        ceve:
          ceve.nombre,

        mes:
          mesInventario,

        fecha_conteo:
          fechaConteo,

        hora_conteo:
          document.getElementById(
            'mHC'
          ).value,

        supervisor:
          document.getElementById(
            'mSup'
          ).value.trim(),

        inventario:
          {},

        firma_op_log:
          document.getElementById(
            'mF1'
          ).value.trim(),

        firma_sup_desp:
          document.getElementById(
            'mF2'
          ).value.trim(),

        firma_contador:
          document.getElementById(
            'mF3'
          ).value.trim(),

        firma_admin:
          document.getElementById(
            'mF4'
          ).value.trim(),

        movimientos: {
          recibidas_dentro:
            leerMov('rd'),

          recibidas_fuera:
            leerMov('rf'),

          despachadas_dentro:
            leerMov('dd'),

          despachadas_fuera:
            leerMov('df')
        },

        observaciones:
          document.getElementById(
            'mObs'
          ).value.trim(),

        motivo_nueva_carga:
          null,

        timestamp:
          new Date().toISOString()
      };

      EQUIPOS.forEach(
        ([codigo, nombre]) => {
          const conProducto =
            Number(
              document.getElementById(
                'm' + codigo + 'CP'
              ).value
            ) || 0;

          const vacios =
            Number(
              document.getElementById(
                'm' + codigo + 'VA'
              ).value
            ) || 0;

          const otros =
            Number(
              document.getElementById(
                'm' + codigo + 'OT'
              ).value
            ) || 0;

          payload.inventario[codigo] = {
            nombre:
              nombre,

            con_producto:
              conProducto,

            vacios:
              vacios,

            otros:
              otros,

            total:
              conProducto +
              vacios +
              otros
          };
        }
      );

      payload.BG =
        payload.inventario.BG.total;

      payload.DO =
        payload.inventario.DOG.total;

      /*
      Esta consulta solamente determina
      si hay que mostrar el modal.

      El cálculo definitivo de versión
      se realiza dentro de la RPC.
      */
      const {
        data: existente,
        error: errorConsulta
      } = await supabaseClient
        .from('inventarios_mensuales')
        .select(
          'id, version, supervisor, ' +
          'created_at, vigente'
        )
        .eq(
          'agencia',
          payload.agencia
        )
        .eq(
          'mes',
          payload.mes
        )
        .eq(
          'vigente',
          true
        )
        .maybeSingle();

      if(errorConsulta){
        msg(
          'msgM',
          'No se pudo verificar si ya existe ' +
          'una presentación: ' +
          errorConsulta.message,
          false
        );

        return;
      }

      if(existente){
        payload.version_anterior =
          Number(existente.version) || 1;

        payload.id_version_anterior =
          existente.id;

        abrirInventarioRepetido(
          payload
        );

        return;
      }

      /*
      Primera presentación.
      No requiere motivo.
      */
      try {
        await guardar(
          payload,
          'msgM',
          'formM',
          mostrarExito
        );
      } catch(error){
        /*
        guardar() ya muestra el mensaje.
        El catch evita un rechazo sin controlar.
        */
      }
    }
  );

function escaparHtml(valor){
  return String(
    valor === null ||
    valor === undefined
      ? ''
      : valor
  )
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function mostrarExito(p){
  document.getElementById('formMensualWrap').style.display='none';
  document.getElementById('postSave').style.display='block';
  const version =
  Number(p.version) || 1;

const tipoPresentacion =
  version > 1
    ? 'Rectificación'
    : 'Primera presentación';

let html =
  '<h2>Inventario Mensual — ' +
  escaparHtml(p.ceve) +
  '</h2>';

html +=
  '<div style="' +
  'background:#e8f0fe;' +
  'border-left:4px solid #0066b3;' +
  'padding:10px;' +
  'margin-bottom:10px;' +
  'font-size:12px' +
  '">' +

  '<b>Versión ' +
  version +
  '</b> · ' +
  tipoPresentacion +

  (
    p.vigente
      ? ' · <b>VIGENTE</b>'
      : ''
  ) +

  '</div>';

html +=
  '<p>' +
  '<b>Mes:</b> ' +
  escaparHtml(p.mes) +
  ' &nbsp; ' +
  '<b>Fecha conteo:</b> ' +
  escaparHtml(p.fecha_conteo) +
  ' ' +
  escaparHtml(p.hora_conteo) +
  '</p>';

html +=
  '<p><b>Supervisor:</b> ' +
  escaparHtml(p.supervisor) +
  '</p>';

if(p.motivo_nueva_carga){
  html +=
    '<h3>Motivo de la nueva presentación</h3>' +
    '<p>' +
    escaparHtml(
      p.motivo_nueva_carga
    ) +
    '</p>';
}

  html += '<h3>Inventario físico</h3><table><tr><th>Equipo</th><th>Con prod.</th><th>Vacíos</th><th>Otros</th><th>TOTAL</th></tr>';
  Object.values(p.inventario).forEach(e=>{
    html += '<tr><td>'+e.nombre+'</td><td>'+e.con_producto+'</td><td>'+e.vacios+'</td><td>'+e.otros+'</td><td><b>'+e.total+'</b></td></tr>';
  });
  html += '</table>';
  const titulos = {recibidas_dentro:'Recibidas (DENTRO del conteo)', recibidas_fuera:'Recibidas (NO dentro del conteo)', despachadas_dentro:'Despachadas (DENTRO del conteo)', despachadas_fuera:'Despachadas (NO dentro del conteo)'};
['recibidas_dentro','recibidas_fuera','despachadas_dentro','despachadas_fuera'].forEach(k=>{
    const movs = (p.movimientos[k] || []).filter(m =>
      (m.fecha || m.chofer || m.conequip || m.remito || m.destino || m.hora ||
       m.band || m.pasc || m.dolly_gr || m.dolly_ch || m.madera || m.plastico)
    );
    if(!movs.length) return;
    html += '<h3>'+titulos[k]+'</h3><table><tr><th>Fecha</th><th>Chofer</th><th>CONEQ</th><th>Remito</th><th>Dest</th><th>Hora</th><th>Band</th><th>Pasc</th><th>D.Gr</th><th>D.Ch</th><th>Mad</th><th>Plast</th></tr>';
    movs.forEach(m=>{ html += '<tr><td>'+m.fecha+'</td><td>'+m.chofer+'</td><td>'+m.conequip+'</td><td>'+m.remito+'</td><td>'+m.destino+'</td><td>'+m.hora+'</td><td>'+m.band+'</td><td>'+m.pasc+'</td><td>'+m.dolly_gr+'</td><td>'+m.dolly_ch+'</td><td>'+m.madera+'</td><td>'+m.plastico+'</td></tr>'; });
    html += '</table>';
  });
  html += '<h3>Firmas</h3>';
  html += '<table><tr><th>Op. Logístico</th><th>Sup. Despacho</th><th>Contador</th><th>Administrador</th></tr><tr><td>'+p.firma_op_log+'</td><td>'+p.firma_sup_desp+'</td><td>'+p.firma_contador+'</td><td>'+p.firma_admin+'</td></tr></table>';
  if(p.observaciones) html += '<h3>Observaciones</h3><p>'+p.observaciones+'</p>';
  document.getElementById('pdfArea').innerHTML = html;
  window._pdfPayload = p;
}

function descargarPDF(){
  const el = document.getElementById('pdfArea');
  const p = window._pdfPayload;
  if(!p){ alert('No hay un inventario preparado para descargar.'); return; }

  const w = window.open('', '_blank', 'width=900,height=900');
  if(!w){ alert('Habilitá las ventanas emergentes para descargar el PDF.'); return; }

  w.document.write(
    '&lt;!DOCTYPE html&gt;&lt;html&gt;&lt;head&gt;&lt;meta charset="UTF-8"&gt;' +
    '&lt;title&gt;Inventario Mensual - ' + escaparHtml(p.ceve) + ' - ' +
    escaparHtml(p.mes) + ' - Versión ' + (Number(p.version) || 1) + '&lt;/title&gt;' +
    '&lt;style&gt;' +
    'body{font-family:Arial,sans-serif;color:#222;padding:20px;max-width:800px;margin:0 auto}' +
    'h2{color:#003e7e;border-bottom:2px solid #003e7e;padding-bottom:6px;font-size:18px}' +
    'h3{color:#003e7e;font-size:13px;margin-top:16px;border-bottom:1px solid #ccc;padding-bottom:3px}' +
    'p{font-size:12px;margin:4px 0}table{width:100%;border-collapse:collapse;margin:8px 0}' +
    'th{background:#003e7e;color:#fff;padding:6px;font-size:11px;text-align:left}' +
    'td{border:1px solid #999;padding:5px;font-size:11px}@media print{body{padding:10px}}' +
    '&lt;/style&gt;&lt;/head&gt;&lt;body&gt;'
  );
  w.document.write(el.innerHTML);
  w.document.write('&lt;p style="text-align:center;margin-top:20px;font-size:11px;color:#666"&gt;Usá Ctrl+P para guardar como PDF&lt;/p&gt;');
  w.document.write('&lt;/body&gt;&lt;/html&gt;');
  w.document.close();
  setTimeout(() =&gt; { w.focus(); w.print(); }, 500);
}

cargar();
</script></body></html>"""
CARGA = (
    CARGA
    .replace('__CDN_PDF__', CDN_PDF)
    .replace('__CDN_SUPABASE__', CDN_SUPABASE)
    .replace('__CSS__', CSS)
    .replace('__NAV__', NAV_CARGA)
    .replace('__SUB_JS__', SUB_JS)
    .replace('__INV_ROWS__', INV_ROWS)
    .replace('__ACCORDIONS__', ACCORDIONS)
    .replace('__FORMSPREE_ABI__', FORMSPREE_MENSUAL_ABI)
    .replace('__FORMSPREE_URI__', FORMSPREE_MENSUAL_URI)
    .replace('__SUPABASE_URL__', SUPABASE_URL)
    .replace('__SUPABASE_KEY__', SUPABASE_PUBLISHABLE_KEY)
)
# Usuarios autorizados para entrar a seguimiento
USUARIOS_SEGUIMIENTO = {
    'abigail.montoya': '123123',
    'uriel.desteffanis': '123123',
}
USR_JS = _json.dumps(USUARIOS_SEGUIMIENTO)

# ========= FACTURACION.HTML =========
FACTURACION = """&lt;!DOCTYPE html&gt;
&lt;html lang="es"&gt;
&lt;head&gt;
&lt;meta charset="UTF-8"&gt;
&lt;title&gt;Carga Facturaciones&lt;/title&gt;
__CDN_SUPABASE__
&lt;style&gt;
__CSS__
.container{
  max-width:1100px;
  margin:0 auto
}
.fact-summary{
  background:#e8f0fe;
  color:#003e7e;
  padding:14px;
  border-radius:8px;
  margin-bottom:16px;
  font-size:13px
}
.fact-summary strong{
  color:#003e7e
}
&lt;/style&gt;
&lt;/head&gt;

&lt;body&gt;

&lt;header&gt;
  &lt;h1&gt;🧾 Carga de Facturaciones - Saldo Primario&lt;/h1&gt;
&lt;/header&gt;

&lt;div class="container"&gt;
__NAV__

&lt;div class="card"&gt;

  &lt;h3&gt;🧾 Carga Facturaciones&lt;/h3&gt;

  &lt;div class="info"&gt;
    Registrá las bandejas o dollys facturados a clientes,
    QSR o transportes.
  &lt;/div&gt;

  &lt;div class="warning"&gt;
    ⚠️ &lt;b&gt;Importante:&lt;/b&gt;
    la cantidad facturada será incorporada al Saldo Primario
    como una devolución manual a Planta Pilar y se sumará a
    los egresos del CeVe seleccionado. Verificá cuidadosamente
    el CeVe, el número de factura, el código y la cantidad.
  &lt;/div&gt;

  &lt;div id="msgF"&gt;&lt;/div&gt;

  &lt;form id="formF"&gt;

    &lt;div class="section-title"&gt;
      📍 Datos de la facturación
    &lt;/div&gt;

    &lt;div class="form-row"&gt;

      &lt;div&gt;
        &lt;label&gt;CeVe&lt;/label&gt;
        &lt;select id="fFacCeve" required&gt;&lt;/select&gt;
      &lt;/div&gt;

      &lt;div&gt;
        &lt;label&gt;Fecha de facturación&lt;/label&gt;
        &lt;input
          type="date"
          id="fFacFecha"
          required
        &gt;
      &lt;/div&gt;

    &lt;/div&gt;

    &lt;div class="form-row"&gt;

      &lt;div&gt;
        &lt;label&gt;Nº de factura&lt;/label&gt;
        &lt;input
          type="text"
          id="fFacNumero"
          required
          placeholder="Ej.: 0005-00012345"
        &gt;
      &lt;/div&gt;

      &lt;div&gt;
        &lt;label&gt;Cliente&lt;/label&gt;
        &lt;input
          type="text"
          id="fFacCliente"
          required
          placeholder="Cliente, QSR o transporte"
        &gt;
      &lt;/div&gt;

    &lt;/div&gt;

    &lt;div class="section-title"&gt;
      📦 Equipo facturado
    &lt;/div&gt;

    &lt;div class="form-row"&gt;

      &lt;div&gt;
        &lt;label&gt;Cód. facturado&lt;/label&gt;
        &lt;select id="fFacCodigo" required&gt;
          &lt;option value="028"&gt;
            028 - BG Bandeja Grande
          &lt;/option&gt;
          &lt;option value="031"&gt;
            031 - DO Dolly
          &lt;/option&gt;
        &lt;/select&gt;
      &lt;/div&gt;

      &lt;div&gt;
        &lt;label&gt;Cantidad facturada&lt;/label&gt;
        &lt;input
          type="number"
          id="fFacCantidad"
          min="1"
          step="1"
          value="1"
          required
        &gt;
      &lt;/div&gt;

    &lt;/div&gt;

    &lt;div class="fact-summary" id="resumenFacturacion"&gt;
      Completá los datos para registrar la facturación.
    &lt;/div&gt;

    &lt;button type="submit"&gt;
      💾 Guardar facturación
    &lt;/button&gt;

  &lt;/form&gt;

&lt;/div&gt;
&lt;/div&gt;

&lt;div class="modal" id="modFacturacion"&gt;
  &lt;div class="modal-box info"&gt;

    &lt;h3&gt;🧾 Confirmar facturación&lt;/h3&gt;

    &lt;p&gt;
      Esta carga será incorporada al Saldo Primario
      como egreso manual a Planta Pilar.
    &lt;/p&gt;

    &lt;div id="confirmacionFacturacion"&gt;&lt;/div&gt;

    &lt;p&gt;
      &lt;b&gt;¿Querés continuar?&lt;/b&gt;
    &lt;/p&gt;

    &lt;div class="modal-actions"&gt;
      &lt;button
        type="button"
        class="btn-cancel"
        onclick="cerrarFacturacion()"
      &gt;
        Cancelar
      &lt;/button&gt;

      &lt;button
        type="button"
        class="btn-confirm"
        id="btnConfirmarFacturacion"
      &gt;
        Sí, guardar
      &lt;/button&gt;
    &lt;/div&gt;

  &lt;/div&gt;
&lt;/div&gt;

&lt;script&gt;

const SUBAG = __SUB_JS__;

const SUPABASE_URL = '__SUPABASE_URL__';
const SUPABASE_KEY = '__SUPABASE_KEY__';

const supabaseClient =
  window.supabase.createClient(
    SUPABASE_URL,
    SUPABASE_KEY
  );

let MAESTRO = null;
let FACTURA_PENDIENTE = null;


async function cargar(){
  const raw = await fetch(
    'data/ceves_maestro.json?ts=' + Date.now(),
    { cache: 'no-store' }
  ).then(r =&gt; r.json());

  const cevesBase = raw.ceves.map(
    c =&gt; ({...c})
  );

  const mdp = cevesBase.find(
    c =&gt;
      String(c.nombre)
        .trim()
        .toLowerCase() === 'mar del plata'
  );

  if(mdp){
    mdp.nombre = 'Mar del Plata + ST';
  }

  MAESTRO = {
    ...raw,
    ceves: cevesBase.filter(
      c =&gt;
        !SUBAG.includes(
          String(c.agencia)
        )
    )
  };

  const selector =
    document.getElementById('fFacCeve');

  selector.innerHTML = '';

  MAESTRO.ceves.forEach(c =&gt; {
    selector.innerHTML +=
      '&lt;option value="' +
      c.agencia +
      '"&gt;' +
      c.nombre +
      '&lt;/option&gt;';
  });

  document.getElementById(
    'fFacFecha'
  ).value =
    new Date()
      .toISOString()
      .slice(0,10);

  actualizarResumen();
}


function datosFormulario(){
  const agencia =
    document.getElementById(
      'fFacCeve'
    ).value;

  const ceve =
    MAESTRO.ceves.find(
      c =&gt;
        String(c.agencia) ===
        String(agencia)
    );

  return {
    agencia: String(agencia),
    ceve: ceve ? ceve.nombre : '',
    fecha_facturacion:
      document.getElementById(
        'fFacFecha'
      ).value,
    numero_factura:
      document.getElementById(
        'fFacNumero'
      ).value.trim(),
    cliente:
      document.getElementById(
        'fFacCliente'
      ).value.trim(),
    codigo_facturado:
      document.getElementById(
        'fFacCodigo'
      ).value,
    cantidad_facturada:
      Number(
        document.getElementById(
          'fFacCantidad'
        ).value
      ) || 0
  };
}


function nombreEquipo(codigo){
  return codigo === '028'
    ? 'BG - Bandeja Grande'
    : 'DO - Dolly';
}


function actualizarResumen(){
  if(!MAESTRO) return;

  const factura = datosFormulario();

  document.getElementById(
    'resumenFacturacion'
  ).innerHTML =
    '&lt;strong&gt;CeVe:&lt;/strong&gt; ' +
    (factura.ceve || '—') +
    ' &amp;nbsp; | &amp;nbsp; ' +
    '&lt;strong&gt;Equipo:&lt;/strong&gt; ' +
    nombreEquipo(
      factura.codigo_facturado
    ) +
    ' &amp;nbsp; | &amp;nbsp; ' +
    '&lt;strong&gt;Cantidad:&lt;/strong&gt; ' +
    factura.cantidad_facturada
      .toLocaleString('es-AR');
}


document.querySelectorAll(
  '#fFacCeve, #fFacFecha, ' +
  '#fFacNumero, #fFacCliente, ' +
  '#fFacCodigo, #fFacCantidad'
).forEach(elemento =&gt; {
  elemento.addEventListener(
    'input',
    actualizarResumen
  );

  elemento.addEventListener(
    'change',
    actualizarResumen
  );
});


document
  .getElementById('formF')
  .addEventListener(
    'submit',
    evento =&gt; {
      evento.preventDefault();

      const formulario =
        document.getElementById('formF');

      if(!formulario.checkValidity()){
        formulario.reportValidity();
        return;
      }

      FACTURA_PENDIENTE =
        datosFormulario();

      document.getElementById(
        'confirmacionFacturacion'
      ).innerHTML =
        '&lt;div class="fact-summary"&gt;' +

        '&lt;div&gt;&lt;b&gt;CeVe:&lt;/b&gt; ' +
        FACTURA_PENDIENTE.ceve +
        '&lt;/div&gt;' +

        '&lt;div&gt;&lt;b&gt;Fecha:&lt;/b&gt; ' +
        FACTURA_PENDIENTE
          .fecha_facturacion
          .split('-')
          .reverse()
          .join('/') +
        '&lt;/div&gt;' +

        '&lt;div&gt;&lt;b&gt;Factura:&lt;/b&gt; ' +
        FACTURA_PENDIENTE.numero_factura +
        '&lt;/div&gt;' +

        '&lt;div&gt;&lt;b&gt;Cliente:&lt;/b&gt; ' +
        FACTURA_PENDIENTE.cliente +
        '&lt;/div&gt;' +

        '&lt;div&gt;&lt;b&gt;Equipo:&lt;/b&gt; ' +
        nombreEquipo(
          FACTURA_PENDIENTE
            .codigo_facturado
        ) +
        '&lt;/div&gt;' +

        '&lt;div&gt;&lt;b&gt;Cantidad:&lt;/b&gt; ' +
        FACTURA_PENDIENTE
          .cantidad_facturada
          .toLocaleString('es-AR') +
        '&lt;/div&gt;' +

        '&lt;div&gt;&lt;b&gt;Destino:&lt;/b&gt; ' +
        'Planta Pilar' +
        '&lt;/div&gt;' +

        '&lt;/div&gt;';

      document.getElementById(
        'modFacturacion'
      ).classList.add('show');
    }
  );


function cerrarFacturacion(){
  document.getElementById(
    'modFacturacion'
  ).classList.remove('show');

  FACTURA_PENDIENTE = null;
}


function mostrarMensaje(texto, correcto){
  document.getElementById(
    'msgF'
  ).innerHTML =
    '&lt;div class="msg ' +
    (correcto ? 'ok' : 'err') +
    '"&gt;' +
    texto +
    '&lt;/div&gt;';
}


document.getElementById(
  'btnConfirmarFacturacion'
).addEventListener(
  'click',
  async () =&gt; {
    if(!FACTURA_PENDIENTE) return;

    const boton =
      document.getElementById(
        'btnConfirmarFacturacion'
      );

    boton.disabled = true;
    boton.textContent =
      'Guardando...';

    try {
      const { error } =
        await supabaseClient
          .from('facturaciones')
          .insert({
            agencia:
              FACTURA_PENDIENTE.agencia,

            ceve:
              FACTURA_PENDIENTE.ceve,

            fecha_facturacion:
              FACTURA_PENDIENTE
                .fecha_facturacion,

            numero_factura:
              FACTURA_PENDIENTE
                .numero_factura,

            cliente:
              FACTURA_PENDIENTE.cliente,

            codigo_facturado:
              FACTURA_PENDIENTE
                .codigo_facturado,

            cantidad_facturada:
              FACTURA_PENDIENTE
                .cantidad_facturada
          });

      if(error) throw error;

      document.getElementById(
        'modFacturacion'
      ).classList.remove('show');

      mostrarMensaje(
        '✅ Facturación registrada ' +
        'correctamente.',
        true
      );

      document.getElementById(
        'formF'
      ).reset();

      document.getElementById(
        'fFacFecha'
      ).value =
        new Date()
          .toISOString()
          .slice(0,10);

      FACTURA_PENDIENTE = null;

      cargar();

    } catch(error) {
      console.error(
        'Error guardando facturación:',
        error
      );

      document.getElementById(
        'modFacturacion'
      ).classList.remove('show');

      const duplicada =
        String(error.message || '')
          .toLowerCase()
          .includes('duplicate');

      mostrarMensaje(
        duplicada
          ? '❌ Esta factura ya fue ' +
            'registrada para el CeVe ' +
            'y el equipo seleccionados.'
          : '❌ No se pudo guardar: ' +
            error.message,
        false
      );

    } finally {
      boton.disabled = false;
      boton.textContent =
        'Sí, guardar';
    }
  }
);


cargar();

&lt;/script&gt;
&lt;/body&gt;
&lt;/html&gt;
"""

FACTURACION = (
    FACTURACION
    .replace(
      '__CDN_SUPABASE__',
      CDN_SUPABASE
    )
    .replace(
      '__CSS__',
      CSS
    )
    .replace(
      '__NAV__',
      NAV_FACTURACION
    )
    .replace(
      '__SUB_JS__',
      SUB_JS
    )
    .replace(
      '__SUPABASE_URL__',
      SUPABASE_URL
    )
    .replace(
      '__SUPABASE_KEY__',
      SUPABASE_PUBLISHABLE_KEY
    )
)

# ========= SEGUIMIENTO.HTML =========
SEG = """<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8"><title>Acceso Control de Equipos</title>
__CDN_XLSX__
<style>__CSS__</style></head><body>
<header><h1>🔐 Acceso Control de Equipos - CeVes Argentina</h1></header>	
<div class="container" id="contenidoSeg" style="display:none">__NAV__
<div class="summary" id="kpis"></div>
<div class="card"><div style="display:flex;align-items:center;margin-bottom:12px;gap:12px"><h3 style="margin:0">📅 Cumplimiento Diario (últimos 7 días)</h3><button class="download-btn" onclick="exp('d')">⬇ Excel</button></div>
<table id="tD"><thead><tr><th>CeVe</th><th>Región</th><th>Supervisor</th><th>Última carga</th><th>Hora</th><th>Días sin cargar</th><th>Estado</th></tr></thead><tbody></tbody></table></div>
<div class="card"><div style="display:flex;align-items:center;margin-bottom:12px;gap:12px"><h3 style="margin:0">📆 Cumplimiento Mensual</h3><button class="download-btn" onclick="exp('m')">⬇ Excel</button></div>
<table id="tM"><thead><tr><th>CeVe</th><th>Región</th><th>Supervisor</th><th>Último mes</th><th>Fecha y hora carga</th><th>Estado</th></tr></thead><tbody></tbody></table></div>
<div class="card"><div style="display:flex;align-items:center;margin-bottom:12px;gap:12px"><h3 style="margin:0">🗓 Calendario (últimos 30 días)</h3><button class="download-btn" onclick="exp('c')">⬇ Excel</button></div>
<div class="cal-wrap"><table class="cal" id="tC"></table></div></div>
<div class="card"><div style="display:flex;align-items:center;margin-bottom:12px;gap:12px"><h3 style="margin:0">📋 Historial (últimas 30)</h3><button class="download-btn" onclick="exp('h')">⬇ Excel</button></div>
<table id="tH"><thead><tr><th>Fecha y hora</th><th>Tipo</th><th>CeVe</th><th>Responsable</th><th>BG</th><th>DO</th><th>Observaciones</th></tr></thead><tbody></tbody></table></div>
</div>

<div class="modal show" id="modLogin">
  <div class="modal-box info">

    <h3>🔒 Acceso exclusivo</h3>

    <p>
      Esta sección está reservada para el equipo de
      Control de Equipos. Ingresá tu usuario y clave
      para continuar.
    </p>
<div class="form-row"><div><label>Usuario</label><input type="text" id="logUsr" placeholder="usuario.apellido"></div><div><label>Clave</label><input type="password" id="logPwd"></div></div>
<div id="msgLog"></div>
<div class="modal-actions">
<button type="button" class="btn-confirm" onclick="login()">Ingresar</button>
</div></div></div>

<script>
const USUARIOS = __USR_JS__;
const SUBAG = __SUB_JS__;
let MAESTRO=null, STOCK={registros_diarios:[],registros_mensuales:[]};

function login(){
  const u = document.getElementById('logUsr').value.trim().toLowerCase();
  const p = document.getElementById('logPwd').value;
  if(USUARIOS[u] && USUARIOS[u] === p){
    document.getElementById('modLogin').classList.remove('show');
    document.getElementById('contenidoSeg').style.display = 'block';
    cargar();
  } else {
    document.getElementById('msgLog').innerHTML = '<div class="msg err">❌ Usuario o clave incorrectos.</div>';
  }
}
document.addEventListener('keydown', e => {
  if(e.key==='Enter' && document.getElementById('modLogin').classList.contains('show')) login();
});
async function cargar(){
  const raw=await fetch('data/ceves_maestro.json').then(r=>r.json());
  MAESTRO={...raw, ceves: raw.ceves.filter(c=>!SUBAG.includes(c.agencia))};
  try{STOCK=await fetch('data/stock_diario.json?ts='+Date.now()).then(r=>r.json())}catch(e){}
  render();
}
function fmt(ts){return ts?ts.slice(0,16).replace('T',' '):'—'}
function fmtH(ts){return ts?ts.slice(11,16):'—'}
function render(){
  const h=new Date(), mA=h.toISOString().slice(0,7);
  const tbD=document.querySelector('#tD tbody'); tbD.innerHTML='';
  let oD=0,wD=0,eD=0; const fD=[];
  MAESTRO.ceves.forEach(c=>{
    const r=(STOCK.registros_diarios||[]).filter(x=>x.agencia===c.agencia).sort((a,b)=>(b.timestamp||'').localeCompare(a.timestamp||''));
    const u=r[0], uf=u?u.fecha:null, ho=u?fmtH(u.timestamp):'—';
    const d=uf?Math.floor((h-new Date(uf))/86400000):null;
    let st,b;
    if(d===null){st='Sin cargar';b='err';eD++}
    else if(d<=1){st='Al día';b='ok';oD++}
    else if(d<=3){st='Atrasado';b='warn';wD++}
    else{st='Crítico';b='err';eD++}
    fD.push({CeVe:c.nombre,Region:c.region,Supervisor:c.supervisor,UltimaCarga:uf||'',Hora:ho,Dias:d!==null?d:'',Estado:st});
    tbD.innerHTML += '<tr><td>'+c.nombre+'</td><td>'+c.region+'</td><td>'+c.supervisor+'</td><td>'+(uf||'—')+'</td><td>'+ho+'</td><td>'+(d!==null?d:'—')+'</td><td><span class="badge '+b+'">'+st+'</span></td></tr>';
  });
  window._d=fD;
  const tbM=document.querySelector('#tM tbody'); tbM.innerHTML='';
  let oM=0,eM=0; const fM=[];
  MAESTRO.ceves.forEach(c=>{
    const r=(STOCK.registros_mensuales||[]).filter(x=>x.agencia===c.agencia).sort((a,b)=>(b.mes||'').localeCompare(a.mes||''));
    const u=r[0], um=u?u.mes:null, ts=u?fmt(u.timestamp):'—';
    const p=um!==mA;
    const st=!um?'Sin cargar':(p?'Pendiente':'Al día');
    const b=!um||p?'err':'ok';
    if(b==='ok') oM++; else eM++;
    fM.push({CeVe:c.nombre,Region:c.region,Supervisor:c.supervisor,UltimoMes:um||'',FechaCarga:ts,Estado:st});
    tbM.innerHTML += '<tr><td>'+c.nombre+'</td><td>'+c.region+'</td><td>'+c.supervisor+'</td><td>'+(um||'—')+'</td><td>'+ts+'</td><td><span class="badge '+b+'">'+st+'</span></td></tr>';
  });
  window._m=fM;
  document.getElementById('kpis').innerHTML =
    '<div class="kpi"><div class="label">Al día (Diario)</div><div class="value">'+oD+'/'+MAESTRO.ceves.length+'</div></div>'+
    '<div class="kpi"><div class="label">Atrasados</div><div class="value">'+wD+'</div></div>'+
    '<div class="kpi"><div class="label">Críticos</div><div class="value">'+eD+'</div></div>'+
    '<div class="kpi"><div class="label">Mes al día</div><div class="value">'+oM+'/'+MAESTRO.ceves.length+'</div></div>';
  renderCal();
  const hi=[...(STOCK.registros_diarios||[]),...(STOCK.registros_mensuales||[])].sort((a,b)=>(b.timestamp||'').localeCompare(a.timestamp||'')).slice(0,30);
  const tbH=document.querySelector('#tH tbody'); tbH.innerHTML=''; const fH=[];
  hi.forEach(r=>{
    fH.push({Fecha:fmt(r.timestamp),Tipo:r.tipo,CeVe:r.ceve,Responsable:r.supervisor||'',BG:r.BG||0,DO:r.DO||0,Observaciones:r.observaciones||''});
    tbH.innerHTML += '<tr><td>'+fmt(r.timestamp)+'</td><td>'+(r.tipo==='diario'?'📅 Diario':'📆 Mensual')+'</td><td>'+r.ceve+'</td><td>'+(r.supervisor||'—')+'</td><td>'+(r.BG||0)+'</td><td>'+(r.DO||0)+'</td><td>'+(r.observaciones||'—')+'</td></tr>';
  });
  window._h=fH;
}
function renderCal(){
  const h=new Date(), ds=[];
  for(let i=29;i>=0;i--){const d=new Date(h);d.setDate(h.getDate()-i);ds.push(d.toISOString().slice(0,10))}
  let html='<thead><tr><th>CeVe</th>';
  ds.forEach(d=>{const dt=new Date(d);html+='<th title="'+d+'">'+String(dt.getDate()).padStart(2,'0')+'/'+String(dt.getMonth()+1).padStart(2,'0')+'</th>'});
  html+='</tr></thead><tbody>';
  const hS=h.toISOString().slice(0,10), fC=[];
  MAESTRO.ceves.forEach(c=>{
    const fi={CeVe:c.nombre};
    html+='<tr><td class="ceve-name">'+c.nombre+'</td>';
    ds.forEach(d=>{
      const cg=(STOCK.registros_diarios||[]).some(r=>r.agencia===c.agencia&&r.fecha===d);
      const fu=d>hS, cl=fu?'c-fut':(cg?'c-ok':'c-no'), tk=cg?'✓':'';
      html+='<td class="'+cl+'">'+tk+'</td>';
      fi[d]=cg?'OK':(fu?'':'NO');
    });
    html+='</tr>'; fC.push(fi);
  });
  html+='</tbody>';
  document.getElementById('tC').innerHTML=html;
  window._c=fC;
}
function exp(t){
  const m={d:'_d',m:'_m',c:'_c',h:'_h'};
  const d=window[m[t]];
  if(!d||!d.length){alert('Sin datos');return}
  const ws=XLSX.utils.json_to_sheet(d);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,t);
  XLSX.writeFile(wb,'Seguimiento_'+t+'_'+new Date().toISOString().slice(0,10)+'.xlsx');
}
cargar();
</script></body></html>"""
SEG = (SEG.replace('__CDN_XLSX__', CDN_XLSX).replace('__CSS__', CSS)
           .replace('__NAV__', NAV_SEG).replace('__SUB_JS__', SUB_JS).replace('__USR_JS__', USR_JS))

for nombre, contenido in [
    ('index.html', INDEX),
    ('carga.html', CARGA),
    ('facturacion.html', FACTURACION),
    ('seguimiento.html', SEG)
]:
    contenido_final = _html.unescape(
        _html.unescape(contenido)
    )

    with open(
        nombre,
        'w',
        encoding='utf-8'
    ) as f:
        f.write(contenido_final)

    print(
        f'✅ {nombre} creado '
        f'({len(contenido_final):,} caracteres)'
    )

print(
    '\n🎉 Listo. Refrescá las pestañas '
    'del navegador con Ctrl+F5.'
)

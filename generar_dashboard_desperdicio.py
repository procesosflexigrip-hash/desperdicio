"""
Flexigrip — Generador de Dashboard de Desperdicio
===================================================
Uso:
    python generar_dashboard_desperdicio.py                      # busca Completo.xlsx en la misma carpeta
    python generar_dashboard_desperdicio.py mi_archivo.xlsx       # especifica otro archivo
    python generar_dashboard_desperdicio.py archivo.xlsx -o reporte.html

Lógica:
    % desperdicio por área = Desperdicio / (KG Real + Desperdicio) x 100
    % desperdicio TOTAL    = suma(Desperdicio de todas las áreas en ruta)
                              / KG Solicitados de la ÚLTIMA área de la ruta x 100

Semáforo (igual para % por área y % total):
    Verde   <= 9.9%
    Amarillo 10% - 19.9%
    Rojo    >= 20%

Requiere: openpyxl
    pip install openpyxl
"""

import sys
import os
import json
import argparse
from datetime import datetime

try:
    import openpyxl
except ImportError:
    print("ERROR: Falta la librería openpyxl.")
    print("Instálala con:  pip install openpyxl")
    sys.exit(1)


# ── Argumentos ────────────────────────────────────────────────────────────────

parser = argparse.ArgumentParser(description="Genera el dashboard HTML de desperdicio Flexigrip.")
parser.add_argument("archivo", nargs="?", default="Completo.xlsx",
                    help="Ruta al archivo Excel (default: Completo.xlsx)")
parser.add_argument("-o", "--output", default=None,
                    help="Nombre del archivo HTML de salida (default: dashboard_desperdicio.html)")
args = parser.parse_args()

excel_path = args.archivo
if not os.path.exists(excel_path):
    print(f"ERROR: No se encontró el archivo '{excel_path}'")
    sys.exit(1)

output_path = args.output or "dashboard_desperdicio.html"


# ── Lectura del Excel ─────────────────────────────────────────────────────────

print(f"Leyendo {excel_path}...")
wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
ws = wb.active

def sd(v):
    if v is None: return None
    if hasattr(v, 'year'): return v.strftime('%Y-%m-%d')
    return None

def sn(v):
    if v is None: return 0
    try: return float(v)
    except: return 0

def pct_desp(kg_real, desp):
    denom = kg_real + desp
    if denom == 0: return None
    return round(desp / denom * 100, 2)

ORDER_SEQ = ['IMP', 'LAM', 'REF', 'POU', 'BOL']

orders = []
for row in ws.iter_rows(min_row=3, values_only=True):
    orden = row[0]
    if not orden:
        continue

    # Detección de áreas en ruta
    imp_in = row[1] is not None
    lam_in = row[7] is not None and sn(row[8]) > 0
    ref_in = row[13] is not None and sn(row[14]) > 0
    pou_in = row[19] is not None and (sn(row[20]) > 0 or sn(row[21]) > 0)
    bol_in = row[25] is not None and (sn(row[26]) > 0 or sn(row[27]) > 0)

    imp_kg_real = sn(row[4]); imp_desp = sn(row[6])
    lam_kg_real = sn(row[10]); lam_desp = sn(row[12])
    ref_kg_real = sn(row[16]); ref_desp = sn(row[18])
    pou_kg_real = sn(row[22]); pou_desp = sn(row[24])
    bol_kg_real = sn(row[28]); bol_desp = sn(row[30])

    areas = {
        "IMP": {
            "f": sd(row[1]), "in": imp_in,
            "kg_sol": sn(row[2]), "kg_real": imp_kg_real, "desp": imp_desp,
            "pct": pct_desp(imp_kg_real, imp_desp) if imp_in else None,
        },
        "LAM": {
            "f": sd(row[7]), "in": lam_in,
            "kg_sol": sn(row[8]), "kg_real": lam_kg_real, "desp": lam_desp,
            "pct": pct_desp(lam_kg_real, lam_desp) if lam_in else None,
        },
        "REF": {
            "f": sd(row[13]), "in": ref_in,
            "kg_sol": sn(row[14]), "kg_real": ref_kg_real, "desp": ref_desp,
            "pct": pct_desp(ref_kg_real, ref_desp) if ref_in else None,
        },
        "POU": {
            "f": sd(row[19]), "in": pou_in,
            "kg_sol": sn(row[20]), "kg_real": pou_kg_real, "desp": pou_desp,
            "pct": pct_desp(pou_kg_real, pou_desp) if pou_in else None,
        },
        "BOL": {
            "f": sd(row[25]), "in": bol_in,
            "kg_sol": sn(row[27]), "kg_real": bol_kg_real, "desp": bol_desp,
            "pct": pct_desp(bol_kg_real, bol_desp) if bol_in else None,
        },
    }

    # % desperdicio total = suma(desp de áreas en ruta) / kg_sol de la ÚLTIMA área en ruta
    route_areas = [k for k in ORDER_SEQ if areas[k]["in"]]
    total_desp = sum(areas[k]["desp"] for k in route_areas)
    last_area = route_areas[-1] if route_areas else None
    kg_sol_last = areas[last_area]["kg_sol"] if last_area else 0
    pct_total = round(total_desp / kg_sol_last * 100, 2) if kg_sol_last > 0 else None

    orders.append({
        "o": str(orden),
        "areas": areas,
        "total_desp": total_desp,
        "kg_sol_last": kg_sol_last,
        "last_area": last_area,
        "pct_total": pct_total,
    })

wb.close()
print(f"  {len(orders)} órdenes leídas.")


# ── Generación del HTML ───────────────────────────────────────────────────────

data_json = json.dumps(orders, separators=(',', ':'))
generated_at = datetime.now().strftime('%d/%m/%Y %H:%M')

HTML_HEAD = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Flexigrip — Dashboard de Desperdicio</title>
<style>
  :root {{
    --bg:#0f1117;--surface:#1a1d27;--surface2:#232636;--border:#2d3149;
    --text:#e2e4f0;--muted:#6b7094;--green:#22c55e;--yellow:#f59e0b;
    --red:#ef4444;--accent:#818cf8;--header-h:64px;
  }}
  *{{box-sizing:border-box;margin:0;padding:0;}}
  body{{background:var(--bg);color:var(--text);font-family:system-ui,-apple-system,sans-serif;font-size:13px;min-height:100vh;}}
  header{{position:sticky;top:0;z-index:100;height:var(--header-h);background:var(--surface);border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;padding:0 24px;gap:16px;}}
  .logo{{font-size:17px;font-weight:700;letter-spacing:-.3px;color:var(--accent);}}
  .logo span{{color:var(--muted);font-weight:400;}}
  .logo small{{color:var(--muted);font-size:11px;font-weight:400;margin-left:8px;}}
  .header-stats{{display:flex;gap:20px;}}
  .stat{{text-align:center;}}
  .stat-val{{font-size:18px;font-weight:700;}}
  .stat-lab{{color:var(--muted);font-size:10px;text-transform:uppercase;letter-spacing:.5px;}}
  .controls{{background:var(--surface);border-bottom:1px solid var(--border);padding:12px 24px;display:flex;gap:12px;align-items:center;flex-wrap:wrap;}}
  .search-wrap{{position:relative;flex:1;min-width:180px;max-width:280px;}}
  .search-wrap input{{width:100%;background:var(--bg);border:1px solid var(--border);color:var(--text);padding:7px 12px 7px 32px;border-radius:6px;font-size:13px;}}
  .search-wrap::before{{content:"\\1F50D";position:absolute;left:9px;top:50%;transform:translateY(-50%);font-size:11px;}}
  select{{background:var(--bg);border:1px solid var(--border);color:var(--text);padding:7px 10px;border-radius:6px;font-size:12px;cursor:pointer;}}
  select:hover{{border-color:var(--accent);}}
  .filter-label{{color:var(--muted);font-size:11px;white-space:nowrap;}}
  .summary-row{{display:flex;gap:12px;padding:12px 24px;background:var(--surface);border-bottom:1px solid var(--border);}}
  .scard{{flex:1;background:var(--bg);border:1px solid var(--border);border-radius:8px;padding:10px 14px;display:flex;align-items:center;gap:10px;}}
  .scard-dot{{width:10px;height:10px;border-radius:50%;}}
  .scard-info strong{{font-size:20px;font-weight:700;display:block;}}
  .scard-info span{{color:var(--muted);font-size:11px;}}
  .sg strong{{color:var(--green);}} .sg .scard-dot{{background:var(--green);}}
  .sy strong{{color:var(--yellow);}} .sy .scard-dot{{background:var(--yellow);}}
  .sr strong{{color:var(--red);}} .sr .scard-dot{{background:var(--red);}}
  .sn strong{{color:var(--muted);}} .sn .scard-dot{{background:var(--muted);}}
  .table-wrap{{overflow-x:auto;padding:0 24px 24px;}}
  table{{width:100%;border-collapse:collapse;}}
  thead th{{position:sticky;top:var(--header-h);background:var(--surface2);border-bottom:1px solid var(--border);padding:10px 8px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);cursor:pointer;user-select:none;white-space:nowrap;}}
  thead th:hover{{color:var(--text);}}
  thead th.sorted{{color:var(--accent);}}
  .th-sub{{font-size:9px;color:var(--muted);font-weight:400;display:block;text-transform:none;letter-spacing:0;margin-top:2px;opacity:.8;}}
  tbody tr{{border-bottom:1px solid var(--border);transition:background .1s;cursor:pointer;}}
  tbody tr:hover{{background:var(--surface2);}}
  td{{padding:9px 8px;vertical-align:middle;}}
  .orden-cell{{font-weight:600;font-family:monospace;font-size:12px;color:var(--accent);}}
  .route{{display:flex;align-items:center;gap:3px;flex-wrap:wrap;}}
  .area-pill{{padding:2px 7px;border-radius:12px;font-size:10px;font-weight:600;letter-spacing:.3px;background:var(--surface2);color:var(--muted);border:1px solid var(--border);}}
  .area-pill.ap-IMP{{border-color:#6366f1;color:#818cf8;background:rgba(99,102,241,.08);}}
  .area-pill.ap-LAM{{border-color:#06b6d4;color:#22d3ee;background:rgba(6,182,212,.08);}}
  .area-pill.ap-REF{{border-color:#a78bfa;color:#c4b5fd;background:rgba(167,139,250,.08);}}
  .area-pill.ap-POU{{border-color:#f97316;color:#fb923c;background:rgba(249,115,22,.08);}}
  .area-pill.ap-BOL{{border-color:#ec4899;color:#f472b6;background:rgba(236,72,153,.08);}}
  .arrow{{color:var(--border);font-size:9px;}}
  .sem-cell{{display:flex;align-items:flex-start;gap:5px;}}
  .sem-dot{{width:7px;height:7px;border-radius:50%;flex-shrink:0;margin-top:3px;}}
  .sem-pct{{font-size:12px;font-weight:600;line-height:1.4;}}
  .sem-line{{color:var(--muted);font-size:10px;line-height:1.4;}}
  .sg-c .sem-dot{{background:var(--green);}} .sg-c .sem-pct{{color:var(--green);}}
  .sy-c .sem-dot{{background:var(--yellow);}} .sy-c .sem-pct{{color:var(--yellow);}}
  .sr-c .sem-dot{{background:var(--red);}} .sr-c .sem-pct{{color:var(--red);}}
  .sn-c .sem-dot{{background:var(--muted);}} .sn-c .sem-pct{{color:var(--muted);}}
  .badge{{display:inline-flex;align-items:center;gap:5px;padding:3px 8px;border-radius:20px;font-size:11px;font-weight:600;}}
  .badge-g{{background:rgba(34,197,94,.15);color:var(--green);}}
  .badge-y{{background:rgba(245,158,11,.15);color:var(--yellow);}}
  .badge-r{{background:rgba(239,68,68,.15);color:var(--red);}}
  .badge-n{{background:rgba(107,112,148,.12);color:var(--muted);}}
  .total-cell{{display:flex;flex-direction:column;align-items:flex-end;gap:2px;}}
  .total-pct{{font-size:14px;font-weight:700;}}
  .total-detail{{font-size:10px;color:var(--muted);}}
  .pagination{{display:flex;gap:6px;justify-content:center;padding:16px 24px;align-items:center;}}
  .pagination button{{min-width:32px;height:32px;background:var(--bg);border:1px solid var(--border);color:var(--text);border-radius:6px;cursor:pointer;font-size:12px;}}
  .pagination button:hover{{border-color:var(--accent);}}
  .pagination .current{{border-color:var(--accent);color:var(--accent);}}
  .pager-info{{color:var(--muted);font-size:12px;margin:0 8px;}}
  .modal-overlay{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.75);z-index:200;align-items:center;justify-content:center;}}
  .modal-overlay.open{{display:flex;}}
  .modal{{background:var(--surface);border:1px solid var(--border);border-radius:12px;width:min(680px,95vw);max-height:90vh;overflow-y:auto;padding:24px;}}
  .modal-header{{display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;}}
  .modal-title{{font-size:16px;font-weight:700;}}
  .modal-close{{background:none;border:none;color:var(--muted);font-size:20px;cursor:pointer;line-height:1;}}
  .modal-close:hover{{color:var(--text);}}
  .total-banner{{background:var(--bg);border:1px solid var(--border);border-radius:8px;padding:14px;margin-bottom:14px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;}}
  .total-banner-label{{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;}}
  .total-banner-value{{font-size:22px;font-weight:800;}}
  .total-banner-sub{{font-size:11px;color:var(--muted);margin-top:2px;}}
  .area-block{{background:var(--bg);border:1px solid var(--border);border-radius:8px;padding:14px;margin-bottom:10px;}}
  .area-block-header{{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;}}
  .area-block-name{{font-weight:700;font-size:13px;}}
  .area-block-skip{{color:var(--muted);font-size:12px;font-style:italic;}}
  .detail-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:8px;}}
  .detail-item label{{display:block;color:var(--muted);font-size:10px;text-transform:uppercase;letter-spacing:.4px;margin-bottom:3px;}}
  .detail-item span{{font-size:13px;font-weight:600;}}
  .prog-bar{{height:5px;background:var(--border);border-radius:3px;overflow:hidden;margin-top:8px;}}
  .prog-fill{{height:100%;border-radius:3px;}}
  .col-right{{text-align:right;}}
  .empty{{text-align:center;color:var(--muted);padding:40px;}}
</style>
</head>
<body>

<header>
  <div class="logo">Flexigrip <span>Desperdicio</span><small>Generado: {generated_at}</small></div>
  <div class="header-stats">
    <div class="stat"><div class="stat-val" id="hd-total">&ndash;</div><div class="stat-lab">Órdenes</div></div>
    <div class="stat"><div class="stat-val" id="hd-areas">5</div><div class="stat-lab">Áreas</div></div>
    <div class="stat"><div class="stat-val" style="color:var(--green)" id="hd-ok">&ndash;</div><div class="stat-lab">En verde</div></div>
  </div>
</header>

<div class="controls">
  <div class="search-wrap"><input id="search" placeholder="Buscar orden…" autocomplete="off"></div>
  <span class="filter-label">Área:</span>
  <select id="f-area">
    <option value="">Todas</option>
    <option value="IMP">Impresión</option>
    <option value="LAM">Laminación</option>
    <option value="REF">Refinado</option>
    <option value="POU">Pouch</option>
    <option value="BOL">Bolseo</option>
  </select>
  <span class="filter-label">Semáforo:</span>
  <select id="f-sem">
    <option value="">Todos (total)</option>
    <option value="G">🟢 Verde</option>
    <option value="Y">🟡 Amarillo</option>
    <option value="R">🔴 Rojo</option>
    <option value="N">⚫ Sin datos</option>
  </select>
  <span class="filter-label">Ruta:</span>
  <select id="f-route">
    <option value="">Cualquiera</option>
    <option value="IMP-BOL">IMP → BOL</option>
    <option value="IMP-LAM-REF">IMP → LAM → REF</option>
    <option value="IMP-LAM-REF-POU">IMP → LAM → REF → POU</option>
    <option value="IMP-REF-BOL">IMP → REF → BOL</option>
    <option value="IMP-LAM-POU">IMP → LAM → POU</option>
    <option value="IMP-POU">IMP → POU</option>
  </select>
  <span class="filter-label">Mostrar:</span>
  <select id="f-page-size">
    <option value="25">25</option>
    <option value="50" selected>50</option>
    <option value="100">100</option>
  </select>
</div>

<div class="summary-row">
  <div class="scard sg"><div class="scard-dot"></div><div class="scard-info"><strong id="cnt-g">&ndash;</strong><span>Verde (&le;9.9% desperdicio total)</span></div></div>
  <div class="scard sy"><div class="scard-dot"></div><div class="scard-info"><strong id="cnt-y">&ndash;</strong><span>Amarillo (10%&ndash;19.9% desperdicio total)</span></div></div>
  <div class="scard sr"><div class="scard-dot"></div><div class="scard-info"><strong id="cnt-r">&ndash;</strong><span>Rojo (&ge;20% desperdicio total)</span></div></div>
  <div class="scard sn"><div class="scard-dot"></div><div class="scard-info"><strong id="cnt-n">&ndash;</strong><span>Sin procesar</span></div></div>
</div>

<div class="table-wrap">
  <table>
    <thead>
      <tr>
        <th onclick="sort('orden')" id="th-orden">ORDEN &#x21D5;</th>
        <th>RUTA</th>
        <th onclick="sort('imp')" id="th-imp">IMPRESIÓN &#x21D5;<span class="th-sub">% desperdicio (kg)</span></th>
        <th onclick="sort('lam')" id="th-lam">LAMINACIÓN &#x21D5;<span class="th-sub">% desperdicio (kg)</span></th>
        <th onclick="sort('ref')" id="th-ref">REFINADO &#x21D5;<span class="th-sub">% desperdicio (kg)</span></th>
        <th onclick="sort('pou')" id="th-pou">POUCH &#x21D5;<span class="th-sub">% desperdicio (kg)</span></th>
        <th onclick="sort('bol')" id="th-bol">BOLSEO &#x21D5;<span class="th-sub">% desperdicio (kg)</span></th>
        <th class="col-right" onclick="sort('total')" id="th-total">% DESPERDICIO TOTAL &#x21D5;<span class="th-sub">suma desp. / kg sol. &uacute;ltima &aacute;rea</span></th>
      </tr>
    </thead>
    <tbody id="table-body"></tbody>
  </table>
  <div id="empty" class="empty" style="display:none">No se encontraron órdenes con los filtros seleccionados.</div>
</div>
<div class="pagination" id="pagination"></div>

<div class="modal-overlay" id="modal-overlay" onclick="if(event.target===this)closeModal()">
  <div class="modal">
    <div class="modal-header">
      <div class="modal-title" id="modal-title">Orden</div>
      <button class="modal-close" onclick="closeModal()">&#x2715;</button>
    </div>
    <div id="modal-body"></div>
  </div>
</div>

<script>
"""

JS_BODY = """
const RAW={data_json};
function semaforo(p){if(p===null||p===undefined)return'N';if(p<=9.9)return'G';if(p<=19.9)return'Y';return'R';}
function getRoute(a){return['IMP','LAM','REF','POU','BOL'].filter(k=>a[k].in).join('-');}
const DATA=RAW.map(d=>({...d,route:getRoute(d.areas),worst:semaforo(d.pct_total),_sems:{IMP:semaforo(d.areas.IMP.pct),LAM:semaforo(d.areas.LAM.pct),REF:semaforo(d.areas.REF.pct),POU:semaforo(d.areas.POU.pct),BOL:semaforo(d.areas.BOL.pct)}}));
let filtered=[...DATA],sortKey=null,sortDir=1,page=1,pageSize=50;
const SC={G:'sg-c',Y:'sy-c',R:'sr-c',N:'sn-c'};
const SL={G:'&#x1F7E2;',Y:'&#x1F7E1;',R:'&#x1F534;',N:'&#x26AB;'};
const SB={G:'badge-g',Y:'badge-y',R:'badge-r',N:'badge-n'};
const BAR={G:'var(--green)',Y:'var(--yellow)',R:'var(--red)',N:'var(--muted)'};
const NAMES={IMP:'Impresi\\u00F3n',LAM:'Laminaci\\u00F3n',REF:'Refinado',POU:'Pouch',BOL:'Bolseo'};
function fN(n){if(!n&&n!==0)return'&ndash;';return n.toLocaleString('es-MX',{maximumFractionDigits:1});}
function fP(p){if(p===null||p===undefined)return'&ndash;';return p.toFixed(1)+'%';}
function routeHtml(r){if(!r)return'&ndash;';return'<div class="route">'+r.split('-').map((p,i)=>`${i?'<span class="arrow">&#x203A;</span>':''}<span class="area-pill ap-${p}">${p}</span>`).join('')+'</div>';}
function cell(k,d){const a=d.areas[k];if(!a.in)return'<td><span style="color:var(--muted);font-size:11px">&ndash;</span></td>';const s=d._sems[k],sc=SC[s];return`<td><div class="sem-cell ${sc}"><div class="sem-dot"></div><div><div class="sem-pct">${fP(a.pct)}</div><div class="sem-line">Desp: ${fN(a.desp)} kg</div><div class="sem-line">Real: ${fN(a.kg_real)} kg &middot; Sol: ${fN(a.kg_sol)} kg</div></div></div></td>`;}
function totalCell(d){const s=d.worst,sc=SC[s];if(d.pct_total===null||d.pct_total===undefined){return`<td class="col-right"><span class="badge badge-n">&#x26AB; Sin datos</span></td>`;}return`<td class="col-right"><div class="total-cell ${sc}"><div class="total-pct">${SL[s]} ${fP(d.pct_total)}</div><div class="total-detail">${fN(d.total_desp)} kg / ${fN(d.kg_sol_last)} kg (${d.last_area})</div></div></td>`;}
function render(){const tb=document.getElementById('table-body'),el=document.getElementById('empty'),sl=filtered.slice((page-1)*pageSize,page*pageSize);if(!sl.length){tb.innerHTML='';el.style.display='block';}else{el.style.display='none';tb.innerHTML=sl.map(d=>`<tr onclick="openModal('${d.o}')"><td class="orden-cell">${d.o}</td><td>${routeHtml(d.route)}</td>${cell('IMP',d)}${cell('LAM',d)}${cell('REF',d)}${cell('POU',d)}${cell('BOL',d)}${totalCell(d)}</tr>`).join('');}renderPagination();updateSummary();}
function renderPagination(){const tot=Math.ceil(filtered.length/pageSize),el=document.getElementById('pagination');if(tot<=1){el.innerHTML='';return;}let h='';if(page>1)h+=`<button onclick="goPage(${page-1})">&#x2039;</button>`;const s=Math.max(1,page-2),e=Math.min(tot,page+2);if(s>1)h+=`<button onclick="goPage(1)">1</button>${s>2?'<span class="pager-info">&hellip;</span>':''}`;for(let i=s;i<=e;i++)h+=`<button onclick="goPage(${i})" class="${i===page?'current':''}">${i}</button>`;if(e<tot)h+=`${e<tot-1?'<span class="pager-info">&hellip;</span>':''}<button onclick="goPage(${tot})">${tot}</button>`;if(page<tot)h+=`<button onclick="goPage(${page+1})">&#x203A;</button>`;h+=`<span class="pager-info">${filtered.length} &oacute;rdenes</span>`;el.innerHTML=h;}
function goPage(p){page=p;render();window.scrollTo(0,0);}
function updateSummary(){const c={G:0,Y:0,R:0,N:0};filtered.forEach(d=>c[d.worst]++);document.getElementById('cnt-g').textContent=c.G;document.getElementById('cnt-y').textContent=c.Y;document.getElementById('cnt-r').textContent=c.R;document.getElementById('cnt-n').textContent=c.N;document.getElementById('hd-total').textContent=filtered.length;document.getElementById('hd-ok').textContent=c.G;}
function applyFilters(){const q=document.getElementById('search').value.trim().toUpperCase(),fa=document.getElementById('f-area').value,fs=document.getElementById('f-sem').value,fr=document.getElementById('f-route').value;pageSize=parseInt(document.getElementById('f-page-size').value);filtered=DATA.filter(d=>{if(q&&!d.o.includes(q))return false;if(fa&&!d.areas[fa].in)return false;if(fs){if((fa?d._sems[fa]:d.worst)!==fs)return false;}if(fr&&d.route!==fr)return false;return true;});if(sortKey)applySort(false);page=1;render();}
function applySort(re=true){const am={imp:'IMP',lam:'LAM',ref:'REF',pou:'POU',bol:'BOL'};filtered.sort((a,b)=>{let va,vb;if(sortKey==='orden'){va=a.o;vb=b.o;}else if(sortKey==='total'){va=a.pct_total??-1;vb=b.pct_total??-1;}else{const ak=am[sortKey];va=a.areas[ak]?.in?(a.areas[ak].pct??-1):-1;vb=b.areas[ak]?.in?(b.areas[ak].pct??-1):-1;}return(va<vb?-1:va>vb?1:0)*sortDir;});if(re)render();}
function sort(key){if(sortKey===key)sortDir*=-1;else{sortKey=key;sortDir=1;}document.querySelectorAll('thead th').forEach(t=>t.classList.remove('sorted'));document.getElementById('th-'+key)?.classList.add('sorted');applySort();}
function openModal(orden){const d=DATA.find(x=>x.o===orden);if(!d)return;document.getElementById('modal-title').textContent='Orden '+orden;const COLORS={IMP:'#818cf8',LAM:'#22d3ee',REF:'#c4b5fd',POU:'#fb923c',BOL:'#f472b6'};const s=d.worst,sc=SC[s];let html=`<div style="margin-bottom:14px">${routeHtml(d.route)}</div>`;html+=`<div class="total-banner"><div><div class="total-banner-label">Desperdicio Total de la Orden</div><div class="total-banner-value ${sc}">${SL[s]} ${fP(d.pct_total)}</div><div class="total-banner-sub">${fN(d.total_desp)} kg desperdiciados / ${fN(d.kg_sol_last)} kg solicitados en ${d.last_area?NAMES[d.last_area]:'&ndash;'} (&uacute;ltima &aacute;rea)</div></div></div>`;['IMP','LAM','REF','POU','BOL'].forEach(k=>{const a=d.areas[k],s2=d._sems[k],sc2=SC[s2],barColor=BAR[s2],barW=a.pct!==null?Math.min(100,a.pct):0;html+=`<div class="area-block" style="${!a.in?'opacity:.4':''}"><div class="area-block-header"><div class="area-block-name" style="color:${COLORS[k]}">${NAMES[k]}</div>${a.in?`<span class="${sc2}" style="font-weight:700;font-size:14px">${SL[s2]} ${fP(a.pct)}</span>`:'<span class="area-block-skip">No incluida en ruta</span>'}</div>`;if(a.in){html+=`<div class="detail-grid"><div class="detail-item"><label>KG Solicitados</label><span>${fN(a.kg_sol)} kg</span></div><div class="detail-item"><label>KG Reales</label><span>${fN(a.kg_real)} kg</span></div><div class="detail-item"><label>Desperdicio</label><span>${fN(a.desp)} kg</span></div><div class="detail-item"><label>Fecha</label><span>${a.f||'&ndash;'}</span></div></div><div class="prog-bar"><div class="prog-fill" style="width:${barW}%;background:${barColor}"></div></div>`;}html+='</div>';});document.getElementById('modal-body').innerHTML=html;document.getElementById('modal-overlay').classList.add('open');}
function closeModal(){document.getElementById('modal-overlay').classList.remove('open');}
document.addEventListener('keydown',e=>{if(e.key==='Escape')closeModal();});
['search','f-area','f-sem','f-route','f-page-size'].forEach(id=>{document.getElementById(id).addEventListener(id==='search'?'input':'change',applyFilters);});
filtered=[...DATA];render();
</script>
</body>
</html>"""

HTML = HTML_HEAD + JS_BODY.replace('{data_json}', data_json)

with open(output_path, 'w', encoding='utf-8') as f:
    f.write(HTML)

print(f"  Dashboard generado: {output_path}")
print(f"  Ábrelo en cualquier navegador.")

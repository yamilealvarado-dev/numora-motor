<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Numora — Compras</title>
<link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<style>
  :root{--navy:#172A37;--petro:#14465B;--gold:#FFC701;--mint:#D8E8E2;--bg:#FAFBFB;--line:#E7ECEC;--muted:#5C6E76;--ok:#1F7A52;--warn:#9A6B00}
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Poppins',sans-serif;background:var(--bg);color:var(--navy);min-height:100vh}
  .top{background:var(--navy);color:#fff;padding:18px 30px;display:flex;align-items:center;gap:12px}
  .mark{width:36px;height:36px;border-radius:10px;background:var(--gold);display:grid;place-items:center}
  .mark svg{width:19px;height:19px}
  .top .name{font-weight:600;font-size:21px;letter-spacing:-.5px}
  .top .sub{font-size:8px;letter-spacing:3px;color:var(--gold);font-weight:500}
  .wrap{max-width:1150px;margin:0 auto;padding:28px 24px}
  .eyebrow{font-size:11px;letter-spacing:2.5px;color:var(--petro);font-weight:600;text-transform:uppercase}
  h1{font-size:25px;font-weight:600;letter-spacing:-.5px;margin:6px 0 3px}
  .hsub{color:var(--muted);font-size:14px;font-weight:300;margin-bottom:24px}
  .card{background:#fff;border:1px solid var(--line);border-radius:16px;padding:22px;margin-bottom:18px}
  .drops{display:grid;grid-template-columns:repeat(2,1fr);gap:14px}
  .drop{border:1.5px dashed #C4D2D2;border-radius:13px;padding:16px;display:flex;flex-direction:column;gap:6px;transition:.15s}
  .drop.has{border-color:var(--ok);background:#F2FAF6}
  .drop label{font-size:13px;font-weight:500}
  .drop small{font-size:11px;color:var(--muted);font-weight:300}
  .drop .opt{font-size:10px;color:var(--petro);font-weight:600;letter-spacing:.5px}
  input[type=file]{font-size:12px;font-family:'Poppins';margin-top:4px}
  .btn{font-family:'Poppins';font-size:14px;font-weight:600;padding:13px 26px;border-radius:11px;border:none;background:var(--gold);color:var(--navy);cursor:pointer;margin-top:18px}
  .btn:disabled{background:var(--mint);color:#8AA39B;cursor:not-allowed}
  .btn.sec{background:#fff;border:1px solid var(--line);color:var(--petro);font-weight:500}
  .summary{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin:18px 0}
  .s{background:#fff;border:1px solid var(--line);border-radius:13px;padding:14px}
  .s .k{font-size:10.5px;color:var(--muted)}
  .s .v{font-size:20px;font-weight:600;margin-top:4px}
  .s.match .v{color:var(--ok)}
  table{width:100%;border-collapse:collapse;font-size:12.5px;background:#fff}
  th{background:var(--petro);color:#fff;font-size:10px;letter-spacing:.5px;text-transform:uppercase;padding:10px 12px;text-align:left}
  td{padding:10px 12px;border-bottom:1px solid #F1F4F4}
  td.num{text-align:right;font-variant-numeric:tabular-nums}
  tr.nuevo{background:#FFF6EC}tr.rev{background:#FFFBEE}tr.div{background:#F1F8F4}
  .pill{font-size:10.5px;font-weight:600;padding:3px 9px;border-radius:12px}
  .pill.ok{background:var(--mint);color:var(--ok)}
  .pill.w{background:#FCEFC7;color:var(--warn)}
  .pill.d{background:#D9EDE1;color:var(--ok)}
  .loading{display:none;text-align:center;padding:40px;color:var(--muted)}
  .spin{width:34px;height:34px;border:3px solid var(--mint);border-top-color:var(--gold);border-radius:50%;animation:r .8s linear infinite;margin:0 auto 14px}
  @keyframes r{to{transform:rotate(360deg)}}
  .err{background:#FDECEC;color:#B3261E;padding:12px 16px;border-radius:10px;font-size:13px;display:none;margin-top:12px}
  #results{display:none}
  .dl{display:flex;gap:10px;margin:8px 0 20px}
</style>
</head>
<body>
  <div class="top">
    <div class="mark"><svg viewBox="0 0 24 24" fill="none"><path d="M5 12.5l4.5 4.5L19 7" stroke="#172A37" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round"/></svg></div>
    <div><div class="name">numora</div><div class="sub">TAXES &amp; FINANCE</div></div>
  </div>
  <div class="wrap">
    <div class="eyebrow">Compras</div>
    <h1>Contabilizar compras</h1>
    <div class="hsub">Sube los archivos y Numora arma los asientos aprendiendo de tu año anterior.</div>

    <div class="card" id="form">
      <div class="drops">
        <div class="drop" id="d1"><span class="opt">OBLIGATORIO</span><label>Auxiliar del año anterior</label><small>Para aprender la cuenta de cada proveedor (.xlsx)</small><input type="file" id="auxiliar" accept=".xlsx"></div>
        <div class="drop" id="d2"><span class="opt">OBLIGATORIO</span><label>Reporte DIAN</label><small>Las compras a contabilizar (.xlsx)</small><input type="file" id="dian" accept=".xlsx"></div>
        <div class="drop" id="d3"><span class="opt">OPCIONAL</span><label>ZIP de XML</label><small>Para dividir facturas por concepto (.zip)</small><input type="file" id="xml" accept=".zip"></div>
        <div class="drop" id="d4"><span class="opt">OPCIONAL</span><label>Plan de cuentas (PUC)</label><small>Para mostrar el nombre de cada cuenta (.xlsx)</small><input type="file" id="puc" accept=".xlsx"></div>
      </div>
      <button class="btn" id="go" disabled>Procesar compras</button>
      <div class="err" id="err"></div>
    </div>

    <div class="loading" id="loading"><div class="spin"></div>Procesando y armando los asientos…</div>

    <div id="results">
      <div class="summary" id="summary"></div>
      <div class="dl">
        <a id="dlx" class="btn sec" href="#">Descargar Excel de revisión</a>
        <a id="dlt" class="btn" href="#">Descargar TXT para ContaI</a>
      </div>
      <div style="overflow-x:auto;border:1px solid var(--line);border-radius:14px">
        <table><thead><tr><th>Fecha</th><th>Factura</th><th>Proveedor</th><th>Total</th><th>Cuentas</th><th>Estado</th></tr></thead><tbody id="tb"></tbody></table>
      </div>
    </div>
  </div>

<script>
const $=id=>document.getElementById(id);
const files={auxiliar:'d1',dian:'d2',xml:'d3',puc:'d4'};
function check(){
  for(const k in files){const f=$(k);$(files[k]).classList.toggle('has',f.files.length>0);}
  $('go').disabled=!($('auxiliar').files.length && $('dian').files.length);
}
Object.keys(files).forEach(k=>$(k).addEventListener('change',check));

$('go').addEventListener('click',async()=>{
  $('err').style.display='none';$('form').style.display='none';$('loading').style.display='block';$('results').style.display='none';
  const fd=new FormData();
  ['auxiliar','dian','xml','puc'].forEach(k=>{if($(k).files.length)fd.append(k,$(k).files[0]);});
  try{
    const r=await fetch('/procesar',{method:'POST',body:fd});
    const d=await r.json();
    $('loading').style.display='none';
    if(d.error){$('form').style.display='block';$('err').textContent=d.error;$('err').style.display='block';return;}
    render(d);
  }catch(e){$('loading').style.display='none';$('form').style.display='block';$('err').textContent='Error de conexión: '+e;$('err').style.display='block';}
});

function render(d){
  const s=d.resumen;
  $('summary').innerHTML=`
    <div class="s"><div class="k">Facturas</div><div class="v">${s.facturas}</div></div>
    <div class="s"><div class="k">Total</div><div class="v">$${s.total.toLocaleString('es-CO',{maximumFractionDigits:0})}</div></div>
    <div class="s"><div class="k">Divididas por XML</div><div class="v">${s.divididas}</div></div>
    <div class="s"><div class="k">A revisar</div><div class="v">${s.nuevos+s.revisar}</div></div>
    <div class="s match"><div class="k">Descuadres</div><div class="v">${s.descuadres}</div></div>`;
  $('dlx').href='/descargar/excel/'+d.token;
  $('dlt').href='/descargar/txt/'+d.token;
  const tb=$('tb');tb.innerHTML='';
  d.asientos.forEach(a=>{
    const cls=a.estado.includes('NUEVO')?'nuevo':(a.estado.includes('Revisar')||a.estado.includes('sin XML'))?'rev':a.dividida?'div':'';
    const pill=a.estado==='OK'?'<span class="pill ok">OK</span>':a.dividida?'<span class="pill d">Dividida</span>':'<span class="pill w">'+a.estado+'</span>';
    const ctas=a.lineas.filter(l=>l.cuenta).map(l=>l.cuenta).join(', ');
    tb.innerHTML+=`<tr class="${cls}"><td>${a.fecha}</td><td>${a.factura}</td><td>${a.proveedor}</td><td class="num">${a.total.toLocaleString('es-CO',{maximumFractionDigits:0})}</td><td>${ctas}</td><td>${pill}</td></tr>`;
  });
  $('results').style.display='block';
}
</script>
</body>
</html>

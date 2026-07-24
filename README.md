# EXTERIOR-V2
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Lector de Operaciones — Capturas a Excel</title>
<script src="https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
  :root{
    --navy:#12324f; --navy-2:#1d4a72; --blue:#2273b8; --blue-soft:#eaf3fa;
    --bg:#f5f7fa; --card:#ffffff; --border:#dde4ec; --text:#22303e; --muted:#64748b;
    --green:#1e8e5a; --amber:#b45309; --red:#b91c1c; --amber-bg:#fef3c7;
  }
  *{box-sizing:border-box; margin:0; padding:0;}
  body{font-family:"Segoe UI", system-ui, -apple-system, Arial, sans-serif; background:var(--bg); color:var(--text); min-height:100vh;}
  header{background:var(--navy); color:#fff; padding:18px 28px; display:flex; align-items:center; gap:14px;}
  header svg{flex-shrink:0;}
  header h1{font-size:1.15rem; font-weight:600; letter-spacing:.2px;}
  header p{font-size:.8rem; color:#b8cbdd; margin-top:2px;}
  main{max-width:1280px; margin:0 auto; padding:24px 20px 60px;}
  .dropzone{
    background:var(--card); border:2px dashed var(--border); border-radius:12px;
    padding:34px 20px; text-align:center; cursor:pointer; transition:border-color .2s, background .2s;
  }
  .dropzone:hover, .dropzone.drag{border-color:var(--blue); background:var(--blue-soft);}
  .dropzone:focus-visible{outline:3px solid var(--blue); outline-offset:2px;}
  .dropzone .big{font-size:1rem; font-weight:600; color:var(--navy);}
  .dropzone .hint{font-size:.83rem; color:var(--muted); margin-top:6px;}
  .dropzone input{display:none;}
  #status{margin:14px 0 4px; font-size:.85rem; color:var(--muted); min-height:22px;}
  #status .ok{color:var(--green); font-weight:600;}
  #status .warn{color:var(--amber); font-weight:600;}
  #status .err{color:var(--red); font-weight:600;}
  .bar{height:6px; background:var(--border); border-radius:3px; overflow:hidden; margin-top:6px; display:none;}
  .bar>div{height:100%; width:0%; background:var(--blue); transition:width .2s;}
  .tablewrap{background:var(--card); border:1px solid var(--border); border-radius:12px; margin-top:18px; overflow-x:auto;}
  table{border-collapse:collapse; width:100%; font-size:.82rem; white-space:nowrap;}
  th{background:var(--navy-2); color:#fff; padding:9px 10px; text-align:left; font-weight:600; position:sticky; top:0;}
  td{padding:7px 10px; border-top:1px solid var(--border);}
  tbody tr:hover{background:var(--blue-soft);}
  td[contenteditable]{cursor:text; min-width:60px;}
  td[contenteditable]:focus{outline:2px solid var(--blue); outline-offset:-2px; background:#fff;}
  td.warnCell{background:var(--amber-bg);}
  td.num{text-align:right; font-variant-numeric:tabular-nums;}
  .rowdel{background:none; border:none; color:var(--muted); cursor:pointer; padding:4px; border-radius:6px; transition:color .15s, background .15s;}
  .rowdel:hover{color:var(--red); background:#fee2e2;}
  .empty{padding:26px; text-align:center; color:var(--muted); font-size:.85rem;}
  .actions{display:flex; gap:10px; margin-top:16px; flex-wrap:wrap;}
  button.btn{
    font:inherit; font-weight:600; font-size:.88rem; padding:11px 20px; border-radius:9px;
    border:1px solid transparent; cursor:pointer; transition:background .2s, box-shadow .2s; display:inline-flex; align-items:center; gap:8px;
  }
  .btn:focus-visible{outline:3px solid var(--blue); outline-offset:2px;}
  .btn-primary{background:var(--blue); color:#fff;}
  .btn-primary:hover{background:var(--navy-2); box-shadow:0 2px 8px rgba(18,50,79,.25);}
  .btn-secondary{background:#fff; color:var(--navy); border-color:var(--border);}
  .btn-secondary:hover{background:var(--blue-soft);}
  .btn:disabled{opacity:.45; cursor:not-allowed; box-shadow:none;}
  .note{font-size:.78rem; color:var(--muted); margin-top:14px; line-height:1.5;}
  @media (prefers-reduced-motion: reduce){ *{transition:none !important;} }
</style>
</head>
<body>
<header>
  <svg width="30" height="30" viewBox="0 0 24 24" fill="none" stroke="#7fb2dd" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
    <rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 9h18M9 9v11"/><path d="m13 13 2 2 3-3"/>
  </svg>
  <div>
    <h1>Lector de Operaciones</h1>
    <p>Capturas de MarketAxess / Bloomberg &rarr; Excel &middot; Cuenta 997 &middot; Comisi&oacute;n 0%</p>
  </div>
</header>

<main>
  <div class="dropzone" id="dropzone" tabindex="0" role="button" aria-label="Cargar capturas de operaciones">
    <div class="big">Arrastr&aacute; las capturas ac&aacute;, hac&eacute; clic para elegirlas, o peg&aacute;las con Ctrl+V</div>
    <div class="hint">JPG o PNG &middot; pod&eacute;s cargar varias a la vez &middot; cada foto agrega una fila a la tabla</div>
    <input type="file" id="fileinput" accept="image/*" multiple>
  </div>

  <div id="status">Cargando motor de OCR&hellip;</div>
  <div class="bar" id="bar"><div id="barfill"></div></div>

  <div class="tablewrap">
    <table id="tabla" aria-label="Operaciones extraídas">
      <thead><tr id="headrow"></tr></thead>
      <tbody id="tbody"></tbody>
    </table>
    <div class="empty" id="empty">Todav&iacute;a no hay operaciones cargadas.</div>
  </div>

  <div class="actions">
    <button class="btn btn-primary" id="btnExcel" disabled>
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="M7 10l5 5 5-5"/><path d="M12 15V3"/></svg>
      Descargar Excel
    </button>
    <button class="btn btn-secondary" id="btnCopiar" disabled>
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
      Copiar tabla (pegar en Excel abierto)
    </button>
    <button class="btn btn-secondary" id="btnLimpiar" disabled>Limpiar tabla</button>
  </div>

  <p class="note">
    Lee capturas de MarketAxess, Bloomberg VCON y ticket Bloomberg en espa&ntilde;ol (detecta el formato solo).
    Los valores se pueden corregir haciendo clic sobre cualquier celda antes de descargar.
    Las celdas resaltadas en amarillo no se pudieron leer con seguridad: revisalas contra la foto.
    El precio dirty se calcula como neto &divide; nominal &times; 100; el monto bruto y el neto son el neto de la pantalla (comisi&oacute;n 0%).
    Plazo: T+0 si liquida el mismo d&iacute;a, T+1 si es a plazo. Las fotos se procesan localmente en tu navegador: no se suben a ning&uacute;n servidor.
    La primera vez necesita internet para descargar el motor de OCR (~7 MB); despu&eacute;s queda en cach&eacute;.
  </p>
</main>

<script>
"use strict";
const COLUMNAS = ["Trade date","Lado","Ticker","Comision","Cuenta","Monto bruto","Monto neto",
                  "Precio dirty","Precio clean","Plazo","Valor nominal","Isin","Contraparte"];
const NUMERICAS = new Set(["Monto bruto","Monto neto","Precio dirty","Precio clean","Valor nominal"]);
const CUENTA = "997", COMISION = "0%";

// ------------------------- utilidades de parseo -------------------------
// Dos convenciones numéricas:
//  - "latam" (MarketAxess): punto = miles, coma = decimal  -> 105.673,89
//  - "us" (Bloomberg VCON / ticket): coma = miles, punto = decimal -> 1,000,110.00
function aNumeroLatam(s){
  if (s == null) return null;
  s = String(s).trim().replace(/\$/g,"").replace(/\s/g,"");
  if (!s) return null;
  s = s.replace(/\./g,"").replace(",",".");
  const v = parseFloat(s);
  return isNaN(v) ? null : v;
}
function aNumeroUS(s){
  if (s == null) return null;
  s = String(s).trim().replace(/\$/g,"").replace(/\s/g,"").replace(/,/g,"");
  if (!s) return null;
  const v = parseFloat(s);
  return isNaN(v) ? null : v;
}
// para releer celdas editadas de la tabla (formato con coma decimal)
function aNumero(s){ return aNumeroLatam(s); }

function buscar(re, texto){ const m = texto.match(re); return m ? m[1].trim() : null; }
function primero(re, textos){ for (const t of textos){ const v = buscar(re, t); if (v) return v; } return null; }

const NUM = "\\$?\\s*([\\d.,]+)";
const FECHA = "(\\d{1,2}\\/\\d{1,2}\\/\\d{4})";
const p2 = n => String(n).padStart(2,"0");

function aFecha(str, orden){  // orden: "dmy" o "mdy"
  if (!str) return null;
  const p = str.split("/").map(Number);
  const [d,m,y] = (orden === "mdy") ? [p[1],p[0],p[2]] : [p[0],p[1],p[2]];
  if (!d || !m || !y || d > 31 || m > 12) return null;
  return new Date(y, m-1, d);
}
const fechaUS = f => f ? `${p2(f.getMonth()+1)}/${p2(f.getDate())}/${f.getFullYear()}` : "";

// Validación de ISIN (algoritmo de Luhn) + autocorrección de errores típicos de OCR
function isinValido(isin){
  if (!/^[A-Z]{2}[A-Z0-9]{9}\d$/.test(isin)) return false;
  let d = "";
  for (const c of isin) d += /\d/.test(c) ? c : String(c.charCodeAt(0) - 55);
  let sum = 0, dbl = false;
  for (let i = d.length - 1; i >= 0; i--){
    let v = +d[i];
    if (dbl){ v *= 2; if (v > 9) v -= 9; }
    sum += v; dbl = !dbl;
  }
  return sum % 10 === 0;
}
const SUBST_OCR = { D:"0", O:"0", Q:"0", U:"0", I:"1", L:"1", Z:"2", S:"5", B:"8", G:"6", T:"7", A:"4" };
function corregirIsin(id){
  if (!id || isinValido(id)) return id;
  for (let i = 2; i < id.length; i++){
    const s = SUBST_OCR[id[i]];
    if (s){
      const cand = id.slice(0, i) + s + id.slice(i + 1);
      if (isinValido(cand)) return cand;
    }
  }
  return id;
}

function detectarLayout(todo){
  if (/YOU\s*BOUGHT|YOU\s*SOLD|Inquiry\s*ID|Trade\s*Proceeds/i.test(todo)) return "mktx";
  if (/YOU\s*BUY|YOU\s*SELL|VCON|Settlement\s*Date/i.test(todo)) return "vcon";
  if (/Compra\s*\/?\s*Venta|Precio\s*limpio|liquidac|Fecha\s*operativa|CUSIP/i.test(todo)) return "ticket";
  return "mktx";
}

// Armado común de la fila + validaciones compartidas
function armarFila(o){
  const { avisos, warnCols } = o;
  if (!o.lado){ avisos.push("No pude determinar compra/venta"); warnCols.add("Lado"); }
  if (!o.fT){ avisos.push("No pude leer el trade date"); warnCols.add("Trade date"); }

  let plazo = "";
  if (o.fT && o.fS) plazo = (o.fT.getTime() === o.fS.getTime()) ? "T+0" : "T+1";
  else { avisos.push("No pude leer la fecha de settlement"); warnCols.add("Plazo"); }

  // nominal: reconstrucción con principal / clean (corrige errores de OCR)
  let nominal = o.nominal;
  if (o.principal && o.clean){
    const calc = o.principal / o.clean * 100;
    if (!nominal || Math.abs(nominal - calc)/calc > 0.001){
      nominal = Math.round(calc*100)/100;
      if (Number.isInteger(nominal)) nominal = Math.trunc(nominal);
    }
  } else if (!nominal){
    avisos.push("No pude leer ni reconstruir el valor nominal"); warnCols.add("Valor nominal");
  }

  let proceeds = o.proceeds;
  if (proceeds == null && o.principal != null && o.accrued != null)
    proceeds = Math.round((o.principal + o.accrued)*100)/100;
  if (proceeds == null){ avisos.push("No pude leer el monto neto"); warnCols.add("Monto bruto"); warnCols.add("Monto neto"); }
  if (o.principal != null && o.accrued != null && proceeds != null &&
      Math.abs((o.principal + o.accrued) - proceeds) > 0.05){
    avisos.push("Principal + interés corrido no coincide con el neto: revisar montos");
    warnCols.add("Monto bruto"); warnCols.add("Monto neto");
  }
  // si principal no cierra contra nominal × clean, la captura puede estar recortada
  if (o.principal != null && nominal && o.clean){
    const calcP = nominal * o.clean / 100;
    if (Math.abs(calcP - o.principal) > 0.011){
      avisos.push("El principal no cierra contra nominal × precio (¿captura recortada?): revisar montos");
      warnCols.add("Monto bruto"); warnCols.add("Monto neto");
    }
  }
  const dirty = (proceeds && nominal) ? Math.round(proceeds/nominal*100*100000)/100000 : null;
  if (o.clean == null){ avisos.push("No pude leer el precio clean"); warnCols.add("Precio clean"); }
  if (dirty == null) warnCols.add("Precio dirty");
  if (!o.id){ avisos.push("No pude leer el ISIN/CUSIP"); warnCols.add("Isin"); }
  else if (o.id.length === 12 && !isinValido(o.id)){
    avisos.push("El ISIN no pasa el dígito verificador: revisar");
    warnCols.add("Isin");
  }
  if (!o.contraparte){ avisos.push("No pude leer la contraparte"); warnCols.add("Contraparte"); }

  return {
    fila: {
      "Trade date": fechaUS(o.fT), "Lado": o.lado || "", "Ticker": "", "Comision": COMISION, "Cuenta": CUENTA,
      "Monto bruto": proceeds, "Monto neto": proceeds, "Precio dirty": dirty, "Precio clean": o.clean,
      "Plazo": plazo, "Valor nominal": nominal, "Isin": o.id || "", "Contraparte": o.contraparte || ""
    },
    avisos, warnCols
  };
}

// --- Layout 1: MarketAxess (números latam, fechas DD/MM/YYYY) ---
function parseMktx(textos){
  const todo = textos.join("\n");
  const avisos = [], warnCols = new Set();
  let lado = "";
  if (/YOU\s*BOUGHT|\bBOT\b/i.test(todo)) lado = "Compra";
  else if (/YOU\s*SOLD|\bSLD\b/i.test(todo)) lado = "Venta";

  const fT = aFecha(primero(new RegExp("Trade\\s*Date[:\\s]*"+FECHA,"i"), textos), "dmy");
  const fS = aFecha(primero(new RegExp("St[il1]\\.?\\s*Date[:\\s]*(?:[A-Za-z]{3}\\s*)?"+FECHA,"i"), textos), "dmy");
  const principal = aNumeroLatam(primero(new RegExp("Princip\\w*[.:\\s]*"+NUM,"i"), textos));
  const accrued   = aNumeroLatam(primero(new RegExp("Accrued[.:\\s]*"+NUM,"i"), textos));
  const proceeds  = aNumeroLatam(primero(new RegExp("Proceeds[.:\\s]*"+NUM,"i"), textos));
  const clean     = aNumeroLatam(primero(/\bPrice[.:\s]*([\d.,]+)/i, textos));
  const nominal   = aNumeroLatam(primero(new RegExp("Size[.:\\s]*"+NUM,"i"), textos));
  const id        = corregirIsin(primero(/ISIN[:\s]*([A-Z]{2}[A-Z0-9]{9}\d)/i, textos));
  const contraparte = primero(/Dealer[.:\s]*([A-Z0-9]{2,})/, textos) || "";
  return armarFila({lado, fT, fS, principal, accrued, proceeds, clean, nominal, id, contraparte, avisos, warnCols});
}

// --- Layout 2: Bloomberg VCON (números US, fechas MM/DD/YYYY) ---
function parseVcon(textos){
  const todo = textos.join("\n");
  const avisos = [], warnCols = new Set();
  let lado = "";
  if (/YOU\s*BUY/i.test(todo)) lado = "Compra";
  else if (/YOU\s*SELL/i.test(todo)) lado = "Venta";

  // encabezado tipo "1.06MM @ 94.3500000"
  let nominal = null, clean = null;
  const mHead = todo.match(/([\d.,]+)\s*(MM|M|K|B)?\s*@\s*([\d.,]+)/i);
  if (mHead){
    const base = aNumeroUS(mHead[1]);
    const mult = { MM:1e6, M:1e3, K:1e3, B:1e9 }[(mHead[2]||"").toUpperCase()] || 1;
    if (base != null) nominal = base * mult;
    clean = aNumeroUS(mHead[3]);
  }
  const fT = aFecha(primero(new RegExp("Trade\\s*Date[^\\d]*"+FECHA,"i"), textos), "mdy");
  const fS = aFecha(primero(new RegExp("Settle\\w*\\s*Date[^\\d]*"+FECHA,"i"), textos), "mdy");
  const principal = aNumeroUS(primero(new RegExp("Princip\\w*[^\\d]*"+NUM,"i"), textos));
  const accrued   = aNumeroUS(primero(new RegExp("Accrued[^\\d]*"+NUM,"i"), textos));
  const proceeds  = aNumeroUS(primero(new RegExp("\\bNet\\b[^\\d]*"+NUM,"i"), textos));
  const id        = corregirIsin(primero(/ISIN[^A-Z0-9]*([A-Z]{2}[A-Z0-9]{9}\d)/i, textos));
  // contraparte: si compramos, es el Seller; si vendemos, el Buyer (lo que va después del @)
  const reCp = (lado === "Venta")
    ? /Buyer[^@\n]*@\s*([A-Z0-9 .&\-]{3,}?)(?:\s+LEI|\s*$)/im
    : /Seller[^@\n]*@\s*([A-Z0-9 .&\-]{3,}?)(?:\s+LEI|\s*$)/im;
  const contraparte = (primero(reCp, textos) || "").trim();
  return armarFila({lado, fT, fS, principal, accrued, proceeds, clean, nominal, id, contraparte, avisos, warnCols});
}

// --- Layout 3: Ticket Bloomberg en español (números US, fechas MM/DD/YYYY) ---
function parseTicket(textos){
  const todo = textos.join("\n");
  const avisos = [], warnCols = new Set();
  let lado = "";
  const mLado = todo.match(/(?:Compra\s*\/?\s*Venta|C\/V)[^A-Za-z]*(Venta|Compra)/i);
  if (mLado) lado = mLado[1][0].toUpperCase() + mLado[1].slice(1).toLowerCase();
  else if (/\bVenta\b/i.test(todo)) lado = "Venta";
  else if (/\bCompra\b/i.test(todo)) lado = "Compra";

  const fT = aFecha(primero(new RegExp("operativa[^\\d]*"+FECHA,"i"), textos), "mdy");
  const fS = aFecha(primero(new RegExp("l?iq[uü]?ida\\w*[^\\d]*"+FECHA,"i"), textos), "mdy");
  const nominal   = aNumeroUS(primero(new RegExp("Cantidad[^\\d]*"+NUM,"i"), textos));
  const clean     = aNumeroUS(primero(/Precio\s*limpio[^\d]*([\d.,]+)/i, textos));
  const principal = aNumeroUS(primero(new RegExp("Princip\\w*[^\\d]*"+NUM,"i"), textos));
  const proceeds  = aNumeroUS(primero(new RegExp("Neto[^\\d]*"+NUM,"i"), textos));
  const accrued   = (principal != null && proceeds != null)
                    ? Math.round((proceeds - principal)*100)/100 : null;
  let id = corregirIsin(primero(/\b([A-Z]{2}[A-Z0-9]{9}\d)\b/, textos)); // ISIN si aparece
  if (!id) id = primero(/CUSIP[^A-Z0-9]*([0-9A-Z]{8,9})\b/i, textos);    // si no, CUSIP
  let contraparte = (primero(/Nombre\s*de\s*B[^A-Z\n]*([A-Z][A-Z0-9 .&\-]{3,})/, textos) || "").trim();
  if (!contraparte) contraparte = (primero(/Br[óo0]?ker[.:\s]*([A-Z]{2,})/i, textos) || "").trim();
  return armarFila({lado, fT, fS, principal, accrued, proceeds, clean, nominal, id, contraparte, avisos, warnCols});
}

function extraerDatos(textos){
  const layout = detectarLayout(textos.join("\n"));
  const r = (layout === "vcon") ? parseVcon(textos)
          : (layout === "ticket") ? parseTicket(textos)
          : parseMktx(textos);
  r.layout = layout;
  return r;
}

// ------------------------- preprocesado de imagen -------------------------
function preprocesar(img){
  const escala = Math.max(1, Math.round(2200 / img.width));
  const w = img.width*escala, h = img.height*escala;

  const c1 = document.createElement("canvas"); c1.width=w; c1.height=h;
  const g1 = c1.getContext("2d");
  g1.imageSmoothingQuality = "high";
  g1.drawImage(img, 0, 0, w, h);
  const d1 = g1.getImageData(0,0,w,h), px = d1.data;
  let suma = 0;
  for (let i=0;i<px.length;i+=4){
    const gray = 0.299*px[i]+0.587*px[i+1]+0.114*px[i+2];
    px[i]=px[i+1]=px[i+2]=gray;
    suma += gray;
  }
  // pantallas oscuras (Bloomberg): invertir para que el texto quede oscuro sobre claro
  const oscura = (suma / (px.length/4)) < 110;
  let min=255, max=0;
  for (let i=0;i<px.length;i+=4){
    const v = oscura ? 255-px[i] : px[i];
    px[i]=px[i+1]=px[i+2]=v;
    if (v<min) min=v; if (v>max) max=v;
  }
  g1.putImageData(d1,0,0);                    // pasada 1: escala de grises (invertida si es oscura)

  const c2 = document.createElement("canvas"); c2.width=w; c2.height=h;
  const g2 = c2.getContext("2d");
  const d2 = g1.getImageData(0,0,w,h), q = d2.data;
  const rango = Math.max(1, max-min);
  for (let i=0;i<q.length;i+=4){
    const v = (q[i]-min)*255/rango;           // autocontraste
    const b = v > 140 ? 255 : 0;              // binarizado
    q[i]=q[i+1]=q[i+2]=b;
  }
  g2.putImageData(d2,0,0);                    // pasada 2: binarizada
  return [c1, c2];
}

// ------------------------- OCR (motor desde CDN, cacheado por el navegador) -------------------------
let workerPromise = null;
function getWorker(){
  if (!workerPromise){
    workerPromise = (async () => {
      const w = await Tesseract.createWorker("eng", 1, {
        workerPath: "https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/worker.min.js",
        corePath:   "https://cdn.jsdelivr.net/npm/tesseract.js-core@5",
        langPath:   "https://cdn.jsdelivr.net/npm/@tesseract.js-data/eng/4.0.0_best_int"
      });
      await w.setParameters({ tessedit_pageseg_mode: "6" });
      return w;
    })();
  }
  return workerPromise;
}

async function ocrImagen(archivo){
  const url = URL.createObjectURL(archivo);
  try{
    const img = await new Promise((res, rej) => {
      const im = new Image();
      im.onload = () => res(im);
      im.onerror = () => rej(new Error("No pude abrir la imagen"));
      im.src = url;
    });
    const canvases = preprocesar(img);
    const worker = await getWorker();
    const textos = [];
    for (const c of canvases){
      const { data } = await worker.recognize(c);
      textos.push(data.text || "");
    }
    return textos;
  } finally {
    URL.revokeObjectURL(url);
  }
}

// ------------------------- UI -------------------------
const $ = id => document.getElementById(id);
const tbody = $("tbody"), statusEl = $("status");

function setStatus(html){ statusEl.innerHTML = html; }
function refrescarBotones(){
  const hay = tbody.children.length > 0;
  $("btnExcel").disabled = !hay; $("btnCopiar").disabled = !hay; $("btnLimpiar").disabled = !hay;
  $("empty").style.display = hay ? "none" : "block";
}

// cabecera
{
  const hr = $("headrow");
  for (const c of COLUMNAS){ const th = document.createElement("th"); th.textContent = c; hr.appendChild(th); }
  const th = document.createElement("th"); th.textContent = ""; hr.appendChild(th);
}

function agregarFila(fila, warnCols){
  const tr = document.createElement("tr");
  for (const c of COLUMNAS){
    const td = document.createElement("td");
    td.contentEditable = "true";
    td.dataset.col = c;
    let v = fila[c];
    if (v == null) v = "";
    if (NUMERICAS.has(c)){ td.classList.add("num"); }
    td.textContent = (typeof v === "number") ? String(v).replace(".", ",") : String(v);
    if (warnCols && warnCols.has(c)){ td.classList.add("warnCell"); td.title = "Revisar contra la foto"; }
    td.addEventListener("input", () => td.classList.remove("warnCell"));
    tr.appendChild(td);
  }
  const tdDel = document.createElement("td");
  const b = document.createElement("button");
  b.className = "rowdel"; b.setAttribute("aria-label","Eliminar fila"); b.title = "Eliminar fila";
  b.innerHTML = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M18 6 6 18M6 6l12 12"/></svg>';
  b.addEventListener("click", () => { tr.remove(); refrescarBotones(); });
  tdDel.appendChild(b);
  tr.appendChild(tdDel);
  tbody.appendChild(tr);
  refrescarBotones();
}

function leerTabla(){
  const filas = [];
  for (const tr of tbody.children){
    const fila = {};
    for (const td of tr.querySelectorAll("td[data-col]")){
      const c = td.dataset.col;
      const txt = td.textContent.trim();
      fila[c] = NUMERICAS.has(c) ? (aNumero(txt) ?? txt) : txt;
    }
    filas.push(fila);
  }
  return filas;
}

// ------------------------- procesamiento -------------------------
let procesando = false;
async function procesarArchivos(archivos){
  const imgs = [...archivos].filter(f => /^image\//.test(f.type) || /\.(jpe?g|png)$/i.test(f.name));
  if (!imgs.length){ setStatus('<span class="err">No encontr&eacute; im&aacute;genes en lo que cargaste.</span>'); return; }
  if (procesando) return;
  procesando = true;
  $("bar").style.display = "block";
  let ok = 0, conAvisos = 0;
  for (let i = 0; i < imgs.length; i++){
    setStatus(`Leyendo <b>${imgs[i].name}</b> (${i+1} de ${imgs.length})&hellip;`);
    $("barfill").style.width = `${Math.round(i/imgs.length*100)}%`;
    try{
      const textos = await ocrImagen(imgs[i]);
      const { fila, avisos, warnCols } = extraerDatos(textos);
      agregarFila(fila, warnCols);
      ok++;
      if (avisos.length){ conAvisos++; console.warn(imgs[i].name, avisos); }
    } catch(e){
      console.error(e);
      setStatus(`<span class="err">Error con ${imgs[i].name}: ${e.message}</span>`);
    }
  }
  $("barfill").style.width = "100%";
  setTimeout(() => { $("bar").style.display = "none"; $("barfill").style.width = "0%"; }, 600);
  const extra = conAvisos ? ` &middot; <span class="warn">${conAvisos} con celdas a revisar (en amarillo)</span>` : "";
  setStatus(`<span class="ok">${ok} operaci&oacute;n(es) cargada(s)</span>${extra}`);
  procesando = false;
}

// ------------------------- export -------------------------
$("btnExcel").addEventListener("click", () => {
  const filas = leerTabla();
  const aoa = [COLUMNAS, ...filas.map(f => COLUMNAS.map(c => f[c] ?? ""))];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = COLUMNAS.map(c => ({wch: Math.max(c.length+2, 12)}));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Operaciones");
  XLSX.writeFile(wb, "operaciones.xlsx");
});

$("btnCopiar").addEventListener("click", async () => {
  const filas = leerTabla();
  const tsv = filas.map(f => COLUMNAS.map(c => {
    let v = f[c] ?? "";
    if (typeof v === "number") v = String(v).replace(".", ",");  // Excel regional AR
    return v;
  }).join("\t")).join("\n");
  try{
    await navigator.clipboard.writeText(tsv);
    setStatus('<span class="ok">Tabla copiada: peg&aacute; con Ctrl+V en tu Excel (sin encabezados).</span>');
  } catch {
    setStatus('<span class="err">El navegador bloque&oacute; el portapapeles: us&aacute; Descargar Excel.</span>');
  }
});

$("btnLimpiar").addEventListener("click", () => {
  if (confirm("¿Borrar todas las filas de la tabla?")){ tbody.innerHTML = ""; refrescarBotones(); setStatus(""); }
});

// ------------------------- entrada de archivos -------------------------
const dz = $("dropzone"), fi = $("fileinput");
dz.addEventListener("click", () => fi.click());
dz.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " "){ e.preventDefault(); fi.click(); } });
fi.addEventListener("change", () => { procesarArchivos(fi.files); fi.value = ""; });
["dragenter","dragover"].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add("drag"); }));
["dragleave","drop"].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove("drag"); }));
dz.addEventListener("drop", e => procesarArchivos(e.dataTransfer.files));
document.addEventListener("paste", e => {
  const items = [...(e.clipboardData?.items || [])].filter(i => i.type.startsWith("image/"));
  if (items.length) procesarArchivos(items.map(i => i.getAsFile()).filter(Boolean));
});

// precargar el motor de OCR
getWorker().then(
  () => setStatus("Motor de OCR listo. Carg&aacute; las capturas."),
  e  => setStatus('<span class="err">No pude descargar el motor de OCR. Verific&aacute; la conexi&oacute;n o si la red bloquea cdn.jsdelivr.net, y recarg&aacute; la p&aacute;gina.</span>')
);
refrescarBotones();
</script>
</body>
</html>

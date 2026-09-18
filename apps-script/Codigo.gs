/************************************************************
 * MAXIM · Control de Maquinado — Web App de datos en vivo
 * ----------------------------------------------------------
 * Lee las 3 hojas de Drive y las devuelve como JSON con la
 * MISMA forma que consume el aplicativo (index.html):
 *   { orders, reports, pending, operators, requesters, updated }
 *
 * DESPLIEGUE (una sola vez):
 *   1) script.google.com  →  Nuevo proyecto  →  pega este código.
 *   2) Guarda.  Ejecuta "buildData" una vez y autoriza los permisos.
 *   3) Implementar  →  Nueva implementación  →  tipo "Aplicación web".
 *        - Ejecutar como:  Yo (tu cuenta)
 *        - Quién tiene acceso:  Cualquier usuario
 *   4) Copia la URL que termina en /exec y pásala a Claude.
 *
 * Nota: "Cualquier usuario" permite que el sitio lea los datos sin
 * login. Cualquiera con esa URL (larga e imposible de adivinar) puede
 * leer estos datos. Es data interna de maquinado; si luego quieres
 * cerrarlo, se hace con el backend de escritura + token.
 ************************************************************/

// ---- IDs de las hojas (carpeta MAQUINADO) ----
var SHEET_OT   = '1XEBZtxP0pnKCe7AIQN-YXn7A_9PoO4phepygqcD7-VI'; // ORDEN DE TRABAJO (OT) (Respuestas)
var SHEET_RES  = '1IGBHc6lH_eQDuEBcrQWXpDxU6g-5sxuixVCAamDyXY8'; // RESULTADOS DE MAQUINADO (respuestas)
var SHEET_PEND = '16FWLB5LIklifCVyULEZWaB-k1o4yAOM-DzRv9HBM5Cg'; // PENDIENTES

function doGet(e) {
  var out = ContentService.createTextOutput(JSON.stringify(buildData()));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}

// ---------- utilidades ----------
function firstSheetValues_(id) {
  // getDisplayValues -> todo como texto tal cual se ve (fechas dd/mm/aaaa, duraciones 0:00:00)
  var sh = SpreadsheetApp.openById(id).getSheets()[0];
  return sh.getDataRange().getDisplayValues();
}
function c_(v) { return (v == null ? '' : String(v)).trim(); }
function d_(v) { // "17/9/2026" o "17/9/2026 11:47:17" -> "2026-09-17"
  var s = c_(v);
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (!m) return '';
  var dd = ('0' + m[1]).slice(-2), mm = ('0' + m[2]).slice(-2);
  return m[3] + '-' + mm + '-' + dd;
}
function norm_(s) {
  s = (s == null ? '' : String(s));
  s = s.normalize('NFD').replace(/[̀-ͯ]/g, '');
  return s.replace(/\s+/g, ' ').trim().toUpperCase();
}
function code_(s) {
  var m = String(s || '').toUpperCase().match(/\b(PR|NP|RE)-?\s*(\d+)/);
  return m ? m[1] + '-' + m[2] : '';
}
function idxOf_(header, name) {
  for (var i = 0; i < header.length; i++) if (c_(header[i]) === name) return i;
  return -1;
}

// ---------- construcción de datos ----------
function buildData() {
  var orders = buildOrders_();
  var reports = buildReports_();
  return {
    orders: orders,
    reports: reports,
    pending: buildPending_(),
    operators: uniqNames_(reports.map(function (r) { return r.op; })),
    requesters: uniqNames_(orders.map(function (o) { return o.sol; })),
    updated: new Date().toISOString()
  };
}

function uniqNames_(vals) {
  var seen = {}, keys = [];
  for (var i = 0; i < vals.length; i++) {
    var v = c_(vals[i]); if (!v) continue;
    var k = norm_(v);
    if (k && !seen[k]) { seen[k] = v; keys.push(k); }
  }
  keys.sort();
  return keys.map(function (k) { return seen[k]; });
}

function buildOrders_() {
  var v = firstSheetValues_(SHEET_OT);
  var h = v[0], out = [];
  // pares Cantidad + Nombre de la parte
  var pairs = [];
  for (var i = 0; i < h.length - 1; i++) {
    if (/^Cantidad/i.test(c_(h[i])) && /^Nombre de la parte/i.test(c_(h[i + 1]))) pairs.push(i);
  }
  var iEnt = -1;
  for (var k = 0; k < h.length; k++) if (/espera la entrega/i.test(c_(h[k]))) iEnt = k;
  for (var r = 1; r < v.length; r++) {
    var row = v[r];
    var fecha = d_(row[0]);
    var herr = c_(row[5]);
    if (!fecha || herr.length < 3) continue;         // descarta filas basura / pie
    var partes = [];
    for (var p = 0; p < pairs.length; p++) {
      var q = c_(row[pairs[p]]), n = c_(row[pairs[p] + 1]);
      if (n) partes.push({ q: q, n: n });
    }
    out.push({
      fecha: fecha, correo: c_(row[1]), sol: c_(row[2]), cc: c_(row[3]),
      tipo: c_(row[4]), herr: herr, cod: c_(row[6]), trab: c_(row[7]),
      ind: c_(row[8]), partes: partes, entrega: iEnt >= 0 ? d_(row[iEnt]) : ''
    });
  }
  return out;
}

function buildReports_() {
  var v = firstSheetValues_(SHEET_RES);
  var h = v[0], out = [];
  var partStarts = [], insStarts = [];
  for (var i = 0; i < h.length; i++) {
    var hi = c_(h[i]);
    if (hi.indexOf('Seleccione la parte de la herramienta a registrar') === 0) partStarts.push(i);
    if (hi.indexOf('Nombre de inserto') === 0) insStarts.push(i);
  }
  var iCons = 1, iOp = idxOf_(h, 'Nombre de operador'), iFec = idxOf_(h, 'Fecha de trabajo');
  var iDesc = idxOf_(h, 'Descripción del trabajo realizado');
  var iEst = idxOf_(h, 'Estado'), iIni = idxOf_(h, 'Hora de inicio del trabajo'), iFin = idxOf_(h, 'Hora fin del trabajo');
  for (var r = 1; r < v.length; r++) {
    var row = v[r];
    if (!code_(row[iCons])) continue;
    var partes = [];
    for (var s = 0; s < partStarts.length; s++) {
      var b = partStarts[s];
      if (!c_(row[b])) continue;
      partes.push({
        p: c_(row[b]), diam: c_(row[b + 1]), long: c_(row[b + 2]), heat: c_(row[b + 3]), ref: c_(row[b + 4]),
        torno: c_(row[b + 5]), tTorno: c_(row[b + 6]), perf: c_(row[b + 7]), mand: c_(row[b + 8]), acab: c_(row[b + 9]),
        pin: c_(row[b + 10]), box: c_(row[b + 11]), fresa: c_(row[b + 12]), tipoF: c_(row[b + 13]), tFresa: c_(row[b + 14])
      });
    }
    var insertos = [];
    for (var j = 0; j < insStarts.length; j++) {
      var ib = insStarts[j];
      if (c_(row[ib])) insertos.push({ n: c_(row[ib]), t: c_(row[ib + 1]) });
    }
    out.push({
      cons: c_(row[iCons]), op: iOp >= 0 ? c_(row[iOp]) : '', fecha: iFec >= 0 ? d_(row[iFec]) : '',
      desc: iDesc >= 0 ? c_(row[iDesc]) : '', partes: partes, insertos: insertos,
      ini: iIni >= 0 ? c_(row[iIni]) : '', fin: iFin >= 0 ? c_(row[iFin]) : '', estado: iEst >= 0 ? c_(row[iEst]) : ''
    });
  }
  return out;
}

function buildPending_() {
  var v = firstSheetValues_(SHEET_PEND), out = [];
  for (var r = 1; r < v.length; r++) {
    var row = v[r];
    if (!code_(row[0])) continue;
    out.push({ cons: c_(row[0]), ind: c_(row[1]), sol: c_(row[2]), fecha: d_(row[3]), entrega: d_(row[4]) });
  }
  return out;
}

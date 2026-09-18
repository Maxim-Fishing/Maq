/************************************************************
 * MAXIM · Maquinado — Web App de datos EN VIVO  (v2)
 * ----------------------------------------------------------
 * Mejora sobre v1: en vez de leer solo las respuestas crudas,
 * lee los REGISTROS REALES del sistema:
 *   - orders      ← "OT ACEPTADAS" + "OT TERMINADAS"  (con consecutivo PR/NP + estado Abierta/Cerrada)
 *   - solicitudes ← "SOLICITUDES OT"                   (con Estado Pendiente/Aprobada/Rechazada + UUID)
 *   - reports     ← "RESULTADOS OT"                    (con Estado + UUID)
 *   - pending     ← OT abiertas SIN reporte (se calcula)
 *   - operators / requesters
 *
 * Devuelve JSON con esta forma; el aplicativo lo consume igual que v1.
 *
 * DESPLIEGUE: pega este código en un proyecto de Apps Script,
 * ejecuta buildData() una vez para autorizar, y publica como
 * Aplicación web (Ejecutar como: tú · Acceso: cualquiera).
 * Copia la URL /exec y pásasela a Claude para fijarla en el app.
 ************************************************************/

// Spreadsheet A (solicitudes / OT) = mismo libro de respuestas del form OT
var SS_OT  = '1XEBZtxP0pnKCe7AIQN-YXn7A_9PoO4phepygqcD7-VI';
// Spreadsheet B (resultados)
var SS_RES = '1IGBHc6lH_eQDuEBcrQWXpDxU6g-5sxuixVCAamDyXY8';
// Spreadsheet C (BALANCE — costos), tab GENERAL (histórico detallado con costos)
var SS_BALANCE = '1Mf6jKf5TkdofcD2TkaiQBGH_zoLOeIant8_Q2Op3BOQ';
var TAB_GENERAL = 'GENERAL';

var TAB_SOLICITUDES = 'SOLICITUDES OT';
var TAB_ACEPTADAS   = 'OT ACEPTADAS';
var TAB_TERMINADAS  = 'OT TERMINADAS';
var TAB_RESULTADOS  = 'RESULTADOS OT';

function doGet(e) {
  var out = ContentService.createTextOutput(JSON.stringify(buildData()));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}

// ---------- utilidades ----------
function displayValues_(ssId, tabName) {
  var ss = SpreadsheetApp.openById(ssId);
  var sh = tabName ? ss.getSheetByName(tabName) : ss.getSheets()[0];
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getDataRange().getDisplayValues();
}
function c_(v) { return (v == null ? '' : String(v)).trim(); }
function d_(v) {
  var s = c_(v), m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (!m) return '';
  return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
}
function norm_(s) {
  s = (s == null ? '' : String(s)).normalize('NFD').replace(/[̀-ͯ]/g, '');
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
function uniqNames_(vals) {
  var seen = {}, keys = [];
  for (var i = 0; i < vals.length; i++) {
    var v = c_(vals[i]); if (!v) continue;
    var k = norm_(v); if (k && !seen[k]) { seen[k] = v; keys.push(k); }
  }
  keys.sort();
  return keys.map(function (k) { return seen[k]; });
}

// ---------- construcción ----------
function buildData() {
  var orders = buildOrders_();
  var solicitudes = buildSolicitudes_();
  var reports = buildReports_();

  // pending = OT abiertas que aún no tienen reporte
  var reportadas = {};
  reports.forEach(function (r) { var k = code_(r.cons); if (k) reportadas[k] = true; });
  var pending = orders
    .filter(function (o) { return o.estado === 'Abierta' && !reportadas[code_(o.cons)]; })
    .map(function (o) { return { cons: o.cons, ind: o.ind, sol: o.sol, fecha: o.fecha, entrega: o.entrega }; });

  return {
    orders: orders,
    solicitudes: solicitudes,
    reports: reports,
    pending: pending,
    costsByMonth: buildCosts_(),
    operators: uniqNames_(reports.map(function (r) { return r.op; })),
    requesters: uniqNames_(solicitudes.map(function (s) { return s.sol; }).concat(orders.map(function (o) { return o.sol; }))),
    updated: new Date().toISOString()
  };
}

// Costos por mes ← hoja GENERAL del libro BALANCE (histórico detallado por parte).
// Columnas GENERAL (0-based): 4 consecutivo, 7 fechaReporte, 10 fechaTrabajo, 11 estado,
// 28 costoMaterial, 29 costoTorno, 30 costoFresa, 31 costoManoObra, 33 costoArriendo, 34 costoTotal.
function money_(s) {
  var n = parseFloat(String(s == null ? '' : s).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}
function mesDe_(s) {
  var m = String(s == null ? '' : s).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? m[3] + '-' + ('0' + m[2]).slice(-2) : '';
}
function buildCosts_() {
  var v;
  try { v = displayValues_(SS_BALANCE, TAB_GENERAL); }
  catch (e) { return []; }
  if (!v.length) return [];
  var acc = {};
  for (var r = 1; r < v.length; r++) {
    var row = v[r];
    if (row.length < 35) continue;
    var mes = mesDe_(row[10]) || mesDe_(row[7]);
    if (!mes) continue;
    if (!acc[mes]) acc[mes] = { mes: mes, material: 0, torno: 0, fresa: 0, mo: 0, arriendo: 0, total: 0, partes: 0, ots: {} };
    var a = acc[mes];
    a.material += money_(row[28]); a.torno += money_(row[29]); a.fresa += money_(row[30]);
    a.mo += money_(row[31]); a.arriendo += money_(row[33]); a.total += money_(row[34]);
    a.partes++; if (row[4]) a.ots[String(row[4]).trim()] = true;
  }
  return Object.keys(acc).sort().map(function (k) {
    var a = acc[k];
    return {
      mes: a.mes,
      material: Math.round(a.material), torno: Math.round(a.torno), fresa: Math.round(a.fresa),
      mo: Math.round(a.mo), arriendo: Math.round(a.arriendo), total: Math.round(a.total),
      partes: a.partes, ots: Object.keys(a.ots).length
    };
  });
}

// OT ACEPTADAS + OT TERMINADAS  (registro real de órdenes con consecutivo y estado)
// Layout (0-based): A0 fechaPet, B1 fechaEntrega, C2 correo, D3 solicitante,
// E4 centroCostos, F5 programacion, G6 consecutivo, H7 estado, I8 herramienta,
// J9 codigo, L11 tipoOT, M12 indicaciones, partes desde N13 (nombre,cantidad).
function leerOT_(tab, estadoPorDefecto) {
  var v = displayValues_(SS_OT, tab), out = [];
  for (var r = 1; r < v.length; r++) {
    var row = v[r];
    var cons = c_(row[6]);
    if (!code_(cons)) continue;
    var partes = [];
    for (var k = 13; k + 1 < row.length && k < 41; k += 2) {
      var nombre = c_(row[k]), cant = c_(row[k + 1]);
      if (nombre) partes.push({ q: cant, n: nombre });
    }
    out.push({
      cons: cons,
      estado: c_(row[7]) || estadoPorDefecto,
      prog: c_(row[5]),
      fecha: d_(row[0]),
      entrega: d_(row[1]),
      correo: c_(row[2]),
      sol: c_(row[3]),
      cc: c_(row[4]),
      herr: c_(row[8]),
      cod: c_(row[9]),
      tipo: c_(row[11]),
      ind: c_(row[12]),
      partes: partes
    });
  }
  return out;
}
function buildOrders_() {
  return leerOT_(TAB_ACEPTADAS, 'Abierta').concat(leerOT_(TAB_TERMINADAS, 'Cerrada'));
}

// SOLICITUDES OT  (respuestas del form OT, con Estado + UUID de control)
// 0 marca,1 correo,2 nombre,3 centroCostos,4 tipoOT,5 herramienta,6 codigo,
// 7 tipoTrabajo,8 indicaciones, pares cant/nombre 10/11,13/14,..., 57 fechaEntrega,
// 58 uuid (col 59), 59 estado (col 60).
function buildSolicitudes_() {
  var v = displayValues_(SS_OT, TAB_SOLICITUDES), out = [];
  for (var r = 1; r < v.length; r++) {
    var row = v[r];
    if (row.length < 58) row = row.concat(new Array(58 - row.length).join('.').split('.'));
    var fecha = d_(row[0]);
    if (!fecha || c_(row[5]).length < 3) continue;
    var partes = [];
    for (var k = 10; k + 1 < 57; k += 3) {
      if (c_(row[k + 1])) partes.push({ q: c_(row[k]), n: c_(row[k + 1]) });
    }
    out.push({
      fecha: fecha, correo: c_(row[1]), sol: c_(row[2]), cc: c_(row[3]),
      tipo: c_(row[4]), herr: c_(row[5]), cod: c_(row[6]), trab: c_(row[7]),
      ind: c_(row[8]), entrega: d_(row[57]), partes: partes,
      uuid: c_(row[58]), estado: c_(row[59]) || 'Pendiente'
    });
  }
  return out;
}

// RESULTADOS OT  (reportes de maquinado con Estado + UUID)
function buildReports_() {
  var v = displayValues_(SS_RES, TAB_RESULTADOS);
  if (!v.length) return [];
  var h = v[0], out = [];
  var partStarts = [], insStarts = [];
  for (var i = 0; i < h.length; i++) {
    var hi = c_(h[i]);
    if (hi.indexOf('Seleccione la parte de la herramienta a registrar') === 0) partStarts.push(i);
    if (hi.indexOf('Nombre de inserto') === 0) insStarts.push(i);
  }
  var iOp = idxOf_(h, 'Nombre de operador'), iFec = idxOf_(h, 'Fecha de trabajo');
  var iDesc = idxOf_(h, 'Descripción del trabajo realizado');
  var iEst = idxOf_(h, 'Estado'), iUuid = idxOf_(h, 'UUID');
  if (iUuid === -1) iUuid = 152; // col 153 (1-based) por defecto
  var iIni = idxOf_(h, 'Hora de inicio del trabajo'), iFin = idxOf_(h, 'Hora fin del trabajo');
  for (var r = 1; r < v.length; r++) {
    var row = v[r];
    if (!code_(row[1])) continue;
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
      cons: c_(row[1]), op: iOp >= 0 ? c_(row[iOp]) : '', fecha: iFec >= 0 ? d_(row[iFec]) : '',
      desc: iDesc >= 0 ? c_(row[iDesc]) : '', partes: partes, insertos: insertos,
      ini: iIni >= 0 ? c_(row[iIni]) : '', fin: iFin >= 0 ? c_(row[iFin]) : '',
      estado: iEst >= 0 ? c_(row[iEst]) : '', uuid: iUuid >= 0 ? c_(row[iUuid]) : ''
    });
  }
  return out;
}

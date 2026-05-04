/**
 * Endpoint para receber OP/Cliente via POST e atualizar a aba op_cliente.
 */
function doPost(e) {
  try {
    var payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    var token = String(payload.token || '');
    var expectedToken = 'true'; // <- altere

    if (token !== expectedToken) {
      return jsonResponse_({ ok: false, error: 'unauthorized' });
    }

    var rows = Array.isArray(payload.rows) ? payload.rows : [];

    var ss = SpreadsheetApp.openById('1Rr8TKB1ypHemCrNzA6EvXsEU-aWjudH0DBf7kmddLf0');
    var sh = ss.getSheetByName('op_cliente');
    if (!sh) {
      throw new Error('Aba "op_cliente" não encontrada.');
    }

    // limpa e escreve cabeçalho
    sh.clearContents();
    sh.getRange(1, 1, 1, 3).setValues([['op', 'cliente', 'atualizado_em']]);

    if (rows.length > 0) {
      var now = new Date();
      var values = rows.map(function (r) {
        return [String(r.op || '').trim(), String(r.cliente || '').trim(), now];
      });
      sh.getRange(2, 1, values.length, 3).setValues(values);
    }

    return jsonResponse_({
      ok: true,
      inserted: rows.length,
      timestamp: new Date()
    });
  } catch (err) {
    return jsonResponse_({
      ok: false,
      error: String(err && err.message ? err.message : err),
      timestamp: new Date()
    });
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
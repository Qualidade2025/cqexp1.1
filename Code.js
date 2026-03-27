/**
 * Renderiza UI web do MVP.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('CQ Expedição - MVP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Retorna catálogos para preencher dropdowns no front-end.
 */
function getCatalogs() {
  return {
    colaboradores: getActiveCollaborators_(),
    posicoes: getActiveCatalogValues_(SHEETS.CAD_POSICOES, 1, 2),
    defeitos: getActiveCatalogValues_(SHEETS.CAD_DEFEITOS, 1, 2),
    origens: getActiveCatalogValues_(SHEETS.CAD_ORIGENS, 1, 2)
  };
}

/**
 * Grava inspeção em lote nas 3 tabelas com lock para concorrência.
 */
function saveInspection(payload) {
  validateInspectionPayload_(payload);

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    var idInspecao = Utilities.getUuid();
    var serverNow = new Date();
    var client = payload.clienteManual || getClientByOP(payload.op) || '';
    var userEmail = Session.getActiveUser().getEmail() || '';

    appendRowsBatch_(SHEETS.INSPECOES, [[
      idInspecao,
      serverNow,
      String(payload.op).trim(),
      Number(payload.qtddRevisada),
      String(payload.origem).trim(),
      String(client).trim(),
      userEmail
    ]]);

    var operatorRows = payload.operadores.map(function (op) {
      return [idInspecao, String(op.id).trim(), String(op.name).trim()];
    });
    appendRowsBatch_(SHEETS.INSPECAO_OPERADORES, operatorRows);

    var defectRows = payload.defeitos.map(function (item, index) {
      return [
        idInspecao,
        index + 1,
        String(item.posicao).trim(),
        String(item.defeito).trim(),
        Number(item.quantidade)
      ];
    });
    appendRowsBatch_(SHEETS.INSPECAO_DEFEITOS, defectRows);

    return {
      ok: true,
      idInspecao: idInspecao,
      serverNow: serverNow
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Stub para integração futura SAP B1 por OP.
 */
function getClientByOP(op) {
  // TODO: integrar com SAP B1 para buscar cliente automaticamente por OP.
  return '';
}

function getServerNow() {
  return new Date();
}

function healthcheck() {
  return {
    ok: true,
    timestamp: new Date(),
    spreadsheetId: SpreadsheetApp.getActiveSpreadsheet().getId()
  };
}
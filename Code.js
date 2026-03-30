/**
 * Renderiza UI web do MVP.
 */
function doGet(e) {
  var page = '';
  try {
    page = String((e && e.parameter && e.parameter.page) || '').trim();
  } catch (error) {
    page = '';
  }

  var templateName = getTemplateNameByPage_(page);
  var template = HtmlService.createTemplateFromFile(templateName);
  template.appBaseUrl = getAppBaseUrl_();

  return template.evaluate()
    .setTitle('Controle de Qualidade - BALDI')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAppBaseUrl_() {
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (error) {
    return '';
  }
}

function getTemplateNameByPage_(page) {
  if (page === 'index' || page === 'home' || page === 'inicio') {
    return 'Index';
  }

  if (page === 'listar-rncs' || page === 'listarcontroles' || page === 'listar-controles') {
    return 'ListarControles';
  }

  if (page === 'editar-colaboradores') {
    return 'EditarColaboradores';
  }

  return 'Index';
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Retorna catálogos para preencher dropdowns no front-end.
 */
function getCatalogs() {
  var defectCatalog = getDefectsByPositionCatalog_();

  return {
    colaboradores: getActiveCollaborators_(),
    posicoes: defectCatalog.posicoes,
    defeitos: defectCatalog.defeitos,
    defeitosPorPosicao: defectCatalog.defeitosPorPosicao,
    origens: getActiveCatalogValues_(SHEETS.CAD_ORIGENS, 1, 2),
    origemObrigatoria: isOriginRequired_()
  };
}

/**
 * Grava inspeção em linha única na aba inspecoes com lock para concorrência.
 */
function saveInspection(payload) {
  validateInspectionPayload_(payload);

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    var idInspecao = getNextInspectionId_();
    var serverNow = new Date();
    var client = payload.clienteManual || getClientByOP(payload.op) || '';
    var userEmail = Session.getActiveUser().getEmail() || '';
    var operators = normalizeOperators_(payload.operadores);
    var defects = normalizeDefects_(normalizePayloadDefects_(payload.defeitos).items);

    appendRowsBatch_(SHEETS.INSPECOES, [[
      idInspecao,
      serverNow,
      String(payload.op).trim(),
      Number(payload.qtddRevisada),
      payload.origem ? String(payload.origem).trim() : '',
      String(client).trim(),
      userEmail,
      operators[0] || '',
      operators[1] || '',
      operators[2] || '',
      operators[3] || '',
      defects.length,
      JSON.stringify(defects),
      defects.map(function (item) {
        return item.posicao + ' / ' + item.defeito + ' / ' + item.quantidade;
      }).join(' | '),
      payload.retrabalho ? 'X' : ''
    ]]);

    return {
      ok: true,
      idInspecao: idInspecao,
      serverNow: serverNow
    };
  } finally {
    lock.releaseLock();
  }
}

function getNextInspectionId_() {
  var sheet = getRequiredSheet_(SHEETS.INSPECOES);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 1;
  }

  var lastValue = sheet.getRange(lastRow, 1).getValue();
  var parsedValue = Number(lastValue);

  if (!isFinite(parsedValue) || parsedValue < 1) {
    return 1;
  }

  return Math.floor(parsedValue) + 1;
}

function normalizeOperators_(operators) {
  return operators.map(function (op) {
    return [String(op.id).trim(), String(op.name).trim()].join(' - ');
  });
}

function normalizeDefects_(defects) {
  return defects.map(function (item, index) {
    return {
      linha: index + 1,
      posicao: String(item.posicao).trim(),
      defeito: String(item.defeito).trim(),
      quantidade: Number(item.quantidade)
    };
  });
}

/**
 * Stub para integração futura SAP B1 por OP.
 */
function getClientByOP(op) {
  // TODO: integrar com SAP B1 para buscar cliente automaticamente por OP.
  return '';
}


function healthcheck() {
  return {
    ok: true,
    timestamp: new Date(),
    spreadsheetId: SpreadsheetApp.getActiveSpreadsheet().getId()
  };
}

function validateCollaboratorsPassword(password) {
  var enteredPassword = String(password || '').trim();
  if (!enteredPassword) {
    throw new Error('Informe a senha.');
  }

  var expectedPassword = getCollaboratorsEditPassword_();
  if (!expectedPassword) {
    throw new Error('Senha de edição não configurada em colaboradores!E2.');
  }

  if (enteredPassword !== expectedPassword) {
    throw new Error('Senha inválida.');
  }

  return {
    ok: true
  };
}

function listCollaboratorsControls() {
  return getCollaboratorsForControlList_();
}

function saveCollaboratorsControls(rows) {
  if (!Array.isArray(rows) || !rows.length) {
    throw new Error('Adicione ao menos um colaborador.');
  }

  var normalized = [];
  var ids = {};

  rows.forEach(function (row, index) {
    var id = row && row.id ? String(row.id).trim() : '';
    var name = row && row.name ? String(row.name).trim() : '';
    var active = !!(row && row.active);

    if (!id) {
      throw new Error('Linha ' + (index + 1) + ': informe o ID.');
    }

    if (!name) {
      throw new Error('Linha ' + (index + 1) + ': informe o nome.');
    }

    if (ids[id]) {
      throw new Error('ID duplicado: ' + id + '.');
    }
    ids[id] = true;

    normalized.push([id, name, active]);
  });

  saveCollaboratorsControlList_(normalized);
  return { ok: true };
}

function listInspectionsControls(params) {
  return listInspectionsControls_(params);
}

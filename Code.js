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
 * Configuração temporária de conexão JDBC para SQL Server.
 * IMPORTANTE: manter apenas provisoriamente; migrar para armazenamento seguro.
 */
var DB_URL = 'jdbc:sqlserver://SEU_HOST:1433;databaseName=SUA_BASE;encrypt=true;trustServerCertificate=true';
var DB_USER = 'SEU_USUARIO';
var DB_PASS = 'SUA_SENHA';

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
 * Busca cliente por OP diretamente no SQL Server.
 * @param {number|string} opNumero
 * @return {{op: string, cliente: string}|null}
 */
function buscarClientePorOP(opNumero) {
  var op = String(opNumero || '').trim();
  if (!op) {
    throw new Error('Informe um número de OP válido.');
  }

  var conn = null;
  var stmt = null;
  var rs = null;

  var etapa = 'validacao';

  try {
    etapa = 'conexao_jdbc';
    conn = Jdbc.getConnection(DB_URL, DB_USER, DB_PASS);

    etapa = 'preparo_sql';
    stmt = conn.prepareStatement('SELECT TOP 1 belnr_id, kndname FROM beas_fthapt WHERE belnr_id = ?');
    stmt.setString(1, op);

    etapa = 'execucao_consulta';
    rs = stmt.executeQuery();

    etapa = 'leitura_resultado';
    if (rs.next()) {
      var result = {
        op: String(rs.getString('belnr_id') || ''),
        cliente: String(rs.getString('kndname') || '')
      };
      Logger.log('[buscarClientePorOP] OP %s encontrada na etapa %s.', op, etapa);
      return result;
    }

    Logger.log('[buscarClientePorOP] OP %s não encontrada.', op);
    return null;
  } catch (error) {
    Logger.log('[buscarClientePorOP] Falha na etapa %s para OP %s. Detalhes: %s', etapa, op, error && error.message ? error.message : error);
    throw new Error('Erro ao consultar cliente por OP na etapa "' + etapa + '": ' + (error && error.message ? error.message : error));
  } finally {
    if (rs) {
      try {
        rs.close();
      } catch (closeRsError) {
        Logger.log('[buscarClientePorOP] Falha ao fechar ResultSet: %s', closeRsError && closeRsError.message ? closeRsError.message : closeRsError);
      }
    }

    if (stmt) {
      try {
        stmt.close();
      } catch (closeStmtError) {
        Logger.log('[buscarClientePorOP] Falha ao fechar Statement: %s', closeStmtError && closeStmtError.message ? closeStmtError.message : closeStmtError);
      }
    }

    if (conn) {
      try {
        conn.close();
      } catch (closeConnError) {
        Logger.log('[buscarClientePorOP] Falha ao fechar Connection: %s', closeConnError && closeConnError.message ? closeConnError.message : closeConnError);
      }
    }
  }
}

/**
 * Mantém compatibilidade com o fluxo atual de gravação.
 */
function getClientByOP(op) {
  var found = buscarClientePorOP(op);
  return found ? found.cliente : '';
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

function updateInspectionControl(payload) {
  return updateInspectionControl_(payload);
}

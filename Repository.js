var SHEETS = {
  INSPECOES: 'inspecoes',
  COLABORADORES: 'colaboradores',
  CAD_DEFEITOS: 'cad_defeitos',
  CAD_ORIGENS: 'cad_origens',
  OP_CLIENTE: 'op_cliente',
  VIEW_RELACOES_ATIVAS: 'view_relacoes_ativas',
  VIEW_POSICOES_ATIVAS: 'view_posicoes_ativas',
  VIEW_DEFEITOS_ATIVOS: 'view_defeitos_ativos'
};

/**
 * Lê catálogo com filtro de ativos e retorno textual.
 */
function getActiveCatalogValues_(sheetName, valueColumnIndex, activeColumnIndex) {
  var sheet = getRequiredSheet_(sheetName);
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  var values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  return values
    .filter(function (row) {
      return isActiveFlag_(row[activeColumnIndex - 1]);
    })
    .map(function (row) {
      return String(row[valueColumnIndex - 1]).trim();
    })
    .filter(function (value) {
      return value.length > 0;
    });
}

/**
 * Retorna colaboradores ativos para seleção de equipe.
 */
function getActiveCollaborators_() {
  var sheet = getRequiredSheet_(SHEETS.COLABORADORES);
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

  return values
    .filter(function (row) {
      return isActiveFlag_(row[2]);
    })
    .map(function (row) {
      return {
        id: String(row[0]).trim(),
        name: String(row[1]).trim()
      };
    })
    .filter(function (c) {
      return c.id && c.name;
    });
}

function getCollaboratorsForControlList_() {
  var sheet = getRequiredSheet_(SHEETS.COLABORADORES);
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

  return values
    .map(function (row) {
      return {
        id: String(row[0] || '').trim(),
        name: String(row[1] || '').trim(),
        active: isActiveFlag_(row[2]),
        isEditing: false,
        isNew: false
      };
    })
    .filter(function (row) {
      return row.id || row.name;
    });
}

function saveCollaboratorsControlList_(rows) {
  var sheet = getRequiredSheet_(SHEETS.COLABORADORES);
  var lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  }

  if (!rows.length) {
    return;
  }

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
}

/**
 * Lê matriz de defeitos x posições da aba cad_defeitos.
 * Formato esperado:
 * - Linha 1 (a partir da coluna B): posições
 * - Coluna A (a partir da linha 2): defeitos
 * - Interseção: "x" para relacionamento ativo
 */
function getDefectsByPositionCatalog_() {
  var viewCatalog = getDefectsByPositionFromViews_();
  if (viewCatalog) {
    return viewCatalog;
  }

  var sheet = getRequiredSheet_(SHEETS.CAD_DEFEITOS);
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 2) {
    return {
      posicoes: [],
      defeitos: [],
      defeitosPorPosicao: {},
      paresAtivos: {}
    };
  }

  var values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  var header = values[0];

  var posicoes = [];
  var defectsByPosition = {};
  var activePairs = {};

  for (var columnIndex = 1; columnIndex < header.length; columnIndex += 1) {
    var posicao = String(header[columnIndex] || '').trim();
    if (!posicao) {
      continue;
    }
    posicoes.push(posicao);
    defectsByPosition[posicao] = [];
  }

  var defectSet = {};

  for (var rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    var row = values[rowIndex];
    var defeito = String(row[0] || '').trim();
    if (!defeito) {
      continue;
    }

    defectSet[defeito] = true;

    for (var matrixColumn = 1; matrixColumn < header.length; matrixColumn += 1) {
      var headerPosicao = String(header[matrixColumn] || '').trim();
      if (!headerPosicao || !defectsByPosition[headerPosicao]) {
        continue;
      }

      if (isActiveMatrixFlag_(row[matrixColumn])) {
        defectsByPosition[headerPosicao].push(defeito);
        activePairs[getPositionDefectKey_(headerPosicao, defeito)] = true;
      }
    }
  }

  return {
    posicoes: posicoes,
    defeitos: Object.keys(defectSet),
    defeitosPorPosicao: defectsByPosition,
    paresAtivos: activePairs
  };
}

function getDefectsByPositionFromViews_() {
  var relationsSheet = getOptionalSheet_(SHEETS.VIEW_RELACOES_ATIVAS);
  if (!relationsSheet || relationsSheet.getLastRow() < 2) {
    return null;
  }

  var values = relationsSheet.getRange(2, 1, relationsSheet.getLastRow() - 1, Math.max(relationsSheet.getLastColumn(), 2)).getValues();
  var defectsByPosition = {};
  var activePairs = {};
  var defectSet = {};
  var hasAtLeastOnePair = false;

  values.forEach(function (row) {
    var posicao = String(row[0] || '').trim();
    var defeito = String(row[1] || '').trim();
    var status = row.length > 2 ? row[2] : 'x';

    if (!posicao || !defeito || !isActiveMatrixFlag_(status)) {
      return;
    }

    if (!defectsByPosition[posicao]) {
      defectsByPosition[posicao] = [];
    }
    defectsByPosition[posicao].push(defeito);
    activePairs[getPositionDefectKey_(posicao, defeito)] = true;
    defectSet[defeito] = true;
    hasAtLeastOnePair = true;
  });

  if (!hasAtLeastOnePair) {
    return null;
  }

  var posicoesFromView = getSimpleListFromView_(SHEETS.VIEW_POSICOES_ATIVAS, 1);
  var defeitosFromView = getSimpleListFromView_(SHEETS.VIEW_DEFEITOS_ATIVOS, 1);

  return {
    posicoes: posicoesFromView.length ? posicoesFromView : Object.keys(defectsByPosition),
    defeitos: defeitosFromView.length ? defeitosFromView : Object.keys(defectSet),
    defeitosPorPosicao: defectsByPosition,
    paresAtivos: activePairs
  };
}

/**
 * Grava linhas em lote no final da aba.
 */
function appendRowsBatch_(sheetName, rows) {
  if (!rows || !rows.length) {
    return;
  }

  var sheet = getRequiredSheet_(sheetName);
  var start = sheet.getLastRow() + 1;
  var width = rows[0].length;
  sheet.getRange(start, 1, rows.length, width).setValues(rows);
}

function getRequiredSheet_(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Aba obrigatória não encontrada: ' + sheetName + '. Rode setupSchema().');
  }
  return sheet;
}

function getOptionalSheet_(sheetName) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

function getSimpleListFromView_(sheetName, valueColumnIndex) {
  var sheet = getOptionalSheet_(sheetName);
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  var values = sheet.getRange(2, valueColumnIndex, sheet.getLastRow() - 1, 1).getValues();
  var result = [];
  var seen = {};

  values.forEach(function (row) {
    var value = String(row[0] || '').trim();
    if (!value || seen[value]) {
      return;
    }
    seen[value] = true;
    result.push(value);
  });

  return result;
}

function getCollaboratorsEditPassword_() {
  var sheet = getRequiredSheet_(SHEETS.COLABORADORES);
  return String(sheet.getRange('E2').getValue() || '').trim();
}

function isOriginRequired_() {
  var sheet = getRequiredSheet_(SHEETS.COLABORADORES);
  var flagValue = sheet.getRange('G2').getValue();
  return flagValue === true || String(flagValue || '').trim().toUpperCase() === 'TRUE';
}

function isOpRequired_() {
  var sheet = getRequiredSheet_(SHEETS.COLABORADORES);
  var flagValue = sheet.getRange('G6').getValue();
  return flagValue === true || String(flagValue || '').trim().toUpperCase() === 'TRUE';
}

function isActiveFlag_(value) {
  if (value === true || value === 1) {
    return true;
  }

  var normalized = String(value || '').trim().toLowerCase();
  return normalized === 'true' || normalized === '1' || normalized === 'sim' || normalized === 'ativo';
}

function isActiveMatrixFlag_(value) {
  var normalized = String(value || '').trim().toLowerCase();
  return normalized === 'x' || isActiveFlag_(value);
}

function getPositionDefectKey_(posicao, defeito) {
  return String(posicao).trim() + '||' + String(defeito).trim();
}

/**
 * Lista inspeções com filtros e paginação para a página de listagem.
 */
function listInspectionsControls_(params) {
  var filters = params && typeof params === 'object' ? params : {};
  var pageSize = normalizePageSize_(filters.pageSize);
  var page = normalizePage_(filters.page);
  var period = normalizePeriodFilter_(filters.dataInicial, filters.dataFinal);
  var opFilter = String(filters.op || '').trim().toLowerCase();
  var clienteFilter = String(filters.cliente || '').trim().toLowerCase();
  var origemFilter = String(filters.origem || '').trim().toLowerCase();
  var operadorFilter = String(filters.operador || '').trim().toLowerCase();
  var retrabalhoFilter = normalizeRetrabalhoFilter_(filters.retrabalho);

  var sheet = getRequiredSheet_(SHEETS.INSPECOES);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {
      items: [],
      total: 0,
      page: 1,
      pageSize: pageSize,
      totalPages: 1
    };
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
  var normalizedItems = [];

  values.forEach(function (row, index) {
    var mapped = mapInspectionRow_(row, index + 2);
    if (!mapped) {
      return;
    }
    if (!matchesInspectionFilters_(mapped, {
      period: period,
      op: opFilter,
      cliente: clienteFilter,
      origem: origemFilter,
      operador: operadorFilter,
      retrabalho: retrabalhoFilter
    })) {
      return;
    }
    normalizedItems.push(mapped);
  });

  normalizedItems.sort(function (a, b) {
    return b.idInspecao - a.idInspecao;
  });

  var aggregates = normalizedItems.reduce(function (acc, item) {
    acc.revisados += Number(item && item.qtddRevisada) || 0;
    acc.defeitos += Number(item && item.totalDefeitos) || 0;
    return acc;
  }, { revisados: 0, defeitos: 0 });

  var total = normalizedItems.length;
  var totalPages = Math.max(1, Math.ceil(total / pageSize));
  var safePage = Math.min(page, totalPages);
  var start = (safePage - 1) * pageSize;
  var end = start + pageSize;

  return {
    items: normalizedItems.slice(start, end),
    total: total,
    page: safePage,
    pageSize: pageSize,
    totalPages: totalPages,
    revisadosFiltrados: aggregates.revisados,
    defeitosFiltrados: aggregates.defeitos
  };
}

function updateInspectionControl_(payload) {
  var data = payload && typeof payload === 'object' ? payload : {};
  var idInspecao = Number(data.idInspecao);
  var sheetRowNumber = Number(data.sheetRowNumber);

  if (!isFinite(idInspecao) || idInspecao < 1) {
    throw new Error('ID da inspeção inválido para edição.');
  }

  if (!isFinite(sheetRowNumber) || sheetRowNumber < 2) {
    throw new Error('Linha da inspeção inválida para edição.');
  }

  var sheet = getRequiredSheet_(SHEETS.INSPECOES);
  var values = sheet.getRange(sheetRowNumber, 1, 1, 19).getValues();
  if (!values.length) {
    throw new Error('Inspeção não encontrada para edição.');
  }

  var row = values[0];
  var currentId = Number(row[0]);
  if (!isFinite(currentId) || Math.floor(currentId) !== Math.floor(idInspecao)) {
    throw new Error('A inspeção foi alterada. Recarregue a listagem e tente novamente.');
  }

  var parsedDate = parseSheetDateValue_(data.dataServidor);
  if (parsedDate) {
    row[1] = parsedDate;
  }

  var normalizedOp = normalizeEditedOp_(data.op);
  var normalizedCliente = String(data.cliente || '').trim();
  var normalizedQtd = normalizePositiveNumberOrFail_(data.qtddRevisada);
  var normalizedOrigem = String(data.origem || '').trim();

  validateEditedInspectionBusinessRules_(normalizedOp, normalizedCliente, normalizedQtd, normalizedOrigem);
  if (normalizedOp && isOpRequired_()) {
    normalizedCliente = getClientByOP(normalizedOp) || normalizedCliente;
  }

  row[2] = normalizedOp;
  row[3] = normalizedQtd;
  row[4] = normalizedOrigem;
  row[5] = normalizedCliente;
  row[6] = String(data.criadoPorEmail || '').trim();

  var operators = normalizeFixedOperators_(data.operadores);
  row[7] = operators[0];
  row[8] = operators[1];
  row[9] = operators[2];
  row[10] = operators[3];
  row[11] = operators[4];
  row[12] = operators[5];
  row[13] = operators[6];
  row[14] = operators[7];

  row[15] = normalizeNonNegativeNumber_(data.totalLancamentosDefeitos);

  var defectsJsonRaw = String(data.defeitosJsonRaw || '').trim();
  validateDefectsJsonRaw_(defectsJsonRaw);
  row[16] = defectsJsonRaw;
  row[17] = String(data.defeitosResumo || '').trim();
  row[18] = data.retrabalho ? 'X' : '';

  sheet.getRange(sheetRowNumber, 1, 1, 19).setValues([row]);

  return {
    ok: true
  };
}

function normalizeEditedOp_(value) {
  var normalized = String(value || '').trim();
  if (normalized.toLowerCase() === 'sem op') {
    return '';
  }
  return normalized;
}

function normalizePositiveNumberOrFail_(value) {
  var parsed = Number(value);
  if (!isFinite(parsed) || parsed <= 0) {
    throw new Error('Qtdd revisada deve ser maior que zero.');
  }
  return Math.floor(parsed);
}

function validateEditedInspectionBusinessRules_(op, cliente, qtd, origem) {
  if (qtd <= 0) {
    throw new Error('Qtdd revisada deve ser maior que zero.');
  }

  var opIsRequired = isOpRequired_();
  if (!op) {
    if (opIsRequired) {
      throw new Error('OP é obrigatória.');
    }
    if (!cliente) {
      throw new Error('Cliente é obrigatório quando marcado como Sem OP.');
    }
  } else if (!/^\d+$/.test(op)) {
    throw new Error('OP deve conter apenas dígitos numéricos.');
  } else if (!getClientByOP(op)) {
    throw new Error('OP ' + op + ' não foi encontrada na aba op_cliente.');
  }

  if (isOriginRequired_() && !origem) {
    throw new Error('Origem é obrigatória.');
  }
}

function parseSheetDateValue_(value) {
  var normalized = String(value || '').trim();
  if (!normalized) {
    return null;
  }

  var parsed = new Date(normalized);
  if (isNaN(parsed.getTime())) {
    throw new Error('Data/Hora inválida para edição.');
  }

  return parsed;
}

function normalizeNonNegativeNumber_(value) {
  var parsed = Number(value);
  if (!isFinite(parsed) || parsed < 0) {
    return 0;
  }
  return Math.floor(parsed);
}

function normalizeFixedOperators_(operators) {
  var list = Array.isArray(operators) ? operators : [];
  var normalized = [];
  for (var i = 0; i < 8; i += 1) {
    normalized.push(String(list[i] || '').trim());
  }
  return normalized;
}

function validateDefectsJsonRaw_(raw) {
  if (!raw) {
    return;
  }

  try {
    JSON.parse(raw);
  } catch (error) {
    throw new Error('JSON de defeitos inválido.');
  }
}

function mapInspectionRow_(row, sheetRowNumber) {
  var id = Number(row[0]);
  var hasId = isFinite(id) && id > 0;
  var dataServidor = row[1];
  var hasDate = dataServidor instanceof Date && !isNaN(dataServidor.getTime());

  if (!hasId || !hasDate) {
    return null;
  }

  var operadores = [row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14]]
    .map(function (value) {
      return String(value || '').trim();
    })
    .filter(function (value) {
      return value.length > 0;
    });

  return {
    idInspecao: Math.floor(id),
    dataServidorIso: toIsoDateTime_(dataServidor),
    dataServidorLabel: Utilities.formatDate(dataServidor, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    op: String(row[2] || '').trim(),
    qtddRevisada: Number(row[3]) || 0,
    origem: String(row[4] || '').trim(),
    cliente: String(row[5] || '').trim(),
    criadoPorEmail: String(row[6] || '').trim(),
    operadores: operadores,
    totalLancamentosDefeitos: Number(row[15]) || 0,
    defeitosJsonRaw: String(row[16] || '').trim(),
    defeitosResumo: String(row[17] || '').trim(),
    totalDefeitos: getTotalDefeitosFromJson_(row[16], row[15]),
    retrabalho: isRetrabalhoFlag_(row[18]),
    sheetRowNumber: sheetRowNumber
  };
}

function getTotalDefeitosFromJson_(rawJson, fallbackValue) {
  var fallback = Number(fallbackValue) || 0;
  var normalized = String(rawJson || '').trim();
  if (!normalized) {
    return fallback;
  }

  try {
    var parsed = JSON.parse(normalized);
    if (!Array.isArray(parsed)) {
      return fallback;
    }

    return parsed.reduce(function (total, defect) {
      var quantidade = Number(defect && defect.quantidade);
      if (!isFinite(quantidade) || quantidade < 0) {
        return total;
      }
      return total + quantidade;
    }, 0);
  } catch (error) {
    return fallback;
  }
}

function matchesInspectionFilters_(item, filters) {
  if (filters.period.start && item.dataServidorIso < filters.period.start) {
    return false;
  }
  if (filters.period.end && item.dataServidorIso > filters.period.end) {
    return false;
  }

  if (filters.op && item.op.toLowerCase().indexOf(filters.op) === -1) {
    return false;
  }

  if (filters.cliente && item.cliente.toLowerCase().indexOf(filters.cliente) === -1) {
    return false;
  }

  if (filters.origem && item.origem.toLowerCase() !== filters.origem) {
    return false;
  }

  if (filters.operador) {
    var hasOperator = item.operadores.some(function (name) {
      return name.toLowerCase().indexOf(filters.operador) > -1;
    });
    if (!hasOperator) {
      return false;
    }
  }

  if (filters.retrabalho === 'sim' && !item.retrabalho) {
    return false;
  }
  if (filters.retrabalho === 'nao' && item.retrabalho) {
    return false;
  }

  return true;
}

function normalizePageSize_(pageSize) {
  var parsed = Number(pageSize);
  if (!isFinite(parsed) || parsed < 1) {
    return 50;
  }
  return Math.min(Math.floor(parsed), 100);
}

function normalizePage_(page) {
  var parsed = Number(page);
  if (!isFinite(parsed) || parsed < 1) {
    return 1;
  }
  return Math.floor(parsed);
}

function normalizePeriodFilter_(startRaw, endRaw) {
  var start = normalizeDateInput_(startRaw, false);
  var end = normalizeDateInput_(endRaw, true);
  return {
    start: start,
    end: end
  };
}

function normalizeDateInput_(rawValue, endOfDay) {
  var raw = String(rawValue || '').trim();
  if (!raw) {
    return '';
  }

  var parsed = parseDateOnlyInput_(raw);
  if (!parsed) {
    return '';
  }

  if (endOfDay) {
    parsed.setHours(23, 59, 59, 999);
  } else {
    parsed.setHours(0, 0, 0, 0);
  }

  return toIsoDateTime_(parsed);
}

function parseDateOnlyInput_(raw) {
  var match = String(raw || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) {
    return null;
  }

  var year = Number(match[1]);
  var month = Number(match[2]);
  var day = Number(match[3]);
  if (!isFinite(year) || !isFinite(month) || !isFinite(day)) {
    return null;
  }

  var parsed = new Date(year, month - 1, day, 0, 0, 0, 0);
  if (!(parsed instanceof Date) || isNaN(parsed.getTime())) {
    return null;
  }

  if (parsed.getFullYear() !== year || parsed.getMonth() !== month - 1 || parsed.getDate() !== day) {
    return null;
  }

  return parsed;
}

function toIsoDateTime_(dateValue) {
  return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

function normalizeRetrabalhoFilter_(value) {
  var normalized = String(value || '').trim().toLowerCase();
  if (normalized === 'sim' || normalized === 'nao') {
    return normalized;
  }
  return 'todos';
}

function isRetrabalhoFlag_(value) {
  var normalized = String(value || '').trim().toLowerCase();
  return normalized === 'x' || normalized === 'true' || normalized === '1' || normalized === 'sim';
}

var SHEETS = {
  INSPECOES: 'inspecoes',
  INSPECAO_OPERADORES: 'inspecao_operadores',
  INSPECAO_DEFEITOS: 'inspecao_defeitos',
  COLABORADORES: 'colaboradores',
  CAD_POSICOES: 'cad_posicoes',
  CAD_DEFEITOS: 'cad_defeitos',
  CAD_ORIGENS: 'cad_origens'
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
      return row[activeColumnIndex - 1] === true;
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
      return row[2] === true;
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
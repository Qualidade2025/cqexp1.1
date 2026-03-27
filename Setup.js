/**
 * Configura e valida o schema mínimo de abas e cabeçalhos.
 */
function setupSchema() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var schema = getSchemaDefinition_();

  Object.keys(schema).forEach(function (sheetName) {
    var headers = schema[sheetName];
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    ensureHeaders_(sheet, headers);
  });
}

/**
 * Popula catálogos iniciais para uso imediato do MVP.
 */
function seedCatalogs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  seedSheet_(ss.getSheetByName('colaboradores'), [
    ['C001', 'Ana Souza', true],
    ['C002', 'Bruno Lima', true],
    ['C003', 'Carla Nunes', true],
    ['C004', 'Diego Alves', true]
  ]);

  seedSheet_(ss.getSheetByName('cad_posicoes'), [
    ['Frente', true],
    ['Costas', true],
    ['Manga Esquerda', true],
    ['Manga Direita', true],
    ['Barra', true]
  ]);

  seedSheet_(ss.getSheetByName('cad_defeitos'), [
    ['Mancha', true],
    ['Furo', true],
    ['Falha de costura', true],
    ['Etiqueta incorreta', true],
    ['Desalinhamento', true]
  ]);

  seedSheet_(ss.getSheetByName('cad_origens'), [
    ['Produção', true],
    ['Retrabalho', true],
    ['Terceirizado', true],
    ['Devolução', true]
  ]);
}

function getSchemaDefinition_() {
  return {
    inspecoes: [
      'id_inspecao',
      'data_servidor',
      'op',
      'qtdd_revisada',
      'origem',
      'cliente',
      'criado_por_email'
    ],
    inspecao_operadores: [
      'id_inspecao',
      'id_colaborador',
      'nome_colaborador'
    ],
    inspecao_defeitos: [
      'id_inspecao',
      'linha',
      'posicao',
      'defeito',
      'quantidade'
    ],
    colaboradores: [
      'id_colaborador',
      'nome_colaborador',
      'ativo'
    ],
    cad_posicoes: ['posicao', 'ativo'],
    cad_defeitos: ['defeito', 'ativo'],
    cad_origens: ['origem', 'ativo']
  };
}

function ensureHeaders_(sheet, headers) {
  var range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function seedSheet_(sheet, rows) {
  if (!sheet) {
    throw new Error('Aba não encontrada para seed. Execute setupSchema() primeiro.');
  }

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}
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
  
  setupCatalogViews();
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

  seedSheet_(ss.getSheetByName('cad_defeitos'), [
    ['Mancha', 'x', 'x', '', '', 'x'],
    ['Furo', 'x', 'x', 'x', 'x', ''],
    ['Falha de costura', '', 'x', 'x', 'x', 'x'],
    ['Etiqueta incorreta', 'x', '', '', '', 'x'],
    ['Desalinhamento', 'x', 'x', 'x', '', 'x']
  ]);

  seedSheet_(ss.getSheetByName('cad_origens'), [
    ['Produção', true],
    ['Retrabalho', true],
    ['Terceirizado', true],
    ['Devolução', true]
  ]);

   setupCatalogViews();
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
      'criado_por_email',
      'operador_1',
      'operador_2',
      'operador_3',
      'operador_4',
      'operador_5',
      'operador_6',
      'operador_7',
      'operador_8',
      'total_lancamentos_defeitos',
      'defeitos_json',
      'defeitos_resumo',
      'retrabalho'
    ],
    colaboradores: [
      'id_colaborador',
      'nome_colaborador',
      'ativo'
    ],
    cad_defeitos: ['defeito', 'Frente', 'Costas', 'Manga Esquerda', 'Manga Direita', 'Barra'],
    cad_origens: ['origem', 'ativo'],
    view_relacoes_ativas: ['posicao', 'defeito', 'status'],
    view_posicoes_ativas: ['posicao'],
    view_defeitos_ativos: ['defeito']
  };
}

/**
 * Configura fórmulas auxiliares para acelerar leitura de catálogos.
 */
function setupCatalogViews() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var relacoes = ss.getSheetByName('view_relacoes_ativas');
  var posicoes = ss.getSheetByName('view_posicoes_ativas');
  var defeitos = ss.getSheetByName('view_defeitos_ativos');

  if (!relacoes || !posicoes || !defeitos) {
    return;
  }

  if (relacoes.getLastRow() > 1) {
    relacoes.getRange(2, 1, relacoes.getLastRow() - 1, Math.max(relacoes.getLastColumn(), 3)).clearContent();
  }
  if (posicoes.getLastRow() > 1) {
    posicoes.getRange(2, 1, posicoes.getLastRow() - 1, Math.max(posicoes.getLastColumn(), 1)).clearContent();
  }
  if (defeitos.getLastRow() > 1) {
    defeitos.getRange(2, 1, defeitos.getLastRow() - 1, Math.max(defeitos.getLastColumn(), 1)).clearContent();
  }

  relacoes.getRange(2, 1).setFormula(
    '=ARRAYFORMULA(QUERY(SPLIT(FLATTEN(IF(cad_defeitos!B2:ZZ<>"",cad_defeitos!B1:ZZ1&"♦"&cad_defeitos!A2:A&"♦"&cad_defeitos!B2:ZZ,"")),"♦"),"select Col1,Col2,Col3 where Col3 is not null",0))'
  );
  posicoes.getRange(2, 1).setFormula(
    '=ARRAYFORMULA(QUERY(UNIQUE(FILTER(view_relacoes_ativas!A2:A,view_relacoes_ativas!A2:A<>"",REGEXMATCH(LOWER(view_relacoes_ativas!C2:C),"^(x|1|true|sim|ativo)$"))),"select Col1",0))'
  );
  defeitos.getRange(2, 1).setFormula(
    '=ARRAYFORMULA(QUERY(UNIQUE(FILTER(view_relacoes_ativas!B2:B,view_relacoes_ativas!B2:B<>"",REGEXMATCH(LOWER(view_relacoes_ativas!C2:C),"^(x|1|true|sim|ativo)$"))),"select Col1",0))'
  );
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

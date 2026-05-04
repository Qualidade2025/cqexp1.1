/**
 * Exporta mensalmente a aba "inspecoes" em .xlsx e .csv,
 * salva os arquivos na pasta do Drive indicada em colaboradores!
 * adiciona os registros no histórico do Power BI
 * e limpa os dados da aba (mantendo o cabeçalho) após sucesso.
 */
function exportMonthlyInspectionsBackup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inspectionsSheet = getRequiredSheet_(SHEETS.INSPECOES);
  var folder = getBackupFolderFromCollaborators_();
  var rowsToArchive = getInspectionsDataRows_(inspectionsSheet);
  var now = new Date();
  var month = Utilities.formatDate(now, 'Etc/GMT-3', 'MM');
  var year = Utilities.formatDate(now, 'Etc/GMT-3', 'yyyy');
  var baseName = 'inspecoes-' + month + '.' + year;

  SpreadsheetApp.flush();
  Utilities.sleep(1500);

  var xlsxBlob = exportSpreadsheetBlob_(ss.getId(), {
    format: 'xlsx',
    filename: baseName + '.xlsx'
  });

  var csvBlob = exportSpreadsheetBlob_(ss.getId(), {
    format: 'csv',
    filename: baseName + '.csv',
    gid: inspectionsSheet.getSheetId()
  });

  ensureNonEmptyBackupBlob_(xlsxBlob, 'xlsx');
  ensureNonEmptyBackupBlob_(csvBlob, 'csv');

  folder.createFile(xlsxBlob);
  folder.createFile(csvBlob);
  appendRowsToPowerBiHistory_(rowsToArchive);

  clearInspectionsDataRows_();

  return {
    ok: true,
    files: [baseName + '.xlsx', baseName + '.csv'],
    folderId: folder.getId()
  };
}

/**
 * Cria (ou recria) o gatilho mensal para executar no dia 1 às 03:00 (GMT+3).
 */
function createMonthlyInspectionsBackupTrigger() {
  var functionName = 'exportMonthlyInspectionsBackup';
  var timezone = 'Etc/GMT-3';

  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(functionName)
    .timeBased()
    .onMonthDay(1)
    .atHour(3)
    .inTimezone(timezone)
    .create();

  return {
    ok: true,
    functionName: functionName,
    schedule: 'Mensal, dia 1 às 03:00 (GMT+3)',
    timezone: timezone
  };
}

function getBackupFolderFromCollaborators_() {
  var collaboratorsSheet = getRequiredSheet_(SHEETS.COLABORADORES);
  var folderLink = String(collaboratorsSheet.getRange('J2').getValue() || '').trim();

  if (!folderLink) {
    throw new Error('Link da pasta de backup não encontrado em colaboradores!J2.');
  }

  var folderIdMatch = folderLink.match(/[-\w]{25,}/);
  if (!folderIdMatch) {
    throw new Error('Não foi possível extrair o ID da pasta a partir de colaboradores!J2.');
  }

  return DriveApp.getFolderById(folderIdMatch[0]);
}

function exportSpreadsheetBlob_(spreadsheetId, options) {
  var format = options.format;
  var filename = options.filename;
  var gid = options.gid;
  var exportUrl = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?format=' + encodeURIComponent(format);

  if (gid || gid === 0) {
    exportUrl += '&gid=' + encodeURIComponent(gid);
  }

  var response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: false
  });

  return response.getBlob().setName(filename);
}

function ensureNonEmptyBackupBlob_(blob, format) {
  if (!blob || blob.getBytes().length === 0) {
    throw new Error('Falha ao gerar backup ' + format.toUpperCase() + ': arquivo vazio.');
  }
}

function clearInspectionsDataRows_() {
  var inspectionsSheet = getRequiredSheet_(SHEETS.INSPECOES);
  var lastRow = inspectionsSheet.getLastRow();

  if (lastRow <= 1) {
    return;
  }

  inspectionsSheet.getRange(2, 1, lastRow - 1, inspectionsSheet.getLastColumn()).clearContent();
}

function getInspectionsDataRows_(inspectionsSheet) {
  var lastRow = inspectionsSheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  return inspectionsSheet
    .getRange(2, 1, lastRow - 1, inspectionsSheet.getLastColumn())
    .getValues();
}

function appendRowsToPowerBiHistory_(rows) {
  if (!rows || !rows.length) {
    return;
  }

  var historySpreadsheet = SpreadsheetApp.openById('1wrkjbWVWGU4ce8Zk_BbgKsb47uSaTdUCUmd0gB5_hs4');
  var historySheet = historySpreadsheet.getSheetByName(SHEETS.INSPECOES);

  if (!historySheet) {
    throw new Error('Aba histórica não encontrada no arquivo CQ Power BI: ' + SHEETS.INSPECOES + '.');
  }

  var startRow = historySheet.getLastRow() + 1;
  historySheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}
// ============================================
// MAKRO TIME INTELLIGENCE - GOOGLE APPS SCRIPT
// ============================================
// Cole este codigo em Extensões > Apps Script da sua planilha

// ---------- SETUP (Execute uma vez) ----------
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Aba Config
  let configSheet = ss.getSheetByName('Config');
  if (!configSheet) configSheet = ss.insertSheet('Config');
  configSheet.clear();
  configSheet.getRange('A1:B1').setValues([['Campo', 'Valor']]).setFontWeight('bold').setBackground('#E3000F').setFontColor('white');
  configSheet.getRange('A2:B5').setValues([
    ['salario', 1000],
    ['entrada', '07:50'],
    ['saida', '17:38'],
    ['diasSemana', '1,2,3,4,5']
  ]);
  configSheet.setColumnWidth(1, 150);
  configSheet.setColumnWidth(2, 200);

  // Aba Ponto
  let pontoSheet = ss.getSheetByName('Ponto');
  if (!pontoSheet) pontoSheet = ss.insertSheet('Ponto');
  pontoSheet.clear();
  pontoSheet.getRange('A1:D1').setValues([['ID', 'Data', 'Entrada', 'Saida']]).setFontWeight('bold').setBackground('#E3000F').setFontColor('white');
  pontoSheet.setColumnWidth(1, 120);
  pontoSheet.setColumnWidth(2, 120);
  pontoSheet.setColumnWidth(3, 100);
  pontoSheet.setColumnWidth(4, 100);

  // Aba Manual
  let manualSheet = ss.getSheetByName('Manual');
  if (!manualSheet) manualSheet = ss.insertSheet('Manual');
  manualSheet.clear();
  manualSheet.getRange('A1:E1').setValues([['ID', 'Ref', 'Horas', 'Tipo', 'Decimal']]).setFontWeight('bold').setBackground('#E3000F').setFontColor('white');
  manualSheet.setColumnWidth(1, 120);
  manualSheet.setColumnWidth(2, 100);
  manualSheet.setColumnWidth(3, 100);
  manualSheet.setColumnWidth(4, 100);
  manualSheet.setColumnWidth(5, 100);

  // Aba Conquistas
  let achSheet = ss.getSheetByName('Conquistas');
  if (!achSheet) achSheet = ss.insertSheet('Conquistas');
  achSheet.clear();
  achSheet.getRange('A1:B1').setValues([['Achievement ID', 'Data']]).setFontWeight('bold').setBackground('#E3000F').setFontColor('white');
  achSheet.setColumnWidth(1, 200);
  achSheet.setColumnWidth(2, 150);

  // Remove aba Sheet1 padrao
  const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Página1');
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);

  SpreadsheetApp.getUi().alert('✅ Planilha configurada com sucesso!\n\nAbas criadas: Config, Ponto, Manual, Conquistas');
}

// ---------- API ----------
function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const action = e.parameter.action;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let result;

  try {
    switch(action) {
      case 'getConfig':
        result = getConfig(ss);
        break;
      case 'saveConfig':
        result = saveConfig(ss, JSON.parse(e.postData.contents));
        break;
      case 'getPonto':
        result = getPonto(ss);
        break;
      case 'savePonto':
        result = savePonto(ss, JSON.parse(e.postData.contents));
        break;
      case 'deletePonto':
        result = deletePonto(ss, e.parameter.id);
        break;
      case 'getManual':
        result = getManual(ss);
        break;
      case 'saveManual':
        result = saveManual(ss, JSON.parse(e.postData.contents));
        break;
      case 'deleteManual':
        result = deleteManual(ss, e.parameter.id);
        break;
      case 'getAchievements':
        result = getAchievements(ss);
        break;
      case 'saveAchievement':
        result = saveAchievement(ss, JSON.parse(e.postData.contents));
        break;
      case 'getAll':
        result = getAll(ss);
        break;
      default:
        result = { error: 'Ação desconhecida: ' + action };
    }
  } catch(err) {
    result = { error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- CONFIG ----------
function getConfig(ss) {
  const sheet = ss.getSheetByName('Config');
  const data = sheet.getRange('A2:B5').getValues();
  const config = {};
  data.forEach(row => {
    if(row[0] === 'diasSemana') {
      config[row[0]] = String(row[1]).split(',').map(Number);
    } else {
      config[row[0]] = row[1];
    }
  });
  return config;
}

function saveConfig(ss, config) {
  const sheet = ss.getSheetByName('Config');
  sheet.getRange('A2:B5').setValues([
    ['salario', config.salario],
    ['entrada', config.entrada],
    ['saida', config.saida],
    ['diasSemana', config.diasSemana.join(',')]
  ]);
  return { success: true };
}

// ---------- PONTO ----------
function getPonto(ss) {
  const sheet = ss.getSheetByName('Ponto');
  const lastRow = sheet.getLastRow();
  if(lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  return data.map(row => ({ id: row[0], date: row[1], entrada: row[2], saida: row[3] }));
}

function savePonto(ss, ponto) {
  const sheet = ss.getSheetByName('Ponto');
  const lastRow = sheet.getLastRow();
  if(lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const idx = data.findIndex(row => String(row[0]) === String(ponto.id));
    if(idx >= 0) {
      sheet.getRange(idx + 2, 1, 1, 4).setValues([[ponto.id, ponto.date, ponto.entrada, ponto.saida]]);
      return { success: true, action: 'updated' };
    }
  }
  sheet.appendRow([ponto.id, ponto.date, ponto.entrada, ponto.saida]);
  return { success: true, action: 'created' };
}

function deletePonto(ss, id) {
  const sheet = ss.getSheetByName('Ponto');
  const lastRow = sheet.getLastRow();
  if(lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const idx = data.findIndex(row => String(row[0]) === String(id));
    if(idx >= 0) {
      sheet.deleteRow(idx + 2);
      return { success: true };
    }
  }
  return { error: 'Registro não encontrado' };
}

// ---------- MANUAL ----------
function getManual(ss) {
  const sheet = ss.getSheetByName('Manual');
  const lastRow = sheet.getLastRow();
  if(lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  return data.map(row => ({ id: row[0], ref: row[1], hrsStr: row[2], tipo: row[3], decimal: row[4] }));
}

function saveManual(ss, manual) {
  const sheet = ss.getSheetByName('Manual');
  const lastRow = sheet.getLastRow();
  if(lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const idx = data.findIndex(row => String(row[0]) === String(manual.id));
    if(idx >= 0) {
      sheet.getRange(idx + 2, 1, 1, 5).setValues([[manual.id, manual.ref, manual.hrsStr, manual.tipo, manual.decimal]]);
      return { success: true, action: 'updated' };
    }
  }
  sheet.appendRow([manual.id, manual.ref, manual.hrsStr, manual.tipo, manual.decimal]);
  return { success: true, action: 'created' };
}

function deleteManual(ss, id) {
  const sheet = ss.getSheetByName('Manual');
  const lastRow = sheet.getLastRow();
  if(lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const idx = data.findIndex(row => String(row[0]) === String(id));
    if(idx >= 0) {
      sheet.deleteRow(idx + 2);
      return { success: true };
    }
  }
  return { error: 'Registro não encontrado' };
}

// ---------- CONQUISTAS ----------
function getAchievements(ss) {
  const sheet = ss.getSheetByName('Conquistas');
  const lastRow = sheet.getLastRow();
  if(lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  return data.map(row => row[0]);
}

function saveAchievement(ss, ach) {
  const sheet = ss.getSheetByName('Conquistas');
  const lastRow = sheet.getLastRow();
  if(lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const exists = data.some(row => String(row[0]) === String(ach.id));
    if(exists) return { success: true, action: 'already_exists' };
  }
  sheet.appendRow([ach.id, ach.date || new Date().toISOString()]);
  return { success: true, action: 'created' };
}

// ---------- ALL ----------
function getAll(ss) {
  return {
    config: getConfig(ss),
    ponto: getPonto(ss),
    manual: getManual(ss),
    achievements: getAchievements(ss)
  };
}
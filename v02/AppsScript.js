// ============================================
// MAKRO TIME INTELLIGENCE - GOOGLE APPS SCRIPT
// ============================================

// ---------- SETUP (Execute uma vez para criar a estrutura) ----------
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // TB_CONFIGURACOES
  var configSheet = ss.getSheetByName('TB_CONFIGURACOES');
  if(!configSheet) configSheet = ss.insertSheet('TB_CONFIGURACOES');
  configSheet.clear();
  configSheet.getRange('A1:B1').setValues([['Chave', 'Valor']]).setFontWeight('bold').setBackground('#E3000F').setFontColor('white');
  configSheet.getRange('A2:B7').setValues([
    ['jornada_inicio', '07:50'],
    ['jornada_fim', '17:38'],
    ['horas_almoco', 1],
    ['dias_uteis', '1, 2, 3, 4, 5'],
    ['saldo_inicial_minutos', 0],
    ['salario', 0]
  ]);
  configSheet.setColumnWidth(1, 200);
  configSheet.setColumnWidth(2, 200);

  // TB_REGISTROS
  var regSheet = ss.getSheetByName('TB_REGISTROS');
  if(!regSheet) regSheet = ss.insertSheet('TB_REGISTROS');
  regSheet.clear();
  regSheet.getRange('A1:D1').setValues([['Data', 'Entrada', 'Saida', 'Saldo']]).setFontWeight('bold').setBackground('#E3000F').setFontColor('white');
  regSheet.setColumnWidth(1, 120);
  regSheet.setColumnWidth(2, 100);
  regSheet.setColumnWidth(3, 100);
  regSheet.setColumnWidth(4, 100);

  // Remove abas padrao
  var defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Página1');
  if(defaultSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet); } catch(e) {}
  }

  SpreadsheetApp.getUi().alert('Planilha configurada!\n\nAbas criadas:\n- TB_CONFIGURACOES\n- TB_REGISTROS');
}

// ---------- TESTE ----------
function testar() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = getAll(ss);
    SpreadsheetApp.getUi().alert('OK!\n\n' + JSON.stringify(result, null, 2));
  } catch(e) {
    SpreadsheetApp.getUi().alert('ERRO: ' + e.message);
  }
}

// ---------- API ----------
function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var action = e.parameter.action;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result;
  try {
    var body = null;
    if(e.parameter.data) {
      try { body = JSON.parse(e.parameter.data); } catch(x) {}
    }
    
    if(action === 'getAll') result = getAll(ss);
    else if(action === 'getConfig') result = getConfig(ss);
    else if(action === 'saveConfig' && body) result = saveConfig(ss, body);
    else if(action === 'getRegistros') result = getRegistros(ss);
    else if(action === 'saveRegistro' && body) result = saveRegistro(ss, body);
    else if(action === 'deleteRegistro' && body) result = deleteRegistro(ss, body.date || body.data);
    else if(action === 'testar') result = {ok: true, abas: ss.getSheets().map(function(s){return s.getName()}), config: getConfig(ss), registros: getRegistros(ss).length};
    else result = {error: 'Acao desconhecida: ' + action};
  } catch(err) {
    result = {error: err.message};
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- CONFIG ----------
function getConfig(ss) {
  var sheet = ss.getSheetByName('TB_CONFIGURACOES');
  if(!sheet) return {};
  var lastRow = sheet.getLastRow();
  if(lastRow < 2) return {};
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var config = {};
  for(var i = 0; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    var val = data[i][1];
    if(key === 'jornada_inicio') config.entrada = formatTime(val);
    else if(key === 'jornada_fim') config.saida = formatTime(val);
    else if(key === 'horas_almoco') config.horasAlmoco = Number(val) || 1;
    else if(key === 'dias_uteis') {
      config.diasSemana = String(val).split(',');
      for(var j = 0; j < config.diasSemana.length; j++) config.diasSemana[j] = Number(config.diasSemana[j].trim());
    }
    else if(key === 'saldo_inicial_minutos') config.saldoInicialMin = Number(val) || 0;
    else if(key === 'salario') config.salario = Number(val) || 0;
  }
  return config;
}

function saveConfig(ss, config) {
  var sheet = ss.getSheetByName('TB_CONFIGURACOES');
  if(!sheet) return {error: 'Aba nao encontrada'};
  var lastRow = sheet.getLastRow();
  if(lastRow < 2) return {error: 'Vazio'};
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for(var i = 0; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    if(key === 'jornada_inicio') sheet.getRange(i+2, 2).setValue(config.entrada);
    else if(key === 'jornada_fim') sheet.getRange(i+2, 2).setValue(config.saida);
    else if(key === 'horas_almoco') sheet.getRange(i+2, 2).setValue(config.horasAlmoco);
    else if(key === 'dias_uteis') sheet.getRange(i+2, 2).setValue(config.diasSemana.join(', '));
    else if(key === 'saldo_inicial_minutos') sheet.getRange(i+2, 2).setValue(config.saldoInicialMin);
    else if(key === 'salario') sheet.getRange(i+2, 2).setValue(config.salario);
  }
  return {success: true};
}

// ---------- REGISTROS ----------
function getRegistros(ss) {
  var sheet = ss.getSheetByName('TB_REGISTROS');
  if(!sheet) return [];
  var lastRow = sheet.getLastRow();
  if(lastRow <= 1) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var result = [];
  for(var i = 0; i < data.length; i++) {
    var d = formatDate(data[i][0]);
    if(d) result.push({date: d, entrada: String(data[i][1]||''), saida: String(data[i][2]||''), saldo: String(data[i][3]||'')});
  }
  return result;
}

function saveRegistro(ss, reg) {
  var sheet = ss.getSheetByName('TB_REGISTROS');
  if(!sheet) return {error: 'Aba nao encontrada'};
  var lastRow = sheet.getLastRow();
  if(lastRow > 1) {
    var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    for(var i = 0; i < data.length; i++) {
      if(formatDate(data[i][0]) === reg.date) {
        sheet.getRange(i+2, 2, 1, 3).setValues([[reg.entrada, reg.saida, reg.saldo]]);
        return {success: true, action: 'updated'};
      }
    }
  }
  sheet.appendRow([reg.date, reg.entrada, reg.saida, reg.saldo]);
  return {success: true, action: 'created'};
}

function deleteRegistro(ss, dateStr) {
  var sheet = ss.getSheetByName('TB_REGISTROS');
  if(!sheet) return {error: 'Aba nao encontrada'};
  var lastRow = sheet.getLastRow();
  if(lastRow > 1) {
    var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    for(var i = 0; i < data.length; i++) {
      if(formatDate(data[i][0]) === dateStr) {
        sheet.deleteRow(i+2);
        return {success: true};
      }
    }
  }
  return {error: 'Nao encontrado'};
}

// ---------- ALL ----------
function getAll(ss) {
  return {config: getConfig(ss), registros: getRegistros(ss)};
}

// ---------- HELPER ----------
function formatDate(val) {
  if(!val) return '';
  if(val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(val).substring(0, 10);
}

function formatTime(val) {
  if(!val) return '';
  if(val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
  var s = String(val);
  if(s.indexOf(':') >= 0) return s.substring(0, 5);
  return s;
}

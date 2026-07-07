// ============================================
// MAKRO TIME INTELLIGENCE - GOOGLE APPS SCRIPT
// Compativel com a planilha BD_bancodehoras
// ============================================

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
      case 'getAll':
        result = getAll(ss);
        break;
      case 'getConfig':
        result = getConfig(ss);
        break;
      case 'saveConfig':
        result = saveConfig(ss, JSON.parse(e.postData.contents));
        break;
      case 'getRegistros':
        result = getRegistros(ss);
        break;
      case 'saveRegistro':
        result = saveRegistro(ss, JSON.parse(e.postData.contents));
        break;
      case 'deleteRegistro':
        result = deleteRegistro(ss, e.parameter.data);
        break;
      default:
        result = { error: 'Acao desconhecida: ' + action };
    }
  } catch(err) {
    result = { error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- CONFIG (TB_CONFIGURACOES) ----------
function getConfig(ss) {
  const sheet = ss.getSheetByName('TB_CONFIGURACOES');
  if(!sheet) return { error: 'Aba TB_CONFIGURACOES nao encontrada' };
  const lastRow = sheet.getLastRow();
  if(lastRow < 2) return {};
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const config = {};
  data.forEach(row => {
    const key = String(row[0]).trim();
    const val = row[1];
    if(key === 'dias_uteis') {
      config.diasSemana = String(val).split(',').map(s => Number(s.trim()));
    } else if(key === 'jornada_inicio') {
      config.entrada = String(val);
    } else if(key === 'jornada_fim') {
      config.saida = String(val);
    } else if(key === 'horas_almoco') {
      config.horasAlmoco = Number(val) || 1;
    } else if(key === 'saldo_inicial_minutos') {
      config.saldoInicialMin = Number(val) || 0;
    } else if(key === 'salario') {
      config.salario = Number(val) || 0;
    }
  });
  return config;
}

function saveConfig(ss, config) {
  const sheet = ss.getSheetByName('TB_CONFIGURACOES');
  if(!sheet) return { error: 'Aba TB_CONFIGURACOES nao encontrada' };
  const lastRow = sheet.getLastRow();
  if(lastRow < 2) return { error: 'Planilha vazia' };
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  const map = {
    'jornada_inicio': config.entrada,
    'jornada_fim': config.saida,
    'horas_almoco': config.horasAlmoco,
    'dias_uteis': config.diasSemana.join(', '),
    'saldo_inicial_minutos': config.saldoInicialMin,
    'salario': config.salario
  };

  data.forEach((row, i) => {
    const key = String(row[0]).trim();
    if(map[key] !== undefined) {
      sheet.getRange(i + 2, 2).setValue(map[key]);
    }
  });
  return { success: true };
}

// ---------- REGISTROS (TB_REGISTROS) ----------
function getRegistros(ss) {
  const sheet = ss.getSheetByName('TB_REGISTROS');
  if(!sheet) return [];
  const lastRow = sheet.getLastRow();
  if(lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  return data.map(row => ({
    date: formatDate(row[0]),
    entrada: String(row[1] || ''),
    saida: String(row[2] || ''),
    saldo: String(row[3] || '')
  })).filter(r => r.date);
}

function saveRegistro(ss, reg) {
  const sheet = ss.getSheetByName('TB_REGISTROS');
  if(!sheet) return { error: 'Aba TB_REGISTROS nao encontrada' };
  const lastRow = sheet.getLastRow();
  if(lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const idx = data.findIndex(row => formatDate(row[0]) === reg.date);
    if(idx >= 0) {
      sheet.getRange(idx + 2, 2, 1, 3).setValues([[reg.entrada, reg.saida, reg.saldo]]);
      return { success: true, action: 'updated' };
    }
  }
  sheet.appendRow([reg.date, reg.entrada, reg.saida, reg.saldo]);
  return { success: true, action: 'created' };
}

function deleteRegistro(ss, dateStr) {
  const sheet = ss.getSheetByName('TB_REGISTROS');
  if(!sheet) return { error: 'Aba nao encontrada' };
  const lastRow = sheet.getLastRow();
  if(lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const idx = data.findIndex(row => formatDate(row[0]) === dateStr);
    if(idx >= 0) {
      sheet.deleteRow(idx + 2);
      return { success: true };
    }
  }
  return { error: 'Registro nao encontrado' };
}

// ---------- ALL ----------
function getAll(ss) {
  return {
    config: getConfig(ss),
    registros: getRegistros(ss)
  };
}

// ---------- TESTE (Execute esta funcao para testar) ----------
function testar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = getAll(ss);
  Logger.log(JSON.stringify(result, null, 2));
  SpreadsheetApp.getUi().alert('Teste OK!\n\n' + JSON.stringify(result, null, 2));
}

// ---------- HELPERS ----------
function formatDate(val) {
  if(!val) return '';
  if(val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val).substring(0, 10);
}

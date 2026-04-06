// ============================================================
//  SULMAK – Google Apps Script para integração com o sistema
//  Cole este código em: Extensions > Apps Script > Code.gs
//  Depois: Deploy > New deployment > Web app
//    - Execute as: Me
//    - Who has access: Anyone
// ============================================================

// Nome da aba onde os pedidos ativos ficam armazenados
var ABA_DADOS = 'Pedidos';

// Colunas da planilha (na ordem exata)
var COLUNAS = ['id','filial','tipo','situacao','empresa','cliente','vendedor','endereco','contato','obs','data','hora','produtos'];

// ---- Ponto de entrada HTTP ----
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var acao    = payload.acao;

    // ── LEITURAS: sem lock (rápidas, não modificam dados) ──────────────────
    if (acao === 'test')        return jsonResponse({ status: 'ok', msg: 'Conexão OK' });
    if (acao === 'buscar')      return jsonResponse(buscar());
    if (acao === 'buscarTodos') return jsonResponse(buscarTodos());

    // ── ESCRITAS: lock exclusivo para evitar race conditions ───────────────
    //    insert já gera o ID internamente (atômico), eliminando a necessidade
    //    de uma chamada prévia a proximoId.
    var lock = LockService.getScriptLock();
    lock.waitLock(20000); // espera até 20 s para obter o lock

    try {
      var result;
      if      (acao === 'insert')    result = inserir(payload);
      else if (acao === 'update')    result = atualizar(payload);
      else if (acao === 'delete')    result = excluir(payload);
      else if (acao === 'proximoId') result = proximoId(); // mantido por compatibilidade
      else result = { status: 'erro', msg: 'Ação desconhecida: ' + acao };

      return jsonResponse(result);
    } catch(err) {
      return jsonResponse({ status: 'erro', msg: err.toString() });
    } finally {
      lock.releaseLock();
    }

  } catch(err) {
    return jsonResponse({ status: 'erro', msg: 'Erro ao processar payload: ' + err.toString() });
  }
}

// Permite chamadas GET simples (ex: testar no navegador)
function doGet(e) {
  return jsonResponse({ status: 'ok', msg: 'SULMAK Sheets API online' });
}

// ---- Helpers de planilha ----
function getAba() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(ABA_DADOS);
  if (!aba) {
    aba = ss.insertSheet(ABA_DADOS);
    aba.appendRow(COLUNAS);
    aba.setFrozenRows(1);
  } else {
    // Corrige o cabeçalho se estiver desatualizado (ex: coluna vendedor faltando)
    var headerRange = aba.getRange(1, 1, 1, Math.max(aba.getLastColumn(), COLUNAS.length));
    var headerAtual = headerRange.getValues()[0];
    var precisaCorrigir = false;
    for (var i = 0; i < COLUNAS.length; i++) {
      if (headerAtual[i] !== COLUNAS[i]) { precisaCorrigir = true; break; }
    }
    if (precisaCorrigir) {
      // Salva dados existentes, reconstrói com cabeçalho correto
      var dadosExistentes = aba.getLastRow() > 1
        ? aba.getRange(2, 1, aba.getLastRow() - 1, headerAtual.filter(String).length).getValues()
        : [];
      var mapeamento = headerAtual.map(function(col) { return COLUNAS.indexOf(col); });
      aba.clearContents();
      aba.appendRow(COLUNAS);
      dadosExistentes.forEach(function(row) {
        var novaRow = COLUNAS.map(function(_, ci) {
          var oldIdx = headerAtual.indexOf(COLUNAS[ci]);
          return oldIdx >= 0 ? row[oldIdx] : '';
        });
        aba.appendRow(novaRow);
      });
      aba.setFrozenRows(1);
    }
  }
  return aba;
}

function getAllRows(aba) {
  var data = aba.getDataRange().getValues();
  if (data.length <= 1) return [];
  var header = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    // Mapeia pelo header real da planilha
    header.forEach(function(col, i) { if (col) obj[col] = row[i]; });
    // Garante que todas as colunas canônicas existem (mesmo que não estejam no header)
    COLUNAS.forEach(function(col) {
      if (obj[col] === undefined) obj[col] = '';
    });
    return obj;
  });
}

function rowToArray(obj) {
  return COLUNAS.map(function(col) { return obj[col] !== undefined ? obj[col] : ''; });
}

function findRowIndex(aba, id) {
  var values = aba.getRange(2, 1, Math.max(aba.getLastRow() - 1, 1), 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).padStart(5,'0') === String(id).padStart(5,'0')) {
      return i + 2;
    }
  }
  return -1;
}

// ---- Formata a data/hora para salvar como texto puro ----
function textoData(val) {
  if (!val) return '';
  return "'" + String(val);
}

// ---- Normaliza uma linha do Sheets para objeto JS ----
function normalizarRow(r) {
  r.id       = String(r.id).padStart(5, '0');
  r.data     = String(r.data     || '').replace(/^'/, '');
  r.hora     = String(r.hora     || '').replace(/^'/, '');
  // Células vazias chegam como 0 ou Date — forçar string em todos os campos
  r.vendedor = String(r.vendedor || '');
  r.empresa  = String(r.empresa  || '');
  r.cliente  = String(r.cliente  || '');
  r.endereco = String(r.endereco || '');
  r.contato  = String(r.contato  || '');
  r.obs      = String(r.obs      || '');
  r.situacao = String(r.situacao || '');
  r.tipo     = String(r.tipo     || '');
  r.filial   = String(r.filial   || '');
  try { if (typeof r.produtos === 'string') r.produtos = JSON.parse(r.produtos); }
  catch(e) { r.produtos = []; }
  return r;
}

// ---- Lê o maior ID existente na planilha (já dentro do lock quando chamado por inserir) ----
function calcularProximoId(aba) {
  var rows = getAllRows(aba);
  var max  = 0;
  rows.forEach(function(r) {
    var n = parseInt(r.id, 10);
    if (!isNaN(n) && n > max) max = n;
  });
  return String(max + 1).padStart(5, '0');
}

// ---- Ações ----

// Retorna todos os pedidos ATIVOS (não Finalizada/Cancelada)
function buscar() {
  var aba  = getAba();
  var rows = getAllRows(aba);
  var sitHist = ['Finalizada', 'Cancelada'];

  var dados = rows
    .filter(function(r) { return r.id && !sitHist.includes(r.situacao); })
    .map(normalizarRow);

  return { status: 'ok', dados: dados };
}

// Retorna TODOS os pedidos: ativos separados do histórico
function buscarTodos() {
  var aba  = getAba();
  var rows = getAllRows(aba);
  var sitHist = ['Finalizada', 'Cancelada'];

  var ativos    = [];
  var historico = [];

  rows.forEach(function(r) {
    if (!r.id) return;
    r = normalizarRow(r);
    if (sitHist.includes(r.situacao)) {
      historico.push(r);
    } else {
      ativos.push(r);
    }
  });

  return { status: 'ok', ativos: ativos, historico: historico };
}

// ── INSERIR com geração atômica de ID ─────────────────────────────────────
//    O ID é gerado aqui, dentro do lock, garantindo unicidade mesmo com
//    múltiplos usuários simultâneos. O ID gerado é devolvido ao frontend.
function inserir(payload) {
  var aba = getAba();

  // Gera o próximo ID de forma atômica (já estamos dentro do lock)
  var id = payload.id ? String(payload.id).padStart(5, '0') : calcularProximoId(aba);

  // ── Proteção anti-duplicata: se o ID já existe na planilha, atualiza em vez de inserir ──
  if (payload.id) {
    var idxExistente = findRowIndex(aba, id);
    if (idxExistente !== -1) {
      payload.id = id;
      return atualizar(payload);
    }
  }

  var obj = {
    id:       id,
    filial:   payload.filial   || '',
    tipo:     payload.tipo     || '',
    situacao: payload.situacao || '',
    empresa:  payload.empresa  || '',
    cliente:  payload.cliente  || '',
    vendedor: payload.vendedor || '',
    endereco: payload.endereco || '',
    contato:  payload.contato  || '',
    obs:      payload.obs      || '',
    data:     String(payload.data || ''),
    hora:     String(payload.hora || ''),
    produtos: JSON.stringify(payload.produtos || [])
  };

  aba.appendRow(rowToArray(obj));

  // Devolve o ID gerado para que o frontend possa usá-lo no card
  return { status: 'ok', msg: 'Inserido: ' + id, id: id };
}

// Atualiza um pedido existente (localiza pelo id)
function atualizar(payload) {
  var aba = getAba();
  var idx = findRowIndex(aba, payload.id);
  if (idx === -1) {
    return inserir(payload);
  }
  var obj = {
    id:       String(payload.id).padStart(5, '0'),
    filial:   payload.filial   || '',
    tipo:     payload.tipo     || '',
    situacao: payload.situacao || '',
    empresa:  payload.empresa  || '',
    cliente:  payload.cliente  || '',
    vendedor: payload.vendedor || '',
    endereco: payload.endereco || '',
    contato:  payload.contato  || '',
    obs:      payload.obs      || '',
    data:     String(payload.data || ''),
    hora:     String(payload.hora || ''),
    produtos: JSON.stringify(payload.produtos || [])
  };
  aba.getRange(idx, 1, 1, COLUNAS.length).setValues([rowToArray(obj)]);
  var colData = COLUNAS.indexOf('data') + 1;
  var colHora = COLUNAS.indexOf('hora') + 1;
  aba.getRange(idx, colData).setNumberFormat('@');
  aba.getRange(idx, colHora).setNumberFormat('@');
  return { status: 'ok', msg: 'Atualizado: ' + obj.id };
}

// Exclui um pedido pelo id
function excluir(payload) {
  var aba = getAba();
  var idx = findRowIndex(aba, payload.id);
  if (idx === -1) return { status: 'ok', msg: 'Não encontrado (já excluído?)' };
  aba.deleteRow(idx);
  return { status: 'ok', msg: 'Excluído: ' + payload.id };
}

// Retorna o próximo ID (mantido para compatibilidade, mas não é mais usado pelo fluxo principal)
function proximoId() {
  var aba  = getAba();
  var proximo = calcularProximoId(aba);
  return { status: 'ok', id: proximo };
}

// ---- Resposta JSON com headers CORS ----
function jsonResponse(obj) {
  var output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ============================================================
//  CONFIGURAÇÃO INICIAL DA PLANILHA
//  Execute esta função UMA VEZ manualmente para formatar a aba
//  Menu: Run > configurarPlanilha
// ============================================================
function configurarPlanilha() {
  var aba = getAba();

  var header = aba.getRange(1, 1, 1, COLUNAS.length);
  header.setFontWeight('bold');
  header.setBackground('#013025');
  header.setFontColor('#ffffff');

  var colData = COLUNAS.indexOf('data') + 1;
  var colHora = COLUNAS.indexOf('hora') + 1;
  aba.getRange(2, colData, 999, 1).setNumberFormat('@STRING@');
  aba.getRange(2, colHora, 999, 1).setNumberFormat('@STRING@');

  aba.setColumnWidth(1, 70);   // id
  aba.setColumnWidth(2, 130);  // filial
  aba.setColumnWidth(3, 110);  // tipo
  aba.setColumnWidth(4, 170);  // situacao
  aba.setColumnWidth(5, 160);  // empresa
  aba.setColumnWidth(6, 160);  // cliente
  aba.setColumnWidth(7, 140);  // vendedor  ← NOVA
  aba.setColumnWidth(8, 200);  // endereco
  aba.setColumnWidth(9, 130);  // contato
  aba.setColumnWidth(10, 200); // obs
  aba.setColumnWidth(11, 100); // data
  aba.setColumnWidth(12, 70);  // hora
  aba.setColumnWidth(13, 300); // produtos

  SpreadsheetApp.getUi().alert('Planilha configurada com sucesso!');
}

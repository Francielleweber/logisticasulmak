// ============================================================
//  SULMAK LOGÍSTICA — Google Apps Script
//  Cole este código em script.google.com e publique como
//  "App da Web" com acesso "Qualquer pessoa".
// ============================================================

// Nomes das abas
var ABA_ENTREGAS      = 'Entregas';
var ABA_SUBSTITUICOES = 'Substituições';
var ABA_DEVOLUCOES    = 'Devoluções';
var ABA_HISTORICO     = 'Histórico';

// Cabeçalhos de cada aba
var CABECALHOS = [
  'ID', 'Filial', 'Tipo', 'Situação', 'Cliente (Responsável)',
  'Empresa / Cliente', 'Endereço', 'Contato', 'Observações',
  'Data', 'Hora', 'Produtos'
];

// ------------------------------------------------------------
//  Ponto de entrada POST (chamado pelo sistema HTML)
// ------------------------------------------------------------
function doPost(e) {
  try {
    var dados = JSON.parse(e.postData.contents);
    var acao  = dados.acao;

    garantirAbas(); // cria abas/cabeçalhos se não existirem

    if (acao === 'test') {
      return respOk('Conexão OK');
    }

    if (acao === 'insert') {
      inserirLinha(dados);
      return respOk('Inserido');
    }

    if (acao === 'update') {
      // Remove a versão antiga (em qualquer aba) e insere a nova
      removerLinha(dados.id);
      inserirLinha(dados);
      return respOk('Atualizado');
    }

    if (acao === 'delete') {
      removerLinha(dados.id);
      return respOk('Excluído');
    }

    return respOk('Ação desconhecida: ' + acao);

  } catch (err) {
    return respErro(err.message);
  }
}

// Também responde a GET (para teste rápido no browser)
function doGet(e) {
  return respOk('Script SULMAK ativo!');
}

// ------------------------------------------------------------
//  Garante que todas as abas existam com cabeçalhos corretos
// ------------------------------------------------------------
function garantirAbas() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var abas = [ABA_ENTREGAS, ABA_SUBSTITUICOES, ABA_DEVOLUCOES, ABA_HISTORICO];

  abas.forEach(function (nomeAba) {
    var aba = ss.getSheetByName(nomeAba);

    if (!aba) {
      // Cria a aba
      aba = ss.insertSheet(nomeAba);
    }

    // Garante cabeçalhos na linha 1
    var primeiraLinha = aba.getRange(1, 1, 1, CABECALHOS.length).getValues()[0];
    var temCabecalho  = primeiraLinha[0] === CABECALHOS[0];

    if (!temCabecalho) {
      aba.getRange(1, 1, 1, CABECALHOS.length).setValues([CABECALHOS]);

      // Formata cabeçalho: fundo verde escuro, texto branco, negrito
      var rangeCab = aba.getRange(1, 1, 1, CABECALHOS.length);
      rangeCab.setBackground('#1a7a3c');
      rangeCab.setFontColor('#ffffff');
      rangeCab.setFontWeight('bold');
      rangeCab.setFontSize(10);

      // Congela a linha de cabeçalho
      aba.setFrozenRows(1);

      // Ajusta largura das colunas automaticamente
      aba.autoResizeColumns(1, CABECALHOS.length);
    }
  });
}

// ------------------------------------------------------------
//  Insere uma linha na aba correta conforme tipo/situação
// ------------------------------------------------------------
function inserirLinha(dados) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = abaParaDemanda(dados, ss);

  var produtos = '';
  if (dados.produtos && dados.produtos.length) {
    produtos = dados.produtos.map(function (p) {
      var partes = [p.nome];
      if (p.periodo)    partes.push(p.periodo);
      if (p.valor)      partes.push('R$ ' + p.valor);
      if (p.patrimonio) partes.push('Pat: ' + p.patrimonio);
      return partes.join(' | ');
    }).join('\n');
  }

  var linha = [
    dados.id        || '',
    dados.filial    || '',
    tipoLabel(dados.tipo),
    dados.situacao  || '',
    dados.cliente   || '',
    dados.empresa   || '',
    dados.endereco  || '',
    dados.contato   || '',
    dados.obs       || '',
    dados.data      || '',
    dados.hora      || '',
    produtos
  ];

  aba.appendRow(linha);

  // Formata a linha recém-inserida
  var ultimaLinha = aba.getLastRow();
  var rangeLinh   = aba.getRange(ultimaLinha, 1, 1, CABECALHOS.length);
  rangeLinh.setWrap(true);
  rangeLinh.setVerticalAlignment('top');

  // Cor de fundo alternada (linhas pares: branco, ímpares: cinza clarinho)
  if (ultimaLinha % 2 === 0) {
    rangeLinh.setBackground('#f9f9f9');
  } else {
    rangeLinh.setBackground('#ffffff');
  }
}

// ------------------------------------------------------------
//  Remove todas as linhas com o ID informado (em todas as abas)
// ------------------------------------------------------------
function removerLinha(id) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var abas = [ABA_ENTREGAS, ABA_SUBSTITUICOES, ABA_DEVOLUCOES, ABA_HISTORICO];

  abas.forEach(function (nomeAba) {
    var aba = ss.getSheetByName(nomeAba);
    if (!aba) return;

    var dados      = aba.getDataRange().getValues();
    // Percorre de baixo para cima para não deslocar índices
    for (var i = dados.length - 1; i >= 1; i--) {
      if (String(dados[i][0]) === String(id)) {
        aba.deleteRow(i + 1); // +1 porque getValues é base-0, Sheets é base-1
      }
    }
  });
}

// ------------------------------------------------------------
//  Escolhe a aba destino:
//  Finalizada/Cancelada → Histórico
//  Entrega              → Entregas
//  Substituição         → Substituições
//  Devolução            → Devoluções
// ------------------------------------------------------------
function abaParaDemanda(dados, ss) {
  var sit = dados.situacao || '';
  if (sit === 'Finalizada' || sit === 'Cancelada') {
    return ss.getSheetByName(ABA_HISTORICO);
  }
  var tipo = dados.tipo || '';
  if (tipo === 'entrega')      return ss.getSheetByName(ABA_ENTREGAS);
  if (tipo === 'substituicao') return ss.getSheetByName(ABA_SUBSTITUICOES);
  if (tipo === 'devolucao')    return ss.getSheetByName(ABA_DEVOLUCOES);
  return ss.getSheetByName(ABA_ENTREGAS); // fallback
}

// ------------------------------------------------------------
//  Helpers
// ------------------------------------------------------------
function tipoLabel(tipo) {
  if (tipo === 'entrega')      return 'Entrega';
  if (tipo === 'substituicao') return 'Substituição';
  if (tipo === 'devolucao')    return 'Devolução';
  return tipo || '';
}

function respOk(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', msg: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

function respErro(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'erro', msg: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

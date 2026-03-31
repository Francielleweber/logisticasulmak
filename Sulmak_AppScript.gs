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
//  Ponto de entrada GET — retorna todos os dados para o front
// ------------------------------------------------------------
function doGet(e) {
  // ?action=dados  →  retorna JSON com todas as demandas ativas
  if (e && e.parameter && e.parameter.action === 'dados') {
    try {
      garantirAbas();
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var abas = [
        { nome: ABA_ENTREGAS,      tipo: 'entrega' },
        { nome: ABA_SUBSTITUICOES, tipo: 'substituicao' },
        { nome: ABA_DEVOLUCOES,    tipo: 'devolucao' }
      ];

      var resultado = [];

      abas.forEach(function(a) {
        var aba = ss.getSheetByName(a.nome);
        if (!aba || aba.getLastRow() < 2) return;

        var rows = aba.getRange(2, 1, aba.getLastRow() - 1, CABECALHOS.length).getValues();
        rows.forEach(function(row) {
          var id = String(row[0]).trim();
          if (!id) return;

          // Reconstrói array de produtos a partir da string
          var prodStr = String(row[11] || '').trim();
          var produtos = [];
          if (prodStr) {
            prodStr.split('\n').forEach(function(linha) {
              var partes = linha.split(' | ');
              var prod = { nome: partes[0] || '' };
              partes.slice(1).forEach(function(p) {
                if (p.startsWith('Pat: '))      prod.patrimonio       = p.replace('Pat: ', '');
                else if (p.startsWith('Ret: ')) prod.patrimonioRetorna = p.replace('Ret: ', '');
                else if (p.startsWith('Novo: '))prod.patrimonioNovo   = p.replace('Novo: ', '');
                else if (p.startsWith('R$ '))   prod.valor            = p.replace('R$ ', '');
                else                            prod.periodo          = p;
              });
              if (prod.nome) produtos.push(prod);
            });
          }

          resultado.push({
            id:        id,
            filial:    String(row[1] || '').trim(),
            tipo:      a.tipo,
            situacao:  String(row[3] || '').trim(),
            cliente:   String(row[4] || '').trim(),
            empresa:   String(row[5] || '').trim(),
            endereco:  String(row[6] || '').trim(),
            contato:   String(row[7] || '').trim(),
            obs:       String(row[8] || '').trim(),
            data:      String(row[9] || '').trim(),
            hora:      String(row[10] || '').trim(),
            produtos:  produtos
          });
        });
      });

      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', dados: resultado }))
        .setMimeType(ContentService.MimeType.JSON);

    } catch(err) {
      return respErro(err.message);
    }
  }

  // GET sem parâmetro → teste simples
  return respOk('Script SULMAK ativo!');
}

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

// ------------------------------------------------------------
//  Garante que todas as abas existam com cabeçalhos corretos
// ------------------------------------------------------------
function garantirAbas() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var abas = [ABA_ENTREGAS, ABA_SUBSTITUICOES, ABA_DEVOLUCOES, ABA_HISTORICO];

  abas.forEach(function (nomeAba) {
    var aba = ss.getSheetByName(nomeAba);

    if (!aba) {
      aba = ss.insertSheet(nomeAba);
    }

    var primeiraLinha = aba.getRange(1, 1, 1, CABECALHOS.length).getValues()[0];
    var temCabecalho  = primeiraLinha[0] === CABECALHOS[0];

    if (!temCabecalho) {
      aba.getRange(1, 1, 1, CABECALHOS.length).setValues([CABECALHOS]);

      var rangeCab = aba.getRange(1, 1, 1, CABECALHOS.length);
      rangeCab.setBackground('#1a7a3c');
      rangeCab.setFontColor('#ffffff');
      rangeCab.setFontWeight('bold');
      rangeCab.setFontSize(10);

      aba.setFrozenRows(1);
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
      if (p.periodo)            partes.push(p.periodo);
      if (p.valor)              partes.push('R$ ' + p.valor);
      if (p.patrimonio)         partes.push('Pat: ' + p.patrimonio);
      if (p.patrimonioRetorna)  partes.push('Ret: ' + p.patrimonioRetorna);
      if (p.patrimonioNovo)     partes.push('Novo: ' + p.patrimonioNovo);
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

  var ultimaLinha = aba.getLastRow();
  var rangeLinh   = aba.getRange(ultimaLinha, 1, 1, CABECALHOS.length);
  rangeLinh.setWrap(true);
  rangeLinh.setVerticalAlignment('top');

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
    for (var i = dados.length - 1; i >= 1; i--) {
      if (String(dados[i][0]) === String(id)) {
        aba.deleteRow(i + 1);
      }
    }
  });
}

// ------------------------------------------------------------
//  Escolhe a aba destino
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
  return ss.getSheetByName(ABA_ENTREGAS);
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

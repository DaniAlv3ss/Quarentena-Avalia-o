/**
 * ==========================================
 * AUDITOR PRO - SERVER SIDE
 * ==========================================
 */

const CONFIG = {
  ID_PLANILHA_DADOS: '10l1w3d3HYSKFgSsnjOZ545efR-bdIECEkOR82IjV3TE', // Planilha Divergência
  NOME_ABA_DADOS: 'Divergência',
  ID_PLANILHA_CUSTOS: '1CEexqCPUyP5b4Qra1tt5qWyBIUW0lfIoYIEvYgBFhtA', // Base de CFs
  NOME_ABA_CUSTOS: 'Base de CFs',
  NOME_ABA_HISTORICO: 'Histórico de Auditorias',
  // NOVA PLANILHA DE TREINAMENTO (Sua base histórica)
  ID_PLANILHA_TREINO: '1tUjt4G45BrRUxn-peLBVR7l6w7td6nDf8rD5hvoDGi8',
  NOME_ABA_TREINO: 'Trei_Hist'
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Auditor Pro - Gestão de Avarias')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- FUNÇÕES DE DADOS ---

function buscarDadosPorListaOS(listaOS) {
  if (!listaOS || listaOS.length === 0) return [];
  const dados = buscarDadosDivergencia();
  const dadosProdutos = carregarMapaPrecos();
  const setOS = new Set(listaOS.map(os => String(os).trim().toUpperCase()));

  const resultado = dados.filter(row => {
    let os = normalizarOS(row);
    return setOS.has(os);
  });

  return enriquecerComCustos(resultado, dadosProdutos);
}

function buscarDadosPorPeriodo(ini, fim) {
  const dados = buscarDadosDivergencia();
  const dadosProdutos = carregarMapaPrecos();
  
  const dtIni = new Date(ini + 'T00:00:00');
  const dtFim = new Date(fim + 'T23:59:59');

  const resultado = dados.filter(row => {
    const valData = row['_idx_data'] || row['CARIMBO DE DATA/HORA']; 
    const dataRow = converterDataPtBr(valData);
    return dataRow && dataRow >= dtIni && dataRow <= dtFim;
  });

  return enriquecerComCustos(resultado, dadosProdutos);
}

// --- INTEGRAÇÃO COM IA E HISTÓRICO ---

function carregarHistoricoTreinamento() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA_TREINO);
    const aba = ss.getSheetByName(CONFIG.NOME_ABA_TREINO);
    if (!aba) return [];

    const dados = aba.getDataRange().getDisplayValues();
    if (dados.length < 2) return [];

    // Mapeamento das colunas da sua planilha Trei_Hist
    // A: OS, B: Data, C: Avaria Customiza, D: Processo, E: Motivos, F: Detalhe, G: Considerações (Decisão)
    const dataset = [];
    
    // Normalizador de labels (Humanos -> Código do Sistema)
    const mapLabel = (txt) => {
      const t = String(txt).toLowerCase().trim();
      if (t.includes('reprovado') || t.includes('não pagar')) return 'nao_pagar';
      if (t.includes('aprovado') || t.includes('pagar') || t.includes('ok')) return 'ok_pagamento';
      if (t.includes('parceiro') || t.includes('avaria')) return 'avaria_parceiro';
      return null;
    };

    for (let i = 1; i < dados.length; i++) {
      const row = dados[i];
      const motivo = row[4]; // Coluna E
      const detalhe = row[5]; // Coluna F
      const decisaoRaw = row[6]; // Coluna G (Considerações)
      
      const textoCompleto = `${motivo} ${detalhe}`.trim();
      const label = mapLabel(decisaoRaw);

      if (textoCompleto.length > 3 && label) {
        dataset.push({
          texto: textoCompleto,
          categoria: label
        });
      }
    }
    return dataset;
  } catch (e) {
    console.error("Erro ao carregar treino: " + e);
    return [];
  }
}

// --- PERSISTÊNCIA E HISTÓRICO ---

function listarSessoesSalvas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(CONFIG.NOME_ABA_HISTORICO);
  if (!aba) return [];
  const dados = aba.getDataRange().getValues();
  if (dados.length < 2) return [];
  const idColIndex = 12; 
  const sessoes = new Set();
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][idColIndex]) sessoes.add(dados[i][idColIndex]);
  }
  return Array.from(sessoes).reverse();
}

function carregarSessaoAntiga(idSessao) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(CONFIG.NOME_ABA_HISTORICO);
  if (!aba) throw new Error("Histórico não encontrado.");
  const dados = aba.getDataRange().getDisplayValues();
  if (dados.length < 2) return [];
  const idColIndex = 12;
  const resultado = [];
  
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][idColIndex] === idSessao) {
      const r = dados[i];
      resultado.push({
        OS: r[0],
        'NOME DO TÉCNICO': r[1],
        'QUAL PROBLEMA?': r[2], 
        'FOTO DA AVARIA!': r[3],
        'CF PRODUTO ANTIGO': r[4],
        _custoAntigo: parseMoney(r[5]),
        'CF PRODUTO NOVO': r[6],
        _custoNovo: parseMoney(r[7]),
        _custoTotal: parseMoney(r[8]),
        _sugestao: r[9] === '-' ? null : r[9],
        _status: r[10],
        _dataAuditoria: r[11]
      });
    }
  }
  return resultado;
}

function salvarSessaoNaPlanilha(nomeSessao, dados, isUpdate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(CONFIG.NOME_ABA_HISTORICO);
  
  const headers = [
    "OS", "Técnico", "Problema/Detalhe", "Evidência", 
    "CF Antigo", "Custo Antigo", "CF Novo", "Custo Novo", "Total Geral",
    "Sugestão IA", "Decisão", "Data Auditoria", "ID Sessão"
  ];

  if (!aba) {
    aba = ss.insertSheet(CONFIG.NOME_ABA_HISTORICO);
    aba.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold").setBackground("#FF6500").setFontColor("#ffffff");
    aba.setFrozenRows(1);
  }

  if (isUpdate) {
    const lastRow = aba.getLastRow();
    if (lastRow > 1) {
      const rangeIds = aba.getRange(2, 13, lastRow - 1, 1).getValues();
      for (let i = rangeIds.length - 1; i >= 0; i--) {
        if (rangeIds[i][0] === nomeSessao) aba.deleteRow(i + 2);
      }
    }
  }

  const linhas = dados.map(item => {
    const subs = Array.isArray(item._subRecords) ? item._subRecords : [];
    const problemasStr = subs.length > 0
      ? subs.map(s => `[${s.id}] ${s.problema}: ${s.detalhe}`).join(" | ")
      : (item['QUAL PROBLEMA?'] || '');

    return [
      item.OS || item['NÚMERO OS'] || 'N/A',
      item['NOME DO TÉCNICO'] || '',
      problemasStr,
      item['FOTO DA AVARIA!'] || '',
      item['CF PRODUTO ANTIGO'] || '',
      item._custoAntigo || 0,
      item['CF PRODUTO NOVO'] || '',
      item._custoNovo || 0,
      item._custoTotal || 0,
      item._sugestao || '-',
      item._status, 
      new Date(),
      nomeSessao
    ];
  });

  if (linhas.length > 0) {
    const lastRow = aba.getLastRow();
    const targetRow = Math.max(lastRow + 1, 2);
    aba.getRange(targetRow, 1, linhas.length, headers.length).setValues(linhas);
    aba.getRange(targetRow, 6, linhas.length, 1).setNumberFormat("R$ #,##0.00");
    aba.getRange(targetRow, 8, linhas.length, 2).setNumberFormat("R$ #,##0.00");
    aba.getRange(targetRow, 12, linhas.length, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
  }
  return { success: true, nomeAba: nomeSessao };
}

// --- FUNÇÃO AJUSTADA PARA O NOVO FORMATO DE TREINAMENTO ---
function salvarAprendizado(dados) {
  // dados = { os, dataInicio, avariaCust, processo, motivo, detalhe, decisao, custoNovo }
  try {
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA_TREINO);
    const aba = ss.getSheetByName(CONFIG.NOME_ABA_TREINO);
    
    if(aba) {
      // Mapeamento de decisão (Sistema -> Humano)
      let labelHumano = 'Em Análise';
      if(dados.decisao === 'ok_pagamento') labelHumano = 'Aprovado Pagamento';
      if(dados.decisao === 'nao_pagar') labelHumano = 'Reprovado Pagamento';
      if(dados.decisao === 'avaria_parceiro') labelHumano = 'Avaria Parceiro';
      
      // Formata custo para R$ se for numérico
      // Se preferir salvar como número puro no Excel e formatar a coluna lá, apenas salve dados.custoNovo
      const valorFormatado = dados.custoNovo 
        ? parseFloat(dados.custoNovo) 
        : 0;

      // ESTRUTURA SOLICITADA:
      // A: Número OS
      // B: Data Início Montagem
      // C: Avaria customiza?
      // D: Processo?
      // E: Motivos
      // F: Detalhe o problema encontrado
      // G: Considerações
      // H: Custo CF Novo
      
      const novaLinha = [
        dados.os || '',
        dados.dataInicio || '', // Salva string como vem (DD/MM/YYYY) ou Date object
        dados.avariaCust || '',
        dados.processo || '',
        dados.motivo || '',
        dados.detalhe || '',
        labelHumano,
        valorFormatado
      ];
      
      aba.appendRow(novaLinha);
      
      // Opcional: Formatar a última célula adicionada como Moeda
      const lastRow = aba.getLastRow();
      if(lastRow > 0) {
         aba.getRange(lastRow, 8).setNumberFormat("R$ #,##0.00");
      }
    }
  } catch(e) {
    console.error("Erro ao salvar aprendizado: " + e);
  }
}

// --- HELPERS (LEITURA CRÍTICA) ---

function buscarDadosDivergencia() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA_DADOS);
    const aba = ss.getSheetByName(CONFIG.NOME_ABA_DADOS);
    if (!aba) throw new Error("Aba Divergência não encontrada");
    
    const lr = aba.getLastRow();
    const lc = aba.getLastColumn();
    if (lr < 2) return [];
    
    const range = aba.getRange(2, 1, lr - 1, lc);
    const values = range.getDisplayValues();
    
    const headers = values[0].map(h => String(h).trim().toUpperCase());
    
    return values.slice(1).map(r => {
      let obj = {};
      headers.forEach((h, i) => obj[h] = r[i]);

      obj['_idx_data'] = r[0];         
      obj['_idx_tecnico'] = r[1];      
      obj['_idx_processo'] = r[2];     
      obj['_idx_problema'] = r[3];     
      obj['_idx_detalhe'] = r[4];      
      obj['_idx_os'] = r[5];           
      obj['_idx_inicio'] = r[6];       
      obj['_idx_fim'] = r[7];          
      obj['_idx_pedido'] = r[8];       
      obj['_idx_prod_div'] = r[9];     
      obj['_idx_tipo_troca'] = r[10];  
      obj['_idx_avaria_cust'] = r[11]; 
      
      const idxFoto = headers.findIndex(h => h.includes('FOTO'));
      if(idxFoto > -1) obj['FOTO DA AVARIA!'] = r[idxFoto];

      const idxCFAnt = headers.findIndex(h => h.includes('CF PRODUTO ANTIGO'));
      const idxCFNov = headers.findIndex(h => h.includes('CF PRODUTO NOVO'));
      if(idxCFAnt > -1) obj['CF PRODUTO ANTIGO'] = r[idxCFAnt];
      if(idxCFNov > -1) obj['CF PRODUTO NOVO'] = r[idxCFNov];

      return obj;
    });
  } catch (e) {
    console.error(e);
    throw e.message;
  }
}

function carregarMapaPrecos() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA_CUSTOS);
    const aba = ss.getSheetByName(CONFIG.NOME_ABA_CUSTOS);
    if (!aba) return {};
    const vals = aba.getDataRange().getDisplayValues();
    if (vals.length < 2) return {};
    
    const headers = vals[0].map(h => String(h).trim().toUpperCase());
    const idxCodigo = headers.findIndex(h => h.includes('CÓDIGO') || h.includes('CODIGO'));
    const idxValor = headers.findIndex(h => h.includes('VALOR'));

    const map = {};
    for (let i = 1; i < vals.length; i++) {
      const r = vals[i];
      if (r[idxCodigo] && r[idxValor]) {
        map[String(r[idxCodigo]).trim().toUpperCase()] = {
          preco: parseMoney(r[idxValor])
        };
      }
    }
    return map;
  } catch (e) { return {}; }
}

function enriquecerComCustos(dados, dadosProdutos) {
  dados.forEach(row => {
    const cfAnt = String(row['CF PRODUTO ANTIGO']||'').trim().toUpperCase();
    const cfNov = String(row['CF PRODUTO NOVO']||'').trim().toUpperCase();
    
    const pAnt = dadosProdutos[cfAnt] || { preco: 0 };
    const pNov = dadosProdutos[cfNov] || { preco: 0 };
    
    row._custoAntigo = pAnt.preco;
    row._custoNovo = pNov.preco;
    // ALTERAÇÃO: Soma apenas o valor do produto NOVO no custo total da troca
    row._custoTotal = pNov.preco; 
  });
  return dados;
}

function normalizarOS(row) {
  if (row['_idx_os']) return String(row['_idx_os']).trim().toUpperCase();
  return String(row['OS'] || row['NÚMERO OS'] || row['ID'] || '').trim().toUpperCase();
}

function converterDataPtBr(dateStr) {
  if (!dateStr) return null;
  const parts = dateStr.split(' ')[0].split('/'); 
  if (parts.length === 3) return new Date(parts[2], parts[1]-1, parts[0]);
  return null;
}

function parseMoney(val) {
  return parseFloat(String(val).replace(/[^\d,-]/g, '').replace(',', '.')) || 0;
}

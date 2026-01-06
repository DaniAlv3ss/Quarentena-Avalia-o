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
  NOME_ABA_HISTORICO: 'Histórico de Auditorias'
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
    // Tenta ler do índice fixo ou do nome da coluna
    const valData = row['_idx_data'] || row['CARIMBO DE DATA/HORA']; 
    const dataRow = converterDataPtBr(valData);
    return dataRow && dataRow >= dtIni && dataRow <= dtFim;
  });

  return enriquecerComCustos(resultado, dadosProdutos);
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
      // Reconstrói o objeto com os campos que salvamos
      // Nota: Ao carregar histórico, alguns campos detalhados do modal original podem não existir se não foram salvos na aba histórica.
      // O histórico salva o "resumo". Se precisar dos detalhes originais ao recarregar, teríamos que salvar mais colunas.
      // Por enquanto, mantemos a estrutura de visualização.
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

// --- APRENDIZADO DE MÁQUINA ---

function buscarBaseConhecimento() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('Treinamento');
  if (!aba) return [];
  const dados = aba.getDataRange().getValues();
  if (dados.length < 2) return [];
  return dados.slice(1).map(r => ({ padrao: r[0], decisao: r[1] }));
}

function salvarAprendizado(texto, decisao) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName('Treinamento');
  if (!aba) {
    aba = ss.insertSheet('Treinamento');
    aba.appendRow(["Texto Padrão", "Decisão", "Data"]);
  }
  aba.appendRow([texto, decisao, new Date()]);
  return { success: true };
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
    
    // Pega todos os dados a partir da linha 2
    const range = aba.getRange(2, 1, lr - 1, lc);
    const values = range.getDisplayValues();
    
    // Mapeamento dinâmico (nomes) + Mapeamento fixo (índices solicitados)
    const headers = values[0].map(h => String(h).trim().toUpperCase());
    
    return values.slice(1).map(r => {
      let obj = {};
      
      // 1. Mapeia por nome do cabeçalho (padrão antigo)
      headers.forEach((h, i) => obj[h] = r[i]);

      // 2. Mapeia por ÍNDICE FIXO (Segurança Máxima conforme solicitado)
      // Índices fornecidos (0-based da linha da planilha):
      // Coluna A = 0.
      obj['_idx_data'] = r[0];         // Carimbo data/hora
      obj['_idx_tecnico'] = r[1];      // Nome Técnico
      obj['_idx_processo'] = r[2];     // Processo?
      obj['_idx_problema'] = r[3];     // Qual Problema?
      obj['_idx_detalhe'] = r[4];      // Detalhe o problema
      obj['_idx_os'] = r[5];           // OS
      obj['_idx_inicio'] = r[6];       // Inicio tratativa
      obj['_idx_fim'] = r[7];          // Fim tratativa
      obj['_idx_pedido'] = r[8];       // Pedido
      obj['_idx_prod_div'] = r[9];     // Produto Divergente
      obj['_idx_tipo_troca'] = r[10];  // Tipo de Troca?
      obj['_idx_avaria_cust'] = r[11]; // Avaria customiza?
      
      // Colunas extras úteis
      // Acha coluna de foto pelo nome aproximado caso mude de lugar
      const idxFoto = headers.findIndex(h => h.includes('FOTO'));
      if(idxFoto > -1) obj['FOTO DA AVARIA!'] = r[idxFoto];

      // Acha colunas de CF
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
    row._custoTotal = pAnt.preco + pNov.preco;
  });
  return dados;
}

function normalizarOS(row) {
  // Prioriza o índice fixo se existir
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

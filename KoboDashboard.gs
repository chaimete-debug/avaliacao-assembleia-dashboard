// =============================================================================
//  27ª Assembleia Distrital – Matola | KoboToolbox → Google Sheets Dashboard
//  GitHub: https://github.com/YOUR_ORG/kobo-assembleia-dashboard
//
//  INSTALAÇÃO:
//  1. Abre um Google Sheets novo
//  2. Extensions → Apps Script → apaga o código existente → cola este ficheiro
//  3. Guarda (Ctrl+S)
//  4. Selecciona syncKobo no dropdown → clica ▶ Run → aceita permissões
//  5. O menu "📊 Kobo Dashboard" aparece no teu Sheets
// =============================================================================

// ── CONFIGURAÇÃO ──────────────────────────────────────────────────────────────
const CONFIG = {
  ASSET_UID : 'aQgBUEPn9wyVBAbWJPUbSQ',
  TOKEN     : '4c6c1683faa457dfc69c8fc743d3374d8b4fd800',
  BASE_URL  : 'https://eu.kobotoolbox.org/api/v2',
  PAGE_SIZE : 30000,
  EVENT_NAME: '27ª Assembleia Distrital – Matola',
};

// ── CORES ─────────────────────────────────────────────────────────────────────
const C = {
  BLUE_DARK  : '#1A3A6B', BLUE_MID   : '#2E6FAD', BLUE_LIGHT : '#D6E8F7',
  GOLD       : '#C9A227', GREEN      : '#1A7A3C', GREEN_LIGHT: '#E6F4EC',
  RED        : '#B71C1C', RED_LIGHT  : '#FDECEA', AMBER      : '#F0B429',
  AMBER_LIGHT: '#FFF3CD', GRAY_DARK  : '#444444', GRAY_MID   : '#888888',
  GRAY_LIGHT : '#F5F7FA', WHITE      : '#FFFFFF',
};

// ── CAMPOS DO FORMULÁRIO ──────────────────────────────────────────────────────
// KoboToolbox devolve campos com prefixo de grupo: "grupo/campo"
// O script normaliza automaticamente para só o nome do campo.
const INDICATORS = [
  { key: 'aval_credenciamento', label: 'Processo de Credenciamento'        },
  { key: 'aval_local',          label: 'Condições do Local (Conforto/Som)' },
  { key: 'aval_alimentacao',    label: 'Qualidade da Alimentação'          },
  { key: 'aval_pontualidade',   label: 'Cumprimento do Horário'            },
  { key: 'aval_temas',          label: 'Relevância dos Temas'              },
  { key: 'aval_metodologia',    label: 'Metodologia Adotada'               },
];

const SCALE   = { '5':'Excelente','4':'Bom','3':'Satisfatório','2':'Mau','1':'Muito Mau' };
const PERFIS  = { delegado:'Delegado', convidado:'Convidado', staff:'Staff/Secretariado', observador:'Observador' };
const GENEROS = { m:'Masculino', f:'Feminino' };

// =============================================================================
//  MENU
// =============================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Kobo Dashboard')
    .addItem('🔄  Sincronizar dados do KoboToolbox', 'syncKobo')
    .addItem('📊  Reconstruir Dashboard',            'buildDashboard')
    .addItem('🗑️   Limpar tudo e re-sincronizar',    'fullReset')
    .addToUi();
}

// =============================================================================
//  ENTRY POINTS
// =============================================================================
function syncKobo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  toast_(ss, '⏳ A ligar ao KoboToolbox…');
  const raw = fetchAllSubmissions_();
  toast_(ss, `✅ ${raw.length} respostas recebidas. A escrever…`);
  writeRawSheet_(ss, raw);
  writeMetaSheet_(ss, raw.length);
  toast_(ss, '📊 A construir o Dashboard…');
  buildDashboard();
  toast_(ss, `✅ Concluído! ${raw.length} respostas sincronizadas.`);
}

function fullReset() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['Raw_Kobo', 'Kobo_Meta', 'Dashboard'].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });
  syncKobo();
}

// =============================================================================
//  KOBO API — paginação completa
// =============================================================================
function fetchAllSubmissions_() {
  const all  = [];
  let   next = `${CONFIG.BASE_URL}/assets/${CONFIG.ASSET_UID}/data/?format=json&limit=${CONFIG.PAGE_SIZE}&start=0`;

  while (next) {
    const resp = UrlFetchApp.fetch(next, {
      headers          : { Authorization: `Token ${CONFIG.TOKEN}` },
      muteHttpExceptions: true,
    });
    if (resp.getResponseCode() !== 200)
      throw new Error(`KoboToolbox ${resp.getResponseCode()}: ${resp.getContentText().slice(0,200)}`);

    const json = JSON.parse(resp.getContentText());
    (json.results || []).forEach(r => all.push(normalise_(r)));
    next = json.next || null;
  }
  return all;
}

// ── Normalisa chaves: remove prefixos de grupo ("grupo/campo" → "campo") ──────
function normalise_(submission) {
  const out = {};
  Object.entries(submission).forEach(([k, v]) => {
    // Preserva campos de sistema com underscore inicial (_id, _uuid, etc.)
    const clean = k.startsWith('_') ? k : k.split('/').pop();
    out[clean]  = (v === null || v === undefined) ? '' : String(v).trim();
  });
  return out;
}

// =============================================================================
//  RAW SHEET
// =============================================================================
function writeRawSheet_(ss, data) {
  let sh = ss.getSheetByName('Raw_Kobo');
  if (!sh) sh = ss.insertSheet('Raw_Kobo');
  sh.clearContents();
  sh.clearFormats();

  if (!data.length) {
    sh.getRange(1,1).setValue('Sem dados recebidos.');
    return;
  }

  // Ordenar colunas: campos conhecidos primeiro, depois o resto
  const priority = [
    'perfil_participante','genero','proveniencia',
    ...INDICATORS.map(i => i.key),
    'espaco_opiniao','pontos_fortes','pontos_melhorar',
    '_id','_uuid','_submission_time',
  ];
  const allKeys = Array.from(new Set([...priority, ...Object.keys(data[0])]));
  const headers = allKeys.filter(k => data.some(r => r[k] !== undefined && r[k] !== ''));

  const rows = [headers, ...data.map(r => headers.map(h => r[h] !== undefined ? r[h] : ''))];
  sh.getRange(1, 1, rows.length, headers.length).setValues(rows);

  // Estilo cabeçalho
  sh.getRange(1, 1, 1, headers.length)
    .setBackground(C.BLUE_DARK).setFontColor(C.WHITE)
    .setFontWeight('bold').setFontSize(10).setFrozenRows(1);
  sh.setFrozenRows(1);

  // Linhas alternadas
  for (let r = 2; r <= rows.length; r++) {
    sh.getRange(r, 1, 1, headers.length)
      .setBackground(r % 2 === 0 ? C.GRAY_LIGHT : C.WHITE)
      .setFontSize(9);
  }

  // Largura automática (max 240px)
  headers.forEach((_, i) => sh.setColumnWidth(i + 1, Math.min(240, 80 + headers[i].length * 3)));
}

// =============================================================================
//  META SHEET
// =============================================================================
function writeMetaSheet_(ss, count) {
  let sh = ss.getSheetByName('Kobo_Meta');
  if (!sh) sh = ss.insertSheet('Kobo_Meta');
  sh.clearContents();
  sh.clearFormats();

  const now  = new Date();
  const rows = [
    ['Campo',                 'Valor'],
    ['Evento',                CONFIG.EVENT_NAME],
    ['Asset UID',             CONFIG.ASSET_UID],
    ['Servidor',              'eu.kobotoolbox.org'],
    ['Total de respostas',    count],
    ['Última sincronização',  Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')],
    ['Versão do script',      '2.0'],
  ];

  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1, 1, 2).setBackground(C.BLUE_DARK).setFontColor(C.WHITE).setFontWeight('bold');
  for (let r = 2; r <= rows.length; r++) {
    sh.getRange(r, 1).setFontWeight('bold').setBackground(C.GRAY_LIGHT);
    sh.getRange(r, 2).setBackground(C.WHITE);
  }
  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 320);
}

// =============================================================================
//  DASHBOARD
// =============================================================================
function buildDashboard() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName('Raw_Kobo');
  if (!raw) { SpreadsheetApp.getUi().alert('Execute primeiro "Sincronizar dados".'); return; }

  let dash = ss.getSheetByName('Dashboard');
  if (!dash) { dash = ss.insertSheet('Dashboard'); ss.moveActiveSheet(1); }
  dash.clearContents();
  dash.clearFormats();
  dash.setTabColor(C.BLUE_MID);

  const data = readNormalisedData_(raw);

  // ── DEBUG: se não há dados, mostrar aviso claro ──────────────────────────
  if (!data.length) {
    dash.getRange('B2').setValue('⚠️  Sem dados na aba Raw_Kobo. Execute "Sincronizar dados" primeiro.');
    return;
  }

  // ── DEBUG: verificar se os campos foram encontrados ──────────────────────
  const debugInfo = debugFields_(data);
  if (debugInfo.missing.length > 0) {
    Logger.log('CAMPOS NÃO ENCONTRADOS: ' + debugInfo.missing.join(', '));
    Logger.log('CAMPOS DISPONÍVEIS: '      + debugInfo.available.join(', '));
  }

  // Larguras das colunas (A=margem, B..S=conteúdo)
  dash.setColumnWidth(1, 18);
  [130,90,90,90,90,130,90,90,90,130,90,90,90,90,90,90,90,90].forEach((w,i) => dash.setColumnWidth(i+2, w));

  let row = 2;
  row = writeBanner_(dash, row);
  row = writeSyncInfo_(dash, row, ss, data.length);
  row++;

  row = writeSectionHeader_(dash, row, '📊  INDICADORES GLOBAIS');
  row = writeKpiCards_(dash, row, data);
  row += 2;

  row = writeSectionHeader_(dash, row, '📋  AVALIAÇÃO POR INDICADOR  (média 1–5)');
  row = writeIndicatorTable_(dash, row, data);
  row += 2;

  row = writeSectionHeader_(dash, row, '📊  DISTRIBUIÇÃO DE RESPOSTAS POR INDICADOR');
  row = writeDistributionTable_(dash, row, data);
  row += 2;

  row = writeSectionHeader_(dash, row, '👥  PERFIL DOS PARTICIPANTES');
  row = writeParticipantProfile_(dash, row, data);
  row += 2;

  row = writeSectionHeader_(dash, row, '🗣️  ESPAÇO DE OPINIÃO');
  row = writeOpinionSpace_(dash, row, data);
  row += 2;

  row = writeSectionHeader_(dash, row, '💬  COMENTÁRIOS ABERTOS');
  row = writeComments_(dash, row, data);
  row++;

  writeFiltersNote_(dash, row);

  ss.setActiveSheet(dash);
}

// =============================================================================
//  LER DADOS NORMALIZADOS DA RAW_KOBO
// =============================================================================
function readNormalisedData_(sh) {
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const headers = vals[0].map(h => String(h).trim());
  return vals.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = String(row[i] || '').trim(); });
    return obj;
  });
}

// ── DEBUG: verifica quais campos estão presentes ──────────────────────────────
function debugFields_(data) {
  if (!data.length) return { missing: [], available: [] };
  const available = Object.keys(data[0]);
  const needed    = [...INDICATORS.map(i => i.key), 'perfil_participante','genero','proveniencia','espaco_opiniao','pontos_fortes','pontos_melhorar'];
  const missing   = needed.filter(k => !available.includes(k));
  return { missing, available };
}

// =============================================================================
//  BANNER
// =============================================================================
function writeBanner_(sh, row) {
  sh.setRowHeight(row, 50);
  sh.getRange(row, 2, 1, 17).merge()
    .setValue(CONFIG.EVENT_NAME.toUpperCase())
    .setBackground(C.BLUE_DARK).setFontColor(C.WHITE)
    .setFontSize(20).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  row++;

  sh.setRowHeight(row, 26);
  sh.getRange(row, 2, 1, 17).merge()
    .setValue('Relatório Dinâmico de Avaliação')
    .setBackground(C.BLUE_MID).setFontColor(C.WHITE)
    .setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle');
  row++;

  sh.setRowHeight(row, 5);
  sh.getRange(row, 2, 1, 17).merge().setBackground(C.GOLD);
  row++;
  return row;
}

// =============================================================================
//  INFO SYNC
// =============================================================================
function writeSyncInfo_(sh, row, ss, count) {
  const meta     = ss.getSheetByName('Kobo_Meta');
  const lastSync = meta ? meta.getRange(6, 2).getValue() : '—';
  sh.setRowHeight(row, 20);
  sh.getRange(row, 2, 1, 17).merge()
    .setValue(`Última sincronização: ${lastSync}   |   Respostas totais: ${count}`)
    .setFontSize(9).setFontColor(C.GRAY_MID).setHorizontalAlignment('right');
  return row + 1;
}

// =============================================================================
//  SECÇÃO HEADER
// =============================================================================
function writeSectionHeader_(sh, row, label) {
  sh.setRowHeight(row, 28);
  sh.getRange(row, 2, 1, 17).merge()
    .setValue(label)
    .setBackground(C.BLUE_DARK).setFontColor(C.WHITE)
    .setFontSize(10).setFontWeight('bold')
    .setVerticalAlignment('middle');
  return row + 1;
}

// =============================================================================
//  KPI CARDS
// =============================================================================
function writeKpiCards_(sh, row, data) {
  const n = data.length;

  // Média global
  let s = 0, c = 0;
  INDICATORS.forEach(ind => data.forEach(r => {
    const v = parseFloat(r[ind.key]);
    if (v >= 1 && v <= 5) { s += v; c++; }
  }));
  const globalAvg = c > 0 ? s / c : 0;

  // Espaço de opinião
  const simCount = data.filter(r => r['espaco_opiniao'] === 'sim').length;
  const simPct   = n > 0 ? Math.round(simCount / n * 100) : 0;

  // Melhor indicador
  let bestLabel = '—', bestScore = 0;
  INDICATORS.forEach(ind => {
    const vals = data.map(r => parseFloat(r[ind.key])).filter(v => v >= 1 && v <= 5);
    const avg  = vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
    if (avg > bestScore) { bestScore = avg; bestLabel = ind.label; }
  });

  const cards = [
    { title: 'Total de Respostas', value: n,                           sub: 'formulários submetidos',    col: C.BLUE_DARK  },
    { title: 'Média Global',       value: globalAvg.toFixed(2) + '/5', sub: 'todos os indicadores',      col: C.GREEN      },
    { title: 'Espaço de Opinião',  value: simPct + '%',                sub: simCount + ' responderam Sim', col: C.BLUE_MID },
    { title: 'Melhor Indicador',   value: bestScore.toFixed(2),        sub: bestLabel,                   col: C.GOLD       },
  ];

  const starts = [2, 6, 11, 15]; // colunas de início de cada card
  const W      = 4;

  sh.setRowHeight(row,   16);
  sh.setRowHeight(row+1, 38);
  sh.setRowHeight(row+2, 22);
  sh.setRowHeight(row+3, 20);
  sh.setRowHeight(row+4, 10);

  cards.forEach((card, ci) => {
    const col = starts[ci];
    sh.getRange(row, col, 5, W).merge().setBackground(C.GRAY_LIGHT);
    sh.getRange(row, col, 5, 1).setBackground(card.col);                           // barra lateral
    sh.getRange(row+1, col+1, 1, W-1).merge()
      .setValue(card.value).setFontSize(24).setFontWeight('bold')
      .setFontColor(card.col).setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.getRange(row+2, col+1, 1, W-1).merge()
      .setValue(card.title).setFontSize(9).setFontWeight('bold')
      .setFontColor(C.GRAY_DARK).setHorizontalAlignment('left');
    sh.getRange(row+3, col+1, 1, W-1).merge()
      .setValue(card.sub).setFontSize(8)
      .setFontColor(C.GRAY_MID).setHorizontalAlignment('left');
  });

  return row + 5;
}

// =============================================================================
//  TABELA DE INDICADORES COM BARRAS
// =============================================================================
function writeIndicatorTable_(sh, row, data) {
  // Cabeçalho
  sh.setRowHeight(row, 22);
  const hCols = [[2,'Indicador',140],[12,'Média',50],[13,'Nº',40],[14,'Classificação',90]];
  hCols.forEach(([c, l]) => sh.getRange(row,c).setValue(l).setBackground(C.BLUE_MID).setFontColor(C.WHITE).setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center'));
  sh.getRange(row, 3, 1, 9).merge().setValue('Barra de progresso').setBackground(C.BLUE_MID).setFontColor(C.WHITE).setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
  row++;

  INDICATORS.forEach((ind, idx) => {
    const vals = data.map(r => parseFloat(r[ind.key])).filter(v => v >= 1 && v <= 5);
    const avg  = vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
    const bg   = idx % 2 === 0 ? C.WHITE : C.GRAY_LIGHT;
    const barC = avg >= 4 ? C.GREEN : avg >= 3 ? C.BLUE_MID : avg >= 2 ? C.AMBER : C.RED;

    sh.setRowHeight(row, 22);
    sh.getRange(row, 2).setValue(ind.label).setFontSize(9).setBackground(bg);

    // Barra (9 células = colunas C..K)
    const filled = avg > 0 ? Math.round((avg / 5) * 9) : 0;
    if (filled > 0) sh.getRange(row, 3, 1, filled).setBackground(barC);
    if (filled < 9) sh.getRange(row, 3 + filled, 1, 9 - filled).setBackground('#E0E0E0');

    sh.getRange(row, 12).setValue(avg > 0 ? avg.toFixed(2) : '—').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center').setBackground(bg);
    sh.getRange(row, 13).setValue(vals.length).setFontSize(9).setHorizontalAlignment('center').setBackground(bg);

    // Badge classificação
    const cls   = avg >= 4.5?'Excelente':avg>=3.5?'Bom':avg>=2.5?'Satisfatório':avg>=1.5?'Mau':avg>0?'Muito Mau':'—';
    const clsBg = avg >= 4 ? C.GREEN_LIGHT : avg >= 3 ? C.BLUE_LIGHT : avg >= 2 ? C.AMBER_LIGHT : avg > 0 ? C.RED_LIGHT : bg;
    const clsFg = avg >= 4 ? C.GREEN : avg >= 3 ? C.BLUE_MID : avg >= 2 ? '#7D4E00' : avg > 0 ? C.RED : C.GRAY_MID;
    sh.getRange(row, 14).setValue(cls).setBackground(clsBg).setFontColor(clsFg).setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center');

    row++;
  });

  sh.setRowHeight(row, 16);
  sh.getRange(row, 2, 1, 13).merge()
    .setValue('■ Verde ≥ 4.0   ■ Azul 3.0–3.9   ■ Âmbar 2.0–2.9   ■ Vermelho < 2.0')
    .setFontSize(8).setFontColor(C.GRAY_MID).setFontStyle('italic');
  return row + 1;
}

// =============================================================================
//  DISTRIBUIÇÃO
// =============================================================================
function writeDistributionTable_(sh, row, data) {
  sh.setRowHeight(row, 30);
  sh.getRange(row, 2).setValue('Indicador').setBackground(C.BLUE_MID).setFontColor(C.WHITE).setFontWeight('bold').setFontSize(9);
  ['5 – Excelente','4 – Bom','3 – Satisfatório','2 – Mau','1 – Muito Mau','Total'].forEach((l, i) =>
    sh.getRange(row, 3+i).setValue(l).setBackground(C.BLUE_MID).setFontColor(C.WHITE).setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center').setWrap(true)
  );
  row++;

  INDICATORS.forEach((ind, idx) => {
    const dist  = {5:0,4:0,3:0,2:0,1:0};
    data.forEach(r => { const v = parseInt(r[ind.key]); if (dist[v] !== undefined) dist[v]++; });
    const total = Object.values(dist).reduce((a,b) => a+b, 0);
    const bg    = idx % 2 === 0 ? C.WHITE : C.GRAY_LIGHT;

    sh.setRowHeight(row, 22);
    sh.getRange(row, 2).setValue(ind.label).setFontSize(9).setBackground(bg);
    [5,4,3,2,1].forEach((k, i) => {
      const pct  = total > 0 ? Math.round(dist[k]/total*100) : 0;
      const cell = sh.getRange(row, 3+i);
      cell.setValue(`${dist[k]}  (${pct}%)`).setFontSize(8).setHorizontalAlignment('center');
      if (total > 0) {
        const bgs = [C.GREEN_LIGHT,'#DCEEFB',C.AMBER_LIGHT,'#FDE8D8',C.RED_LIGHT];
        cell.setBackground(dist[k] > 0 ? bgs[i] : bg);
      } else { cell.setBackground(bg); }
    });
    sh.getRange(row, 8).setValue(total).setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center').setBackground(bg);
    row++;
  });
  return row;
}

// =============================================================================
//  PERFIL DOS PARTICIPANTES
// =============================================================================
function writeParticipantProfile_(sh, row, data) {
  const n = data.length;

  // ── Por Perfil ─────────────────────────────────────────────────────────────
  sh.getRange(row,2,1,5).merge().setValue('Por Perfil').setBackground(C.BLUE_LIGHT).setFontColor(C.BLUE_DARK).setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
  sh.getRange(row,8,1,3).merge().setValue('Por Género').setBackground(C.BLUE_LIGHT).setFontColor(C.BLUE_DARK).setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
  row++;

  ['Perfil','Nº','%'].forEach((h,i) => sh.getRange(row,2+i).setValue(h).setFontWeight('bold').setFontSize(8).setBackground(C.GRAY_LIGHT));
  sh.getRange(row,5,1,2).merge().setValue('Barra').setFontWeight('bold').setFontSize(8).setBackground(C.GRAY_LIGHT).setHorizontalAlignment('center');
  ['Género','Nº','%'].forEach((h,i) => sh.getRange(row,8+i).setValue(h).setFontWeight('bold').setFontSize(8).setBackground(C.GRAY_LIGHT));
  row++;

  const perfilStart = row;
  Object.entries(PERFIS).forEach(([k,v], idx) => {
    const cnt = data.filter(r => r['perfil_participante'] === k).length;
    const pct = n > 0 ? Math.round(cnt/n*100) : 0;
    const bg  = idx%2===0?C.WHITE:C.GRAY_LIGHT;
    sh.setRowHeight(row, 22);
    sh.getRange(row,2).setValue(v).setFontSize(9).setBackground(bg);
    sh.getRange(row,3).setValue(cnt).setFontSize(9).setHorizontalAlignment('center').setBackground(bg);
    sh.getRange(row,4).setValue(pct+'%').setFontSize(9).setHorizontalAlignment('center').setBackground(bg);
    const f = Math.round(pct/100*3);
    sh.getRange(row,5,1,3).setBackground('#E0E0E0');
    if (f > 0) sh.getRange(row,5,1,f).setBackground(C.BLUE_MID);
    row++;
  });
  // Total perfil
  const totalPerfil = Object.keys(PERFIS).reduce((s,k) => s + data.filter(r=>r['perfil_participante']===k).length, 0);
  sh.getRange(row,2).setValue('TOTAL').setFontWeight('bold').setFontSize(9).setBackground(C.BLUE_LIGHT);
  sh.getRange(row,3).setValue(n).setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center').setBackground(C.BLUE_LIGHT);
  sh.getRange(row,4).setValue('100%').setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center').setBackground(C.BLUE_LIGHT);
  row++;

  // ── Por Género (ao lado) ──────────────────────────────────────────────────
  let gr = perfilStart;
  Object.entries(GENEROS).forEach(([k,v], idx) => {
    const cnt = data.filter(r => r['genero'] === k).length;
    const pct = n > 0 ? Math.round(cnt/n*100) : 0;
    const bg  = idx%2===0?C.WHITE:C.GRAY_LIGHT;
    sh.getRange(gr,8).setValue(v).setFontSize(9).setBackground(bg);
    sh.getRange(gr,9).setValue(cnt).setFontSize(9).setHorizontalAlignment('center').setBackground(bg);
    sh.getRange(gr,10).setValue(pct+'%').setFontSize(9).setHorizontalAlignment('center').setBackground(bg);
    gr++;
  });
  sh.getRange(gr,8).setValue('TOTAL').setFontWeight('bold').setFontSize(9).setBackground(C.BLUE_LIGHT);
  sh.getRange(gr,9).setValue(n).setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center').setBackground(C.BLUE_LIGHT);
  sh.getRange(gr,10).setValue('100%').setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center').setBackground(C.BLUE_LIGHT);

  // ── Por Igreja ────────────────────────────────────────────────────────────
  row++;
  sh.getRange(row,2,1,8).merge().setValue('Por Igreja / Proveniência').setBackground(C.BLUE_LIGHT).setFontColor(C.BLUE_DARK).setFontWeight('bold').setFontSize(9);
  row++;
  ['Igreja','Nº','%'].forEach((h,i) => sh.getRange(row,2+i).setValue(h).setFontWeight('bold').setFontSize(8).setBackground(C.GRAY_LIGHT));
  row++;

  const churchMap = {};
  data.forEach(r => {
    const ch = (r['proveniencia'] || '').trim() || '(não indicado)';
    churchMap[ch] = (churchMap[ch]||0)+1;
  });
  Object.entries(churchMap).sort((a,b)=>b[1]-a[1]).forEach(([ch, cnt], idx) => {
    const pct = n > 0 ? Math.round(cnt/n*100) : 0;
    const bg  = idx%2===0?C.WHITE:C.GRAY_LIGHT;
    sh.setRowHeight(row, 20);
    sh.getRange(row,2).setValue(ch).setFontSize(9).setBackground(bg);
    sh.getRange(row,3).setValue(cnt).setFontSize(9).setHorizontalAlignment('center').setBackground(bg);
    sh.getRange(row,4).setValue(pct+'%').setFontSize(9).setHorizontalAlignment('center').setBackground(bg);
    row++;
  });

  return row;
}

// =============================================================================
//  ESPAÇO DE OPINIÃO
// =============================================================================
function writeOpinionSpace_(sh, row, data) {
  const n        = data.length;
  const simCount = data.filter(r => r['espaco_opiniao'] === 'sim').length;
  const naoCount = data.filter(r => r['espaco_opiniao'] === 'nao').length;
  const simPct   = n > 0 ? Math.round(simCount/n*100) : 0;

  // Totais
  ['Resposta','Nº','%','Barra'].forEach((h,i) => sh.getRange(row,2+i).setValue(h).setFontWeight('bold').setFontSize(9).setBackground(C.GRAY_LIGHT));
  sh.getRange(row,5,1,5).merge().setValue('Barra').setFontWeight('bold').setFontSize(9).setBackground(C.GRAY_LIGHT).setHorizontalAlignment('center');
  row++;

  [['Sim', simCount, simPct,      C.GREEN,    C.GREEN_LIGHT],
   ['Não', naoCount, 100-simPct,  C.RED,      C.RED_LIGHT  ]].forEach(([lbl,cnt,pct,barC,bgL]) => {
    sh.setRowHeight(row, 24);
    sh.getRange(row,2).setValue(lbl).setFontSize(10).setFontWeight('bold').setFontColor(barC).setBackground(bgL);
    sh.getRange(row,3).setValue(cnt).setFontSize(10).setHorizontalAlignment('center').setBackground(bgL);
    sh.getRange(row,4).setValue(pct+'%').setFontSize(10).setHorizontalAlignment('center').setBackground(bgL);
    const f = Math.round(pct/100*5);
    sh.getRange(row,5,1,5).setBackground('#E0E0E0');
    if (f>0) sh.getRange(row,5,1,f).setBackground(barC);
    row++;
  });

  // Por perfil
  row++;
  sh.getRange(row,2,1,7).merge().setValue('Espaço de Opinião por Perfil').setBackground(C.BLUE_LIGHT).setFontColor(C.BLUE_DARK).setFontWeight('bold').setFontSize(9);
  row++;
  ['Perfil','Total','Sim','Não','% Sim'].forEach((h,i) => sh.getRange(row,2+i).setValue(h).setFontWeight('bold').setFontSize(8).setBackground(C.GRAY_LIGHT).setHorizontalAlignment('center'));
  row++;

  Object.entries(PERFIS).forEach(([k,v], idx) => {
    const sub = data.filter(r => r['perfil_participante'] === k);
    const s   = sub.filter(r => r['espaco_opiniao'] === 'sim').length;
    const ns  = sub.filter(r => r['espaco_opiniao'] === 'nao').length;
    const p   = sub.length > 0 ? Math.round(s/sub.length*100) : 0;
    const bg  = idx%2===0?C.WHITE:C.GRAY_LIGHT;
    sh.getRange(row,2).setValue(v).setFontSize(9).setBackground(bg);
    sh.getRange(row,3).setValue(sub.length).setFontSize(9).setHorizontalAlignment('center').setBackground(bg);
    sh.getRange(row,4).setValue(s).setFontSize(9).setHorizontalAlignment('center').setFontColor(C.GREEN).setBackground(bg);
    sh.getRange(row,5).setValue(ns).setFontSize(9).setHorizontalAlignment('center').setFontColor(C.RED).setBackground(bg);
    sh.getRange(row,6).setValue(p+'%').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center').setBackground(bg);
    row++;
  });

  return row;
}

// =============================================================================
//  COMENTÁRIOS
// =============================================================================
function writeComments_(sh, row, data) {
  const cols  = [['#',30],['Perfil',120],['Género',80],['Igreja',130],['O que correu melhor',280],['O que deve ser melhorado',280]];
  cols.forEach(([h,w], i) => {
    sh.getRange(row, 2+i).setValue(h).setBackground(C.BLUE_MID).setFontColor(C.WHITE).setFontWeight('bold').setFontSize(9).setWrap(true);
    sh.setColumnWidth(2+i, w);
  });
  row++;

  let n = 0;
  data.forEach(r => {
    const f = (r['pontos_fortes']   || '').trim();
    const m = (r['pontos_melhorar'] || '').trim();
    if (!f && !m) return;
    n++;
    const bg = n%2===0?C.GRAY_LIGHT:C.WHITE;
    sh.setRowHeight(row, 60);
    sh.getRange(row,2).setValue(n).setFontSize(8).setHorizontalAlignment('center').setBackground(bg);
    sh.getRange(row,3).setValue(PERFIS[r['perfil_participante']]||'—').setFontSize(8).setBackground(bg);
    sh.getRange(row,4).setValue(GENEROS[r['genero']]||'—').setFontSize(8).setBackground(bg);
    sh.getRange(row,5).setValue(r['proveniencia']||'—').setFontSize(8).setBackground(bg);
    sh.getRange(row,6).setValue(f||'—').setFontSize(8).setWrap(true).setBackground(bg);
    sh.getRange(row,7).setValue(m||'—').setFontSize(8).setWrap(true).setBackground(bg);
    row++;
  });

  if (n === 0) {
    sh.getRange(row,2,1,6).merge().setValue('Sem comentários submetidos ainda.')
      .setFontSize(9).setFontColor(C.GRAY_MID).setFontStyle('italic');
    row++;
  }
  return row;
}

// =============================================================================
//  NOTA DE FILTROS
// =============================================================================
function writeFiltersNote_(sh, row) {
  sh.setRowHeight(row, 44);
  sh.getRange(row, 2, 1, 17).merge()
    .setValue('💡  COMO FILTRAR:  Vai ao separador Raw_Kobo → selecciona os dados → Dados → Criar filtro. Filtra por: perfil_participante, genero ou proveniencia. Depois volta ao Dashboard e executa  📊 Kobo Dashboard → Reconstruir Dashboard.')
    .setBackground(C.AMBER_LIGHT).setFontColor('#7D4E00')
    .setFontSize(8).setWrap(true).setFontStyle('italic');
}

// =============================================================================
//  HELPER
// =============================================================================
function toast_(ss, msg) {
  ss.toast(msg, CONFIG.EVENT_NAME, 5);
}

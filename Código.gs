// ============================================================================
// 1. CONFIGURAÇÃO E ESTRUTURA DA PLANILHA
// ============================================================================
const FONTE_ID = '1GtYG4Ahy5XJyJjE37S27u8RyELdRkct8nDAVGIBRI-w';
const FONTE_ABA = ''; // vazio = detecta a aba pelo cabeçalho
const FUSO = 'America/Fortaleza';
const SYNC_INTERVALO_MS = 10 * 60 * 1000;
const SESSAO_SEGUNDOS = 6 * 60 * 60;
const MAX_TENTATIVAS = 5;
const BLOQUEIO_SEGUNDOS = 15 * 60;
const EMAIL_RE = /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/;
const TOKEN_ENVIO_RE = /^[a-f0-9]{32}$/;

const COLS_FONTE = {
  cartela: 'CARTELA',
  cliente: 'CLIENTE',
  cod: 'COD. CLIENTE',
  desc: 'DESCRICAO',
  tam: 'TAMANHO',
  oc: 'ORD. COMPRA'
};

const CABECALHOS = {
  Itens: ['ID', 'Cliente', 'Descrição do Item', 'Cód. Cliente', 'Tamanho', 'Ordem de Compra', 'Status', 'Resposta', 'Observação', 'Data Atualização'],
  Emails: ['Cliente', 'Nome', 'Email', 'Tipo'],
  Usuarios: ['Usuario', 'Salt', 'SenhaHash', 'Ativo'],
  Envios: ['Token', 'Cliente', 'IDs', 'Data', 'Tipo', 'Usuario'],
  Historico: ['Data', 'Evento', 'Cliente', 'ID', 'Item', 'Resposta', 'Observacao', 'Origem'],
  Config: ['Parametro', 'Texto']
};
const CORES = { Itens: '#d9ead3', Emails: '#cfe2f3', Usuarios: '#f4cccc', Envios: '#ead1dc', Historico: '#e6e6e6', Config: '#fff2cc' };
const TEXTO_DISPARO = 'Texto do disparo';
const TEXTO_REENVIO = 'Texto do reenvio';
const TEXTOS_PADRAO = [
  [TEXTO_DISPARO, 'Olá! Segue abaixo a lista de itens que necessitamos de cartela. Por favor, clique no botão do e-mail para informar a disponibilidade de cada item.'],
  [TEXTO_REENVIO, 'Precisamos de uma revisão ou nova posição sobre os itens abaixo:']
];

function configurarPlanilhaAutomaticamente() {
  CacheService.getScriptCache().remove('abas_ok');
  garantirAbas_();
}

function garantirAbas_() {
  const cache = CacheService.getScriptCache();
  if (cache.get('abas_ok')) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(CABECALHOS).forEach(nome => {
    const cab = CABECALHOS[nome];
    let aba = ss.getSheetByName(nome);
    if (aba && aba.getLastRow() > 0 && String(aba.getRange(1, 2).getValue()) !== cab[1]) {
      aba.setName(`${nome}_antigo_${Utilities.formatDate(new Date(), FUSO, 'yyyyMMddHHmmss')}`);
      aba = null;
    }
    if (!aba) aba = ss.insertSheet(nome);
    if (aba.getLastRow() === 0) {
      aba.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground(CORES[nome]);
      aba.setFrozenRows(1);
      if (nome === 'Config') gravar_(aba, 2, 1, TEXTOS_PADRAO);
    }
  });

  const abaEmails = ss.getSheetByName('Emails');
  abaEmails.getRange('A1').setNote('Nome do cliente ou início do nome. Ex.: DAKOTA atende DAKOTA RUSSAS, DAKOTA IGUATU e filiais futuras.\nO destinatário recebe o formulário; a cópia recebe só a lista de itens e as respostas.');
  abaEmails.getRange('D2:D1000').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Destinatário', 'Cópia'], true).setAllowInvalid(false).build()
  );

  let abaRespostas = ss.getSheetByName('Respostas');
  if (!abaRespostas) {
    abaRespostas = ss.insertSheet('Respostas');
    abaRespostas.appendRow(['Opções de Resposta']);
    abaRespostas.getRange('A1').setFontWeight('bold').setBackground('#fff2cc');
    abaRespostas.setFrozenRows(1);
    const frasesPadrao = [
      ['Cartela disponível - Envio imediato'],
      ['Em produção - Envio em até 5 dias'],
      ['Aguardando tingimento / lote'],
      ['Necessário confirmar cor / referência'],
      ['Item fora de linha / Indisponível']
    ];
    abaRespostas.getRange(2, 1, frasesPadrao.length, 1).setValues(frasesPadrao);
  }

  cache.put('abas_ok', '1', 21600);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Cartelas')
    .addItem('Definir usuário e senha', 'definirSenhaUsuario')
    .addItem('Sincronizar com planilha de produção', 'sincronizarAgora')
    .addItem('Ativar sincronização horária', 'instalarGatilhoSincronizacao')
    .addToUi();
}

// ============================================================================
// 2. UTILITÁRIOS
// ============================================================================
function normalizar_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function hex_(bytes) {
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function sha256_(texto) {
  return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto, Utilities.Charset.UTF_8));
}

function novoToken_(partes) {
  let t = '';
  for (let i = 0; i < partes; i++) t += Utilities.getUuid().replace(/-/g, '');
  return t.slice(0, partes * 32);
}

function agora_() {
  return Utilities.formatDate(new Date(), FUSO, 'dd/MM/yyyy HH:mm');
}

function comLock_(fn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    throw new Error('Sistema ocupado. Tente novamente em instantes.');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function gravar_(aba, linha, coluna, valores) {
  if (!valores.length) return;
  aba.getRange(linha, coluna, valores.length, valores[0].length).setNumberFormat('@').setValues(valores);
}

function registrarHistorico_(linhas) {
  if (!linhas.length) return;
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Historico');
  gravar_(aba, aba.getLastRow() + 1, 1, linhas);
}

// ============================================================================
// 3. AUTENTICAÇÃO (APENAS PAINEL ADMINISTRATIVO)
// ============================================================================
function pepper_() {
  const props = PropertiesService.getScriptProperties();
  let p = props.getProperty('PEPPER');
  if (!p) {
    p = novoToken_(2);
    props.setProperty('PEPPER', p);
  }
  return p;
}

function hashSenha_(senha, salt) {
  return sha256_(salt + pepper_() + senha);
}

function compararSeguro_(a, b) {
  if (a.length !== b.length) return false;
  let r = 0;
  for (let i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}

function login(usuario, senha) {
  garantirAbas_();
  const u = String(usuario || '').trim().toLowerCase();
  if (!u || u.length > 100 || typeof senha !== 'string' || !senha || senha.length > 200) {
    throw new Error('Usuário ou senha inválidos.');
  }
  const cache = CacheService.getScriptCache();
  const chaveFalhas = 'falhas_' + sha256_(u).slice(0, 32);
  const falhas = Number(cache.get(chaveFalhas) || 0);
  if (falhas >= MAX_TENTATIVAS) throw new Error('Muitas tentativas. Aguarde 15 minutos.');

  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
  const ult = aba.getLastRow();
  const linhas = ult > 1 ? aba.getRange(2, 1, ult - 1, 4).getDisplayValues() : [];
  const reg = linhas.find(r => r[0].trim().toLowerCase() === u && normalizar_(r[3]) === 'SIM');
  const calculado = hashSenha_(senha, reg ? reg[1] : '-');

  if (!reg || !compararSeguro_(calculado, reg[2])) {
    cache.put(chaveFalhas, String(falhas + 1), BLOQUEIO_SEGUNDOS);
    throw new Error('Usuário ou senha inválidos.');
  }
  cache.remove(chaveFalhas);
  const token = novoToken_(2);
  cache.put('sess_' + token, reg[0].trim(), SESSAO_SEGUNDOS);
  return { token, usuario: reg[0].trim() };
}

function logout(token) {
  if (typeof token === 'string' && /^[a-f0-9]{64}$/.test(token)) CacheService.getScriptCache().remove('sess_' + token);
}

function exigirSessao_(token) {
  const usuario = typeof token === 'string' && /^[a-f0-9]{64}$/.test(token)
    ? CacheService.getScriptCache().get('sess_' + token)
    : null;
  if (!usuario) throw new Error('SESSAO_EXPIRADA');
  return usuario;
}

function definirSenhaUsuario() {
  const ui = SpreadsheetApp.getUi();
  const pu = ui.prompt('Usuário', 'Nome do usuário (novo ou existente):', ui.ButtonSet.OK_CANCEL);
  if (pu.getSelectedButton() !== ui.Button.OK) return;
  const usuario = pu.getResponseText().trim();
  if (!usuario) return;
  const ps = ui.prompt('Senha', 'Nova senha para "' + usuario + '" (mínimo 8 caracteres):', ui.ButtonSet.OK_CANCEL);
  if (ps.getSelectedButton() !== ui.Button.OK) return;
  const senha = ps.getResponseText();
  if (senha.length < 8) {
    ui.alert('A senha precisa ter pelo menos 8 caracteres.');
    return;
  }

  garantirAbas_();
  comLock_(() => {
    const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
    const ult = aba.getLastRow();
    const nomes = ult > 1 ? aba.getRange(2, 1, ult - 1, 1).getDisplayValues() : [];
    const idx = nomes.findIndex(r => r[0].trim().toLowerCase() === usuario.toLowerCase());
    const salt = novoToken_(1);
    gravar_(aba, idx >= 0 ? idx + 2 : ult + 1, 1, [[usuario, salt, hashSenha_(senha, salt), 'SIM']]);
  });
  ui.alert('Usuário "' + usuario + '" salvo. Para desativar, escreva NÃO na coluna Ativo da aba Usuarios.');
}

// ============================================================================
// 4. SINCRONIZAÇÃO COM A PLANILHA DE PRODUÇÃO
// ============================================================================
function localizarCabecalho_(topo) {
  for (let r = 0; r < topo.length; r++) {
    const norm = topo[r].map(normalizar_);
    const idx = {};
    const completo = Object.keys(COLS_FONTE).every(k => {
      idx[k] = norm.indexOf(COLS_FONTE[k]);
      return idx[k] >= 0;
    });
    if (completo) return { linha: r, idx };
  }
  return null;
}

function lerFonte_() {
  const ss = SpreadsheetApp.openById(FONTE_ID);
  const props = PropertiesService.getScriptProperties();
  const preferida = FONTE_ABA || props.getProperty('FONTE_ABA');
  let abas = ss.getSheets();
  if (FONTE_ABA) abas = abas.filter(a => a.getName() === FONTE_ABA);
  abas.sort((a, b) => (b.getName() === preferida) - (a.getName() === preferida));

  for (const aba of abas) {
    const lc = aba.getLastColumn();
    const lr = aba.getLastRow();
    if (lc < 1 || lr < 2) continue;
    const cabecalho = localizarCabecalho_(aba.getRange(1, 1, Math.min(10, lr), lc).getDisplayValues());
    if (!cabecalho) continue;
    props.setProperty('FONTE_ABA', aba.getName());
    const primeira = cabecalho.linha + 2;
    const linhas = lr >= primeira ? aba.getRange(primeira, 1, lr - primeira + 1, lc).getDisplayValues() : [];
    return { cabecalho, linhas };
  }
  throw new Error('Cabeçalho (CARTELA, CLIENTE, CÓD. CLIENTE, DESCRIÇÃO, TAMANHO, ORD. COMPRA) não localizado na planilha de produção.');
}

function adicionarUnico_(lista, valor) {
  const v = String(valor == null ? '' : valor).trim();
  if (v && lista.indexOf(v) < 0) lista.push(v);
}

function extrairPendentes_(cabecalho, linhas) {
  const idx = cabecalho.idx;
  const grupos = new Map();
  let ignoradas = 0;
  linhas.forEach(r => {
    if (normalizar_(r[idx.cartela]) !== 'NAO') return;
    const cliente = String(r[idx.cliente]).trim();
    const desc = String(r[idx.desc]).trim();
    if (!cliente || !desc) {
      ignoradas++;
      return;
    }
    const chave = normalizar_(cliente) + '|' + normalizar_(desc);
    let g = grupos.get(chave);
    if (!g) {
      g = { chave, cliente, desc, cods: [], tams: [], ocs: [] };
      grupos.set(chave, g);
    }
    adicionarUnico_(g.cods, r[idx.cod]);
    adicionarUnico_(g.tams, r[idx.tam]);
    adicionarUnico_(g.ocs, r[idx.oc]);
  });
  return { itens: Array.from(grupos.values()), ignoradas };
}

function sincronizarItens_() {
  const fonte = lerFonte_();
  if (!fonte.linhas.length) throw new Error('Planilha de produção sem linhas de dados. Sincronização abortada.');
  const pend = extrairPendentes_(fonte.cabecalho, fonte.linhas);

  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Itens');
  const ult = aba.getLastRow();
  const atuais = ult > 1 ? aba.getRange(2, 1, ult - 1, 10).getDisplayValues().filter(r => r[0]) : [];
  const porId = new Map(atuais.map(r => [r[0], r]));

  const novas = pend.itens.map(p => {
    const id = sha256_(p.chave).slice(0, 12);
    const ant = porId.get(id);
    return [id, p.cliente, p.desc, p.cods.join(' / '), p.tams.join(' / '), p.ocs.join(' / '),
      ant && ant[6] ? ant[6] : 'Não enviado', ant ? ant[7] : '', ant ? ant[8] : '', ant ? ant[9] : ''];
  });

  const idsNovos = new Set(novas.map(r => r[0]));
  const resumo = {
    total: novas.length,
    novos: novas.filter(r => !porId.has(r[0])).length,
    removidos: atuais.filter(r => !idsNovos.has(r[0])).length,
    ignoradas: pend.ignoradas
  };

  if (JSON.stringify(atuais) !== JSON.stringify(novas)) {
    if (ult > 1) aba.getRange(2, 1, ult - 1, 10).clearContent();
    gravar_(aba, 2, 1, novas);
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty('ULTIMA_SYNC', String(Date.now()));
  props.setProperty('ULTIMA_SYNC_TXT', agora_());
  props.deleteProperty('ERRO_SYNC');
  return resumo;
}

function registrarErroSync_(err) {
  PropertiesService.getScriptProperties().setProperty('ERRO_SYNC', `${agora_()} - ${err.message}`);
}

function sincronizarGatilho() {
  const ult = Number(PropertiesService.getScriptProperties().getProperty('ULTIMA_SYNC') || 0);
  if (Date.now() - ult < 5 * 60 * 1000) return;
  try {
    garantirAbas_();
    comLock_(sincronizarItens_);
  } catch (err) {
    registrarErroSync_(err);
  }
}

function sincronizarAgora() {
  const ui = SpreadsheetApp.getUi();
  try {
    garantirAbas_();
    const r = comLock_(sincronizarItens_);
    ui.alert(`Sincronizado: ${r.total} itens (${r.novos} novos, ${r.removidos} removidos${r.ignoradas ? ', ' + r.ignoradas + ' linhas sem cliente/descrição ignoradas' : ''}).`);
  } catch (err) {
    registrarErroSync_(err);
    ui.alert('Falha na sincronização: ' + err.message);
  }
}

function instalarGatilhoSincronizacao() {
  const ui = SpreadsheetApp.getUi();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sincronizarGatilho')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sincronizarGatilho').timeBased().everyHours(1).create();
  ui.alert('Sincronização automática ativada (a cada 1 hora).');
}

// ============================================================================
// 5. LEITURA DE DADOS E CLIENTES
// ============================================================================
function lerItens_() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Itens');
  const ult = aba.getLastRow();
  if (ult < 2) return [];
  return aba.getRange(2, 1, ult - 1, 10).getDisplayValues()
    .filter(r => r[0])
    .map(r => ({ id: r[0], cliente: r[1], desc: r[2], cod: r[3], tam: r[4], oc: r[5], status: r[6] || 'Não enviado', resposta: r[7], obs: r[8], data: r[9] }));
}

function lerEmails_(avisos) {
  const aviso = avisos || [];
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Emails');
  const ult = aba.getLastRow();
  const linhas = ult > 1 ? aba.getRange(2, 1, ult - 1, 4).getDisplayValues() : [];
  const lista = [];
  linhas.forEach((r, n) => {
    const cliente = r[0].trim();
    const emails = r[2].split(/[;,\s]+/).filter(Boolean);
    if (!cliente && !emails.length) return;
    const onde = `Aba Emails, linha ${n + 2}: `;
    const tipoN = normalizar_(r[3]);
    const tipo = (tipoN === '' || tipoN.indexOf('DEST') === 0) ? 'D' : (tipoN.indexOf('COP') === 0 || /^CC/.test(tipoN)) ? 'C' : null;
    if (!cliente) { aviso.push(onde + 'Cliente vazio (destinatários e cópias são por cliente). Linha ignorada.'); return; }
    if (!tipo) { aviso.push(`${onde}Tipo "${r[3]}" não reconhecido. Use Destinatário ou Cópia. Linha ignorada.`); return; }
    if (!emails.length) { aviso.push(onde + 'sem e-mail. Linha ignorada.'); return; }
    emails.forEach(email => {
      if (EMAIL_RE.test(email)) lista.push({ chave: normalizar_(cliente), nome: cliente, email, tipo });
      else aviso.push(`${onde}e-mail "${email}" inválido. Ignorado.`);
    });
  });
  return lista;
}

function mapaClientes_(avisos) {
  const mapa = new Map();
  lerEmails_(avisos).forEach(e => {
    let g = mapa.get(e.chave);
    if (!g) {
      g = { nome: e.nome, dest: [], copia: [] };
      mapa.set(e.chave, g);
    }
    const alvo = e.tipo === 'D' ? g.dest : g.copia;
    if (!alvo.some(x => x.toLowerCase() === e.email.toLowerCase())) alvo.push(e.email);
  });
  mapa.forEach(g => {
    g.copia = g.copia.filter(c => !g.dest.some(d => d.toLowerCase() === c.toLowerCase()));
  });
  return mapa;
}

function criarResolvedor_(mapa) {
  const chaves = Array.from(mapa.keys()).sort((a, b) => b.length - a.length);
  return cliente => {
    const c = normalizar_(cliente);
    return chaves.find(k => c === k || c.indexOf(k + ' ') === 0) || null;
  };
}

function obterOpcoes_() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Respostas');
  let opcoes = [];
  if (aba && aba.getLastRow() > 1) {
    opcoes = aba.getRange(2, 1, aba.getLastRow() - 1, 1).getDisplayValues().map(r => r[0].trim()).filter(Boolean);
  }
  return opcoes.length ? opcoes : ['Cartela disponível - Envio imediato', 'Indisponível', 'Verificar estoque'];
}

// ============================================================================
// 6. ROTEADOR DO WEBAPP
// ============================================================================
function doGet(e) {
  garantirAbas_();
  const params = (e && e.parameter) || {};

  if (params.t !== undefined || params.ids !== undefined) {
    const template = HtmlService.createTemplateFromFile('Resposta');
    template.t = TOKEN_ENVIO_RE.test(String(params.t || '')) ? params.t : '';
    return template.evaluate()
      .setTitle('Responder Solicitação de Cartelas')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Gestão e Solicitação de Cartelas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================================
// 7. PAINEL ADMINISTRATIVO
// ============================================================================
function montarDadosPainel_() {
  const props = PropertiesService.getScriptProperties();
  const avisosEmails = [];
  const mapa = mapaClientes_(avisosEmails);
  const resolver = criarResolvedor_(mapa);
  const itens = lerItens_().map(i => {
    const k = resolver(i.cliente);
    const g = k ? mapa.get(k) : null;
    return {
      id: i.id,
      cliente: i.cliente,
      descricao: i.desc,
      codCliente: i.cod,
      tamanho: i.tam,
      oc: i.oc,
      status: i.status,
      resposta: i.resposta,
      observacao: i.obs,
      data: i.data,
      grupo: g ? g.nome : '',
      destinatarios: g ? g.dest.length : 0,
      copias: g ? g.copia.length : 0
    };
  });
  return {
    itens,
    avisosEmails,
    ultimaSync: props.getProperty('ULTIMA_SYNC_TXT') || '',
    erroSync: props.getProperty('ERRO_SYNC') || ''
  };
}

function obterDadosPainel(token) {
  exigirSessao_(token);
  garantirAbas_();
  return montarDadosPainel_();
}

function sincronizarFonte(token, forcar) {
  exigirSessao_(token);
  garantirAbas_();
  const ult = Number(PropertiesService.getScriptProperties().getProperty('ULTIMA_SYNC') || 0);
  if (forcar || Date.now() - ult > SYNC_INTERVALO_MS) {
    try {
      comLock_(sincronizarItens_);
    } catch (err) {
      registrarErroSync_(err);
    }
  }
  return montarDadosPainel_();
}

// ============================================================================
// 8. DISPARO DE E-MAILS (POR CLIENTE)
// ============================================================================
function htmlEmail_(o) {
  const cab = o.cabecalhos.map(h => `<th style="padding: 10px; border-bottom: 2px solid #ddd;">${esc_(h)}</th>`).join('');
  const lin = o.linhas.map(l => '<tr>' + l.map((c, n) =>
    `<td style="padding: 10px; border-bottom: 1px solid #eee;${n === 0 ? ' font-weight: bold;' : ''}">${esc_(c)}</td>`).join('') + '</tr>').join('');
  const botao = o.link ? `
        <table role="presentation" align="center" cellpadding="0" cellspacing="0" style="margin: 16px auto;">
          <tr>
            <td align="center" bgcolor="#28a745" style="border-radius: 8px; padding: 18px 40px;">
              <a href="${esc_(o.link)}" style="color: #ffffff; font-size: 18px; font-weight: bold; text-decoration: none; display: block;">👉 CLIQUE AQUI PARA RESPONDER</a>
            </td>
          </tr>
        </table>` : '';
  const destaque = o.link ? botao : (o.faixa ? `
        <div style="background-color: #eceff1; border-radius: 8px; padding: 12px; margin: 0 0 20px; text-align: center; font-size: 14px; font-weight: bold; color: #455a64;">${esc_(o.faixa)}</div>` : '');
  return `
    <div style="font-family: Arial, sans-serif; max-width: 700px; color: #333; border: 1px solid #ddd; border-radius: 8px; overflow: hidden;">
      <div style="background-color: ${o.cor}; color: white; padding: 16px;">
        <h2 style="margin: 0; font-size: 18px;">${esc_(o.titulo)}</h2>
      </div>
      <div style="padding: 20px;">
        <p style="font-size: 15px; line-height: 1.5;">${esc_(o.mensagem).replace(/\n/g, '<br>')}</p>${destaque}
        <table style="width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 14px;">
          <thead><tr style="background-color: #f7f7f7; text-align: left;">${cab}</tr></thead>
          <tbody>${lin}</tbody>
        </table>${o.link ? botao : ''}
        <p style="font-size: 12px; color: #888; margin-top: 20px; text-align: center;">${esc_(o.rodape)}</p>
      </div>
    </div>`;
}

function emailSolicitacao_(itens, mensagem, ehReenvio, link) {
  const cabecalhos = ['Item', 'Cliente / Filial', 'Tamanho', 'Ordem de compra', 'Cód. Cliente'];
  if (ehReenvio) cabecalhos.push('Retorno anterior');
  const linhas = itens.map(i => {
    const l = [i.desc, i.cliente, i.tam, i.oc, i.cod];
    if (ehReenvio) l.push((i.resposta || 'Sem resposta') + (i.obs ? ' - ' + i.obs : ''));
    return l;
  });
  return htmlEmail_({
    titulo: ehReenvio ? 'Revisão solicitada: lista de itens para cartela' : 'Solicitação: lista de itens que necessitamos de cartela',
    cor: ehReenvio ? '#d9534f' : '#0275d8',
    mensagem,
    cabecalhos,
    linhas,
    link,
    faixa: link ? '' : '📋 CÓPIA INFORMATIVA: nenhuma ação necessária',
    rodape: link
      ? 'Clique no botão acima para selecionar a resposta pré-definida e preencher observações para cada item.'
      : 'Cópia informativa: nenhuma ação é necessária. As respostas também serão enviadas a você por e-mail.'
  });
}

function assunto_(ehReenvio, nome, copia) {
  const base = ehReenvio ? '[REVISÃO SOLICITADA] Lista de Itens para Cartela' : '[SOLICITAÇÃO] Lista de Itens que necessitamos de cartela';
  return (copia ? '[CÓPIA] ' : '') + base + ' - ' + nome.replace(/[\r\n]+/g, ' ');
}

function obterTexto_(parametro) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName('Config');
  if (!aba) {
    CacheService.getScriptCache().remove('abas_ok');
    garantirAbas_();
    aba = ss.getSheetByName('Config');
  }
  const ult = aba.getLastRow();
  const linhas = ult > 1 ? aba.getRange(2, 1, ult - 1, 2).getDisplayValues() : [];
  const reg = linhas.find(r => normalizar_(r[0]) === normalizar_(parametro) && r[1].trim());
  const padrao = TEXTOS_PADRAO.find(p => p[0] === parametro)[1];
  return (reg ? reg[1].trim() : padrao).slice(0, 2000);
}

function dispararSolicitacao(token, ids, ehReenvio) {
  const usuario = exigirSessao_(token);
  if (!Array.isArray(ids) || !ids.length) throw new Error('Selecione pelo menos um item.');
  const urlWebApp = ScriptApp.getService().getUrl();
  if (!urlWebApp) throw new Error('Publique o Web App antes de disparar.');
  ehReenvio = !!ehReenvio;
  if (ehReenvio && ids.length !== 1) throw new Error('O reenvio é individual: use o botão Reenviar do item.');
  const mensagem = obterTexto_(ehReenvio ? TEXTO_REENVIO : TEXTO_DISPARO);
  const selecionados = new Set(ids.map(String));
  const dataAtual = agora_();
  const evento = ehReenvio ? 'Reenvio' : 'Disparo';
  const plano = [];
  const pulados = [];

  comLock_(() => {
    const mapa = mapaClientes_();
    const resolver = criarResolvedor_(mapa);
    const grupos = new Map();
    const semCadastro = new Map();

    const itensSel = lerItens_().filter(i => selecionados.has(i.id));
    const aptos = itensSel.filter(i => (i.status !== 'Não enviado') === ehReenvio);
    if (aptos.length < itensSel.length) {
      pulados.push(ehReenvio
        ? `${itensSel[0].desc}: ainda não foi enviado. Use "Disparar Selecionados".`
        : `${itensSel.length - aptos.length} item(ns) já enviado(s) ignorado(s). Para enviar de novo, use Reenviar no item.`);
    }

    aptos.forEach(i => {
      const k = resolver(i.cliente);
      if (!k) {
        semCadastro.set(i.cliente, (semCadastro.get(i.cliente) || 0) + 1);
        return;
      }
      if (!grupos.has(k)) grupos.set(k, []);
      grupos.get(k).push(i);
    });
    semCadastro.forEach((n, nome) => pulados.push(`${nome}: cliente sem e-mail cadastrado (${n} item(ns))`));

    let cota = MailApp.getRemainingDailyQuota();
    const registros = [];
    grupos.forEach((lista, k) => {
      const g = mapa.get(k);
      if (!g.dest.length) {
        pulados.push(`${g.nome}: sem destinatário cadastrado (${lista.length} item(ns))`);
        return;
      }
      const necessario = g.dest.length + g.copia.length;
      if (necessario > cota) {
        pulados.push(`${g.nome}: cota diária de e-mails insuficiente`);
        return;
      }
      cota -= necessario;
      const tk = novoToken_(1);
      registros.push([tk, g.nome, lista.map(i => i.id).join(','), dataAtual, evento, usuario]);
      plano.push({ g, lista, tk });
    });

    const abaEnvios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Envios');
    gravar_(abaEnvios, abaEnvios.getLastRow() + 1, 1, registros);
  });

  const concluidos = [];
  const enviados = [];
  plano.forEach(p => {
    try {
      MailApp.sendEmail({
        to: p.g.dest.join(','),
        subject: assunto_(ehReenvio, p.g.nome, false),
        htmlBody: emailSolicitacao_(p.lista, mensagem, ehReenvio, `${urlWebApp}?t=${p.tk}`)
      });
    } catch (err) {
      pulados.push(`${p.g.nome}: falha no envio (${err.message})`);
      return;
    }
    concluidos.push(p);
    let copiasEnviadas = 0;
    if (p.g.copia.length) {
      try {
        MailApp.sendEmail({
          to: p.g.copia.join(','),
          subject: assunto_(ehReenvio, p.g.nome, true),
          htmlBody: emailSolicitacao_(p.lista, mensagem, ehReenvio, '')
        });
        copiasEnviadas = p.g.copia.length;
      } catch (err) {
        pulados.push(`${p.g.nome}: enviado ao cliente, mas falhou o envio às cópias (${err.message})`);
      }
    }
    enviados.push({ cliente: p.g.nome, itens: p.lista.length, destinatarios: p.g.dest.length, copias: copiasEnviadas });
  });

  if (concluidos.length) {
    comLock_(() => {
      const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Itens');
      const ult = aba.getLastRow();
      if (ult < 2) return;
      const bloco = aba.getRange(2, 1, ult - 1, 10).getDisplayValues();
      const nomePorId = new Map();
      concluidos.forEach(p => p.lista.forEach(i => nomePorId.set(i.id, p.g.nome)));
      const historico = [];
      bloco.forEach(r => {
        if (!nomePorId.has(r[0])) return;
        r[6] = ehReenvio ? 'Reenviado - Aguardando' : 'Aguardando Resposta';
        r[9] = dataAtual;
        historico.push([dataAtual, evento, nomePorId.get(r[0]), r[0], r[2], '', '', usuario]);
      });
      gravar_(aba, 2, 7, bloco.map(r => r.slice(6, 10)));
      registrarHistorico_(historico);
    });
  }

  return { dados: montarDadosPainel_(), enviados, pulados };
}

// ============================================================================
// 9. TELA DE RESPOSTA DO DESTINATÁRIO (ACESSO POR TOKEN DE ENVIO)
// ============================================================================
function buscarEnvio_(token) {
  if (!TOKEN_ENVIO_RE.test(String(token || ''))) return null;
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Envios');
  const ult = aba.getLastRow();
  if (ult < 2) return null;
  const achado = aba.getRange(2, 1, ult - 1, 1).createTextFinder(token).matchEntireCell(true).findNext();
  if (!achado) return null;
  const r = aba.getRange(achado.getRow(), 1, 1, 6).getDisplayValues()[0];
  return { token: r[0], cliente: r[1], ids: r[2].split(','), data: r[3], tipo: r[4] };
}

function obterItensPorToken(tokenEnvio) {
  garantirAbas_();
  const envio = buscarEnvio_(tokenEnvio);
  if (!envio) return { invalido: true };
  const ids = new Set(envio.ids);
  const itens = lerItens_().filter(i => ids.has(i.id)).map(i => ({
    id: i.id,
    nome: i.desc,
    cliente: i.cliente,
    codCliente: i.cod,
    tamanho: i.tam,
    oc: i.oc,
    resposta: i.resposta,
    observacao: i.obs
  }));
  return { itens, opcoes: obterOpcoes_(), grupo: envio.cliente };
}

function salvarRespostasFornecedor(tokenEnvio, listaRespostas) {
  if (!Array.isArray(listaRespostas) || listaRespostas.length > 500) throw new Error('Dados inválidos.');
  const alteradas = [];
  let grupo = '';

  comLock_(() => {
    const envio = buscarEnvio_(tokenEnvio);
    if (!envio) throw new Error('Link inválido ou expirado.');
    grupo = envio.cliente;
    const permitidos = new Set(envio.ids);
    const opcoes = obterOpcoes_();
    const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Itens');
    const ult = aba.getLastRow();
    if (ult < 2) return;
    const bloco = aba.getRange(2, 1, ult - 1, 10).getDisplayValues();
    const porId = new Map(bloco.map(r => [r[0], r]));
    const dataAtual = agora_();

    listaRespostas.forEach(resp => {
      const id = String(resp && resp.id);
      const r = porId.get(id);
      if (!r || !permitidos.has(id)) return;
      const resposta = String(resp.resposta || '').trim();
      const obs = String(resp.observacao || '').trim().slice(0, 500);
      if (opcoes.indexOf(resposta) < 0) throw new Error('Resposta inválida.');
      if (r[6] === 'Respondido' && r[7] === resposta && r[8] === obs) return;
      r[6] = 'Respondido';
      r[7] = resposta;
      r[8] = obs;
      r[9] = dataAtual;
      alteradas.push(r.slice());
    });

    gravar_(aba, 2, 7, bloco.map(r => r.slice(6, 10)));
    registrarHistorico_(alteradas.map(r => [dataAtual, 'Resposta', grupo, r[0], r[2], r[7], r[8], 'Formulário']));
  });

  if (alteradas.length) notificarCopiasResposta_(grupo, alteradas);
  return { salvos: alteradas.length };
}

function notificarCopiasResposta_(grupo, linhasItens) {
  try {
    const g = mapaClientes_().get(normalizar_(grupo));
    if (!g || !g.copia.length) return;
    MailApp.sendEmail({
      to: g.copia.join(','),
      subject: `[RESPOSTA] ${grupo.replace(/[\r\n]+/g, ' ')} - ${linhasItens.length} item(ns) respondido(s)`,
      htmlBody: htmlEmail_({
        titulo: 'Resposta recebida: itens de cartela',
        cor: '#28a745',
        mensagem: `${grupo} respondeu à solicitação de cartelas em ${linhasItens[0][9]}.`,
        cabecalhos: ['Item', 'Cliente / Filial', 'Tamanho', 'Ordem de compra', 'Cód. Cliente', 'Resposta', 'Observação'],
        linhas: linhasItens.map(r => [r[2], r[1], r[4], r[5], r[3], r[7], r[8] || '-']),
        link: '',
        rodape: 'Cópia informativa: nenhuma ação é necessária.'
      })
    });
  } catch (err) {
    console.error('Falha ao notificar cópias: ' + err.message);
  }
}

// ============================================================================
// 1. GERADOR AUTOMÁTICO DE ABAS E FRASES
// (Pode rodar pelo botão "Executar" no topo ou apenas abrir o WebApp)
// ============================================================================
function configurarPlanilhaAutomaticamente() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Cria e formata a aba "Itens"
  let abaItens = ss.getSheetByName('Itens');
  if (!abaItens) {
    abaItens = ss.insertSheet('Itens');
    abaItens.appendRow(['ID', 'Nome do Item', 'Status', 'Resposta', 'Observacao', 'Data Atualizacao']);
    abaItens.getRange('A1:F1').setFontWeight('bold').setBackground('#d9ead3');
    abaItens.setFrozenRows(1);
    // Exemplo inicial (você pode apagar ou editar na planilha)
    abaItens.appendRow(['101', 'Item de Teste 01 (Apague ou Edite)', 'Não enviado', '', '', '']);
  }

  // 2. Cria e formata a aba "Emails"
  let abaEmails = ss.getSheetByName('Emails');
  if (!abaEmails) {
    abaEmails = ss.insertSheet('Emails');
    abaEmails.appendRow(['Nome', 'Email']);
    abaEmails.getRange('A1:B1').setFontWeight('bold').setBackground('#cfe2f3');
    abaEmails.setFrozenRows(1);
  }

  // 3. Cria e formata a aba "Respostas" já com as frases iniciais
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

  // Ajusta largura das colunas automaticamente
  abaItens.autoResizeColumns(1, 6);
  abaEmails.autoResizeColumns(1, 2);
  abaRespostas.autoResizeColumns(1, 1);
}

// ============================================================================
// 2. ROTEADOR DO WEBAPP (Painel Principal vs Tela do Fornecedor)
// ============================================================================
function doGet(e) {
  // Garante que as abas existam sempre que o sistema for aberto
  configurarPlanilhaAutomaticamente();
  
  const params = e.parameter;
  
  if (params.ids) {
    const template = HtmlService.createTemplateFromFile('Resposta');
    template.ids = params.ids;
    return template.evaluate()
      .setTitle('Responder Solicitação de Cartelas')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Gestão e Solicitação de Cartelas')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================================
// 3. FUNÇÕES DO PAINEL PRINCIPAL E DISPARO DE E-MAIL
// ============================================================================
function obterDadosPainel() {
  configurarPlanilhaAutomaticamente();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaItens = ss.getSheetByName('Itens');
  const abaEmails = ss.getSheetByName('Emails');
  
  const dadosItens = abaItens.getLastRow() > 1 
    ? abaItens.getRange(2, 1, abaItens.getLastRow() - 1, 6).getDisplayValues() 
    : [];
    
  const dadosEmails = abaEmails.getLastRow() > 1 
    ? abaEmails.getRange(2, 1, abaEmails.getLastRow() - 1, 2).getDisplayValues() 
    : [];
  
  const itens = dadosItens
    .filter(r => r[0] || r[1]) // Ignora linhas em branco
    .map(r => ({
      id: r[0],
      nome: r[1],
      status: r[2] || 'Não enviado',
      resposta: r[3] || '',
      observacao: r[4] || '',
      data: r[5] || ''
    }));
  
  const emails = dadosEmails
    .map(r => ({ nome: r[0].trim(), email: r[1].trim() }))
    .filter(e => e.email);
  
  return { itens, emails };
}

function adicionarNovoItem(nomeItem) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('Itens');
  const novoId = new Date().getTime().toString().slice(-5);
  aba.appendRow([novoId, nomeItem, 'Não enviado', '', '', '']);
  return obterDadosPainel();
}

function dispararSolicitacao(idsSelecionados, mensagemPersonalizada, ehReenvio) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaItens = ss.getSheetByName('Itens');
  const abaEmails = ss.getSheetByName('Emails');
  const urlWebApp = ScriptApp.getService().getUrl();
  
  if (abaEmails.getLastRow() < 2) {
    throw new Error('Preencha pelo menos um e-mail na aba "Emails" antes de disparar.');
  }

  const dadosEmails = abaEmails.getRange(2, 2, abaEmails.getLastRow() - 1, 1).getDisplayValues();
  const listaEmails = dadosEmails.map(r => r[0].trim()).filter(Boolean).join(',');
  
  if (!listaEmails) {
    throw new Error('Nenhum endereço de e-mail válido encontrado na Coluna B da aba "Emails".');
  }
  
  const rangeItens = abaItens.getRange(2, 1, abaItens.getLastRow() - 1, 6);
  const valoresItens = rangeItens.getValues();
  const itensParaEnviar = [];
  const dataAtual = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  
  valoresItens.forEach(linha => {
    if (idsSelecionados.includes(String(linha[0]))) {
      linha[2] = ehReenvio ? 'Reenviado - Aguardando' : 'Aguardando Resposta';
      linha[5] = dataAtual;
      itensParaEnviar.push({
        id: linha[0],
        nome: linha[1],
        respAnterior: linha[3],
        obsAnterior: linha[4]
      });
    }
  });
  
  rangeItens.setValues(valoresItens);
  
  const linkFormulario = `${urlWebApp}?ids=${idsSelecionados.join(',')}`;
  const assunto = ehReenvio 
    ? `[REVISÃO SOLICITADA] Lista de Itens para Cartela` 
    : `[SOLICITAÇÃO] Lista de Itens que necessitamos de cartela`;

  let htmlCorpo = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; color: #333; border: 1px solid #ddd; border-radius: 8px; overflow: hidden;">
      <div style="background-color: ${ehReenvio ? '#d9534f' : '#0275d8'}; color: white; padding: 16px;">
        <h2 style="margin: 0; font-size: 18px;">${assunto}</h2>
      </div>
      <div style="padding: 20px;">
        <p style="font-size: 15px; line-height: 1.5;">${mensagemPersonalizada.replace(/\n/g, '<br>')}</p>
        <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
          <thead>
            <tr style="background-color: #f7f7f7; text-align: left;">
              <th style="padding: 10px; border-bottom: 2px solid #ddd;">Item Solicitado</th>
              ${ehReenvio ? '<th style="padding: 10px; border-bottom: 2px solid #ddd;">Retorno Anterior</th>' : ''}
            </tr>
          </thead>
          <tbody>
  `;
  
  itensParaEnviar.forEach(item => {
    htmlCorpo += `
      <tr>
        <td style="padding: 10px; border-bottom: 1px solid #eee;"><strong>${item.nome}</strong> (Cód: ${item.id})</td>
        ${ehReenvio ? `<td style="padding: 10px; border-bottom: 1px solid #eee; color: #777;">${item.respAnterior || 'Sem resposta'} ${item.obsAnterior ? `- <em>${item.obsAnterior}</em>` : ''}</td>` : ''}
      </tr>
    `;
  });
  
  htmlCorpo += `
          </tbody>
        </table>
        <div style="margin-top: 25px; text-align: center;">
          <a href="${linkFormulario}" style="background-color: #28a745; color: white; padding: 14px 24px; text-decoration: none; font-weight: bold; border-radius: 6px; display: inline-block; font-size: 15px;">
            Responder Itens e Adicionar Observações
          </a>
        </div>
        <p style="font-size: 12px; color: #888; margin-top: 20px; text-align: center;">
          Clique no botão acima para selecionar a resposta pré-definida e preencher observações para cada item.
        </p>
      </div>
    </div>
  `;
  
  MailApp.sendEmail({
    to: listaEmails,
    subject: assunto,
    htmlBody: htmlCorpo
  });
  
  return obterDadosPainel();
}

// ============================================================================
// 4. FUNÇÕES DA TELA DE RESPOSTA DO DESTINATÁRIO
// ============================================================================
function obterItensPorIds(idsString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ids = idsString.split(',');
  const dadosItens = obterDadosPainel().itens;
  const itensFiltrados = dadosItens.filter(item => ids.includes(String(item.id)));
  
  const abaRespostas = ss.getSheetByName('Respostas');
  let opcoes = [];
  
  if (abaRespostas && abaRespostas.getLastRow() > 1) {
    opcoes = abaRespostas
      .getRange(2, 1, abaRespostas.getLastRow() - 1, 1)
      .getDisplayValues()
      .map(r => r[0].trim())
      .filter(Boolean);
  }
  
  if (opcoes.length === 0) {
    opcoes = ["Cartela disponível - Envio imediato", "Indisponível", "Verificar estoque"];
  }
  
  return {
    itens: itensFiltrados,
    opcoes: opcoes
  };
}

function salvarRespostasFornecedor(listaRespostas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaItens = ss.getSheetByName('Itens');
  const rangeItens = abaItens.getRange(2, 1, abaItens.getLastRow() - 1, 6);
  const valores = rangeItens.getValues();
  const dataAtual = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  
  listaRespostas.forEach(resp => {
    for (let i = 0; i < valores.length; i++) {
      if (String(valores[i][0]) === String(resp.id)) {
        valores[i][2] = 'Respondido';
        valores[i][3] = resp.resposta;
        valores[i][4] = resp.observacao;
        valores[i][5] = dataAtual;
        break;
      }
    }
  });
  
  rangeItens.setValues(valores);
  return true;
}

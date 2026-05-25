const CONFIG = {
  spreadsheetId: '1r7VuxUTRUzbTiAGmrhex5L5L2kSpSUCcatTCKLZAM9k',
  sheetName: 'Maio/2026',
  timezone: 'America/Sao_Paulo',
  sendHour: 8,
  senderName: 'Central de Operacoes e Performance - Devs',
  subjectPrefix: 'Resumo diario de demandas',
};

const DEV_EMAILS = {
  // Preencha com o nome exatamente como aparece na coluna "Dev Responsavel".
  // 'Nome do Dev': 'email@empresa.com',
};

const STATUS_ENTREGUES = [
  'APROVADO',
  'OMOLOGACAO CLIENTE',
  'HOMOLOGACAO CLIENTE',
  'LIBERADO DEPLOY',
  'AGUARDANDO PRODUCAO',
  'LIBERADO FORMALIZACAO',
  'CONCLUIDO',
];

const CORES_STATUS = {
  'EM QA': '#d5a6bd',
  'NECESSITA CORRECAO': '#ea4335',
  'EM DESENVOLVIMENTO': '#ffe599',
  'APROVADO': '#b6d7a8',
  'HOMOLOGACAO CLIENTE': '#6fa8dc',
  'OMOLOGACAO CLIENTE': '#6fa8dc',
  'LIBERADO DEPLOY': '#999999',
  'AGUARDANDO PENDENCIAS': '#674ea7',
  'PARA FAZER': '#f6b26b',
  'AGUARDANDO PRODUCAO': '#e6b8af',
  'LIBERADO FORMALIZACAO': '#c9daf8',
  'IMPEDIMENTOS': '#783f04',
  'CONCLUIDO': '#10b981',
  'PAUSADA': '#f59e0b',
};

function enviarRelatoriosDiariosDevs() {
  const dados = carregarDadosTratados_();
  const devs = [...new Set(dados.map(item => item['Dev Responsável']).filter(Boolean))].sort();
  const hoje = Utilities.formatDate(new Date(), CONFIG.timezone, 'dd/MM/yyyy');

  devs.forEach(dev => {
    const email = DEV_EMAILS[dev];
    if (!email) {
      console.warn(`E-mail nao configurado para o dev: ${dev}`);
      return;
    }

    const dadosDev = dados.filter(item => item['Dev Responsável'] === dev);
    const htmlBody = gerarCorpoEmailDev_(dev, dadosDev, hoje);
    const plainBody = gerarTextoSimples_(dev, dadosDev, hoje);

    GmailApp.sendEmail(
      email,
      `${CONFIG.subjectPrefix} - ${dev} - ${hoje}`,
      plainBody,
      {
        htmlBody,
        name: CONFIG.senderName,
      }
    );
  });
}

function instalarGatilhoDiario() {
  removerGatilhosRelatorio_();

  ScriptApp.newTrigger('enviarRelatoriosDiariosDevs')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.sendHour)
    .inTimezone(CONFIG.timezone)
    .create();
}

function removerGatilhosRelatorio_() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'enviarRelatoriosDiariosDevs') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function carregarDadosTratados_() {
  const sheet = SpreadsheetApp.openById(CONFIG.spreadsheetId).getSheetByName(CONFIG.sheetName);
  if (!sheet) {
    throw new Error(`Aba nao encontrada: ${CONFIG.sheetName}`);
  }

  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.shift().map(header => header.trim());

  return values
    .map(row => objetoPorCabecalho_(headers, row))
    .filter(item => item.Projeto && item.Projeto.trim() !== '')
    .map(tratarLinha_);
}

function objetoPorCabecalho_(headers, row) {
  return headers.reduce((acc, header, index) => {
    acc[header] = row[index];
    return acc;
  }, {});
}

function tratarLinha_(item) {
  item.HorasTratadas = parseFloat(String(item.Horas || '0').replace(',', '.')) || 0;
  item.StatusPadronizado = normalizarTexto_(item.Status || '');

  const dataInicial = extrairData_(item['Data Prometida Inicial']);
  const dataFinal = extrairData_(item['Data Prometida Final']);
  const dataEntrega = extrairData_(item['Data Entregue']) || extrairData_(item['Data QA ok']) || extrairData_(item['Data Cliente OK']);
  const dataPrazoAlvo = dataFinal || dataInicial;

  item.DataPrazoFormatada = dataPrazoAlvo
    ? Utilities.formatDate(dataPrazoAlvo, CONFIG.timezone, 'dd/MM/yyyy')
    : '-';

  item.IsEntregueSucesso = STATUS_ENTREGUES.includes(item.StatusPadronizado);
  item.AtrasoTratado = calcularDiasAtraso_(dataPrazoAlvo, dataEntrega, item.IsEntregueSucesso);

  const colPostergado = normalizarTexto_(item['Postergação'] || '');
  item.FoiPostergado = (dataInicial && dataFinal && dataInicial.getTime() !== dataFinal.getTime()) || colPostergado === 'SIM';

  return item;
}

function gerarCorpoEmailDev_(dev, dados, dataReferencia) {
  const metricas = calcularMetricas_(dados);
  const rankingAtrasos = dados
    .filter(item => item.AtrasoTratado > 0)
    .sort((a, b) => b.AtrasoTratado - a.AtrasoTratado)
    .slice(0, 8);

  const statusResumo = agruparContagem_(dados, 'Status');
  const projetosResumo = agruparContagem_(dados, 'Projeto');
  const backlog = dados.filter(item => !item.IsEntregueSucesso);

  return `
    <div style="font-family:Segoe UI,Arial,sans-serif;background:#f8fafc;padding:24px;color:#1e293b;">
      <div style="max-width:860px;margin:0 auto;background:#ffffff;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;">
        <div style="padding:22px 24px;border-bottom:1px solid #e2e8f0;">
          <h1 style="margin:0;font-size:22px;color:#0f172a;">Resumo diario de demandas</h1>
          <p style="margin:6px 0 0;color:#64748b;">${escapeHtml_(dev)} - ${dataReferencia}</p>
        </div>

        ${renderMetricas_(metricas)}

        <div style="padding:0 24px 22px;">
          <h2 style="font-size:16px;margin:22px 0 10px;color:#0f172a;">Demandas em atraso</h2>
          ${renderTabelaAtrasos_(rankingAtrasos)}

          <h2 style="font-size:16px;margin:24px 0 10px;color:#0f172a;">Distribuicao de status</h2>
          ${renderTabelaStatus_(statusResumo)}

          <h2 style="font-size:16px;margin:24px 0 10px;color:#0f172a;">Projetos por volume</h2>
          ${renderTabelaSimples_(projetosResumo, 'Projeto', 'Qtd.')}

          <h2 style="font-size:16px;margin:24px 0 10px;color:#0f172a;">Backlog pendente</h2>
          ${renderTabelaBacklog_(backlog)}
        </div>
      </div>
    </div>
  `;
}

function calcularMetricas_(dados) {
  const totalDemandas = dados.length;
  const entregues = dados.filter(item => item.IsEntregueSucesso);
  const atrasadas = dados.filter(item => item.AtrasoTratado > 0);
  const totalHoras = somarHoras_(dados);
  const horasEntregues = somarHoras_(entregues);
  const mediaAtraso = atrasadas.length
    ? (atrasadas.reduce((acc, item) => acc + item.AtrasoTratado, 0) / atrasadas.length).toFixed(1)
    : '0';
  const postergadas = dados.filter(item => item.FoiPostergado).length;

  return {
    totalDemandas,
    entregues: entregues.length,
    pendentes: totalDemandas - entregues.length,
    totalHoras: totalHoras.toFixed(1),
    horasEntregues: horasEntregues.toFixed(1),
    mediaAtraso,
    taxaPostergacao: totalDemandas ? Math.round((postergadas / totalDemandas) * 100) : 0,
  };
}

function renderMetricas_(metricas) {
  const cards = [
    ['Demandas totais', metricas.totalDemandas],
    ['Entregues', metricas.entregues],
    ['Pendentes', metricas.pendentes],
    ['Horas previstas', `${metricas.totalHoras}h`],
    ['Horas entregues', `${metricas.horasEntregues}h`],
    ['Media atraso', `${metricas.mediaAtraso}d`],
    ['Postergacao', `${metricas.taxaPostergacao}%`],
  ];

  return `
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="padding:18px 18px 4px;">
      ${cards.reduce((html, card, index) => {
        if (index % 3 === 0) html += '<tr>';
        html += `
          <td width="33.33%" style="padding:6px;">
            <div style="border:1px solid #e2e8f0;border-left:4px solid #3b82f6;border-radius:8px;padding:12px;background:#ffffff;">
              <div style="font-size:11px;text-transform:uppercase;color:#64748b;font-weight:700;">${card[0]}</div>
              <div style="font-size:22px;font-weight:700;color:#0f172a;margin-top:4px;">${card[1]}</div>
            </div>
          </td>
        `;
        if (index % 3 === 2 || index === cards.length - 1) html += '</tr>';
        return html;
      }, '')}
    </table>
  `;
}

function renderTabelaAtrasos_(itens) {
  if (!itens.length) {
    return '<p style="margin:0;color:#10b981;font-weight:700;">Nenhuma demanda em atraso.</p>';
  }

  const rows = itens.map(item => `
    <tr>
      <td style="${td_()}">${escapeHtml_(item.Tarefa || '-')}</td>
      <td style="${td_()}">${escapeHtml_(item.Projeto || '-')}</td>
      <td style="${td_()}">${item.DataPrazoFormatada}</td>
      <td style="${td_()}">${renderBadgeStatus_(item.Status || '-')}</td>
      <td style="${td_()}font-weight:700;color:#ef4444;">${item.AtrasoTratado} dias</td>
    </tr>
  `).join('');

  return `
    <table width="100%" cellpadding="0" cellspacing="0" style="${table_()}">
      <tr>
        <th style="${th_()}">Tarefa</th>
        <th style="${th_()}">Projeto</th>
        <th style="${th_()}">Prazo</th>
        <th style="${th_()}">Status</th>
        <th style="${th_()}">Atraso</th>
      </tr>
      ${rows}
    </table>
  `;
}

function renderTabelaStatus_(itens) {
  const rows = itens.map(([status, qtd]) => `
    <tr>
      <td style="${td_()}">${renderBadgeStatus_(status)}</td>
      <td style="${td_()}font-weight:700;">${qtd}</td>
    </tr>
  `).join('');

  return `
    <table width="100%" cellpadding="0" cellspacing="0" style="${table_()}">
      <tr><th style="${th_()}">Status</th><th style="${th_()}">Qtd.</th></tr>
      ${rows || `<tr><td style="${td_()}" colspan="2">Sem dados.</td></tr>`}
    </table>
  `;
}

function renderTabelaSimples_(itens, colunaA, colunaB) {
  const rows = itens.slice(0, 10).map(([label, value]) => `
    <tr>
      <td style="${td_()}">${escapeHtml_(label || '-')}</td>
      <td style="${td_()}font-weight:700;">${value}</td>
    </tr>
  `).join('');

  return `
    <table width="100%" cellpadding="0" cellspacing="0" style="${table_()}">
      <tr><th style="${th_()}">${colunaA}</th><th style="${th_()}">${colunaB}</th></tr>
      ${rows || `<tr><td style="${td_()}" colspan="2">Sem dados.</td></tr>`}
    </table>
  `;
}

function renderTabelaBacklog_(itens) {
  const rows = itens
    .sort((a, b) => b.HorasTratadas - a.HorasTratadas)
    .slice(0, 10)
    .map(item => `
      <tr>
        <td style="${td_()}">${escapeHtml_(item.Tarefa || '-')}</td>
        <td style="${td_()}">${escapeHtml_(item.Projeto || '-')}</td>
        <td style="${td_()}">${item.HorasTratadas.toFixed(1)}h</td>
        <td style="${td_()}">${renderBadgeStatus_(item.Status || '-')}</td>
      </tr>
    `).join('');

  return `
    <table width="100%" cellpadding="0" cellspacing="0" style="${table_()}">
      <tr>
        <th style="${th_()}">Tarefa</th>
        <th style="${th_()}">Projeto</th>
        <th style="${th_()}">Horas</th>
        <th style="${th_()}">Status</th>
      </tr>
      ${rows || `<tr><td style="${td_()}" colspan="4">Sem backlog pendente.</td></tr>`}
    </table>
  `;
}

function renderBadgeStatus_(status) {
  const statusNormalizado = normalizarTexto_(status);
  const bg = CORES_STATUS[statusNormalizado] || '#f1f5f9';
  const fg = ['IMPEDIMENTOS', 'AGUARDANDO PENDENCIAS', 'NECESSITA CORRECAO', 'LIBERADO DEPLOY'].includes(statusNormalizado)
    ? '#ffffff'
    : '#1e293b';

  return `<span style="display:inline-block;background:${bg};color:${fg};padding:4px 8px;border-radius:6px;font-size:11px;font-weight:700;text-transform:uppercase;">${escapeHtml_(status)}</span>`;
}

function gerarTextoSimples_(dev, dados, dataReferencia) {
  const metricas = calcularMetricas_(dados);
  return [
    `Resumo diario de demandas - ${dev} - ${dataReferencia}`,
    `Demandas totais: ${metricas.totalDemandas}`,
    `Entregues: ${metricas.entregues}`,
    `Pendentes: ${metricas.pendentes}`,
    `Horas previstas: ${metricas.totalHoras}h`,
    `Horas entregues: ${metricas.horasEntregues}h`,
    `Media atraso: ${metricas.mediaAtraso}d`,
    `Postergacao: ${metricas.taxaPostergacao}%`,
  ].join('\n');
}

function agruparContagem_(dados, campo) {
  const agrupado = dados.reduce((acc, item) => {
    const chave = item[campo] || `Sem ${campo}`;
    acc[chave] = (acc[chave] || 0) + 1;
    return acc;
  }, {});

  return Object.entries(agrupado).sort((a, b) => b[1] - a[1]);
}

function somarHoras_(dados) {
  return dados.reduce((acc, item) => acc + item.HorasTratadas, 0);
}

function calcularDiasAtraso_(dataPrazoAlvo, dataEntrega, isEntregue) {
  if (!dataPrazoAlvo) return 0;

  const dataComparacao = dataEntrega || (!isEntregue ? hojeSemHora_() : null);
  if (!dataComparacao) return 0;

  const diff = dataComparacao.getTime() - dataPrazoAlvo.getTime();
  const dias = Math.round(diff / (1000 * 60 * 60 * 24));
  return dias > 0 ? dias : 0;
}

function extrairData_(valor) {
  if (!valor || String(valor).trim() === '' || String(valor).trim() === '-') return null;

  const texto = String(valor).trim();
  const partes = texto.split('/');
  let data;

  if (partes.length === 3) {
    data = new Date(Number(partes[2]), Number(partes[1]) - 1, Number(partes[0]));
  } else {
    data = new Date(texto);
  }

  if (!data || isNaN(data.getTime())) return null;
  data.setHours(0, 0, 0, 0);
  return data;
}

function hojeSemHora_() {
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  return hoje;
}

function normalizarTexto_(texto) {
  return String(texto || '')
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function table_() {
  return 'border-collapse:collapse;border:1px solid #e2e8f0;font-size:13px;';
}

function th_() {
  return 'text-align:left;background:#f8fafc;color:#64748b;text-transform:uppercase;font-size:11px;letter-spacing:.4px;padding:9px;border-bottom:1px solid #e2e8f0;';
}

function td_() {
  return 'padding:9px;border-bottom:1px solid #f1f5f9;color:#1e293b;vertical-align:top;';
}

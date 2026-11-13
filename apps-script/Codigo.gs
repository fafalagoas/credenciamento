const CONFIG = {
  ABA: 'Respostas',
  COL_EMAIL: 2,
  COL_NOME: 3,
  COL_FUNCAO: 6,
  COL_VEICULO: 7,
  COL_JOGO: 9,
  COL_STATUS: 10,
  COL_EMAIL_ENVIADO: 11,
  COL_DATA_ENVIO: 12
};


// ======================================================
// AO EDITAR O STATUS
// ======================================================

function processarCredenciamento(e) {
 if (!e || !e.range) return;
 const range=e.range, sheet=range.getSheet();
 if (sheet.getName()!==CONFIG.ABA || range.getLastColumn()<CONFIG.COL_STATUS || range.getColumn()>CONFIG.COL_STATUS) return;
 const lock=LockService.getScriptLock(); lock.waitLock(30000);
 try {
  for(let row=Math.max(2,range.getRow()); row<=range.getLastRow(); row++) {
   const status=String(sheet.getRange(row,CONFIG.COL_STATUS).getValue()).trim().toUpperCase();
   try {
    if(status==='APROVADO') enviarAprovacaoDaLinha_(sheet,row);
    else if(status==='NÃO APROVADO') enviarNaoAprovacaoDaLinha_(sheet,row);
   } catch(err) { console.error('Comunicação pendente na linha '+row+': '+err.message); }
  }
 } finally { SpreadsheetApp.flush(); lock.releaseLock(); }
}


function inicializarNovaInscricao(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();

  if (sheet.getName() !== CONFIG.ABA) return;

  const row = e.range.getRow();

  if (row <= 1) return;

  sheet
    .getRange(row, CONFIG.COL_STATUS)
    .setValue('PENDENTE');

  sheet
    .getRange(row, CONFIG.COL_EMAIL_ENVIADO)
    .setValue('NÃO');

  sheet
    .getRange(row, CONFIG.COL_DATA_ENVIO)
    .clearContent();
}


// ======================================================
// APROVAÇÃO
// ======================================================

function enviarAprovacaoDaLinha_(sheet, row) {

  const dados = obterDadosDaLinha_(sheet, row);

  if (dados.status !== 'APROVADO') return;

  validarDados_(dados, row, true);

  // Impede duplicidade
  if (dados.emailEnviado === 'SIM') return;
  if (!['', 'NÃO'].includes(dados.emailEnviado)) throw new Error('Envio pendente de conferência. Não reenviar automaticamente.');

  const primeiroNome = obterPrimeiroNome_(dados.nome);
  const detalhesJogo = obterDetalhesJogo_(dados.jogo, sheet.getRange(row, 1, 1, COL_IDS).getValues()[0]);

  const assunto =
    'Credenciamento aprovado | ' +
    detalhesJogo.assunto +
    ' | ' + eventoDaLinha_(sheet, row);


  const texto = `
CREDENCIAMENTO APROVADO

Olá, ${primeiroNome}!

Sua solicitação de credenciamento para a ${eventoDaLinha_(sheet, row)} foi aprovada pela Federação Alagoana de Futebol (FAF).

CREDENCIADO

Nome: ${dados.nome}
Função: ${dados.funcao}
Veículo: ${dados.veiculo}

PARTIDA(S) CREDENCIADA(S)

${detalhesJogo.texto}

O credenciamento é pessoal e intransferível.

No acesso ao estádio, apresente documento oficial com foto e siga as orientações da equipe de Comunicação da FAF.

A aprovação é válida exclusivamente para a(s) partida(s) informada(s) neste e-mail.

Diretoria de Comunicação e Marketing
Federação Alagoana de Futebol (FAF)
`.trim();


  const html = `
  <div style="
    margin:0;
    padding:0;
    background:#f3f4f6;
    font-family:Arial,Helvetica,sans-serif;
    color:#181818;
  ">

    <div style="
      max-width:620px;
      margin:0 auto;
      padding:32px 16px;
    ">

      <div style="
        background:#ffffff;
        border-radius:12px;
        overflow:hidden;
        border:1px solid #e5e5e5;
      ">

        <div style="
          background:#111111;
          padding:30px 32px;
          color:#ffffff;
        ">

          <div style="
            font-size:11px;
            font-weight:bold;
            letter-spacing:1.5px;
            color:#cfcfcf;
            margin-bottom:12px;
          ">
            FEDERAÇÃO ALAGOANA DE FUTEBOL
          </div>

          <div style="
            font-size:27px;
            font-weight:bold;
            line-height:1.2;
          ">
            Credenciamento<br>
            aprovado
          </div>

        </div>


        <div style="padding:32px;">

          <div style="margin-bottom:26px;">

            <span style="
              display:inline-block;
              background:#e8f5e9;
              color:#166534;
              border:1px solid #bbdfc1;
              border-radius:20px;
              padding:7px 14px;
              font-size:11px;
              font-weight:bold;
              letter-spacing:.7px;
            ">
              APROVADO
            </span>

          </div>


          <p style="
            font-size:16px;
            margin:0 0 12px 0;
          ">
            Olá, <strong>${escaparHtml_(primeiroNome)}</strong>!
          </p>


          <p style="
            font-size:15px;
            line-height:1.65;
            color:#444444;
            margin:0 0 28px 0;
          ">
            Sua solicitação de credenciamento para a
            <strong>${escaparHtml_(eventoDaLinha_(sheet, row))}</strong>
            foi aprovada pela
            <strong>Federação Alagoana de Futebol (FAF)</strong>.
          </p>


          <div style="
            background:#f7f7f7;
            border-radius:9px;
            padding:20px;
            margin-bottom:30px;
          ">

            <div style="
              font-size:11px;
              font-weight:bold;
              letter-spacing:1px;
              color:#777777;
              margin-bottom:13px;
            ">
              CREDENCIADO
            </div>

            <div style="
              font-size:18px;
              font-weight:bold;
              margin-bottom:6px;
            ">
              ${escaparHtml_(dados.nome)}
            </div>

            <div style="
              font-size:14px;
              color:#555555;
            ">
              ${escaparHtml_(dados.funcao)}
              ${
                dados.veiculo
                  ? ' &nbsp;•&nbsp; ' + escaparHtml_(dados.veiculo)
                  : ''
              }
            </div>

          </div>


          <div style="
            font-size:11px;
            font-weight:bold;
            letter-spacing:1px;
            color:#777777;
            margin-bottom:12px;
          ">
            PARTIDA(S) CREDENCIADA(S)
          </div>

          ${detalhesJogo.html}


          <div style="
            margin-top:30px;
            padding-top:24px;
            border-top:1px solid #eeeeee;
          ">

            <p style="
              font-size:14px;
              line-height:1.65;
              color:#444444;
              margin:0 0 12px 0;
            ">
              O credenciamento é
              <strong>pessoal e intransferível</strong>.
            </p>

            <p style="
              font-size:14px;
              line-height:1.65;
              color:#444444;
              margin:0;
            ">
              No acesso ao estádio, apresente
              <strong>documento oficial com foto</strong>
              e siga as orientações da equipe de Comunicação da FAF.
            </p>

          </div>

        </div>


        <div style="
          background:#f7f7f7;
          border-top:1px solid #eeeeee;
          padding:22px 32px;
        ">

          <div style="
            font-size:13px;
            font-weight:bold;
            color:#222222;
          ">
            Diretoria de Comunicação e Marketing
          </div>

          <div style="
            font-size:12px;
            color:#777777;
            margin-top:4px;
          ">
            Federação Alagoana de Futebol (FAF)
          </div>

        </div>

      </div>

    </div>

  </div>`;


  enviarEmailControlado_(sheet, row,
    dados.email,
    assunto,
    texto,
    {
      htmlBody: html,
      name: 'Federação Alagoana de Futebol'
    }
  );


  registrarEnvio_(sheet, row);
}


// ======================================================
// NÃO APROVAÇÃO
// ======================================================

function enviarNaoAprovacaoDaLinha_(sheet, row) {

  const dados = obterDadosDaLinha_(sheet, row);

  if (dados.status !== 'NÃO APROVADO') return;

  validarDados_(dados, row, false);

  // Impede duplicidade
  if (dados.emailEnviado === 'SIM') return;
  if (!['', 'NÃO'].includes(dados.emailEnviado)) throw new Error('Envio pendente de conferência. Não reenviar automaticamente.');

  const primeiroNome = obterPrimeiroNome_(dados.nome);

  const assunto =
    'Credenciamento de imprensa | ' + eventoDaLinha_(sheet, row);


  const texto = `
CREDENCIAMENTO DE IMPRENSA

Olá, ${primeiroNome}!

Agradecemos o seu interesse na cobertura da ${eventoDaLinha_(sheet, row)}.

Após análise da solicitação, informamos que o credenciamento não foi aprovado para esta oportunidade.

O envio do formulário representa uma solicitação de credenciamento e não garante automaticamente o acesso ao evento.

Agradecemos a compreensão e o interesse na cobertura das competições promovidas pela Federação Alagoana de Futebol.

Diretoria de Comunicação e Marketing
Federação Alagoana de Futebol (FAF)
`.trim();


  const html = `
  <div style="
    margin:0;
    padding:0;
    background:#f3f4f6;
    font-family:Arial,Helvetica,sans-serif;
    color:#181818;
  ">

    <div style="
      max-width:620px;
      margin:0 auto;
      padding:32px 16px;
    ">

      <div style="
        background:#ffffff;
        border-radius:12px;
        overflow:hidden;
        border:1px solid #e5e5e5;
      ">

        <div style="
          background:#111111;
          padding:30px 32px;
          color:#ffffff;
        ">

          <div style="
            font-size:11px;
            font-weight:bold;
            letter-spacing:1.5px;
            color:#cfcfcf;
            margin-bottom:12px;
          ">
            FEDERAÇÃO ALAGOANA DE FUTEBOL
          </div>

          <div style="
            font-size:27px;
            font-weight:bold;
            line-height:1.2;
          ">
            Credenciamento<br>
            de imprensa
          </div>

        </div>


        <div style="padding:32px;">

          <div style="margin-bottom:26px;">

            <span style="
              display:inline-block;
              background:#f3f4f6;
              color:#555555;
              border:1px solid #dddddd;
              border-radius:20px;
              padding:7px 14px;
              font-size:11px;
              font-weight:bold;
              letter-spacing:.7px;
            ">
              NÃO APROVADO
            </span>

          </div>


          <p style="
            font-size:16px;
            margin:0 0 12px 0;
          ">
            Olá, <strong>${escaparHtml_(primeiroNome)}</strong>!
          </p>


          <p style="
            font-size:15px;
            line-height:1.65;
            color:#444444;
            margin:0 0 18px 0;
          ">
            Agradecemos o seu interesse na cobertura da
            <strong>${escaparHtml_(eventoDaLinha_(sheet, row))}</strong>.
          </p>


          <p style="
            font-size:15px;
            line-height:1.65;
            color:#444444;
            margin:0 0 18px 0;
          ">
            Após análise da solicitação, informamos que o
            <strong>credenciamento não foi aprovado</strong>
            para esta oportunidade.
          </p>


          <div style="
            background:#f7f7f7;
            border-radius:9px;
            padding:18px 20px;
            margin:26px 0;
          ">

            <div style="
              font-size:14px;
              line-height:1.6;
              color:#555555;
            ">
              O envio do formulário representa uma solicitação
              de credenciamento e não garante automaticamente
              o acesso ao evento.
            </div>

          </div>


          <p style="
            font-size:14px;
            line-height:1.65;
            color:#555555;
            margin:0;
          ">
            Agradecemos a compreensão e o interesse na cobertura
            das competições promovidas pela Federação Alagoana de Futebol.
          </p>

        </div>


        <div style="
          background:#f7f7f7;
          border-top:1px solid #eeeeee;
          padding:22px 32px;
        ">

          <div style="
            font-size:13px;
            font-weight:bold;
            color:#222222;
          ">
            Diretoria de Comunicação e Marketing
          </div>

          <div style="
            font-size:12px;
            color:#777777;
            margin-top:4px;
          ">
            Federação Alagoana de Futebol (FAF)
          </div>

        </div>

      </div>

    </div>

  </div>`;


  enviarEmailControlado_(sheet, row,
    dados.email,
    assunto,
    texto,
    {
      htmlBody: html,
      name: 'Federação Alagoana de Futebol'
    }
  );


  registrarEnvio_(sheet, row);
}


// ======================================================
// DADOS DA LINHA
// ======================================================

function obterDadosDaLinha_(sheet, row) {

  return {

    email: String(
      sheet.getRange(row, CONFIG.COL_EMAIL).getValue()
    ).trim(),

    nome: String(
      sheet.getRange(row, CONFIG.COL_NOME).getValue()
    ).trim(),

    funcao: String(
      sheet.getRange(row, CONFIG.COL_FUNCAO).getValue()
    ).trim(),

    veiculo: String(
      sheet.getRange(row, CONFIG.COL_VEICULO).getValue()
    ).trim(),

    jogo: String(
      sheet.getRange(row, CONFIG.COL_JOGO).getValue()
    ).trim(),

    status: String(
      sheet.getRange(row, CONFIG.COL_STATUS).getValue()
    ).trim().toUpperCase(),

    emailEnviado: String(
      sheet.getRange(row, CONFIG.COL_EMAIL_ENVIADO).getValue()
    ).trim().toUpperCase()

  };
}


// ======================================================
// VALIDAÇÃO
// ======================================================

function validarDados_(dados, row, exigirJogo) {

  if (!dados.email) {
    throw new Error(
      'A inscrição da linha ' +
      row +
      ' não possui endereço de e-mail.'
    );
  }

  if (exigirJogo && !dados.jogo) {
    throw new Error(
      'A inscrição da linha ' +
      row +
      ' não possui partida selecionada.'
    );
  }
}


// ======================================================
// REGISTRA ENVIO
// ======================================================

function registrarEnvio_(sheet, row) {

  sheet
    .getRange(row, CONFIG.COL_EMAIL_ENVIADO)
    .setValue('SIM');

  sheet
    .getRange(row, CONFIG.COL_DATA_ENVIO)
    .setValue(new Date())
    .setNumberFormat('dd/MM/yyyy HH:mm:ss');
}


// ======================================================
// PRIMEIRO NOME
// ======================================================

function obterPrimeiroNome_(nome) {

  if (!nome) return 'profissional';

  return nome.split(/\s+/)[0];
}



// ======================================================
// SEGURANÇA HTML
// ======================================================

function escaparHtml_(valor) {

  return String(valor || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');

}

// =====================================================================
// CENTRAL FAF: adaptações sobre o fluxo validado do Alagoano Sub-20
// O fluxo acima (e-mails, disparo ao editar STATUS, trava) é o original.
// ATENÇÃO: no app web, toda função sem "_" no fim pode ser chamada por qualquer visitante.
// =====================================================================
const ABA_PARTIDAS = 'Partidas';
const ABA_CHECKINS = 'Checkins';
const CABECALHO_RESPOSTAS = ['Carimbo de data/hora','Endereço de e-mail','Nome completo','CPF','Telefone/WhatsApp','Função/Atividade','Nome do veículo de comunicação','Tipo de veículo/trabalho (colete)','Vai para qual jogo?','STATUS','E-MAIL ENVIADO','DATA DO ENVIO','PROTOCOLO','FOTO','ASSOCIACAO','FOTO_CREDENCIAL','IDS_PARTIDAS'];
const COL_PROTOCOLO = 13;
const COL_FOTO = 14;
const COL_ASSOCIACAO = 15;
const COL_FOTO_CREDENCIAL = 16;
const COL_IDS = 17;
const DIAS_RETENCAO_FOTO = 30;
const MESES_ = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];

// Rodar uma vez pelo editor: cria a aba de respostas e liga o disparo de e-mails ao editar o STATUS.
function instalar_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(CONFIG.ABA) || ss.insertSheet(CONFIG.ABA);
  if (aba.getLastRow() === 0) aba.appendRow(CABECALHO_RESPOSTAS);
  const existe = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'processarCredenciamento');
  if (!existe) ScriptApp.newTrigger('processarCredenciamento').forSpreadsheet(ss).onEdit().create();
  [[COL_FOTO, 'FOTO'], [COL_ASSOCIACAO, 'ASSOCIACAO'], [COL_FOTO_CREDENCIAL, 'FOTO_CREDENCIAL'], [COL_IDS, 'IDS_PARTIDAS']]
    .forEach(([c, n]) => { if (aba.getRange(1, c).getValue() !== n) aba.getRange(1, c).setValue(n); });
  if (!ss.getSheetByName(ABA_COMPETICOES)) {
    const c = ss.insertSheet(ABA_COMPETICOES);
    c.appendRow(['ID_URANO', 'NOME_EXIBICAO', 'MODELO', 'AJUSTES', 'ATIVA', 'ULTIMA_SINCRONIZACAO']);
    c.appendRow(['COPAALAGOAS20', 'Final da Copa Alagoas Sub-20 2026', 'BASE', '', 'SIM', '']);
  }
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'rotinaDiaria')) ScriptApp.newTrigger('rotinaDiaria').timeBased().everyDays(1).atHour(3).create();
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'registrarHistorico')) ScriptApp.newTrigger('registrarHistorico').forSpreadsheet(ss).onEdit().create();
  protegerColunas_(aba);
}

// Colunas que só o sistema preenche: quem editar à mão recebe um aviso antes (STATUS continua livre).
function protegerColunas_(aba) {
  const jaTem = aba.getProtections(SpreadsheetApp.ProtectionType.RANGE).some(p => p.getDescription() === 'Sistema de credenciamento');
  if (jaTem) return;
  [[1, 9], [11, 7]].forEach(([coluna, qtd]) => {
    aba.getRange(1, coluna, aba.getMaxRows(), qtd).protect().setDescription('Sistema de credenciamento').setWarningOnly(true);
  });
  aba.hideColumns(4); // CPF fora da vista no dia a dia; quem precisar, reexibe
}

function linhas_(nome) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
  if (!aba || aba.getLastRow() < 2) return [];
  const v = aba.getDataRange().getValues();
  const cab = v.shift();
  return v.map(r => Object.fromEntries(cab.map((c, k) => [c, r[k]])));
}

function fmt_(v, padrao) {
  return v instanceof Date ? Utilities.formatDate(v, 'America/Maceio', padrao) : String(v);
}

// Partida com campos prontos para texto: "11/11", "15h", "14h45", "11 de novembro de 2026".
function partida_(p) {
  const data = fmt_(p.DATA, 'dd/MM/yyyy'), hora = fmt_(p.HORA, 'HH:mm');
  const [d, m, a] = data.split('/'), [h, min] = hora.split(':');
  return {
    id: String(p.ID_PARTIDA || '').trim(), competicao: String(p.COMPETICAO), jogo: p.JOGO, mandante: p.MANDANTE, visitante: p.VISITANTE,
    local: p.LOCAL, aberta: String(p.INSCRICOES_ABERTAS).toUpperCase() === 'SIM' && dataValida_(data) && /^([01]\d|2[0-3]):[0-5]\d$/.test(hora),
    data, hora, chave: d + '/' + m, horaCurta: h + 'h' + (min === '00' ? '' : min),
    dataExtenso: `${+d} de ${MESES_[+m - 1]} de ${a}`,
    escudoMandante: String(p.ESCUDO_MANDANTE || ''), escudoVisitante: String(p.ESCUDO_VISITANTE || ''),
    logo: String(p.LOGO_COMPETICAO || ''), competicaoId: String(p.COMPETICAO_ID || '')
  };
}

function partidas_() { return linhas_(ABA_PARTIDAS).map(partida_); }

// Jogo pertence ao texto da inscrição se bate data E confronto (duas partidas podem cair no mesmo dia).
function pertence_(texto, p) {
  const t = String(texto || '').toLowerCase();
  return t.includes(p.chave) && t.includes(`${p.mandante} x ${p.visitante}`.toLowerCase());
}

// Nome do evento usado nos e-mails (coluna COMPETICAO da aba Partidas).
function evento_() {
  const p = partidas_().find(x => x.aberta) || partidas_()[0];
  return p ? p.competicao : 'Federação Alagoana de Futebol';
}

// Substitui a versão com datas fixas: monta os cartões a partir da aba Partidas, no mesmo visual.
function obterDetalhesJogo_(jogo, linhaOrigem) {
  const texto = String(jogo || '').toLowerCase();
  const todas = partidas_();
  // o fluxo original só passa o texto; acha a linha com esse texto para usar os códigos dos jogos
  // Usa a linha efetivamente aprovada, nunca a primeira inscrição com texto semelhante.
  const achadas = linhaOrigem ? jogosDaLinha_(linhaOrigem, todas) : todas.filter(p => pertence_(texto, p));
  if (!achadas.length) return {
    assunto: evento_(), texto: jogo,
    html: `<div style="border:1px solid #e5e5e5;border-radius:9px;padding:18px 20px;margin-bottom:12px;">${escaparHtml_(jogo)}</div>`
  };
  return {
    assunto: achadas.map(p => `${p.mandante} x ${p.visitante}`).join(' e '),
    texto: achadas.map(p => `${p.mandante} x ${p.visitante}\n${p.dataExtenso} • ${p.horaCurta}\n${p.local}`).join('\n\n'),
    html: achadas.map(p => `
<div style="border:1px solid #e5e5e5;border-radius:9px;padding:18px 20px;margin-bottom:12px;">
<div style="font-size:18px;font-weight:bold;margin-bottom:9px;">${escaparHtml_(p.mandante)} x ${escaparHtml_(p.visitante)}</div>
<div style="font-size:14px;color:#555555;line-height:1.6;">${escaparHtml_(p.dataExtenso)} &nbsp;•&nbsp; ${escaparHtml_(p.horaCurta)}<br>${escaparHtml_(p.local)}</div>
</div>`).join('')
  };
}

// Contador que nunca volta atrás: apagar linha da planilha não faz protocolo repetir. Chamar dentro do lock.
function proximoProtocolo_(existentes) {
 const props=PropertiesService.getScriptProperties();
 const maior=existentes.reduce((m,r)=>Math.max(m,+(String(r[12]).match(/^FAF-(\d+)$/)||[,0])[1]),0);
 const salvo=Number(props.getProperty('ULTIMO_PROTOCOLO')||0);
 if(!Number.isSafeInteger(salvo)||salvo<0) throw new Error('Contador de protocolos inválido.');
 const n=Math.max(salvo,maior)+1;
 if(!Number.isSafeInteger(n)) throw new Error('Contador excedido.');
 props.setProperty('ULTIMO_PROTOCOLO',String(n)); return 'FAF-'+String(n).padStart(4,'0');
}


function jogosDaLinha_(r, partidas) {
  const ids = String(r[COL_IDS - 1] || '').split(',').map(x=>x.trim()).filter(Boolean);
  return ids.length ? partidas.filter(p => ids.includes(p.id)) : partidas.filter(p => pertence_(r[8], p));
}

function cpfValido_(cpf) {
  if (!/^\d{11}$/.test(cpf) || /^(\d)\1{10}$/.test(cpf)) return false;
  const dv = n => { let s = 0; for (let k = 0; k < n; k++) s += +cpf[k] * (n + 1 - k); const r = (s * 10) % 11; return r === 10 ? 0 : r; };
  return dv(9) === +cpf[9] && dv(10) === +cpf[10];
}

// ---------- Formulário ----------
// Todas as partidas: as encerradas aparecem como "Encerrado" na grade.
function partidasParaFormulario() {
  const campos = {};
  return partidas_().map(({ id, competicao, jogo, mandante, visitante, data, hora, local, aberta, escudoMandante, escudoVisitante, logo, competicaoId }) =>
    ({ id, competicao, jogo, mandante, visitante, data, hora, local, aberta, escudoMandante, escudoVisitante, logo, competicaoId,
       campos: campos[competicaoId] = campos[competicaoId] || camposDaCompeticao_(competicaoId) }));
}

// Grava no mesmo formato de linha do Google Forms antigo, já como PENDENTE / NÃO.
// Os campos exigidos dependem da competição (aba Competicoes: MODELO + AJUSTES).
function registrarInscricao(d) {
  if (!d || typeof d !== 'object' || !Array.isArray(d.partidas) || d.partidas.length > 10 || d.partidas.some(x => typeof x !== 'string' || x.length > 250))
    return { ok: false, erro: 'Solicitação inválida.' };
  if (d.ciencia !== undefined && d.ciencia !== true) return { ok: false, erro: 'Confirme a ciência das regras e do aviso de privacidade.' };
  const t = s => String(s || '').trim();
  const dados = {
    nome: t(d.nome), email: t(d.email).toLowerCase(), cpf: t(d.cpf).replace(/\D/g, ''), telefone: t(d.telefone),
    funcao: t(d.funcao), veiculo: t(d.veiculo), tipo: t(d.tipo), associacao: t(d.associacao).slice(0, 60),
    partidas: [...new Set(d.partidas || [])]
  };
  for (const c of ['nome','email','cpf','telefone','funcao','veiculo','tipo','associacao']) {
    if (d[c] != null && typeof d[c] !== 'string') return { ok: false, erro: 'Campo inválido.' };
    if (String(d[c] || '').length > (c === 'email' ? 254 : 250)) return { ok: false, erro: 'Campo excede o limite.' };
  }
  if (!dados.partidas.length) return { ok: false, erro: 'Escolha pelo menos uma partida.' };
  const abertas = partidas_().filter(p => p.aberta);
  const escolhidas = dados.partidas.map(id => abertas.find(p => p.id === id));
  if (escolhidas.some(p => !p)) return { ok: false, erro: 'Uma das partidas escolhidas não está com inscrições abertas.' };
  if (new Set(escolhidas.map(p => p.competicaoId)).size > 1) return { ok: false, erro: 'Escolha partidas de uma competição por vez.' };
  const campos = camposDaCompeticao_(escolhidas[0].competicaoId);
  const usa = c => campos.includes(c);
  const jpeg = jpegValido_;

  if (!dados.nome) return { ok: false, erro: 'Preencha o nome completo.' };
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(dados.email)) return { ok: false, erro: 'E-mail inválido.' };
  if (usa('cpf') && !cpfValido_(dados.cpf)) return { ok: false, erro: 'CPF inválido.' };
  if (usa('telefone') && dados.telefone.replace(/\D/g, '').length < 10) return { ok: false, erro: 'Informe o celular com DDD.' };
  for (const [c, n] of [['funcao', 'a função'], ['veiculo', 'o veículo'], ['tipo', 'o tipo de veículo'], ['associacao', 'a credencial de associação']])
    if (usa(c) && !dados[c]) return { ok: false, erro: `Informe ${n}.` };
  if (usa('fotoCredencial') && !jpeg(d.fotoCredencial)) return { ok: false, erro: 'Envie a foto da credencial de associação.' };
  if (usa('selfie') && !jpeg(d.foto)) return { ok: false, erro: 'Envie uma foto do rosto.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // evita protocolo repetido com envios simultâneos
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(CONFIG.ABA);
    const existentes = aba.getLastRow() < 2 ? [] : aba.getRange(2, 1, aba.getLastRow() - 1, COL_IDS).getValues();
    const todas = partidas_();
    const camposAgora = camposDaCompeticao_(escolhidas[0].competicaoId);
    if (JSON.stringify(camposAgora)!==JSON.stringify(campos)) return {ok:false,erro:'As regras foram atualizadas. Recarregue a página.'};
    if (escolhidas.some(p=>!todas.some(x=>x.id===p.id && x.competicaoId===p.competicaoId))) return {ok:false,erro:'Partida atualizada. Recarregue a página.'};
    // Revalida a abertura após adquirir o lock para impedir envio no fechamento.
    if (escolhidas.some(p => !todas.some(atual => atual.id === p.id && atual.aberta)))
      return { ok: false, erro: 'As inscrições para uma das partidas foram encerradas. Atualize a página.' };
    // sem CPF na competição, a trava de duplicidade usa o e-mail
    const mesmaPessoa = r => usa('cpf') ? String(r[3]).replace(/\D/g, '') === dados.cpf : String(r[1]).trim().toLowerCase() === dados.email;
    const dup = escolhidas.find(p => existentes.some(r => mesmaPessoa(r) && jogosDaLinha_(r, todas).some(x => x.id === p.id)));
    if (dup) return { ok: false, erro: `Já existe inscrição sua para ${dup.jogo}: ${dup.mandante} x ${dup.visitante}.` };

    const protocolo = proximoProtocolo_(existentes);
    const arquivos=[];
    try {
    const selfieId = usa('selfie') ? salvarFoto_(d.foto, protocolo) : '';
    if(selfieId) arquivos.push(selfieId);
    const credId = usa('fotoCredencial') ? salvarFoto_(d.fotoCredencial, protocolo + '-credencial') : '';
    if(credId) arquivos.push(credId);
    const jogo = escolhidas.map(p => `${p.jogo}: ${p.mandante} x ${p.visitante} (${p.chave} - ${p.horaCurta})`).join(', ');
    aba.appendRow([new Date(), textoCelula_(dados.email), textoCelula_(dados.nome), usa('cpf') ? "'" + dados.cpf : '', "'" + dados.telefone, textoCelula_(dados.funcao), textoCelula_(dados.veiculo), textoCelula_(dados.tipo),
      jogo, 'PENDENTE', 'NÃO', '', protocolo, selfieId, usa('associacao') ? textoCelula_(dados.associacao) : '', credId, escolhidas.map(p => p.id).join(','), (d.ciencia === true ? 'CIENTE v1 ' : 'CIÊNCIA NÃO INFORMADA — CLIENTE LEGADO ')+new Date().toISOString()]);
    SpreadsheetApp.flush();
    return { ok: true, protocolo, partidas: escolhidas.map(p => `${p.jogo}: ${p.mandante} x ${p.visitante}`) };
    } catch(err) {
      // Append may have succeeded before an exception. Never remove a referenced photo.
      const gravadas=aba.getLastRow()<2?[]:aba.getRange(2,COL_PROTOCOLO,aba.getLastRow()-1,1).getValues();
      if(!gravadas.some(r=>r[0]===protocolo)) arquivos.forEach(id=>{try{DriveApp.getFileById(id).setTrashed(true);}catch(e){console.error('Arquivo pendente de limpeza: '+id);}});
      throw err;
    }
  } finally {
    lock.releaseLock();
  }
}

// ---------- App web ----------
function doGet(e) {
  const pagina = { checkin: 'Checkin', admin: 'Admin' }[e && e.parameter && e.parameter.p] || 'Formulario';
  return HtmlService.createHtmlOutputFromFile(pagina).setTitle('Credenciamento FAF')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------- Check-in (PIN nas Propriedades do script: PIN_CHECKIN) ----------
function pinOk_(pin) {
  const certo = PropertiesService.getScriptProperties().getProperty('PIN_CHECKIN');
  if (typeof pin !== 'string' || pin.length>128 || !certo || pin !== certo) throw new Error('PIN inválido.');
}

function partidasCheckin(pin) {
  pinOk_(pin);
  return partidas_().map(({ id, jogo, mandante, visitante, data, hora, local }) => ({ id, jogo, mandante, visitante, data, hora, local }));
}

function checkinsFeitos_() {
  return Object.fromEntries(linhas_(ABA_CHECKINS).map(c => [c.ID, fmt_(c.DATA_HORA, 'HH:mm')]));
}

// Aprovados de uma partida, sem CPF, e-mail ou telefone.
function aprovados_(idPartida) {
  const p = partidas_().find(x => x.id === idPartida);
  if (!p) return [];
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ABA);
  if (aba.getLastRow() < 2) return [];
  const todas = partidas_();
  return aba.getRange(2, 1, aba.getLastRow() - 1, COL_IDS).getValues()
    .filter(r => String(r[9]).trim().toUpperCase() === 'APROVADO' && jogosDaLinha_(r, todas).some(x => x.id === p.id))
    .map(r => ({ id: (r[12] || 'L' + r[0]) + '|' + p.id, nome: String(r[2]).trim(), funcao: r[5], veiculo: r[6], partida: p.id }));
}

function listarAprovados(pin, idPartida) {
  pinOk_(pin);
  const feitos = checkinsFeitos_();
  return aprovados_(idPartida).map(a => ({ ...a, entrada: feitos[a.id] || '' }))
    .sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
}

function registrarCheckin(pin, id, operador) {
  pinOk_(pin);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // dois celulares no mesmo portão não registram em dobro
  try {
    const chave = String(id);
    const separador = chave.indexOf('|');
    if (separador < 1) return { ok: false, erro: 'Credencial inválida.' };
    const a = aprovados_(chave.slice(separador + 1)).find(x => x.id === chave);
    if (!a) return { ok: false, erro: 'Credencial não aprovada para esta partida.' };
    const ja = checkinsFeitos_()[id];
    if (ja) return { ok: false, erro: `Entrada já registrada às ${ja}.` };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let aba = ss.getSheetByName(ABA_CHECKINS);
    if (!aba) { aba = ss.insertSheet(ABA_CHECKINS); aba.appendRow(['ID','ID_PARTIDA','NOME','OPERADOR','DATA_HORA']); }
    const agora = new Date();
    aba.appendRow([textoCelula_(id), textoCelula_(a.partida), textoCelula_(a.nome), textoCelula_(String(operador || '').trim().slice(0,120) || 'não informado'), agora]);
    return { ok: true, hora: fmt_(agora, 'HH:mm') };
  } finally {
    lock.releaseLock();
  }
}

// ---------- Fotos (LGPD: pasta privada, só para conferência no acesso, apagadas após o prazo) ----------
function pastaFotos_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOTOS_PASTA_ID');
  if (id) { const pasta=DriveApp.getFolderById(id); if(pasta.getSharingAccess()!==DriveApp.Access.PRIVATE) throw new Error('Pasta de fotos deve ser privada.'); return pasta; }
  const pasta = DriveApp.createFolder('Credenciamento FAF - Fotos (RESTRITO)');
  props.setProperty('FOTOS_PASTA_ID', pasta.getId());
  return pasta;
}

function salvarFoto_(dataUrl, protocolo) {
  const blob = Utilities.newBlob(Utilities.base64Decode(dataUrl.split(',')[1]), 'image/jpeg', protocolo + '.jpg');
  return pastaFotos_().createFile(blob).getId();
}

// Foto em base64 para check-in/painel (sempre atrás de senha).
function fotoBase64_(fotoId) {
  if (!fotoId || String(fotoId).indexOf('EXCL') === 0) return '';
  try { return 'data:image/jpeg;base64,' + Utilities.base64Encode(DriveApp.getFileById(fotoId).getBlob().getBytes()); } catch (e) { return ''; }
}

// Diária: manda para a lixeira as fotos de quem teve o último jogo há mais de DIAS_RETENCAO_FOTO dias.
function rotinaLGPD_() {
 try {sincronizarUrano_();}catch(e){console.error('Sincronização falhou: '+e.message);}
 const lock=LockService.getScriptLock(); lock.waitLock(30000);
 try{return reterFotos_();}finally{SpreadsheetApp.flush();lock.releaseLock();}
}
function reterFotos_() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ABA);
  if (!aba || aba.getLastRow() < 2) return 0;
  const datas = partidas_().map(p => ({ p, quando: new Date(p.data.split('/').reverse().join('-') + 'T23:59:59-03:00') }));
  const limite = new Date(Date.now() - DIAS_RETENCAO_FOTO * 864e5);
  const v = aba.getRange(2, 1, aba.getLastRow() - 1, COL_IDS).getValues();
  let apagadas = 0;
  v.forEach((r, k) => {
    const meus = jogosDaLinha_(r, datas.map(x => x.p)).map(p => p.id);
    const jogos = datas.filter(x => meus.includes(x.p.id));
    if (!jogos.length || jogos.some(p => !dataValida_(p.p.data) || !Number.isFinite(p.quando.getTime()) || p.quando > limite)) return;
    [COL_FOTO, COL_FOTO_CREDENCIAL].forEach(col => {
      const fotoId = String(r[col - 1] || '');
      if (!fotoId || fotoId.indexOf('EXCL') === 0) return;
      try { DriveApp.getFileById(fotoId).setTrashed(true); } catch (e) { console.error('Falha ao remover foto na linha ' + (k + 2)); return; }
      aba.getRange(k + 2, col).setValue('EXCLUÍDA ' + fmt_(new Date(), 'dd/MM/yyyy'));
      apagadas++;
    });
  });
  return apagadas;
}

// ---------- Competições e campos por competição ----------
const ABA_COMPETICOES = 'Competicoes';
const MODELOS_ = {
  PROFISSIONAL: ['cpf', 'telefone', 'associacao', 'fotoCredencial', 'funcao', 'veiculo', 'tipo'],
  BASE: ['cpf', 'telefone', 'funcao', 'veiculo', 'tipo']
};
const CAMPOS_VALIDOS_ = ['cpf', 'telefone', 'associacao', 'fotoCredencial', 'selfie', 'funcao', 'veiculo', 'tipo'];
// Partidas cadastradas à mão antes da aba Competicoes seguem o formulário que já estava no ar.
const CAMPOS_LEGADO_ = ['cpf', 'telefone', 'funcao', 'veiculo', 'tipo', 'selfie'];

// MODELO (PROFISSIONAL ou BASE) + AJUSTES como "+selfie, -cpf".
function camposDaCompeticao_(id) {
  const c = id && linhas_(ABA_COMPETICOES).find(x => String(x.ID_URANO).trim() === id);
  if (!c) return CAMPOS_LEGADO_.slice();
  const campos = new Set(MODELOS_[String(c.MODELO).trim().toUpperCase()] || MODELOS_.BASE);
  String(c.AJUSTES || '').split(',').map(x => x.trim()).filter(Boolean).forEach(a => {
    const nome = a.slice(1);
    if (!CAMPOS_VALIDOS_.includes(nome)) return;
    if (a[0] === '+') campos.add(nome); else if (a[0] === '-') campos.delete(nome);
  });
  return CAMPOS_VALIDOS_.filter(x => campos.has(x));
}

// ---------- Sincronização com o Urano (somente leitura) ----------
const URANO_API = 'https://velguljjckaolxquzidr.supabase.co/rest/v1/rpc/';
const URANO_CHAVE = 'sb_publishable_cAeiVdDPrSCzeLiC3GhErQ_IE3OP6CZ'; // chave pública: só chama as funções de leitura
const URANO_ASSETS = 'https://gerador-artes.vercel.app/assets/';

function urano_(funcao, corpo) {
  const r = UrlFetchApp.fetch(URANO_API + funcao, { method: 'post', contentType: 'application/json',
    headers: { apikey: URANO_CHAVE }, payload: JSON.stringify(corpo || {}), muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error(`Urano respondeu ${r.getResponseCode()}`);
  return JSON.parse(r.getContentText());
}

// Escudo: link do Urano se houver; senão o arquivo assets/escudos/<id>.png que existir (resultado guardado).
function escudo_(clubeId, link) {
  if (/^https:\/\//.test(String(link || ''))) return link;
  if (!clubeId) return '';
  const props = PropertiesService.getScriptProperties(), chave = 'ESCUDO_' + clubeId;
  const salvo = props.getProperty(chave);
  if (salvo) return salvo;
  const nomes = [...new Set([clubeId, clubeId.replace(/-/g, '_'), clubeId.split('-')[0]])];
  let achado = '';
  for (const n of nomes) {
    const u = URANO_ASSETS + 'escudos/' + n + '.png';
    if (UrlFetchApp.fetch(u, { muteHttpExceptions: true }).getResponseCode() === 200) { achado = u; break; }
  }
  props.setProperty(chave, achado);
  return achado;
}

function logo_(logo) {
  const l = String(logo || '');
  if (/^https:\/\//.test(l)) return l;
  return /^[\w.-]+\.(png|svg|jpg)$/i.test(l) ? URANO_ASSETS + 'logos/' + l : ''; // logos embutidas no banco não cabem na planilha
}

// Traz para a aba Partidas os jogos (de ontem em diante) das competições ativas na aba Competicoes.
// Atualiza confronto, data, hora e estádio; nunca mexe em INSCRICOES_ABERTAS de jogo já existente.
function sincronizarUrano_() {
 const lock=LockService.getScriptLock(); lock.waitLock(30000);
 try{return sincronizarUranoComLock_();}finally{SpreadsheetApp.flush();lock.releaseLock();}
}
function sincronizarUranoComLock_() {
 const ss=SpreadsheetApp.getActiveSpreadsheet();
 const comps=linhas_(ABA_COMPETICOES).map((c,k)=>({...c,linha:k+2})).filter(c=>String(c.ATIVA).trim().toUpperCase()==='SIM' && String(c.ID_URANO).trim());
 if(!comps.length)return {novos:0,atualizados:0};
 const lista=urano_('credenciamento_competicoes');
 if(!Array.isArray(lista))throw new Error('Resposta de competições inválida.');
 const infos=Object.fromEntries(lista.map(c=>[c.id,c]));
 const aba=ss.getSheetByName(ABA_PARTIDAS); if(!aba)throw new Error('Aba Partidas ausente.');
 const tabela=aba.getDataRange().getValues(); const cab=tabela[0].map(String);
 const obrigatorias=['ID_PARTIDA','COMPETICAO','JOGO','MANDANTE','VISITANTE','DATA','HORA','LOCAL','INSCRICOES_ABERTAS','ESCUDO_MANDANTE','ESCUDO_VISITANTE','LOGO_COMPETICAO','COMPETICAO_ID'];
 if(obrigatorias.some(n=>!cab.includes(n)))throw new Error('Cabeçalhos de Partidas incompatíveis.');
 const col=n=>cab.indexOf(n), linhas=tabela.slice(1).filter(r=>String(r[col('ID_PARTIDA')]).trim());
 const indice=new Map(linhas.map((r,k)=>[String(r[col('ID_PARTIDA')]),k]));
 if(indice.size!==linhas.length)throw new Error('IDs de partida duplicados. Corrija antes de sincronizar.');
 const ontem=new Date(Date.now()-864e5); ontem.setHours(0,0,0,0);
 const res={novos:0,atualizados:0}; const planos=[];
 // Validate every response before writing any cells.
 comps.forEach(c=>{
  const id=String(c.ID_URANO).trim(),info=infos[id]; if(!info)throw new Error('Competição não encontrada: '+id);
  const jogos=urano_('credenciamento_jogos',{p_competicao:id});
  if(!Array.isArray(jogos))throw new Error('Resposta de jogos inválida.');
  const ids=new Set();
  jogos.forEach(j=>{
   if(!j || typeof j.id!=='string' || !j.id.trim() || j.id.length>250 || ids.has(j.id) || !j.mandante || !j.visitante || !dataValida_(j.data) || !/^([01]\d|2[0-3]):[0-5]\d$/.test(String(j.hora||'')))throw new Error('Jogo Urano inválido ou duplicado.');
   ids.add(j.id);
  });
  jogos.forEach(j=>{
   const [d,m,a]=j.data.split('/').map(Number); if(new Date(a,m-1,d)<ontem)return;
   // IDs composed from dates are unstable. Do not silently create a new credential target on rescheduling.
   if(!indice.has(j.id) && linhas.some(r=>String(r[col('COMPETICAO_ID')])===id && r[col('MANDANTE')]===j.mandante && r[col('VISITANTE')]===j.visitante && !ids.has(String(r[col('ID_PARTIDA')]))))throw new Error('Possível reagendamento com novo ID. Reconciliar manualmente antes de sincronizar.');
   planos.push({c,j,info});
  });
 });
 planos.forEach(({c,j,info})=>{
  const id=String(c.ID_URANO).trim();
  const valores={ID_PARTIDA:j.id,COMPETICAO:String(c.NOME_EXIBICAO||'').trim()||info.nome,JOGO:j.fase||(j.rodada?'Rodada '+j.rodada:'Jogo'),MANDANTE:j.mandante,VISITANTE:j.visitante,DATA:j.data,HORA:j.hora,LOCAL:j.estadio||'',ESCUDO_MANDANTE:escudo_(j.mandante_id,j.escudo_mandante),ESCUDO_VISITANTE:escudo_(j.visitante_id,j.escudo_visitante),LOGO_COMPETICAO:logo_(info.logo),COMPETICAO_ID:id};
  const existe=indice.has(j.id),k=existe?indice.get(j.id):linhas.length;
  const r=existe?linhas[k].slice():cab.map(()=> '');
  Object.entries(valores).forEach(([n,v])=>r[col(n)]=typeof v==='string'?textoCelula_(v):v);
  if(!existe){r[col('INSCRICOES_ABERTAS')]='NÃO';indice.set(j.id,k);res.novos++;}else res.atualizados++;
  linhas[k]=r;
 });
 if(linhas.length)aba.getRange(2,1,linhas.length,cab.length).setValues(linhas);
 comps.forEach(c=>ss.getSheetByName(ABA_COMPETICOES).getRange(c.linha,6).setValue(new Date()));
 return res;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Credenciamento')
    .addItem('Sincronizar jogos do Urano', 'sincronizarMenu')
    .addToUi();
}

function sincronizarMenu() {
  const ui = SpreadsheetApp.getUi(); // só existe dentro da planilha: bloqueia chamada pelo app web
  const r = sincronizarUrano_();
  ui.alert(`Jogos novos: ${r.novos}\nJogos atualizados: ${r.atualizados}\n\nJogos novos entram com INSCRICOES_ABERTAS = NÃO. Mude para SIM para abrir o credenciamento.`);
}

// ---------- API para o site público (GitHub Pages) ----------
// Só estas ações ficam abertas para fora; painel, check-in e e-mails continuam só no Google.
const ACOES_PUBLICAS_ = {
  partidasParaFormulario: () => partidasParaFormulario(),
  registrarInscricao: args => registrarInscricao(args[0] || {})
};

function doPost(e) {
  let saida;
  try {
    if(!e || !e.postData || typeof e.postData.contents!=='string' || e.postData.contents.length>850000) throw new Error('Solicitação inválida.');
    const req = JSON.parse(e.postData.contents);
    const acao = req && Object.prototype.hasOwnProperty.call(ACOES_PUBLICAS_, req.acao) ? ACOES_PUBLICAS_[req.acao] : null;
    if (!acao) throw new Error('Ação inválida.');
    saida = { ok: true, dados: acao(Array.isArray(req.args) ? req.args : []) };
  } catch (err) {
    saida = { ok: false, erro: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(saida)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- Histórico de alterações (auditoria) ----------
// Gatilho instalável "ao editar": registra quem mudou o quê. Não grava o conteúdo da coluna de CPF.
function registrarHistorico(e) {
  if (!e || !e.range) return;
  const aba = e.range.getSheet();
  if (aba.getName() === 'Historico') return;
  const ss = aba.getParent();
  const h = ss.getSheetByName('Historico') || (() => { const n = ss.insertSheet('Historico'); n.appendRow(['DATA_HORA', 'USUARIO', 'ABA', 'CELULA', 'ANTES', 'DEPOIS']); return n; })();
  const cpf = aba.getName() === CONFIG.ABA && e.range.getColumn() <= 4 && e.range.getLastColumn() >= 4;
  const valor = v => cpf ? '(dado sensível omitido)' : (v === undefined ? (e.range.getNumRows() * e.range.getNumColumns() > 1 ? '(várias células)' : '') : String(v));
  const usuario = (e.user && e.user.getEmail && e.user.getEmail()) || 'não identificado';
  h.appendRow([new Date(), textoCelula_(usuario), aba.getName(), e.range.getA1Notation(), textoCelula_(valor(e.oldValue)), textoCelula_(valor(e.value))]);
}

// ---------- Admin (PIN nas Propriedades do script: PIN_ADMIN) ----------
function pinAdmin_(pin) {
  const certo = PropertiesService.getScriptProperties().getProperty('PIN_ADMIN');
  if (!certo || String(pin) !== certo) throw new Error('Senha inválida.');
}

function inscricoesAdmin_() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ABA);
  if (aba.getLastRow() < 2) return [];
  const todas=partidas_();
  return aba.getRange(2, 1, aba.getLastRow() - 1, COL_IDS).getValues().map((r, k) => ({
    idsPartidas: jogosDaLinha_(r,todas).map(p=>p.id), linha: k + 2, protocolo: String(r[12] || ''), data: fmt_(r[0], 'dd/MM/yyyy HH:mm'),
    email: String(r[1]), nome: String(r[2]).trim(), telefone: String(r[4]), funcao: String(r[5]), veiculo: String(r[6]),
    tipo: String(r[7]), jogo: String(r[8]), status: String(r[9]).trim().toUpperCase() || 'PENDENTE',
    emailEnviado: String(r[10]).trim().toUpperCase(), dataEnvio: r[11] ? fmt_(r[11], 'dd/MM HH:mm') : ''
  }));
}

// Lista sem CPF.
function adminListar(pin) {
  pinAdmin_(pin);
  return { partidas: partidas_().map(p => ({ id: p.id, chave: p.chave, rotulo: `${p.jogo}: ${p.mandante} x ${p.visitante}` })), inscricoes: inscricoesAdmin_() };
}

// Aprova ou recusa e envia o e-mail do fluxo original. Mudança feita por script não dispara o gatilho de edição.
function adminDecidir(pin, protocolo, status) {
  pinAdmin_(pin);
  if (!['APROVADO', 'NÃO APROVADO'].includes(status)) return { ok: false, erro: 'Decisão inválida.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // dois admins clicando juntos não enviam em dobro
  try {
    const i = inscricoesAdmin_().find(x => x.protocolo && x.protocolo === protocolo);
    if (!i) return { ok: false, erro: 'Inscrição não encontrada.' };
    if (!['', 'NÃO'].includes(i.emailEnviado)) return { ok: false, erro: `Esta inscrição já foi comunicada como ${i.status}. Para mudar, fale com o responsável pela planilha.` };
    const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ABA);
    aba.getRange(i.linha, CONFIG.COL_STATUS).setValue(status);
    SpreadsheetApp.flush();
    try {
      if (status === 'APROVADO') enviarAprovacaoDaLinha_(aba, i.linha); else enviarNaoAprovacaoDaLinha_(aba, i.linha);
    } catch (e) {
      return { ok: false, erro: 'Status salvo. Confira o registro de comunicação antes de tentar novamente: ' + e.message, inscricao: inscricoesAdmin_().find(x => x.protocolo === protocolo) };
    }
    return { ok: true, inscricao: inscricoesAdmin_().find(x => x.protocolo === protocolo) };
  } finally {
    lock.releaseLock();
  }
}

function textoCelula_(valor) {
 const s = String(valor == null ? '' : valor);
 return /^[=+@\-\t\r\n]/.test(s) ? "'" + s : s;
}

// Safe default: mail requires an explicit production switch set only at approved release.
function enviarEmailControlado_(sheet,row,email,assunto,texto,opcoes) {
 const props=PropertiesService.getScriptProperties();
 const modo=props.getProperty('MODO_EMAIL')||'BLOQUEADO';
 if(modo!=='PRODUCAO') throw new Error('Envio bloqueado neste ambiente.');
 const estado=String(sheet.getRange(row,CONFIG.COL_EMAIL_ENVIADO).getValue()).trim().toUpperCase();
 if(!['','NÃO'].includes(estado)) throw new Error('Comunicação já enviada ou pendente de conferência.');
 // Persist and flush BEFORE Gmail. Any later exception must not cause an automatic retry.
 sheet.getRange(row,CONFIG.COL_EMAIL_ENVIADO).setValue('EM PROCESSAMENTO');
 SpreadsheetApp.flush();
 try { GmailApp.sendEmail(email,assunto,texto,opcoes); }
 catch(err) {
  try {sheet.getRange(row,CONFIG.COL_EMAIL_ENVIADO).setValue('INCERTO — CONFERIR'); SpreadsheetApp.flush();} catch(e) {}
  throw new Error('Entrega incerta. Conferir mensagens enviadas antes de liberar nova tentativa.');
 }
}
function eventoDaLinha_(sheet,row) {
 const linha=sheet.getRange(row,1,1,COL_IDS).getValues()[0];
 const jogos=jogosDaLinha_(linha,partidas_());
 return [...new Set(jogos.map(p=>p.competicao))].join(' / ') || 'Federação Alagoana de Futebol';
}
function dataValida_(s) {
 const m=/^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(s)); if(!m)return false;
 const d=new Date(+m[3],+m[2]-1,+m[1]); return d.getFullYear()===+m[3] && d.getMonth()===+m[2]-1 && d.getDate()===+m[1];
}
function jpegValido_(v) {
 if(typeof v!=='string' || v.length>400000 || !/^data:image\/jpeg;base64,[A-Za-z0-9+/]+={0,2}$/.test(v))return false;
 try {const b=Utilities.base64Decode(v.split(',')[1]); return b.length>=4 && (b[0]&255)===255 && (b[1]&255)===216 && (b[b.length-2]&255)===255 && (b[b.length-1]&255)===217;} catch(e){return false;}
}

// ---------- Gatilho diário ----------
// Gatilhos não podem chamar funções privadas (terminadas em "_"); esta é a porta pública só para o gatilho.
// Recusa qualquer chamada que não venha de um gatilho deste projeto (ex.: pelo site).
function rotinaDiaria(e) {
  const uid = e && e.triggerUid;
  if (!uid || !ScriptApp.getProjectTriggers().some(t => t.getUniqueId() === String(uid))) throw new Error('Uso exclusivo do gatilho diário.');
  return rotinaLGPD_();
}

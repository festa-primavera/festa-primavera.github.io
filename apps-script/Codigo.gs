/**
 * Festa da Primavera 2026 — E.M. Dep. Federal Ulisses Guimarães
 * Backend da página de venda de mesas: grava e lê as vendas na planilha do Google Drive.
 *
 * Como publicar: Implantar → Nova implantação → Tipo "App da Web"
 *   Executar como: Eu  |  Quem pode acessar: Qualquer pessoa
 * Copie a URL que termina em /exec e cole em config.js (APPS_SCRIPT_URL).
 */

const PLANILHA_ID = '1JR07wt93lP6XxJo4z5V_T-UZM1Bh23vTunRRo8XBuxg';
const TOTAL_MESAS = 50;
const PAGAMENTOS = ['Dinheiro', 'PIX'];
const TURMAS = ['A', 'B', 'C'];

function aba_() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  // Datas e horários sempre no horário de Manaus
  if (ss.getSpreadsheetTimeZone() !== 'America/Manaus') ss.setSpreadsheetTimeZone('America/Manaus');
  const sh = ss.getSheets()[0];
  sh.getRange('I:K').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  return sh;
}

function lerVendas_() {
  const linhas = aba_().getDataRange().getValues();
  linhas.shift(); // cabeçalho
  return linhas
    .filter(function (r) { return r[0] !== '' && r[0] !== null; })
    .map(function (r) {
      return {
        mesa: Number(r[0]), comprador: String(r[1]), aluno: String(r[2]),
        serie: String(r[3]), turma: String(r[4]), pagamento: String(r[5]),
        valor: Number(r[6]) || 0, vendedor: String(r[7]),
        criadoEm: iso_(r[8]),
        chegada: iso_(r[9]),
        pixOk: iso_(r[10]),
        telefone: String(r[11] || ''),
      passe: String(r[12] || ''),
      };
    });
}

// Datas podem vir como Date ou como texto (planilha em formato texto)
function iso_(x) {
  if (x instanceof Date) return x.toISOString();
  if (!x) return '';
  const d = new Date(String(x).replace(/\s*\(.*\)\s*$/, ''));
  return isNaN(d.getTime()) ? String(x) : d.toISOString();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Evita que um texto digitado seja interpretado como fórmula na planilha
function texto_(s) {
  s = String(s || '').trim().slice(0, 100);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

// PIN de cada vendedor, guardado só como hash SHA-256 de "festa-primavera-2026:" + PIN.
// Os valores reais ficam APENAS no projeto do Apps Script (nunca no GitHub).
const PIN_HASH = { /* preenchido no Apps Script */ };

function hash_(txt) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, txt, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function autentica_(nome, pin) {
  if (!nome || !PIN_HASH[nome] || PIN_HASH[nome] !== hash_('festa-primavera-2026:' + String(pin || '').trim()))
    throw new Error('PIN incorreto para ' + (nome || 'o vendedor') + '.');
}

// ---- Prestação de contas: quem recebe o dinheiro dos outros vendedores
const TESOUREIRO = 'CARLOS ANDRE ALONSO DA SILVA';

function abaPrest_() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  let sh = ss.getSheetByName('Prestações');
  if (!sh) {
    sh = ss.insertSheet('Prestações', ss.getNumSheets()); // sempre no fim: a aba de vendas continua sendo a primeira
    sh.appendRow(['ID', 'Data/hora', 'Vendedor', 'Valor', 'Forma', 'Observação', 'Recebido por']);
  }
  sh.getRange('B:B').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  return sh;
}

function lerPrest_() {
  const v = abaPrest_().getDataRange().getValues();
  v.shift();
  return v.filter(function (r) { return r[0] !== ''; }).map(function (r) {
    return { id: String(r[0]), quando: iso_(r[1]), vendedor: String(r[2]), valor: Number(r[3]) || 0,
      forma: String(r[4]), obs: String(r[5]), recebidoPor: String(r[6]) };
  });
}

// ---- PIX: a chave fica APENAS no projeto do Apps Script (nunca no GitHub).
// O código PIX (copia e cola / QR Code) é montado aqui no servidor e só é entregue
// a um vendedor autenticado com o PIN, para uma venda PIX dele mesmo.
const PIX_CHAVE = '';
const PIX_NOME = 'ALESSON MIRANDA PEREIRA';
const PIX_CIDADE = 'MANAUS';
function emv_(id, v) { v = String(v); return id + ('0' + v.length).slice(-2) + v; }
function crc16_(s) {
  let c = 0xFFFF; const b = Utilities.newBlob(s).getBytes();
  for (let i = 0; i < b.length; i++) {
    c ^= (b[i] & 0xff) << 8;
    for (let j = 0; j < 8; j++) { c = (c & 0x8000) ? ((c << 1) ^ 0x1021) : (c << 1); c &= 0xFFFF; }
  }
  return ('000' + c.toString(16).toUpperCase()).slice(-4);
}
// Chave de quem recebe a prestação de contas (também só aqui no Apps Script)
const PIX_CHAVE_RECEBEDOR = '';
const PIX_NOME_RECEBEDOR = 'CARLOS ANDRE ALONSO DA SILVA';
// Nome no QR tem no máximo 25 letras: abrevia os nomes do meio
function nome25_(n) {
  if (n.length <= 25) return n;
  const p = n.split(' ');
  return [p[0]].concat(p.slice(1, -1).map(function (x) { return x.length <= 2 ? x : x[0]; }), [p[p.length - 1]]).join(' ').slice(0, 25);
}
function pixCom_(chave, nome, valor, txid) {
  if (!chave) throw new Error('Chave PIX não configurada.');
  txid = String(txid).replace(/[^A-Za-z0-9]/g, '').slice(0, 25);
  const s = emv_('00', '01') + emv_('01', '11') +
    emv_('26', emv_('00', 'br.gov.bcb.pix') + emv_('01', chave)) +
    emv_('52', '0000') + emv_('53', '986') + emv_('54', Number(valor).toFixed(2)) + emv_('58', 'BR') +
    emv_('59', nome25_(nome)) + emv_('60', PIX_CIDADE.slice(0, 15)) + emv_('62', emv_('05', txid)) + '6304';
  return { payload: s + crc16_(s), valor: Number(valor), recebedor: nome };
}
// Vendedores com chave PIX própria: o QR das vendas deles cai na conta deles
// (as chaves ficam só aqui no Apps Script, nunca no GitHub)
const PIX_POR_VENDEDOR = { /* preenchido no Apps Script */ };
function pix_(valor, mesas, vendedor) {
  const k = PIX_POR_VENDEDOR[vendedor];
  return k ? pixCom_(k, vendedor, valor, 'FP26M' + mesas.join('M'))
           : pixCom_(PIX_CHAVE, PIX_NOME, valor, 'FP26M' + mesas.join('M'));
}
// Quanto um vendedor ainda tem para prestar de contas
function pendente_(vendas, nome) {
  const vendido = vendas.filter(function (x) { return x.vendedor === nome; }).reduce(function (s, x) { return s + x.valor; }, 0);
  const recebido = lerPrest_().filter(function (p) { return p.vendedor === nome; }).reduce(function (s, p) { return s + p.valor; }, 0);
  return Math.round((vendido - recebido) * 100) / 100;
}
// Confere se as mesas são uma venda PIX do próprio vendedor
function vendaPix_(vendas, nome, mesas) {
  const g = vendas.filter(function (x) { return mesas.indexOf(x.mesa) >= 0; });
  if (!mesas.length || g.length !== mesas.length) throw new Error('Mesa sem venda registrada.');
  if (g.some(function (x) { return x.vendedor !== nome; })) throw new Error('Essa venda é de outro vendedor.');
  if (g.some(function (x) { return x.pagamento !== 'PIX'; })) throw new Error('Essa venda não foi paga por PIX.');
  return g;
}

// Telefones dos compradores nunca saem na lista pública
function semTel_(l) { return l.map(function (x) { const y = Object.assign({}, x); delete y.telefone; return y; }); }
function telefone_(t) {
  const d = String(t || '').replace(/\D/g, '');
  if (d.length !== 10 && d.length !== 11) throw new Error('Telefone inválido: informe DDD + número.');
  return '(' + d.slice(0, 2) + ') ' + (d.length === 11 ? d.slice(2, 7) + '-' + d.slice(7) : d.slice(2, 6) + '-' + d.slice(6));
}

// ---- Ano da festa e reinício anual (só o administrador)
const ADMIN = 'ALESSON MIRANDA PEREIRA';
function config_() {
  const p = PropertiesService.getScriptProperties();
  return { ano: Number(p.getProperty('ANO')) || 2026, data: p.getProperty('DATA_FESTA') || '2026-11-12' };
}
// Guarda uma cópia da aba (com todos os dados) antes de zerar
function arquivar_(ss, sh, nome) {
  const copia = sh.copyTo(ss);
  copia.setName(nome);
  ss.setActiveSheet(copia); ss.moveActiveSheet(ss.getNumSheets());
  const n = sh.getLastRow();
  if (n > 1) sh.deleteRows(2, n - 1);
}

// ---- Passe livre: cada compra tem um código secreto (coluna M) que vai no QR Code do comprovante
const INGRESSOS_POR_MESA = 4;
function chavePasse_(r) { return String(r[1]).trim().toUpperCase() + '|' + String(r[2]).trim().toUpperCase(); }
// Dá um código às vendas que ainda não têm (as mesas do mesmo comprador/aluno usam o mesmo código)
function garantirPasses_(sh) {
  const n = sh.getLastRow(); if (n < 2) return;
  const r = sh.getRange(2, 1, n - 1, 13).getValues(), cod = {};
  r.forEach(function (x) { if (x[0] !== '' && x[12]) cod[chavePasse_(x)] = String(x[12]); });
  let mudou = false;
  const col = r.map(function (x) {
    if (x[0] === '' || x[12]) return [x[12]];
    const k = chavePasse_(x);
    if (!cod[k]) cod[k] = Utilities.getUuid().replace(/-/g, '');
    mudou = true; return [cod[k]];
  });
  if (!mudou) return;
  if (sh.getRange(1, 13).getValue() === '') sh.getRange(1, 13).setValue('Passe');
  sh.getRange(2, 13, n - 1, 1).setNumberFormat('@').setValues(col);
}
// Comprovante online: só os dados daquela compra, e só para quem tem o código do QR
function passe_(cod) {
  if (!/^[a-f0-9]{32}$/.test(cod)) return { ok: false, erro: 'Comprovante inválido.' };
  const g = lerVendas_().filter(function (x) { return x.passe === cod; });
  if (!g.length) return { ok: false, erro: 'Comprovante não encontrado.', config: config_() };
  const ms = g.map(function (x) { return x.mesa; }).sort(function (a, b) { return a - b; });
  return { ok: true, config: config_(), passe: {
    comprador: g[0].comprador, mesas: ms, ingressos: INGRESSOS_POR_MESA * ms.length,
    criadoEm: g[0].criadoEm, chegada: g.some(function (x) { return x.chegada; }) } };
}

function doGet(e) {
  if (e && e.parameter && e.parameter.passe) return json_(passe_(String(e.parameter.passe)));
  // Sem PIN: só quais mesas estão vendidas (nenhum dado de comprador ou de venda)
  return json_({ ok: true, mesas: lerVendas_().map(function (x) { return x.mesa; }), config: config_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // uma venda por vez: impede vender a mesma mesa duas vezes
  try {
    const d = JSON.parse(e.postData.contents);
    const sh = aba_();
    const vendas = lerVendas_();
    let pixResp = null, telResp = null;

    if (d.acao === 'entrar') {
      autentica_(d.vendedor, d.pin);
    } else if (d.acao === 'vender') {
      const v = d.venda || {};
      autentica_(v.vendedor, d.pin); // a venda só sai no nome de quem tem o PIN
      // Uma venda pode ter várias mesas: grava todas ou nenhuma
      const mesas = (Array.isArray(v.mesas) ? v.mesas : [v.mesa]).map(Number)
        .filter(function (m, i, a) { return a.indexOf(m) === i; });
      if (!mesas.length || mesas.some(function (m) { return !(m >= 1 && m <= TOTAL_MESAS); }))
        throw new Error('Mesa inválida.');
      const ja = mesas.filter(function (m) { return vendas.some(function (x) { return x.mesa === m; }); });
      if (ja.length) throw new Error((ja.length > 1 ? 'As mesas ' + ja.join(', ') + ' acabaram de ser vendidas'
        : 'A mesa ' + ja[0] + ' acabou de ser vendida') + ' por outra pessoa.');
      if (!(Number(v.valor) > 0)) throw new Error('Informe o valor.');
      if (PAGAMENTOS.indexOf(v.pagamento) < 0) throw new Error('Forma de pagamento inválida.');
      // Professor(a) ou funcionário(a) compra sem aluno vinculado
      const prof = v.tipo === 'professor';
      if (prof) { v.aluno = ''; v.serie = 'Professor(a)'; v.turma = ''; }
      else if (v.tipo === 'eja') { v.aluno = v.comprador; v.serie = 'EJA'; v.turma = ''; } // aluno(a) da EJA compra para si
      if (!prof && v.tipo !== 'eja' && TURMAS.indexOf(v.turma) < 0 && !(v.turma === 'D' && v.serie === '7º Ano')) throw new Error('Turma inválida.'); // 7º Ano também tem turma D
      if (!String(v.comprador || '').trim() || (!prof && !String(v.aluno || '').trim()))
        throw new Error(prof ? 'Preencha o nome do comprador.' : 'Preencha comprador e aluno.');
      const tel = v.telefone ? telefone_(v.telefone) : ''; // telefone é opcional
      const agora = new Date();
      const linhas = mesas.map(function (m) {
        return [m, texto_(v.comprador), texto_(v.aluno), texto_(v.serie), v.turma,
          v.pagamento, Number(v.valor), texto_(v.vendedor), agora];
      });
      const ini = sh.getLastRow() + 1;
      sh.getRange(ini, 1, linhas.length, 9).setValues(linhas);
      if (sh.getRange(1, 12).getValue() === '') sh.getRange(1, 12).setValue('Telefone');
      sh.getRange(ini, 12, linhas.length, 1).setNumberFormat('@').setValues(linhas.map(function () { return [tel]; }));
      if (v.pagamento === 'PIX' && PIX_CHAVE) pixResp = pix_(Number(v.valor) * mesas.length, mesas, v.vendedor);
    } else if (d.acao === 'chegada') {
      autentica_(d.vendedor, d.pin); // só vendedores confirmam a chegada
      // Recepção no dia da festa: marca (ou desmarca) a chegada do comprador
      const alvo = (Array.isArray(d.mesas) ? d.mesas : [d.mesa]).map(Number);
      if (sh.getRange(1, 10).getValue() === '') sh.getRange(1, 10).setValue('Chegada');
      const col = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
      let achou = 0;
      for (let i = 1; i < col.length; i++) {
        if (alvo.indexOf(Number(col[i][0])) >= 0) { sh.getRange(i + 1, 10).setValue(d.chegou ? new Date() : ''); achou++; }
      }
      if (!achou) throw new Error('Mesa sem venda registrada.');
    } else if (d.acao === 'pixPrestacao') {
      autentica_(TESOUREIRO, d.pin); // só o recebedor gera o QR da prestação de contas
      const nome = String(d.vendedor || '');
      if (nome === TESOUREIRO || !PIN_HASH[nome]) throw new Error('Vendedor inválido.');
      const valor = Math.round(Number(d.valor) * 100) / 100;
      if (!(valor > 0)) throw new Error('Informe o valor.');
      const pend = pendente_(vendas, nome);
      if (valor > pend + 0.001) throw new Error('O valor passa do pendente de ' + nome + ' (R$ ' + pend.toFixed(2).replace('.', ',') + ').');
      pixResp = pixCom_(PIX_CHAVE_RECEBEDOR, PIX_NOME_RECEBEDOR, valor, 'FP26P' + Date.now().toString(36).toUpperCase());
    } else if (d.acao === 'resetar') {
      autentica_(ADMIN, d.pin); // só o Alesson
      if (d.confirmacao !== 'RESETAR') throw new Error('Digite RESETAR para confirmar.');
      const ss = SpreadsheetApp.openById(PLANILHA_ID), cfg = config_();
      const quando = Utilities.formatDate(new Date(), 'America/Manaus', 'dd-MM-yyyy HH.mm');
      // o arquivo leva o ano das vendas (mesmo que o ano da festa já tenha sido trocado)
      const anoDados = vendas.length && vendas[0].criadoEm ? new Date(vendas[0].criadoEm).getFullYear() : cfg.ano;
      arquivar_(ss, sh, 'Vendas ' + anoDados + ' (arquivo ' + quando + ')');
      arquivar_(ss, abaPrest_(), 'Prestações ' + anoDados + ' (arquivo ' + quando + ')');
      ss.setActiveSheet(ss.getSheets()[0]);
    } else if (d.acao === 'mudarAno') {
      autentica_(ADMIN, d.pin); // só o Alesson
      const ano = Number(d.ano), data = String(d.data || '');
      if (!(ano >= 2026 && ano <= 2100)) throw new Error('Ano inválido.');
      if (!/^\d{4}-\d{2}-\d{2}$/.test(data) || Number(data.slice(0, 4)) !== ano) throw new Error('A data da festa precisa ser do ano ' + ano + '.');
      PropertiesService.getScriptProperties().setProperties({ ANO: String(ano), DATA_FESTA: data });
    } else if (d.acao === 'vendas') {
      autentica_(d.vendedor, d.pin); // dados das vendas só para vendedores com PIN
    } else if (d.acao === 'telefones') {
      autentica_(d.vendedor, d.pin); // cada vendedor vê só os telefones dos próprios compradores
      telResp = {};
      vendas.forEach(function (x) { if (x.vendedor === d.vendedor && x.telefone) telResp[x.mesa] = x.telefone; });
    } else if (d.acao === 'pix') {
      autentica_(d.vendedor, d.pin);
      const ms = (d.mesas || []).map(Number);
      const g = vendaPix_(vendas, d.vendedor, ms);
      pixResp = pix_(g.reduce(function (s, x) { return s + x.valor; }, 0), ms, d.vendedor);
    } else if (d.acao === 'confirmarPix') {
      autentica_(d.vendedor, d.pin); // só quem vendeu confirma o recebimento
      const ms = (d.mesas || []).map(Number);
      vendaPix_(vendas, d.vendedor, ms);
      if (sh.getRange(1, 11).getValue() === '') sh.getRange(1, 11).setValue('PIX confirmado');
      const col = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
      for (let i = 1; i < col.length; i++) {
        if (ms.indexOf(Number(col[i][0])) >= 0) sh.getRange(i + 1, 11).setValue(d.desfazer ? '' : new Date());
      }
    } else if (d.acao === 'minhas') {
      autentica_(d.vendedor, d.pin); // cada vendedor vê só as próprias prestações
    } else if (d.acao === 'prestacoes') {
      autentica_(TESOUREIRO, d.pin);
    } else if (d.acao === 'receber') {
      autentica_(TESOUREIRO, d.pin);
      const nome = String(d.vendedor || '');
      if (nome === TESOUREIRO || !PIN_HASH[nome]) throw new Error('Vendedor inválido.');
      const valor = Math.round(Number(d.valor) * 100) / 100;
      if (!(valor > 0)) throw new Error('Informe o valor recebido.');
      if (PAGAMENTOS.indexOf(d.forma) < 0) throw new Error('Forma inválida.');
      const vendido = vendas.filter(function (x) { return x.vendedor === nome; })
        .reduce(function (s, x) { return s + x.valor; }, 0);
      const recebido = lerPrest_().filter(function (p) { return p.vendedor === nome; })
        .reduce(function (s, p) { return s + p.valor; }, 0);
      const pend = Math.round((vendido - recebido) * 100) / 100;
      if (valor > pend + 0.001) throw new Error('O valor passa do pendente de ' + nome + ' (R$ ' + pend.toFixed(2).replace('.', ',') + ').');
      abaPrest_().appendRow([Utilities.getUuid(), new Date(), nome, valor, d.forma, texto_(d.obs), TESOUREIRO]);
    } else if (d.acao === 'estornar') {
      autentica_(TESOUREIRO, d.pin);
      const sp = abaPrest_(), ids = sp.getRange(1, 1, sp.getLastRow(), 1).getValues();
      for (let i = ids.length - 1; i >= 1; i--) {
        if (String(ids[i][0]) === String(d.id)) { sp.deleteRow(i + 1); break; }
      }
    } else if (d.acao === 'cancelar') {
      const mesa = Number(d.mesa);
      const venda = vendas.filter(function (x) { return x.mesa === mesa; })[0];
      if (!venda) throw new Error('A mesa ' + mesa + ' não tem venda registrada.');
      autentica_(venda.vendedor, d.pin); // só quem vendeu pode cancelar
      const col = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
      for (let i = col.length - 1; i >= 1; i--) {
        if (Number(col[i][0]) === mesa) { sh.deleteRow(i + 1); break; }
      }
    } else {
      throw new Error('Ação desconhecida.');
    }
    SpreadsheetApp.flush();
    let atual = lerVendas_();
    if (atual.some(function (x) { return !x.passe; })) { garantirPasses_(sh); SpreadsheetApp.flush(); atual = lerVendas_(); }
    const resp = { ok: true, vendas: semTel_(atual), mesas: atual.map(function (x) { return x.mesa; }) };
    if (pixResp) resp.pix = pixResp;
    if (telResp) resp.telefones = telResp;
    resp.config = config_();
    if (['prestacoes', 'receber', 'estornar'].indexOf(d.acao) >= 0) resp.prestacoes = lerPrest_();
    if (d.acao === 'minhas') resp.prestacoes = lerPrest_().filter(function (p) { return p.vendedor === d.vendedor; });
    return json_(resp);
  } catch (err) {
    return json_({ ok: false, erro: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

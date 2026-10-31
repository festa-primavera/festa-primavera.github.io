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
  sh.getRange('I:J').setNumberFormat('dd/MM/yyyy HH:mm:ss');
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

function doGet() {
  return json_({ ok: true, vendas: lerVendas_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // uma venda por vez: impede vender a mesma mesa duas vezes
  try {
    const d = JSON.parse(e.postData.contents);
    const sh = aba_();
    const vendas = lerVendas_();

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
      if (TURMAS.indexOf(v.turma) < 0) throw new Error('Turma inválida.');
      if (!String(v.comprador || '').trim() || !String(v.aluno || '').trim())
        throw new Error('Preencha comprador e aluno.');
      const agora = new Date();
      const linhas = mesas.map(function (m) {
        return [m, texto_(v.comprador), texto_(v.aluno), texto_(v.serie), v.turma,
          v.pagamento, Number(v.valor), texto_(v.vendedor), agora];
      });
      sh.getRange(sh.getLastRow() + 1, 1, linhas.length, 9).setValues(linhas);
    } else if (d.acao === 'chegada') {
      // Recepção no dia da festa: marca (ou desmarca) a chegada do comprador
      const alvo = (Array.isArray(d.mesas) ? d.mesas : [d.mesa]).map(Number);
      if (sh.getRange(1, 10).getValue() === '') sh.getRange(1, 10).setValue('Chegada');
      const col = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
      let achou = 0;
      for (let i = 1; i < col.length; i++) {
        if (alvo.indexOf(Number(col[i][0])) >= 0) { sh.getRange(i + 1, 10).setValue(d.chegou ? new Date() : ''); achou++; }
      }
      if (!achou) throw new Error('Mesa sem venda registrada.');
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
    const resp = { ok: true, vendas: lerVendas_() };
    if (['prestacoes', 'receber', 'estornar'].indexOf(d.acao) >= 0) resp.prestacoes = lerPrest_();
    return json_(resp);
  } catch (err) {
    return json_({ ok: false, erro: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

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
  return SpreadsheetApp.openById(PLANILHA_ID).getSheets()[0];
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

    if (d.acao === 'vender') {
      const v = d.venda || {};
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
    } else if (d.acao === 'cancelar') {
      const mesa = Number(d.mesa);
      const col = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
      for (let i = col.length - 1; i >= 1; i--) {
        if (Number(col[i][0]) === mesa) { sh.deleteRow(i + 1); break; }
      }
    } else {
      throw new Error('Ação desconhecida.');
    }
    SpreadsheetApp.flush();
    return json_({ ok: true, vendas: lerVendas_() });
  } catch (err) {
    return json_({ ok: false, erro: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

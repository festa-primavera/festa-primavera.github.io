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
        criadoEm: r[8] instanceof Date ? r[8].toISOString() : String(r[8] || ''),
        chegada: r[9] instanceof Date ? r[9].toISOString() : String(r[9] || ''),
      };
    });
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
      const mesa = Number(v.mesa);
      if (!(mesa >= 1 && mesa <= TOTAL_MESAS)) throw new Error('Mesa inválida.');
      if (vendas.some(function (x) { return x.mesa === mesa; }))
        throw new Error('A mesa ' + mesa + ' acabou de ser vendida por outra pessoa.');
      if (!(Number(v.valor) > 0)) throw new Error('Informe o valor.');
      if (PAGAMENTOS.indexOf(v.pagamento) < 0) throw new Error('Forma de pagamento inválida.');
      if (TURMAS.indexOf(v.turma) < 0) throw new Error('Turma inválida.');
      if (!String(v.comprador || '').trim() || !String(v.aluno || '').trim())
        throw new Error('Preencha comprador e aluno.');
      sh.appendRow([mesa, texto_(v.comprador), texto_(v.aluno), texto_(v.serie), v.turma,
        v.pagamento, Number(v.valor), texto_(v.vendedor), new Date()]);
    } else if (d.acao === 'chegada') {
      // Recepção no dia da festa: marca (ou desmarca) a chegada do comprador
      const mesa = Number(d.mesa);
      if (sh.getRange(1, 10).getValue() === '') sh.getRange(1, 10).setValue('Chegada');
      const col = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
      let achou = false;
      for (let i = 1; i < col.length; i++) {
        if (Number(col[i][0]) === mesa) { sh.getRange(i + 1, 10).setValue(d.chegou ? new Date() : ''); achou = true; break; }
      }
      if (!achou) throw new Error('A mesa ' + mesa + ' não tem venda registrada.');
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

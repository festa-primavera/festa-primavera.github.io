// Camada de dados: planilha do Google Drive (via Google Apps Script) ou modo demonstração
import { APPS_SCRIPT_URL, INTERVALO_ATUALIZACAO } from "./config.js?v=20261105-1610";

export async function criarStore() {
  return APPS_SCRIPT_URL ? planilhaStore() : demoStore();
}

// ------------------------------------------------- Google Drive (planilha)
function planilhaStore() {
  let vendas = [], timer = null;
  const ouvPub = new Set(), ouvVen = new Set(), ouvErr = new Set();

  const normal = (l) => l.map((v) => ({ ...v, mesa: Number(v.mesa), valor: Number(v.valor),
    criadoEm: v.criadoEm ? new Date(v.criadoEm) : new Date() }));
  function publicar(l) {
    vendas = normal(l);
    ouvPub.forEach((cb) => cb(vendas.map((v) => v.mesa).sort((a, b) => a - b)));
    ouvVen.forEach((cb) => cb(vendas));
  }
  async function buscar() {
    try {
      const r = await fetch(APPS_SCRIPT_URL + "?t=" + Date.now(), { cache: "no-store" });
      const d = await r.json();
      if (!d.ok) throw new Error(d.erro || "Erro na planilha");
      publicar(d.vendas);
    } catch (e) { ouvErr.forEach((cb) => cb(e)); }
  }
  async function enviar(corpo) {
    // text/plain evita a verificação CORS prévia, que o Apps Script não responde
    const r = await fetch(APPS_SCRIPT_URL, { method: "POST", body: JSON.stringify(corpo),
      headers: { "Content-Type": "text/plain;charset=utf-8" } });
    const d = await r.json();
    if (!d.ok) throw new Error(d.erro || "Não foi possível salvar.");
    publicar(d.vendas);
    return d;
  }
  function iniciar() {
    if (timer) return;
    buscar();
    timer = setInterval(() => { if (!document.hidden) buscar(); }, INTERVALO_ATUALIZACAO * 1000);
    document.addEventListener("visibilitychange", () => { if (!document.hidden) buscar(); });
  }

  return {
    modo: "planilha",
    onPublico(cb, onErr) { ouvPub.add(cb); onErr && ouvErr.add(onErr); iniciar(); },
    onVendas(cb, onErr) { ouvVen.add(cb); onErr && ouvErr.add(onErr); iniciar(); },
    entrar: (vendedor, pin) => enviar({ acao: "entrar", vendedor, pin }),
    vender: (venda, pin) => enviar({ acao: "vender", venda, pin }),
    cancelar: (mesa, pin) => enviar({ acao: "cancelar", mesa, pin }),
    chegada: (mesas, chegou) => enviar({ acao: "chegada", mesas: [].concat(mesas), chegou }),
    // Prestação de contas (só o recebedor, com o PIN dele)
    prestacoes: async (pin) => (await enviar({ acao: "prestacoes", pin })).prestacoes || [],
    receber: async (r, pin) => (await enviar({ acao: "receber", ...r, pin })).prestacoes || [],
    estornar: async (id, pin) => (await enviar({ acao: "estornar", id, pin })).prestacoes || [],
    // PIX: o código (QR) vem do servidor; só quem vendeu confirma o recebimento
    pix: (vendedor, pin, mesas) => enviar({ acao: "pix", vendedor, pin, mesas }),
    confirmarPix: (vendedor, pin, mesas, desfazer) => enviar({ acao: "confirmarPix", vendedor, pin, mesas, desfazer: !!desfazer }),
    // Cada vendedor vê só as próprias prestações (com o PIN dele)
    minhas: async (vendedor, pin) => (await enviar({ acao: "minhas", vendedor, pin })).prestacoes || [],
  };
}

// ---------------------------------------------------------- Demonstração
function demoStore() {
  const KEY = "festa-primavera-demo";
  const ler = () => { try { return JSON.parse(localStorage.getItem(KEY)) || []; } catch { return []; } };
  const gravar = (l) => { try { localStorage.setItem(KEY, JSON.stringify(l)); } catch {} avisar(); };
  const ouvintesPub = new Set(), ouvintesVen = new Set();
  const lerP = () => { try { return JSON.parse(localStorage.getItem(KEY + "-prest")) || []; } catch { return []; } };
  const gravarP = (l) => { try { localStorage.setItem(KEY + "-prest", JSON.stringify(l)); } catch {} };

  function avisar() {
    const l = ler().map((v) => ({ ...v, criadoEm: new Date(v.criadoEm) }));
    ouvintesPub.forEach((cb) => cb(l.map((v) => v.mesa).sort((a, b) => a - b)));
    ouvintesVen.forEach((cb) => cb(l));
  }
  window.addEventListener("storage", (e) => { if (e.key === KEY) avisar(); });

  return {
    modo: "demo",
    onPublico(cb) { ouvintesPub.add(cb); setTimeout(avisar); },
    onVendas(cb) { ouvintesVen.add(cb); setTimeout(avisar); },
    async vender(v) {
      const l = ler(), mesas = v.mesas || [v.mesa];
      const ja = mesas.filter((m) => l.some((x) => x.mesa === m));
      if (ja.length) throw new Error(`A mesa ${ja.join(", ")} já foi vendida.`);
      const { mesas: _m, ...base } = v;
      mesas.forEach((mesa) => l.push({ ...base, mesa, criadoEm: new Date().toISOString() })); gravar(l);
    },
    async entrar(nome, pin) { if (!/^\d{4,6}$/.test(pin)) throw new Error("PIN incorreto."); },
    async cancelar(mesa) { gravar(ler().filter((x) => x.mesa !== mesa)); },
    async prestacoes() { return lerP(); },
    async receber(r) { const l = lerP(); l.push({ ...r, id: String(Date.now()), quando: new Date().toISOString() }); gravarP(l); return l; },
    async pix(nome, pin, mesas) { return { pix: { payload: "DEMONSTRACAO-SEM-CHAVE-PIX", valor: 50 * mesas.length, recebedor: "(modo demonstração)" } }; },
    async confirmarPix(nome, pin, mesas) { gravar(ler().map((x) => mesas.includes(x.mesa) ? { ...x, pixOk: new Date().toISOString() } : x)); },
    async minhas(nome) { return lerP().filter((p) => p.vendedor === nome); },
    async estornar(id) { const l = lerP().filter((p) => p.id !== id); gravarP(l); return l; },
    async chegada(mesas, chegou) {
      const lista = [].concat(mesas);
      gravar(ler().map((x) => lista.includes(x.mesa) ? { ...x, chegada: chegou ? new Date().toISOString() : "" } : x));
    },
  };
}

// Camada de dados: planilha do Google Drive (via Google Apps Script) ou modo demonstração
import { APPS_SCRIPT_URL, INTERVALO_ATUALIZACAO } from "./config.js";

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
    vender: (venda) => enviar({ acao: "vender", venda }),
    cancelar: (mesa) => enviar({ acao: "cancelar", mesa }),
    chegada: (mesa, chegou) => enviar({ acao: "chegada", mesa, chegou }),
  };
}

// ---------------------------------------------------------- Demonstração
function demoStore() {
  const KEY = "festa-primavera-demo";
  const ler = () => { try { return JSON.parse(localStorage.getItem(KEY)) || []; } catch { return []; } };
  const gravar = (l) => { try { localStorage.setItem(KEY, JSON.stringify(l)); } catch {} avisar(); };
  const ouvintesPub = new Set(), ouvintesVen = new Set();

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
      const l = ler();
      if (l.some((x) => x.mesa === v.mesa)) throw new Error(`A mesa ${v.mesa} já foi vendida.`);
      l.push({ ...v, criadoEm: new Date().toISOString() }); gravar(l);
    },
    async cancelar(mesa) { gravar(ler().filter((x) => x.mesa !== mesa)); },
    async chegada(mesa, chegou) {
      gravar(ler().map((x) => x.mesa === mesa ? { ...x, chegada: chegou ? new Date().toISOString() : "" } : x));
    },
  };
}

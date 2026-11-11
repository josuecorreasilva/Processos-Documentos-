const STORAGE_KEY = "painel-dap-v1";
const PASTA_MEMORANDOS = "https://drive.google.com/drive/folders/1OoepixxhTFxVzkzwqMBn5KqsOEFRrVJ3?usp=drive_link";
const STATUS = ["Aguardando Entrega", "Em Andamento", "Liquidado", "Pago", "Cancelado"];
const SETORES_BASE = ["Compras", "DAF", "Licitações", "Emendas", "Atenção Primária", "Planejamento", "Empenhos", "COPES", "CCGOF"];
const VIEWS = ["painel", "registros", "empenhos", "emendas", "compras", "materiais", "busca", "dados"];
const SITUACOES_RP = ["Em elaboração", "DFD pronto", "ETP pronto", "TR pronto", "Encaminhado"];
const PECAS_RP = [
  { chave: "dfd", rotulo: "DFD", nome: "Documento de formalização da demanda" },
  { chave: "etp", rotulo: "ETP", nome: "Estudo técnico preliminar" },
  { chave: "tr", rotulo: "TR", nome: "Termo de referência" },
];

const ui = {
  view: "painel",
  q: "",
  tab: "controle",
  filtros: {
    empenhos: { texto: "", status: "", fonte: "", ano: "" },
    itens: { texto: "", tipo: "", ano: "" },
    emendas: { texto: "", ano: "", tipo: "" },
    compras: { texto: "", setor: "", ano: "" },
    arquivos: { texto: "", tipo: "" },
    materiais: { ano: "", texto: "" },
    registros: { ano: "", texto: "" },
  },
  aberto: null,
  form: null,
  editId: null,
  _mat: null,
  _keep: false,
};

function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
function clone(v) { return JSON.parse(JSON.stringify(v)); }
function uid() { return crypto.randomUUID ? crypto.randomUUID() : `id-${Date.now()}-${Math.random()}`; }
function fold(s) {
  return String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
}
function tem(texto, q) { return fold(texto).includes(fold(q)); }
function soma(list, campo = "valor") {
  const t = list.reduce((a, x) => a + (Number(x[campo]) || 0), 0);
  return Math.round(t * 100) / 100;
}
function brl(n) {
  if (n == null || n === "" || Number.isNaN(Number(n))) return "—";
  return Number(n).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
}
function brl4(n) {
  if (n == null || n === "" || Number.isNaN(Number(n))) return "—";
  return Number(n).toLocaleString("pt-BR", { style: "currency", currency: "BRL", minimumFractionDigits: 2, maximumFractionDigits: 4 });
}
function qtdFmt(n) {
  if (n == null || n === "" || Number.isNaN(Number(n))) return "—";
  return Number(n).toLocaleString("pt-BR", { maximumFractionDigits: 2 });
}
function dataBR(iso) {
  const m = String(iso || "").match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? `${m[3]}/${m[2]}/${m[1]}` : "";
}
function mostrarData(iso, texto) {
  if (iso) return dataBR(iso);
  const t = String(texto || "").trim();
  return t || "—";
}
function clip(s, n) {
  const t = String(s || "").replace(/\s+/g, " ").trim();
  if (!t) return "—";
  return t.length > n ? `${t.slice(0, n - 1)}…` : t;
}
function prose(s) {
  const t = String(s || "").trim();
  return t ? `<div class="prose">${esc(t)}</div>` : `<span class="muted">—</span>`;
}
function parseValor(v) {
  const s = String(v ?? "").trim();
  if (!s) return null;
  const n = s.includes(",") ? Number(s.replace(/\./g, "").replace(",", ".")) : Number(s);
  return Number.isFinite(n) ? Math.round(n * 10000) / 10000 : null;
}
function normSis(s) {
  const m = String(s || "").match(/(\d{2,6})\s*\/\s*(\d{4})/);
  return m ? `${m[1]}/${m[2]}` : "";
}
function normEmenda(s) {
  const m = String(s || "").match(/(\d{1,4})\s*\/\s*(\d{4})/);
  return m ? `${String(Number(m[1])).padStart(3, "0")}/${m[2]}` : "";
}
function normEmp(s) { return String(s || "").toUpperCase().replace(/\s+/g, ""); }
function anoPlausivel(s) {
  return /^20(1[8-9]|2\d|3[0-5])$/.test(String(s || "").trim());
}
function anoDoNumero(s) {
  const m = String(s || "").match(/\/(20\d{2})\s*$/);
  return m ? m[1] : "";
}
function anoCriacao(e) {
  const dataAno = String((e && e.data) || "").slice(0, 4);
  if (anoPlausivel(dataAno)) return dataAno;
  return anoDoNumero(e && (e.numero || e.numeroEmpenho)) || (anoPlausivel(e && e.ano) ? String(e.ano).trim() : "");
}
function anoEmpenho(e) {
  if (e && e.anoManual && anoPlausivel(e.ano)) return String(e.ano).trim();
  return anoCriacao(e);
}
function emendaMencionada(texto) {
  const m = String(texto || "").match(/emenda(?:\s+impositiva)?\s*(?:n[º°.]\s*)?(\d{1,4})\s*\/\s*(\d{4})/i);
  return m ? normEmenda(`${m[1]}/${m[2]}`) : "";
}
function sisesNoTexto(texto) {
  const out = [];
  const re = /SIS\s*n?[º°.]?\s*(\d{2,6})\s*\/\s*(\d{4})/gi;
  let m;
  const src = String(texto || "");
  while ((m = re.exec(src))) out.push(`${m[1]}/${m[2]}`);
  return out;
}
function fonteDe(texto) {
  const m = String(texto || "").match(/Fonte:\s*([^;\n]+)/i);
  return m ? m[1].trim() : "";
}
function fonteCurta(texto) {
  const f = fonteDe(texto).replace(/\s+/g, " ").trim();
  if (!f) return "Não informada";
  const m = f.match(/^([\d.]+)/);
  if (!m) return f.length > 42 ? `${f.slice(0, 41)}…` : f;
  return m[1].replace(/\.$/, "");
}
function tipoItem(t) {
  const s = String(t || "").replace(/\s+/g, " ").trim();
  if (!s) return "Não informado";
  return s.replace(/:\s*/, ": ");
}
function classeStatus(s) {
  const k = fold(s);
  if (k.includes("aguard")) return "st-wait";
  if (k.includes("andamento")) return "st-run";
  if (k.includes("liquid") || k.includes("pago")) return "st-ok";
  if (k.includes("cancel")) return "st-bad";
  return "st-muted";
}
function pill(s) {
  const label = String(s || "").trim() || "Sem status";
  return `<span class="pill ${classeStatus(label)}">${esc(label)}</span>`;
}
function unique(list) {
  return [...new Set(list.filter(Boolean))].sort((a, b) => a.localeCompare(b, "pt-BR"));
}
function opcoesStatus(atual) {
  const lista = [{ value: "", label: "Sem status" }, ...STATUS.map((s) => ({ value: s, label: s }))];
  if (atual && !lista.some((o) => o.value === atual)) lista.push({ value: atual, label: atual });
  return lista;
}

function sane(data) {
  const base = data || {};
  base.empenhos = base.empenhos || [];
  base.itens = base.itens || [];
  base.liberacoes = base.liberacoes || [];
  base.emendas = (base.emendas || []).filter((e) => /^\d{3}\/\d{4}$/.test(String(e.numero || "")));
  base.processos = base.processos || [];
  base.atas = base.atas || [];
  base.lotes = base.lotes || [];
  (base.lotes || []).forEach((l) => {
    if (!(l.anoManual && anoPlausivel(l.ano))) l.ano = "2026";
  });
  base.consolidados = base.consolidados || [];
  base.registros = Array.isArray(base.registros) ? base.registros : [];
  base.registros.forEach((r) => {
    if (!r.id) r.id = uid();
    if (!anoPlausivel(r.ano)) r.ano = "2026";
    if (!SITUACOES_RP.includes(r.situacao)) r.situacao = "Em elaboração";
    r.materiais = Array.isArray(r.materiais) ? r.materiais : [];
    r.materiais.forEach((m) => { if (!m.id) m.id = uid(); });
  });
  const idsRp = new Set(base.registros.map((r) => r.id));
  const nomesRp = new Set(base.registros.map((r) => fold(r.nome)));
  (window.ACERVO_RP || []).forEach((r) => {
    if (!r || !r.id || idsRp.has(r.id) || nomesRp.has(fold(r.nome))) return;
    base.registros.push(clone(r));
    idsRp.add(r.id);
    nomesRp.add(fold(r.nome));
  });
  base.atalhos = Object.assign({ documentosRelacao: "", pastaEmendas: "" }, base.atalhos || {});
  if (!base.atalhos.pastaEmendas) base.atalhos.pastaEmendas = PASTA_MEMORANDOS;
  const mapaDoc = new Map();
  ((window.SEED && window.SEED.empenhos) || []).forEach((e) => {
    if (e && e.documento && e.numero) mapaDoc.set(String(e.numero).toUpperCase().replace(/\s+/g, ""), e.documento);
  });
  (base.empenhos || []).forEach((e) => {
    if (!e.documento) {
      const url = mapaDoc.get(String(e.numero || "").toUpperCase().replace(/\s+/g, ""));
      if (url) e.documento = url;
    }
    if (!e.anoManual) {
      const calculado = anoCriacao(e);
      if (calculado) e.ano = calculado;
    }
    const oficial = anoDoNumero(e.numero);
    const guardado = String(e.ano ?? "");
    if (!anoPlausivel(guardado) || (oficial && guardado !== oficial && String(e.numero || "").includes(guardado))) {
      e.anoManual = false;
      e.ano = anoCriacao(e) || oficial;
    }
  });
  const presentes = new Set((base.empenhos || []).map((e) => String(e.numero || "").toUpperCase().replace(/\s+/g, "")));
  ((window.SEED && window.SEED.empenhos) || []).forEach((e) => {
    const chave = String((e && e.numero) || "").toUpperCase().replace(/\s+/g, "");
    if (!chave || presentes.has(chave)) return;
    base.empenhos.push(clone(e));
    presentes.add(chave);
  });
  return base;
}
function load() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return sane(clone(window.SEED || {}));
    const parsed = JSON.parse(raw);
    if (!parsed.empenhos || !parsed.emendas || !parsed.processos) return sane(clone(window.SEED || {}));
    return sane(parsed);
  } catch {
    return sane(clone(window.SEED || {}));
  }
}
function save() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch {
    avisar("Não foi possível salvar neste navegador. Baixe uma cópia em Cópia dos dados.");
  }
}
let state = load();
let anexos = [];

function abrirDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open("painel-dap-arquivos", 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains("arquivos")) {
        const store = db.createObjectStore("arquivos", { keyPath: "id" });
        store.createIndex("empenhoId", "empenhoId", { unique: false });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
function pedir(req) {
  return new Promise((resolve, reject) => {
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
function esperarTx(tx) {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error);
  });
}
async function listarAnexos() {
  const db = await abrirDb();
  const todos = await pedir(db.transaction("arquivos").objectStore("arquivos").getAll());
  db.close();
  anexos = todos
    .map(({ arquivo, ...meta }) => meta)
    .sort((a, b) => String(b.incluidoEm || "").localeCompare(String(a.incluidoEm || "")));
}
function anexosDo(empenhoId) {
  return anexos.filter((a) => a.empenhoId === empenhoId);
}
function tamanhoArquivo(n) {
  const v = Number(n) || 0;
  if (v < 1024) return `${v} B`;
  if (v < 1048576) return `${Math.round(v / 1024)} KB`;
  return `${(v / 1048576).toLocaleString("pt-BR", { maximumFractionDigits: 1 })} MB`;
}
function rotuloAnexo(tipo) {
  if (tipo === "nota") return "Nota fiscal";
  if (tipo === "dfd") return "DFD";
  if (tipo === "etp") return "ETP";
  if (tipo === "tr") return "TR";
  return "Documento do empenho";
}
function marcaAnexos(e) {
  const lista = anexosDo(e.id);
  if (!lista.length) return "";
  const doc = lista.filter((a) => a.tipo !== "nota").length;
  const nf = lista.filter((a) => a.tipo === "nota").length;
  const partes = [];
  if (doc) partes.push(doc === 1 ? "1 documento" : `${doc} documentos`);
  if (nf) partes.push(nf === 1 ? "1 nota" : `${nf} notas`);
  return `<div class="caption">${partes.join(" · ")}</div>`;
}
async function gravarAnexo(refId, tipo, file) {
  preservarRegistro();
  const e = empenhoPorId(refId);
  const reg = registroPorId(refId);
  if ((!e && !reg) || !file) return;
  if (reg) save();
  const tipos = ["nota", "empenho", "dfd", "etp", "tr"];
  const rec = {
    id: uid(),
    empenhoId: e ? e.id : "",
    registroId: reg ? reg.id : "",
    numero: e ? e.numero : reg.nome,
    tipo: tipos.includes(tipo) ? tipo : "empenho",
    nome: file.name || "arquivo",
    mime: file.type || "application/octet-stream",
    tamanho: file.size || 0,
    incluidoEm: new Date().toISOString(),
    arquivo: file,
  };
  try {
    const db = await abrirDb();
    const tx = db.transaction("arquivos", "readwrite");
    tx.objectStore("arquivos").put(rec);
    await esperarTx(tx);
    db.close();
    await listarAnexos();
    render();
    avisar(rec.tipo === "nota" ? "Nota fiscal registrada neste navegador." : "Documento registrado neste navegador.");
  } catch {
    avisar("Não foi possível guardar o arquivo neste navegador.");
  }
}
async function abrirAnexo(id) {
  try {
    const db = await abrirDb();
    const rec = await pedir(db.transaction("arquivos").objectStore("arquivos").get(id));
    db.close();
    if (!rec || !rec.arquivo) {
      avisar("Não encontrei esse arquivo.");
      return;
    }
    const url = URL.createObjectURL(rec.arquivo);
    window.open(url, "_blank", "noopener");
    setTimeout(() => URL.revokeObjectURL(url), 120000);
  } catch {
    avisar("Não foi possível abrir o arquivo.");
  }
}
async function removerAnexo(id) {
  const item = anexos.find((a) => a.id === id);
  if (!item || !confirm(`Remover ${item.nome}?`)) return;
  try {
    const db = await abrirDb();
    const tx = db.transaction("arquivos", "readwrite");
    tx.objectStore("arquivos").delete(id);
    await esperarTx(tx);
    db.close();
    await listarAnexos();
    render();
    avisar("Arquivo removido deste navegador.");
  } catch {
    avisar("Não foi possível remover o arquivo.");
  }
}
function anexosDoRegistro(registroId, tipo) {
  return anexos.filter((a) => a.registroId === registroId && (!tipo || a.tipo === tipo));
}
async function removerAnexosDoEmpenho(empenhoId) {
  const ids = anexosDo(empenhoId).map((a) => a.id);
  if (!ids.length) return;
  const db = await abrirDb();
  const tx = db.transaction("arquivos", "readwrite");
  const store = tx.objectStore("arquivos");
  ids.forEach((id) => store.delete(id));
  await esperarTx(tx);
  db.close();
  await listarAnexos();
}
async function removerAnexosDoRegistro(registroId) {
  const ids = anexosDoRegistro(registroId).map((a) => a.id);
  if (!ids.length) return;
  const db = await abrirDb();
  const tx = db.transaction("arquivos", "readwrite");
  const store = tx.objectStore("arquivos");
  ids.forEach((id) => store.delete(id));
  await esperarTx(tx);
  db.close();
  await listarAnexos();
}
function blobParaBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const s = String(reader.result || "");
      resolve(s.slice(s.indexOf(",") + 1));
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}
async function exportarCopia() {
  try {
    const db = await abrirDb();
    const todos = await pedir(db.transaction("arquivos").objectStore("arquivos").getAll());
    db.close();
    const lista = [];
    for (const rec of todos) {
      lista.push({
        id: rec.id,
        empenhoId: rec.empenhoId,
        registroId: rec.registroId || "",
        numero: rec.numero,
        tipo: rec.tipo,
        nome: rec.nome,
        mime: rec.mime,
        tamanho: rec.tamanho,
        incluidoEm: rec.incluidoEm,
        base64: await blobParaBase64(rec.arquivo),
      });
    }
    baixar("painel-compras-empenhos.json", JSON.stringify({ ...state, anexos: lista }, null, 2), "application/json");
    avisar(lista.length ? `Cópia baixada, com ${lista.length} arquivo(s).` : "Cópia baixada.");
  } catch {
    avisar("Não foi possível montar a cópia.");
  }
}
async function restaurarAnexos(lista) {
  const db = await abrirDb();
  const tx = db.transaction("arquivos", "readwrite");
  const store = tx.objectStore("arquivos");
  store.clear();
  lista.forEach((a) => {
    if (!a || !a.id || !a.base64) return;
    const bin = atob(a.base64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i += 1) bytes[i] = bin.charCodeAt(i);
    store.put({
      id: a.id,
      empenhoId: a.empenhoId || "",
      registroId: a.registroId || "",
      numero: a.numero || "",
      tipo: ["nota", "empenho", "dfd", "etp", "tr"].includes(a.tipo) ? a.tipo : "empenho",
      nome: a.nome || "arquivo",
      mime: a.mime || "application/octet-stream",
      tamanho: a.tamanho || bytes.length,
      incluidoEm: a.incluidoEm || new Date().toISOString(),
      arquivo: new Blob([bytes], { type: a.mime || "application/octet-stream" }),
    });
  });
  await esperarTx(tx);
  db.close();
  await listarAnexos();
}

function avisar(msg) {
  const el = document.getElementById("toast");
  el.hidden = false;
  el.textContent = msg;
  clearTimeout(avisar._t);
  avisar._t = setTimeout(() => { el.hidden = true; }, 3400);
}

function htmlAtalho(chave, rotulo, placeholder) {
  const url = String((state.atalhos && state.atalhos[chave]) || "").trim();
  const ok = /^https?:\/\//i.test(url);
  return `<div class="atalho">
    ${ok ? `<a class="btn btn-small" href="${esc(url)}" target="_blank" rel="noopener noreferrer">${esc(rotulo)}</a>` : `<span class="caption">${esc(rotulo)}</span>`}
    <input class="input" data-act="atalho" data-chave="${esc(chave)}" value="${esc(url)}" placeholder="${esc(placeholder || "Cole o link do Drive")}">
  </div>`;
}
function urlDoc(e) {
  const url = String((e && e.documento) || "").trim();
  return /^https?:\/\//i.test(url) ? url : "";
}
function linkDoc(e, rotulo) {
  const url = urlDoc(e);
  const texto = rotulo || (e && e.numero) || "";
  if (!url) return esc(texto);
  return `<a class="doc" href="${esc(url)}" target="_blank" rel="noopener noreferrer" title="Abrir documento">${esc(texto)}</a>`;
}
function empenhoPorId(id) { return state.empenhos.find((e) => e.id === id); }
function registroPorId(id) { return (state.registros || []).find((r) => r.id === id); }
function linkHttp(v) {
  const s = String(v || "").trim();
  if (!s) return "";
  return /^https?:\/\//i.test(s) ? s : null;
}
function linkDocRp(v) {
  const s = String(v || "").trim();
  if (!s) return "";
  if (/^https?:\/\//i.test(s)) return s;
  if (/^documentos\/[\w./-]+$/i.test(s) && !s.includes("..")) return s;
  return null;
}
function anosDasEmendas(texto) {
  return unique(emendasCitadas(texto).map((n) => String(n).slice(-4)));
}
function emendaPorId(id) { return state.emendas.find((e) => e.id === id); }
function emendaPorNumero(n) {
  const k = normEmenda(n);
  return k ? state.emendas.find((e) => e.numero === k) : null;
}
function procPorId(id) { return state.processos.find((p) => p.id === id); }
function procPorSis(sis) {
  const k = normSis(sis);
  return k ? state.processos.find((p) => p.sis === k) : null;
}
function sisesEmpenho(e) {
  return unique([normSis(e.sis), ...sisesNoTexto(`${e.observacoes || ""} ${e.objeto || ""}`)]);
}
function emendaDoEmpenho(e) {
  return emendaPorNumero(emendaMencionada(e.observacoes) || e.emenda);
}
function itensDoNumero(numero) {
  const k = normEmp(numero);
  return state.itens.filter((i) => normEmp(i.numeroEmpenho) === k);
}
function gruposItens() {
  const map = new Map();
  for (const item of state.itens) {
    const k = normEmp(item.numeroEmpenho) || item.id;
    if (!map.has(k)) map.set(k, []);
    map.get(k).push(item);
  }
  return [...map.values()].sort((a, b) => String(b[0].data || "").localeCompare(String(a[0].data || "")));
}
function emendasCitadas(texto) {
  const keys = [];
  const re = /(\d{1,4})\s*\/\s*(\d{4})/g;
  let m;
  const src = String(texto || "");
  while ((m = re.exec(src))) {
    const k = normEmenda(`${m[1]}/${m[2]}`);
    if (k && !keys.includes(k)) keys.push(k);
  }
  return keys;
}
function ehLegenda(l) {
  const nome = String(l.item || "").trim();
  if (/^(Observações|Emendas Agrupadas|Licitação de Computadores|Kits Odontológicos)\b/i.test(nome)) return true;
  return !l.sis && !String(l.destinos || "").trim() && !String(l.descricao || "").trim();
}
function materiaisDaEmenda(numero) {
  return state.lotes.filter((l) => !ehLegenda(l) && emendasCitadas(l.emendas).includes(numero));
}
function hojeISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
function situacaoAtual(p) {
  return String((p && (p.descricao || p.observacao)) || "").trim();
}
function opcoesSetor(atual) {
  return unique([...SETORES_BASE, atual, ...state.processos.map((p) => p.setor)].filter(Boolean));
}
function htmlSetor(atual, id) {
  const opcoes = [{ value: "", label: "Sem setor" }, ...opcoesSetor(atual).map((s) => ({ value: s, label: s }))];
  const html = opcoes.map((o) => `<option value="${esc(o.value)}"${o.value === (atual || "") ? " selected" : ""}>${esc(o.label)}</option>`).join("");
  return `<select class="select status-select" data-act="setor-proc" data-id="${esc(id)}">${html}</select>`;
}
function processosPorSisCurto(sis) {
  const n = String(sis || "").trim();
  if (!n) return [];
  return state.processos.filter((p) => p.sis === n || String(p.sis || "").startsWith(`${n}/`));
}

function topo(titulo, lede, acoes = "") {
  return `<header class="topbar"><div><h1>${titulo}</h1><p class="lede">${lede}</p></div><div class="top-actions">${acoes}</div></header>`;
}
function campo(label, name, value, opts = {}) {
  const wide = opts.wide ? "wide" : opts.half ? "half" : "";
  if (opts.area) {
    return `<label class="field ${wide}"><span>${esc(label)}</span><textarea class="textarea" name="${name}">${esc(value ?? "")}</textarea></label>`;
  }
  if (opts.options) {
    const cur = value ?? "";
    const html = opts.options.map((o) => {
      const val = o && typeof o === "object" ? o.value : o;
      const rotulo = o && typeof o === "object" ? o.label : o;
      return `<option value="${esc(val)}"${val === cur ? " selected" : ""}>${esc(rotulo)}</option>`;
    }).join("");
    return `<label class="field ${wide}"><span>${esc(label)}</span><select class="select" name="${name}">${html}</select></label>`;
  }
  const type = opts.type || "text";
  const step = opts.step ? ` step="${opts.step}"` : "";
  return `<label class="field ${wide}"><span>${esc(label)}</span><input class="input" type="${type}" name="${name}" value="${esc(value ?? "")}"${step}></label>`;
}
function chip(tipo, id, label) {
  if (!id) return "";
  return `<button type="button" class="chip" data-act="open" data-tipo="${esc(tipo)}" data-id="${esc(id)}">${esc(label)}</button>`;
}
function tabela(cols, rows, vazio) {
  const head = cols.map((c) => `<th class="${c.cls || ""}">${c.label}</th>`).join("");
  const body = rows.length ? rows.join("") : `<tr><td class="empty" colspan="${cols.length}">${vazio || "Nenhum registro com esse filtro."}</td></tr>`;
  return `<div class="table-wrap"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
}
function td(html, cls = "") { return `<td class="${cls}">${html}</td>`; }

function barras(grupos, modo) {
  const max = Math.max(...grupos.map((g) => (modo === "n" ? g.n : g.valor)), 1);
  return grupos.map((g) => {
    const w = Math.max(2, ((modo === "n" ? g.n : g.valor) / max) * 100);
    const direita = modo === "n" ? `<strong>${g.n}</strong>` : `<strong>${brl(g.valor)}</strong><div class="caption">${g.n}</div>`;
    return `<div class="bar-row"><div>${esc(g.nome)}</div><div class="bar-track"><div class="bar-fill" style="width:${w}%"></div></div><div class="bar-amt">${direita}</div></div>`;
  }).join("");
}

function viewPainel() {
  const emps = state.empenhos.filter((e) => e.origem !== "relacao");
  const total = soma(emps);
  const por = { aguardando: [], andamento: [], liquidado: [], cancelado: [], outro: [] };
  emps.forEach((e) => {
    const k = fold(e.status);
    if (k.includes("aguard")) por.aguardando.push(e);
    else if (k.includes("andamento")) por.andamento.push(e);
    else if (k.includes("liquid") || k.includes("pago")) por.liquidado.push(e);
    else if (k.includes("cancel")) por.cancelado.push(e);
    else por.outro.push(e);
  });
  const fontes = new Map();
  emps.forEach((e) => {
    const nome = fonteCurta(e.observacoes);
    if (!fontes.has(nome)) fontes.set(nome, { nome, valor: 0, n: 0 });
    const g = fontes.get(nome);
    g.valor += Number(e.valor) || 0;
    g.n += 1;
  });
  const fonteLista = [...fontes.values()].sort((a, b) => b.valor - a.valor);
  const setores = new Map();
  state.processos.forEach((p) => {
    const nome = p.setor || "Sem setor";
    if (!setores.has(nome)) setores.set(nome, { nome, n: 0, valor: 0 });
    setores.get(nome).n += 1;
  });
  const setorLista = [...setores.values()].sort((a, b) => b.n - a.n).slice(0, 8);
  const anos = new Map();
  state.emendas.forEach((e) => {
    const nome = e.ano || "—";
    if (!anos.has(nome)) anos.set(nome, { nome, n: 0, valor: 0 });
    const g = anos.get(nome);
    g.n += 1;
    g.valor += Number(e.valor) || 0;
  });
  const espera = [...por.aguardando].sort((a, b) => String(b.data || "").localeCompare(String(a.data || ""))).slice(0, 6);
  const linhasEspera = espera.map((e) => `<tr data-tipo="empenho" data-id="${esc(e.id)}">${td(mostrarData(e.data, e.dataTexto), "nowrap")}${td(linkDoc(e), "nowrap")}${td(esc(clip(e.fornecedor, 42)))}${td(brl(e.valor), "num")}</tr>`).join("");
  const cardsAno = [...anos.values()].sort((a, b) => b.nome.localeCompare(a.nome)).map((g) => (
    `<div class="mix-card a"><div class="caption">Emendas ${esc(g.nome)}</div><div class="n">${g.n}</div><div class="caption">${brl(g.valor)} previstos</div></div>`
  )).join("");

  return `
    ${topo("Painel", "O processo entra em Compras. Quando o empenho sai, a entrega se acompanha em Empenhos. A emenda é a origem do recurso.")}
    <section class="kpi-row">
      <div class="kpi hero"><div class="lbl">Valor empenhado</div><div class="val">${brl(total)}</div><div class="hint">${emps.length} registros do controle de 2026</div></div>
      <button class="kpi" data-act="goto" data-arg="empenhos"><div class="lbl">Aguardando entrega</div><div class="val">${por.aguardando.length}</div><div class="hint">${brl(soma(por.aguardando))}</div></button>
      <button class="kpi" data-act="goto" data-arg="empenhos"><div class="lbl">Em andamento</div><div class="val">${por.andamento.length}</div><div class="hint">${brl(soma(por.andamento))}</div></button>
      <button class="kpi" data-act="goto" data-arg="empenhos"><div class="lbl">Liquidados ou pagos</div><div class="val">${por.liquidado.length}</div><div class="hint">${brl(soma(por.liquidado))}</div></button>
      <button class="kpi" data-act="goto" data-arg="compras"><div class="lbl">Processos de compra</div><div class="val">${state.processos.length}</div><div class="hint">${state.processos.filter((p) => p.ano === "2026").length} com ano 2026</div></button>
    </section>
    <section class="mix">${cardsAno}
      <div class="mix-card b"><div class="caption">Empenhos da relação</div><div class="n">${state.empenhos.filter((e) => e.origem === "relacao").length}</div><div class="caption">${state.empenhos.filter((e) => e.origem === "relacao" && anoEmpenho(e) === "2025").length} de 2025, fora do controle</div></div>
      <div class="mix-card b"><div class="caption">Itens da relação</div><div class="n">${state.itens.length}</div><div class="caption">${brl(soma(state.itens, "valorTotal"))} em destinos</div></div>
      <div class="mix-card c"><div class="caption">Liberações de empenho</div><div class="n">${state.liberacoes.length}</div><div class="caption">${brl(soma(state.liberacoes))} com valor informado</div></div>
    </section>
    <section class="split">
      <div class="panel"><h2>Por código de fonte</h2>${barras(fonteLista, "valor")}</div>
      <div class="panel"><h2>Compras por setor</h2>${barras(setorLista, "n")}</div>
    </section>
    <section class="panel">
      <div class="panel-head"><h2>Aguardando entrega</h2><span class="caption">Os seis lançamentos mais recentes</span></div>
      ${tabela(
        [{ label: "Data" }, { label: "Empenho" }, { label: "Fornecedor" }, { label: "Valor", cls: "num" }],
        linhasEspera ? [linhasEspera] : []
      )}
    </section>`;
}

function empenhosFiltrados() {
  const f = ui.filtros.empenhos;
  return state.empenhos.filter((e) => {
    if (f.status && e.status !== f.status) return false;
    if (f.ano && anoEmpenho(e) !== f.ano) return false;
    if (f.fonte && fonteCurta(e.observacoes) !== f.fonte) return false;
    if (f.texto) {
      const blob = [e.numero, e.fornecedor, e.cnpj, e.objeto, e.observacoes, e.unidade, ...sisesEmpenho(e)].join(" ");
      if (!tem(blob, f.texto)) return false;
    }
    return true;
  }).sort((a, b) => anoEmpenho(b).localeCompare(anoEmpenho(a)) || String(b.data || "").localeCompare(String(a.data || "")) || String(b.numero).localeCompare(String(a.numero)));
}

function formEmpenho() {
  const atual = ui.editId ? empenhoPorId(ui.editId) : null;
  const e = atual || { unidade: "Secretaria Municipal de Saúde / Atenção Primária em Saúde", status: "Aguardando Entrega", ano: "2026" };
  const anos = unique(["2024", "2025", "2026", "2027", anoEmpenho(e), ...state.empenhos.map(anoEmpenho)].filter(Boolean)).sort();
  return `<form class="editor" data-form="empenho">
    <h3>${atual ? "Editar empenho" : "Novo empenho"}</h3>
    <div class="form-grid">
      ${campo("Número", "numero", e.numero)}
      ${campo("Ano", "ano", anoEmpenho(e) || "2026", { options: anos })}
      ${campo("Documento", "documento", e.documento, { wide: true })}
      ${campo("Data", "data", e.data, { type: "date" })}
      ${campo("Status", "status", e.status || "", { options: opcoesStatus(e.status || "") })}
      ${campo("Fornecedor", "fornecedor", e.fornecedor, { half: true })}
      ${campo("CNPJ", "cnpj", e.cnpj)}
      ${campo("Valor (R$)", "valor", e.valor ?? "", { type: "number", step: "0.01" })}
      ${campo("Unidade demandante", "unidade", e.unidade, { half: true })}
      ${campo("Objeto", "objeto", e.objeto, { area: true, wide: true })}
      ${campo("Observações, fonte, SIS, emenda", "observacoes", e.observacoes, { area: true, wide: true })}
    </div>
    <div class="form-actions"><button class="btn btn-primary" type="submit">Salvar</button><button class="btn" type="button" data-act="fechar-form">Cancelar</button></div>
  </form>`;
}

function detalheEmpenho(e) {
  const sises = sisesEmpenho(e);
  const procs = sises.map(procPorSis).filter(Boolean);
  const emenda = emendaDoEmpenho(e);
  const itens = itensDoNumero(e.numero);
  const chips = [
    ...procs.map((p) => chip("processo", p.id, `SIS ${p.sis}`)),
    emenda ? chip("emenda", emenda.id, `Emenda ${emenda.numero}`) : "",
    itens.length ? chip("grupo", normEmp(e.numero), `${itens.length} itens de destino`) : "",
  ].join("");
  return `<article class="detail">
    <div class="panel-head"><h2>${esc(e.numero)}</h2><div class="top-actions">${urlDoc(e) ? `<a class="btn btn-small" href="${esc(urlDoc(e))}" target="_blank" rel="noopener noreferrer">Abrir no Drive</a>` : ""}${pill(e.status)}<button class="btn btn-small" type="button" data-act="edit" data-tipo="empenho" data-id="${esc(e.id)}">Editar</button><button class="btn btn-small btn-danger" type="button" data-act="del" data-tipo="empenho" data-id="${esc(e.id)}">Excluir</button><button class="btn btn-small" type="button" data-act="fechar">Fechar</button></div></div>
    <p class="lede">${esc(e.fornecedor || "Fornecedor não informado")}${e.cnpj ? " · " + esc(e.cnpj) : ""}</p>
    <div class="kv">
      <div class="k">Ano</div><div>${esc(anoEmpenho(e) || "—")}</div>
      <div class="k">Data</div><div>${esc(mostrarData(e.data, e.dataTexto))}</div>
      <div class="k">Valor</div><div>${brl(e.valor)}</div>
      <div class="k">Fonte</div><div>${esc(fonteDe(e.observacoes) || "—")}</div>
      <div class="k">Unidade</div><div>${esc(e.unidade || "—")}</div>
      <div class="k">Objeto</div><div>${prose(e.objeto)}</div>
      <div class="k">Observações</div><div>${prose(e.observacoes)}</div>
    </div>
    ${chips ? `<div class="chips">${chips}</div>` : ""}
    ${htmlAnexos(e)}
    ${ui.form === "empenho" && ui.editId === e.id ? formEmpenho() : ""}
  </article>`;
}

function htmlAnexos(e) {
  const lista = anexosDo(e.id);
  const linha = (a) => `<div class="anexo-row">
    <div><strong>${esc(rotuloAnexo(a.tipo))}</strong> · ${esc(a.nome)}<div class="caption">${esc(tamanhoArquivo(a.tamanho))} · ${esc(dataBR(a.incluidoEm) || "hoje")}</div></div>
    <div class="acoes">
      <button class="btn btn-small" type="button" data-act="abrir-anexo" data-id="${esc(a.id)}">Abrir</button>
      <button class="btn btn-small btn-danger" type="button" data-act="del-anexo" data-id="${esc(a.id)}">Remover</button>
    </div>
  </div>`;
  return `<section class="anexos">
    <div class="panel-head"><h3>Arquivos deste empenho</h3><div class="top-actions">
      <label class="btn btn-small">Documento do empenho<input type="file" hidden data-anexo="empenho" data-id="${esc(e.id)}" accept=".pdf,application/pdf,image/*"></label>
      <label class="btn btn-small">Nota fiscal<input type="file" hidden data-anexo="nota" data-id="${esc(e.id)}" accept=".pdf,application/pdf,image/*"></label>
    </div></div>
    ${lista.length ? lista.map(linha).join("") : `<p class="caption">Nenhum arquivo neste empenho. O link do Drive, quando existe, continua no número. Aqui fica a cópia guardada neste navegador.</p>`}
  </section>`;
}

function viewEmpenhos() {
  const f = ui.filtros.empenhos;
  const lista = empenhosFiltrados();
  const fontes = unique(state.empenhos.map((e) => fonteCurta(e.observacoes)));
  const statuses = unique(state.empenhos.map((e) => e.status).filter(Boolean));
  const anos = unique(state.empenhos.map(anoEmpenho)).filter(Boolean).sort((a, b) => b.localeCompare(a));
  const opts = (arr, atual, todos) => [`<option value="">${todos}</option>`, ...arr.map((o) => `<option value="${esc(o)}"${o === atual ? " selected" : ""}>${esc(o)}</option>`)].join("");
  let extra = "";
  if (ui.form === "empenho" && !ui.editId) extra = formEmpenho();
  if (ui.aberto?.tipo === "empenho") {
    const e = empenhoPorId(ui.aberto.id);
    if (e) extra += detalheEmpenho(e);
  }
  if (ui.aberto?.tipo === "grupo") extra += detalheGrupo(ui.aberto.id);

  const tab = ui.tab === "itens" || ui.tab === "arquivos" ? ui.tab : "controle";
  let corpo = "";
  if (tab === "controle") {
    const rows = [];
    let anoAtual = "";
    lista.forEach((e) => {
      const ano = anoEmpenho(e) || "Sem ano";
      if (ano !== anoAtual) {
        anoAtual = ano;
        rows.push(`<tr class="ano-grupo"><td colspan="6">${esc(ano)}</td></tr>`);
      }
      const sel = ui.aberto?.tipo === "empenho" && ui.aberto.id === e.id ? "sel" : "";
      const statusOpts = opcoesStatus(e.status || "").map((s) => `<option value="${esc(s.value)}"${s.value === (e.status || "") ? " selected" : ""}>${esc(s.label)}</option>`).join("");
      rows.push(`<tr class="${sel}" data-tipo="empenho" data-id="${esc(e.id)}">
        ${td(`${esc(mostrarData(e.data, e.dataTexto))}<div class="caption">${esc(ano)}</div>`, "nowrap")}
        ${td(`${linkDoc(e)}${marcaAnexos(e)}`, "nowrap")}
        ${td(esc(clip(e.fornecedor, 36)))}
        ${td(esc(clip(e.objeto, 72)))}
        ${td(brl(e.valor), "num")}
        ${td(`<select class="select status-select" data-act="status" data-id="${esc(e.id)}">${statusOpts}</select>`)}
      </tr>`);
    });
    corpo = `
      <div class="filters">
        <input id="filtro-emp-texto" class="input" data-live="1" data-grupo="empenhos" data-filtro="texto" placeholder="Número, fornecedor, objeto, SIS" value="${esc(f.texto)}">
        <select class="select" data-filtro="ano" data-grupo="empenhos">${opts(anos, f.ano, "Todos os anos")}</select>
        <select class="select" data-filtro="status" data-grupo="empenhos">${opts(statuses, f.status, "Todos os status")}</select>
        <select class="select" data-filtro="fonte" data-grupo="empenhos">${opts(fontes, f.fonte, "Todas as fontes")}</select>
        <div class="count">${lista.length} de ${state.empenhos.length}</div>
      </div>
      ${tabela(
        [{ label: "Data" }, { label: "Empenho" }, { label: "Fornecedor" }, { label: "Objeto" }, { label: "Valor", cls: "num" }, { label: "Status" }],
        rows
      )}`;
  } else if (tab === "itens") {
    corpo = viewItens();
  } else {
    corpo = viewArquivos();
  }

  return `
    ${topo("Empenhos", "Cada empenho fica no ano da data em que foi criado. Selecione o ano para ver só esse grupo. Os antigos ficam na pasta Notas de empenho.", `<button class="btn btn-primary" type="button" data-act="novo" data-tipo="empenho">Novo empenho</button><button class="btn" type="button" data-act="csv" data-arg="empenhos">Exportar CSV</button>`)}
    ${htmlAtalho("documentosRelacao", "Abrir pasta Notas de empenho", "Cole o link da pasta Notas de empenho")}
    <div class="tabs">
      <button type="button" class="${tab === "controle" ? "on" : ""}" data-act="tab" data-arg="controle">Por ano</button>
      <button type="button" class="${tab === "itens" ? "on" : ""}" data-act="tab" data-arg="itens">Itens e destinos</button>
      <button type="button" class="${tab === "arquivos" ? "on" : ""}" data-act="tab" data-arg="arquivos">Arquivos</button>
    </div>
    ${corpo}
    ${extra}`;
}

function viewArquivos() {
  const f = ui.filtros.arquivos;
  const lista = anexos.filter((a) => {
    if (f.tipo && a.tipo !== f.tipo) return false;
    if (!f.texto) return true;
    const emp = empenhoPorId(a.empenhoId);
    return tem([a.nome, a.numero, emp && emp.numero, emp && emp.fornecedor].join(" "), f.texto);
  });
  const rows = lista.map((a) => {
    const emp = empenhoPorId(a.empenhoId);
    const numero = (emp && emp.numero) || a.numero || "Empenho removido";
    const abre = emp ? ` data-tipo="empenho" data-id="${esc(emp.id)}"` : "";
    return `<tr${abre}>
      ${td(esc(numero), "nowrap")}
      ${td(esc(rotuloAnexo(a.tipo)))}
      ${td(esc(a.nome))}
      ${td(esc(tamanhoArquivo(a.tamanho)), "nowrap")}
      ${td(esc(dataBR(a.incluidoEm) || "—"), "nowrap")}
      ${td(`<button class="btn btn-small" type="button" data-act="abrir-anexo" data-id="${esc(a.id)}">Abrir</button> <button class="btn btn-small btn-danger" type="button" data-act="del-anexo" data-id="${esc(a.id)}">Remover</button>`)}
    </tr>`;
  });
  return `
    <div class="filters">
      <input id="filtro-arq-texto" class="input" data-live="1" data-grupo="arquivos" data-filtro="texto" placeholder="Empenho, fornecedor ou nome do arquivo" value="${esc(f.texto)}">
      <select class="select" data-filtro="tipo" data-grupo="arquivos">
        <option value="">Todos os tipos</option>
        <option value="empenho"${f.tipo === "empenho" ? " selected" : ""}>Documento do empenho</option>
        <option value="nota"${f.tipo === "nota" ? " selected" : ""}>Nota fiscal</option>
      </select>
      <div class="count">${lista.length} arquivo(s)</div>
    </div>
    ${lista.length ? tabela(
      [{ label: "Empenho" }, { label: "Tipo" }, { label: "Arquivo" }, { label: "Tamanho" }, { label: "Incluído" }, { label: "" }],
      rows
    ) : `<p class="caption">Nenhum arquivo registrado. Abra um empenho e use Documento do empenho ou Nota fiscal.</p>`}`;
}

function viewItens() {
  const f = ui.filtros.itens;
  const tipos = unique(state.itens.map((i) => tipoItem(i.tipo)));
  const anos = unique(state.itens.map((i) => anoEmpenho({ numero: i.numeroEmpenho, data: i.data }))).filter(Boolean).sort((a, b) => b.localeCompare(a));
  const grupos = gruposItens().filter((g) => {
    const tipo = tipoItem(g[0].tipo);
    if (f.tipo && tipo !== f.tipo) return false;
    if (f.ano && anoEmpenho({ numero: g[0].numeroEmpenho, data: g[0].data }) !== f.ano) return false;
    if (!f.texto) return true;
    const blob = g.map((i) => [i.numeroEmpenho, i.fornecedor, i.descricao, i.destino, i.sis, i.emenda, i.autor, i.observacoes].join(" ")).join(" ");
    return tem(blob, f.texto);
  });
  const opts = [`<option value="">Todos os tipos</option>`, ...tipos.map((o) => `<option value="${esc(o)}"${o === f.tipo ? " selected" : ""}>${esc(o)}</option>`)].join("");
  const rows = grupos.map((g) => {
    const id = normEmp(g[0].numeroEmpenho);
    const sel = ui.aberto?.tipo === "grupo" && ui.aberto.id === id ? "sel" : "";
    const destinos = unique(g.map((i) => String(i.destino || "").replace(/\s+/g, " ").trim()));
    return `<tr class="${sel}" data-tipo="grupo" data-id="${esc(id)}">
      ${td(esc(g[0].numeroEmpenho), "nowrap")}
      ${td(esc(mostrarData(g[0].data, g[0].dataTexto)), "nowrap")}
      ${td(esc(tipoItem(g[0].tipo)))}
      ${td(esc(clip(g[0].fornecedor, 32)))}
      ${td(String(g.length), "num")}
      ${td(brl(soma(g, "valorTotal")), "num")}
      ${td(esc(clip(destinos.join(", "), 48)))}
    </tr>`;
  });
  return `
    <div class="filters">
      <input id="filtro-item-texto" class="input" data-live="1" data-grupo="itens" data-filtro="texto" placeholder="Empenho, destino, emenda, fornecedor" value="${esc(f.texto)}">
      <select class="select" data-filtro="tipo" data-grupo="itens">${opts}</select>
      <select class="select" data-filtro="ano" data-grupo="itens">${[`<option value="">Todos os anos</option>`, ...anos.map((o) => `<option value="${esc(o)}"${o === f.ano ? " selected" : ""}>${esc(o)}</option>`)].join("")}</select>
      <div class="count">${grupos.length} empenhos · ${state.itens.length} itens</div>
    </div>
    ${tabela(
      [{ label: "Empenho" }, { label: "Data" }, { label: "Tipo" }, { label: "Fornecedor" }, { label: "Itens", cls: "num" }, { label: "Valor", cls: "num" }, { label: "Destinos" }],
      rows
    )}`;
}

function formItem(base) {
  const item = ui.editId ? state.itens.find((i) => i.id === ui.editId) : {
    numeroEmpenho: base.numeroEmpenho,
    sis: base.sis,
    emenda: base.emenda,
    tipo: base.tipo,
    fornecedor: base.fornecedor,
    autor: base.autor,
    solicitante: base.solicitante,
  };
  if (!item) return "";
  return `<form class="editor" data-form="item">
    <h3>${ui.editId ? "Editar item" : "Incluir destino"}</h3>
    <div class="form-grid">
      ${campo("Descrição", "descricao", item.descricao, { area: true, wide: true })}
      ${campo("Quantidade", "quantidade", item.quantidade ?? "", { type: "number", step: "0.01" })}
      ${campo("Unidade", "unidade", item.unidade)}
      ${campo("Valor total (R$)", "valorTotal", item.valorTotal ?? "", { type: "number", step: "0.01" })}
      ${campo("Destino", "destino", item.destino, { half: true })}
      ${campo("Previsão de entrega", "previsaoEntrega", item.previsaoEntrega, { type: "date" })}
      ${campo("SIS", "sis", item.sis)}
      ${campo("Emenda", "emenda", item.emenda)}
      ${campo("Observações", "observacoes", item.observacoes, { area: true, wide: true })}
    </div>
    <div class="form-actions"><button class="btn btn-primary" type="submit">Salvar item</button><button class="btn" type="button" data-act="fechar-form">Cancelar</button></div>
  </form>`;
}

function detalheGrupo(id) {
  const itens = state.itens.filter((i) => normEmp(i.numeroEmpenho) === id);
  if (!itens.length) return "";
  const base = itens[0];
  const emp = state.empenhos.find((e) => normEmp(e.numero) === id);
  const proc = procPorSis(base.sis);
  const emenda = emendaPorNumero(base.emenda);
  const chips = [
    emp ? chip("empenho", emp.id, `Controle ${emp.numero}`) : "",
    proc ? chip("processo", proc.id, `SIS ${proc.sis}`) : "",
    emenda ? chip("emenda", emenda.id, `Emenda ${emenda.numero}`) : "",
  ].join("");
  const rows = itens.map((i) => `<tr>
    ${td(esc(clip(i.descricao, 80)))}
    ${td(qtdFmt(i.quantidade), "num")}
    ${td(esc(i.unidade || "—"))}
    ${td(brl(i.valorTotal), "num")}
    ${td(esc(clip(i.destino, 40)))}
    ${td(esc(mostrarData(i.previsaoEntrega, i.previsaoTexto)))}
    ${td(`<button class="btn btn-small" type="button" data-act="edit" data-tipo="item" data-id="${esc(i.id)}">Editar</button> <button class="btn btn-small btn-danger" type="button" data-act="del" data-tipo="item" data-id="${esc(i.id)}">Excluir</button>`)}
  </tr>`).join("");
  return `<article class="detail">
    <div class="panel-head"><h2>${esc(base.numeroEmpenho)}</h2><div class="top-actions"><button class="btn btn-small" type="button" data-act="novo" data-tipo="item">Incluir destino</button><button class="btn btn-small" type="button" data-act="fechar">Fechar</button></div></div>
    <p class="lede">${esc(tipoItem(base.tipo))}${base.autor ? " · " + esc(base.autor) : ""}${base.fornecedor ? " · " + esc(base.fornecedor) : ""}</p>
    ${chips ? `<div class="chips">${chips}</div>` : ""}
    ${emp ? htmlAnexos(emp) : ""}
    ${tabela(
      [{ label: "Descrição" }, { label: "Qtd", cls: "num" }, { label: "Un." }, { label: "Valor", cls: "num" }, { label: "Destino" }, { label: "Previsão" }, { label: "" }],
      [rows]
    )}
    ${ui.form === "item" ? formItem(base) : ""}
  </article>`;
}

function viewLiberacoes() {
  const lista = [...state.liberacoes].sort((a, b) => String(b.data || "").localeCompare(String(a.data || "")));
  const rows = lista.map((l) => {
    const sel = ui.aberto?.tipo === "liberacao" && ui.aberto.id === l.id ? "sel" : "";
    return `<tr class="${sel}" data-tipo="liberacao" data-id="${esc(l.id)}">
      ${td(esc(l.sis || "—"), "nowrap")}
      ${td(esc(l.numero), "nowrap")}
      ${td(brl(l.valor), "num")}
      ${td(esc(mostrarData(l.data, l.dataTexto)), "nowrap")}
      ${td(esc(clip(l.observacao, 60)))}
    </tr>`;
  });
  return tabela(
    [{ label: "SIS" }, { label: "Empenho" }, { label: "Valor", cls: "num" }, { label: "Liberação" }, { label: "Observação" }],
    rows
  );
}

function detalheLiberacao(id) {
  const l = state.liberacoes.find((x) => x.id === id);
  if (!l) return "";
  const proc = procPorSis(l.sis);
  return `<article class="detail">
    <div class="panel-head"><h2>${esc(l.numero)}</h2><button class="btn btn-small" type="button" data-act="fechar">Fechar</button></div>
    <div class="kv">
      <div class="k">SIS</div><div>${esc(l.sis || "—")}</div>
      <div class="k">Valor</div><div>${brl(l.valor)}</div>
      <div class="k">Liberação</div><div>${esc(mostrarData(l.data, l.dataTexto))}</div>
      <div class="k">Observação</div><div>${prose(l.observacao)}</div>
    </div>
    ${proc ? `<div class="chips">${chip("processo", proc.id, `Abrir processo ${proc.sis}`)}</div>` : ""}
  </article>`;
}

function emendasFiltradas() {
  const f = ui.filtros.emendas;
  return state.emendas.filter((e) => {
    if (f.ano && e.ano !== f.ano) return false;
    if (f.tipo === "execucao" && !e.statusExecucao) return false;
    if (f.tipo === "impositiva" && !(e.ementa || e.objeto)) return false;
    if (f.texto) {
      const blob = [e.numero, e.vereador, e.partido, e.destino, e.ubs, e.objeto, e.ementa, e.statusExecucao, e.observacao, e.sis, e.materiaisDefinidos].join(" ");
      if (!tem(blob, f.texto)) return false;
    }
    return true;
  }).sort((a, b) => {
    const [na, aa] = a.numero.split("/").map(Number);
    const [nb, ab] = b.numero.split("/").map(Number);
    return ab - aa || nb - na;
  });
}

function formEmenda() {
  const atual = ui.editId ? emendaPorId(ui.editId) : null;
  const e = atual || {};
  return `<form class="editor" data-form="emenda">
    <h3>${atual ? "Atualizar emenda" : "Nova emenda"}</h3>
    <div class="form-grid">
      ${campo("Número (ex.: 033/2025)", "numero", e.numero)}
      ${campo("Documento", "documento", e.documento, { wide: true })}
      ${campo("Vereador", "vereador", e.vereador)}
      ${campo("Partido", "partido", e.partido)}
      ${campo("Valor previsto (R$)", "valor", e.valor ?? "", { type: "number", step: "0.01" })}
      ${campo("Destino", "destino", e.destino, { half: true })}
      ${campo("Natureza", "natureza", e.natureza)}
      ${campo("Objeto", "objeto", e.objeto, { area: true, wide: true })}
      ${campo("Ementa", "ementa", e.ementa, { area: true, wide: true })}
      ${campo("Justificativa", "justificativa", e.justificativa, { area: true, wide: true })}
      ${campo("O que ficou definido", "definicao", e.definicao, { area: true, wide: true })}
      ${campo("Materiais", "materiaisDefinidos", e.materiaisDefinidos, { area: true, wide: true })}
      ${campo("UBS", "ubs", e.ubs, { area: true, wide: true })}
      ${campo("Andamento", "statusExecucao", e.statusExecucao, { area: true, wide: true })}
      ${campo("Observação", "observacao", e.observacao, { area: true, wide: true })}
      ${campo("Memorando", "memorando", e.memorando)}
      ${campo("SIS", "sis", e.sis)}
    </div>
    <div class="form-actions"><button class="btn btn-primary" type="submit">Salvar</button><button class="btn" type="button" data-act="fechar-form">Cancelar</button></div>
  </form>`;
}

function detalheEmenda(e) {
  const itens = state.itens.filter((i) => normEmenda(i.emenda) === e.numero);
  const emps = state.empenhos.filter((emp) => emendaMencionada(emp.observacoes) === e.numero);
  const procs = e.sis ? [procPorSis(e.sis)].filter(Boolean) : [];
  const materiais = materiaisDaEmenda(e.numero);
  const chips = [
    ...emps.map((emp) => chip("empenho", emp.id, emp.numero)),
    ...procs.map((p) => chip("processo", p.id, `SIS ${p.sis}`)),
    itens.length ? chip("grupo", normEmp(itens[0].numeroEmpenho), `${itens.length} itens na relação`) : "",
  ].join("");
  return `<article class="detail">
    <div class="panel-head"><h2>Emenda ${esc(e.numero)}</h2><div class="top-actions">${urlDoc(e) ? `<a class="btn btn-small" href="${esc(urlDoc(e))}" target="_blank" rel="noopener noreferrer">Abrir documento</a>` : ""}<a class="btn btn-small" href="${esc(state.atalhos.pastaEmendas || PASTA_MEMORANDOS)}" target="_blank" rel="noopener noreferrer">Memorandos</a><button class="btn btn-small" type="button" data-act="edit" data-tipo="emenda" data-id="${esc(e.id)}">Editar</button><button class="btn btn-small btn-danger" type="button" data-act="del" data-tipo="emenda" data-id="${esc(e.id)}">Excluir</button><button class="btn btn-small" type="button" data-act="fechar">Fechar</button></div></div>
    <p class="lede">${esc(e.vereador || "Autor não informado")}${e.partido ? " · " + esc(e.partido) : ""} · ${brl(e.valor)}</p>
    <div class="kv">
      <div class="k">Destino</div><div>${prose(e.destino || e.ubs)}</div>
      <div class="k">Natureza</div><div>${esc(e.natureza || "—")}${e.obra ? " · obra/ref.: " + esc(e.obra) : ""}</div>
      <div class="k">Órgão</div><div>${esc([e.smsTerceiro, e.orgao, e.secretaria].filter(Boolean).join(" · ") || "—")}</div>
      <div class="k">Objeto</div><div>${prose(e.objeto)}</div>
      <div class="k">Ementa</div><div>${prose(e.ementa)}</div>
      <div class="k">Justificativa</div><div>${prose(e.justificativa)}</div>
      <div class="k">O que ficou definido</div><div>${prose(e.definicao)}</div>
      <div class="k">Andamento</div><div>${prose(e.statusExecucao)}</div>
      <div class="k">Observação</div><div>${prose(e.observacao)}</div>
      <div class="k">Materiais</div><div>${prose(e.materiaisDefinidos)}</div>
      <div class="k">Memorando</div><div>${esc(e.memorando || "—")}</div>
    </div>
    ${chips ? `<div class="chips">${chips}</div>` : ""}
    ${materiais.length ? `<h3 class="bloco-titulo">Materiais ligados a esta emenda</h3>${materiais.map((l) => cartaoMaterial(l, { naEmenda: true })).join("")}` : ""}
    ${ui.form === "emenda" && ui.editId === e.id ? formEmenda() : ""}
  </article>`;
}

function viewEmendas() {
  const f = ui.filtros.emendas;
  const lista = emendasFiltradas();
  const anos = unique(state.emendas.map((e) => e.ano));
  const opt = (arr, atual, todos) => [`<option value="">${todos}</option>`, ...arr.map((o) => `<option value="${esc(o)}"${o === atual ? " selected" : ""}>${esc(o)}</option>`)].join("");
  const rows = lista.map((e) => {
    const sel = ui.aberto?.tipo === "emenda" && ui.aberto.id === e.id ? "sel" : "";
    return `<tr class="${sel}" data-tipo="emenda" data-id="${esc(e.id)}">
      ${td(linkDoc(e), "nowrap")}
      ${td(esc(clip(e.vereador, 28)))}
      ${td(brl(e.valor), "num")}
      ${td(esc(clip(e.destino || e.ubs || e.local, 42)))}
      ${td(esc(clip(e.statusExecucao || e.observacao || "Sem andamento", 80)))}
    </tr>`;
  });
  let extra = "";
  if (ui.form === "emenda" && !ui.editId) extra += formEmenda();
  if (ui.aberto?.tipo === "emenda") {
    const e = emendaPorId(ui.aberto.id);
    if (e) extra += detalheEmenda(e);
  }
  return `
    ${topo("Emendas", "A origem do recurso. Os memorandos ficam na pasta do Drive. Cada emenda também pode ter o link do próprio arquivo.", `<button class="btn btn-primary" type="button" data-act="novo" data-tipo="emenda">Nova emenda</button><button class="btn" type="button" data-act="csv" data-arg="emendas">Exportar CSV</button>`)}
    ${htmlAtalho("pastaEmendas", "Abrir memorandos", "Link da pasta de memorandos")}
    <div class="filters">
      <input id="filtro-emenda-texto" class="input" data-live="1" data-grupo="emendas" data-filtro="texto" placeholder="Número, vereador, destino, andamento" value="${esc(f.texto)}">
      <select class="select" data-filtro="ano" data-grupo="emendas">${opt(anos, f.ano, "Todos os anos")}</select>
      <select class="select" data-filtro="tipo" data-grupo="emendas">
        <option value=""${f.tipo === "" ? " selected" : ""}>Todas</option>
        <option value="execucao"${f.tipo === "execucao" ? " selected" : ""}>Com andamento</option>
        <option value="impositiva"${f.tipo === "impositiva" ? " selected" : ""}>Com objeto ou ementa</option>
      </select>
      <div class="count">${lista.length} de ${state.emendas.length} · ${brl(soma(lista))}</div>
    </div>
    ${tabela(
      [{ label: "Emenda" }, { label: "Autor" }, { label: "Valor previsto", cls: "num" }, { label: "Destino" }, { label: "Andamento" }],
      rows
    )}
    ${extra}`;
}

function comprasFiltradas() {
  const f = ui.filtros.compras;
  return state.processos.filter((p) => {
    if (f.setor && (p.setor || "Sem setor") !== f.setor) return false;
    if (f.ano && p.ano !== f.ano) return false;
    if (f.texto) {
      const blob = [p.sis, p.objeto, p.descricao, p.observacao, p.setor, p.setorOriginal].join(" ");
      if (!tem(blob, f.texto)) return false;
    }
    return true;
  }).sort((a, b) => String(b.data || "").localeCompare(String(a.data || "")));
}

function formProcesso() {
  const setores = [{ value: "", label: "Sem setor" }, ...opcoesSetor("").map((s) => ({ value: s, label: s }))];
  return `<form class="editor" data-form="processo">
    <h3>Novo processo</h3>
    <div class="form-grid">
      ${campo("SIS", "sis", "")}
      ${campo("Data", "data", hojeISO(), { type: "date" })}
      ${campo("Setor", "setor", "", { options: setores })}
      ${campo("Objeto", "objeto", "", { area: true, wide: true })}
      ${campo("Situação", "descricao", "", { area: true, wide: true })}
    </div>
    <div class="form-actions"><button class="btn btn-primary" type="submit">Salvar</button><button class="btn" type="button" data-act="fechar-form">Cancelar</button></div>
  </form>`;
}

function detalheProcesso(p) {
  const libs = state.liberacoes.filter((l) => l.sis && l.sis === p.sis);
  const emps = state.empenhos.filter((e) => sisesEmpenho(e).includes(p.sis));
  const itens = state.itens.filter((i) => normSis(i.sis) === p.sis);
  const chips = [
    ...emps.map((e) => chip("empenho", e.id, e.numero)),
    itens.length ? chip("grupo", normEmp(itens[0].numeroEmpenho), `${itens.length} itens com este SIS`) : "",
  ].join("");
  const anterior = p.observacao && p.descricao && p.observacao.trim() !== p.descricao.trim() ? p.observacao : "";
  return `<article class="detail">
    <div class="panel-head"><h2>${esc(p.sis || "Processo sem SIS")}</h2><div class="top-actions"><button class="btn btn-small btn-danger" type="button" data-act="del" data-tipo="processo" data-id="${esc(p.id)}">Excluir</button><button class="btn btn-small" type="button" data-act="fechar">Fechar</button></div></div>
    <p class="lede">${esc(p.objeto || "Sem objeto")} · aberto em ${esc(mostrarData(p.data, p.dataTexto))}</p>
    <form class="editor" data-form="andamento">
      <div class="form-grid">
        ${campo("Situação", "descricao", situacaoAtual(p), { area: true, wide: true })}
        ${campo("Setor", "setor", p.setor || "", { options: [{ value: "", label: "Sem setor" }, ...opcoesSetor(p.setor).map((s) => ({ value: s, label: s }))] })}
        ${campo("SIS", "sis", p.sis)}
        ${campo("Data de abertura", "data", p.data, { type: "date" })}
        ${campo("Objeto", "objeto", p.objeto, { area: true, wide: true })}
      </div>
      <div class="form-actions">
        <button class="btn btn-primary" type="submit">Salvar</button>
        <button class="btn" type="button" data-act="verificar" data-id="${esc(p.id)}">Verificado hoje</button>
      </div>
    </form>
    ${anterior ? `<p class="caption">Anotação anterior, só consulta</p>${prose(anterior)}` : ""}
    ${chips ? `<div class="chips">${chips}</div>` : ""}
    ${libs.length ? `<p class="caption">${libs.length} liberação(ões) neste SIS, somando ${brl(soma(libs))}.</p>` : ""}
  </article>`;
}

function viewCompras() {
  const f = ui.filtros.compras;
  const lista = comprasFiltradas();
  const setores = unique(state.processos.map((p) => p.setor || "Sem setor"));
  const anos = unique(state.processos.map((p) => p.ano));
  const opt = (arr, atual, todos) => [`<option value="">${todos}</option>`, ...arr.map((o) => `<option value="${esc(o)}"${o === atual ? " selected" : ""}>${esc(o)}</option>`)].join("");
  const rows = [];
  lista.forEach((p) => {
    const sel = ui.aberto?.tipo === "processo" && ui.aberto.id === p.id ? "sel" : "";
    const dataAntiga = !p.ultimaVerificacao && p.ultimaTexto ? `<div class="caption">${esc(p.ultimaTexto)}</div>` : "";
    rows.push(`<tr class="${sel}" data-tipo="processo" data-id="${esc(p.id)}">
      ${td(esc(p.sis || "—"), "nowrap")}
      ${td(esc(clip(p.objeto, 48)))}
      ${td(htmlSetor(p.setor, p.id))}
      ${td(`<input class="input date-inline" type="date" data-act="verificacao" data-id="${esc(p.id)}" value="${esc(p.ultimaVerificacao || "")}">${dataAntiga}`)}
      ${td(esc(clip(situacaoAtual(p), 90)))}
    </tr>`);
    if (sel) rows.push(`<tr class="detail-row"><td colspan="5">${detalheProcesso(p)}</td></tr>`);
  });
  const aberto = ui.aberto?.tipo === "processo" ? procPorId(ui.aberto.id) : null;
  const foraDaLista = aberto && !lista.some((p) => p.id === aberto.id) ? detalheProcesso(aberto) : "";
  let extra = "";
  if (ui.form === "processo" && !ui.editId) extra += formProcesso();
  return `
    ${topo("Compras", "Clique na linha para editar a situação, o setor, o SIS e o objeto. A edição abre logo abaixo dela.", `<button class="btn btn-primary" type="button" data-act="novo" data-tipo="processo">Novo processo</button><button class="btn" type="button" data-act="csv" data-arg="compras">Exportar CSV</button>`)}
    ${extra}
    <div class="filters">
      <input id="filtro-compra-texto" class="input" data-live="1" data-grupo="compras" data-filtro="texto" placeholder="SIS, objeto, descrição" value="${esc(f.texto)}">
      <select class="select" data-filtro="setor" data-grupo="compras">${opt(setores, f.setor, "Todos os setores")}</select>
      <select class="select" data-filtro="ano" data-grupo="compras">${opt(anos, f.ano, "Todos os anos")}</select>
      <div class="count">${lista.length} de ${state.processos.length}</div>
    </div>
    ${foraDaLista}
    ${tabela(
      [{ label: "SIS" }, { label: "Objeto" }, { label: "Setor" }, { label: "Verificação" }, { label: "Situação" }],
      rows
    )}`;
}

function cartaoMaterial(l, opts = {}) {
  const citadas = opts.naEmenda ? [] : emendasCitadas(l.emendas);
  const chipsEmenda = citadas.map((n) => {
    const em = emendaPorNumero(n);
    return em ? chip("emenda", em.id, `Emenda ${em.numero}${em.vereador ? " · " + em.vereador : ""}`) : `<span class="pill st-muted">${esc(n)}</span>`;
  }).join("");
  const procs = processosPorSisCurto(l.sis);
  const chipProc = procs.length === 1 ? chip("processo", procs[0].id, `Processo ${procs[0].sis}`) : "";
  const editando = ui.form === "lote" && ui.editId === l.id;
  return `<article class="mat-item${editando ? " editando" : ""}">
    ${opts.naEmenda && l.lote ? `<p class="caption">${esc(l.lote)}</p>` : ""}
    <div class="panel-head"><h3>${esc(l.item)}</h3><div class="top-actions"><span class="caption">Aquisição ${esc(l.ano || "2026")}</span>${anosDasEmendas(l.emendas).length ? `<span class="caption">Emenda ${esc(anosDasEmendas(l.emendas).join(", "))}</span>` : ""}${l.sis ? `<span class="caption">SIS ${esc(l.sis)}</span>` : ""}<button class="btn btn-small" type="button" data-act="edit" data-tipo="lote" data-id="${esc(l.id)}">Editar</button><button class="btn btn-small btn-danger" type="button" data-act="del" data-tipo="lote" data-id="${esc(l.id)}">Excluir</button></div></div>
    ${editando ? formLote({ embutido: true }) : `${l.descricao ? `<div class="prose mat-desc">${esc(l.descricao)}</div>` : ""}
    <div class="kv">
      <div class="k">Destino</div><div>${prose(l.destinos)}</div>
      <div class="k">Referência</div><div>${esc(l.emendas || "—")}</div>
    </div>
    ${chipsEmenda || chipProc ? `<div class="chips">${chipsEmenda}${chipProc}</div>` : ""}`}
  </article>`;
}

function formLote(opts = {}) {
  const l = (ui.editId && state.lotes.find((x) => x.id === ui.editId)) || {};
  const nomes = unique(state.lotes.map((x) => x.lote).filter(Boolean));
  const ano = l.ano || ui.filtros.materiais.ano || "2026";
  const anos = unique([ano, "2024", "2025", "2026", "2027", ...state.lotes.map((x) => x.ano)]).filter(anoPlausivel).sort();
  return `<form class="editor" data-form="lote" data-id="${esc(l.id || "")}">
    ${opts.embutido ? "" : `<h3>${l.id ? "Editar material" : "Novo material"}</h3>`}
    <p class="caption">Altere o destino quando a entrega for diferente do que a emenda indica. A alteração fica salva neste navegador.</p>
    <div class="form-grid">
      ${campo("Destino", "destinos", l.destinos, { area: true, wide: true })}
      ${campo("Ano da aquisição", "ano", ano, { options: anos })}
      <label class="field"><span>Lote</span><input class="input" name="lote" list="lotes-conhecidos" value="${esc(l.lote || "")}"></label>
      ${campo("Item", "item", l.item, { half: true })}
      ${campo("SIS", "sis", l.sis)}
      ${campo("Descrição", "descricao", l.descricao, { area: true, wide: true })}
      ${campo("Emenda ou referência", "emendas", l.emendas, { wide: true })}
    </div>
    <datalist id="lotes-conhecidos">${nomes.map((n) => `<option value="${esc(n)}"></option>`).join("")}</datalist>
    <div class="form-actions"><button class="btn btn-primary" type="submit">Salvar material</button><button class="btn" type="button" data-act="fechar-form">Cancelar</button></div>
  </form>`;
}

function viewMateriais() {
  const f = ui.filtros.materiais;
  const anos = unique(["2024", "2025", "2026", "2027", ...state.lotes.map((l) => l.ano)]).filter(anoPlausivel).sort((a, b) => b.localeCompare(a));
  const visiveis = state.lotes.filter((l) => {
    if (f.ano && (l.ano || "2026") !== f.ano) return false;
    if (f.texto && !tem([l.item, l.lote, l.descricao, l.destinos, l.emendas, l.sis, l.ano].join(" "), f.texto)) return false;
    return true;
  });
  const porAno = new Map();
  visiveis.forEach((l) => {
    const ano = l.ano || "2026";
    if (!porAno.has(ano)) porAno.set(ano, []);
    porAno.get(ano).push(l);
  });
  const blocos = [...porAno.keys()].sort((a, b) => b.localeCompare(a)).map((ano) => {
    const porLote = new Map();
    const legendas = [];
    porAno.get(ano).forEach((l) => {
      if (ehLegenda(l)) {
        const texto = String(l.item || "").trim().replace(/^Observações:\s*/i, "");
        if (texto && !/^Observações:?$/i.test(texto)) legendas.push(texto);
        return;
      }
      const k = l.lote || "Sem lote";
      if (!porLote.has(k)) porLote.set(k, []);
      porLote.get(k).push(l);
    });
    const lotesHtml = [...porLote.entries()].map(([nome, itens]) => `
      <section class="panel lote-bloco">
        <h2>${esc(nome)}</h2>
        ${itens.map((l) => cartaoMaterial(l)).join("")}
      </section>`).join("");
    const legendaHtml = legendas.length ? `<aside class="legendas"><h2>Legendas</h2><ul>${legendas.map((t) => `<li>${esc(t)}</li>`).join("")}</ul></aside>` : "";
    return `<h2 class="bloco-titulo">Materiais para aquisição em ${esc(ano)}</h2>${lotesHtml}${legendaHtml}`;
  }).join("");
  const optAno = [`<option value="">Todos os anos</option>`, ...anos.map((o) => `<option value="${esc(o)}"${o === f.ano ? " selected" : ""}>${esc(o)}</option>`)].join("");
  return `
    ${topo("Materiais", "O ano é o da aquisição. Os itens já lançados são para comprar em 2026, com recurso das emendas de 2025. Novo material entra no ano escolhido.", `<button class="btn btn-primary" type="button" data-act="novo" data-tipo="lote">Novo material</button>`)}
    <div class="filters">
      <input id="filtro-mat-texto" class="input" data-live="1" data-grupo="materiais" data-filtro="texto" placeholder="Item, lote, emenda, SIS" value="${esc(f.texto)}">
      <label class="field"><span>Ano da aquisição</span><select class="select" data-filtro="ano" data-grupo="materiais">${optAno}</select></label>
      <div class="count">${visiveis.filter((l) => !ehLegenda(l)).length} materiais</div>
    </div>
    ${ui.form === "lote" && !ui.editId ? formLote() : ""}
    ${blocos || `<p class="caption">Nenhum material neste ano. Use Novo material para incluir.</p>`}`;
}

function pecaPreenchida(r, chave) {
  return Boolean(linkDocRp(r[chave])) || anexosDoRegistro(r.id, chave).length > 0;
}
function htmlPecas(r) {
  return `<div class="pecas">${PECAS_RP.map((p) => {
    const url = linkDocRp(r[p.chave]);
    const arquivos = anexosDoRegistro(r.id, p.chave);
    const linhas = arquivos.map((a) => `<div class="anexo-row">
      <div>${esc(a.nome)}<div class="caption">${esc(tamanhoArquivo(a.tamanho))}</div></div>
      <div class="acoes">
        <button class="btn btn-small" type="button" data-act="abrir-anexo" data-id="${esc(a.id)}">Abrir</button>
        <button class="btn btn-small btn-danger" type="button" data-act="del-anexo" data-id="${esc(a.id)}">Remover</button>
      </div>
    </div>`).join("");
    return `<section class="peca">
      <h3>${esc(p.rotulo)}</h3>
      <p class="caption">${esc(p.nome)}</p>
      ${url ? `<a class="btn btn-small" href="${esc(url)}" target="_blank" rel="noopener noreferrer">${/^https?:/i.test(url) ? "Abrir link" : "Abrir arquivo"}</a>` : `<p class="caption">Sem documento. Cole um link https:// ou anexe o arquivo.</p>`}
      <label class="btn btn-small">Anexar arquivo<input type="file" hidden data-anexo="${esc(p.chave)}" data-id="${esc(r.id)}" accept=".pdf,application/pdf,.doc,.docx,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document,image/*"></label>
      ${linhas || `<p class="caption">Nenhum arquivo deste documento neste navegador.</p>`}
    </section>`;
  }).join("")}</div>`;
}
function formRegistro(r) {
  const atual = r || {};
  const ano = atual.ano || ui.filtros.registros.ano || "2026";
  const anos = unique([ano, "2024", "2025", "2026", "2027", ...state.registros.map((x) => x.ano)]).filter(anoPlausivel).sort((a, b) => b.localeCompare(a));
  const situacao = SITUACOES_RP.includes(atual.situacao) ? atual.situacao : "Em elaboração";
  return `<form class="editor" data-form="registro" data-id="${esc(atual.id || "")}">
    <h3>${atual.id ? "Grupo" : "Novo grupo"}</h3>
    <div class="form-grid">
      ${campo("Ano", "ano", ano, { options: anos })}
      ${campo("Nome do grupo", "nome", atual.nome, { half: true })}
      ${campo("Situação", "situacao", situacao, { options: SITUACOES_RP })}
      ${campo("Objeto", "objeto", atual.objeto, { area: true, wide: true })}
      ${campo("Link do DFD", "dfd", atual.dfd)}
      ${campo("Link do ETP", "etp", atual.etp)}
      ${campo("Link do TR", "tr", atual.tr)}
      ${campo("Observações", "observacoes", atual.observacoes, { area: true, wide: true })}
    </div>
    <div class="form-actions"><button class="btn btn-primary" type="submit">Salvar grupo</button>${atual.id ? "" : `<button class="btn" type="button" data-act="fechar-form">Cancelar</button>`}</div>
  </form>`;
}
function formMaterialRp(reg) {
  const salvo = (ui.editId && (reg.materiais || []).find((m) => m.id === ui.editId)) || {};
  const m = ui._mat || salvo;
  return `<form class="editor" data-form="material-rp" data-id="${esc(ui.editId || "")}">
    <h3>${ui.editId ? "Editar material do grupo" : "Material do grupo"}</h3>
    <div class="form-grid">
      ${campo("Item", "item", m.item, { half: true })}
      ${campo("Unidade", "unidade", m.unidade)}
      ${campo("Quantidade", "quantidade", m.quantidade ?? "")}
      ${campo("Descrição", "descricao", m.descricao, { area: true, wide: true })}
    </div>
    <div class="form-actions"><button class="btn btn-primary" type="submit">Salvar material</button><button class="btn" type="button" data-act="fechar-form">Cancelar</button></div>
  </form>`;
}
function detalheRegistro(r) {
  const materiais = r.materiais || [];
  const linhas = materiais.map((m) => `<tr>
    ${td(esc(m.item || "—"))}
    ${td(esc(m.unidade || "—"))}
    ${td(m.quantidade == null || m.quantidade === "" ? "—" : esc(qtdFmt(m.quantidade)), "num")}
    ${td(esc(clip(m.descricao, 80)))}
    ${td(`<button class="btn btn-small" type="button" data-act="edit" data-tipo="material-rp" data-id="${esc(m.id)}">Editar</button> <button class="btn btn-small btn-danger" type="button" data-act="del" data-tipo="material-rp" data-id="${esc(m.id)}">Excluir</button>`)}
  </tr>`);
  return `<article class="detail">
    <div class="panel-head"><h2>${esc(r.nome)}</h2><div class="top-actions"><button class="btn btn-small btn-danger" type="button" data-act="del" data-tipo="registro" data-id="${esc(r.id)}">Excluir grupo</button><button class="btn btn-small" type="button" data-act="fechar">Fechar</button></div></div>
    <p class="lede">Ano ${esc(r.ano)} · ${esc(r.situacao || "Em elaboração")}</p>
    ${formRegistro(r)}
    ${htmlPecas(r)}
    <div class="panel-head" style="margin-top:16px"><h3>Materiais do grupo</h3><button class="btn btn-small" type="button" data-act="novo" data-tipo="material-rp">Incluir material</button></div>
    ${ui.form === "material-rp" ? formMaterialRp(r) : ""}
    ${tabela(
      [{ label: "Item" }, { label: "Unidade" }, { label: "Qtd.", cls: "num" }, { label: "Descrição" }, { label: "" }],
      linhas,
      "Nenhum material neste grupo."
    )}
  </article>`;
}
function viewRegistros() {
  const f = ui.filtros.registros;
  const anos = unique(["2024", "2025", "2026", "2027", ...state.registros.map((r) => r.ano)]).filter(anoPlausivel).sort((a, b) => b.localeCompare(a));
  const lista = state.registros.filter((r) => {
    if (f.ano && r.ano !== f.ano) return false;
    if (f.texto) {
      const mats = (r.materiais || []).map((m) => [m.item, m.descricao, m.unidade].join(" ")).join(" ");
      if (!tem([r.nome, r.objeto, r.observacoes, r.situacao, r.ano, mats].join(" "), f.texto)) return false;
    }
    return true;
  }).sort((a, b) => String(b.ano).localeCompare(String(a.ano)) || String(a.nome).localeCompare(String(b.nome), "pt-BR"));
  const optAno = [`<option value="">Todos os anos</option>`, ...anos.map((o) => `<option value="${esc(o)}"${o === f.ano ? " selected" : ""}>${esc(o)}</option>`)].join("");
  const rows = [];
  lista.forEach((r) => {
    const sel = ui.aberto?.tipo === "registro" && ui.aberto.id === r.id ? "sel" : "";
    const docs = PECAS_RP.map((p) => `<span class="pill ${pecaPreenchida(r, p.chave) ? "st-ok" : "st-muted"}">${esc(p.rotulo)}</span>`).join(" ");
    rows.push(`<tr class="${sel}" data-tipo="registro" data-id="${esc(r.id)}">
      ${td(esc(r.ano), "nowrap")}
      ${td(esc(r.nome))}
      ${td(esc(r.situacao || "Em elaboração"))}
      ${td(String((r.materiais || []).length), "num")}
      ${td(docs)}
    </tr>`);
    if (sel) rows.push(`<tr class="detail-row"><td colspan="5">${detalheRegistro(r)}</td></tr>`);
  });
  const aberto = ui.aberto && ui.aberto.tipo === "registro" ? registroPorId(ui.aberto.id) : null;
  const foraDaLista = aberto && !lista.some((r) => r.id === aberto.id) ? detalheRegistro(aberto) : "";
  const extra = ui.form === "registro" && !ui.aberto ? formRegistro({}) : "";
  return `
    ${topo("Registros de preços", "Os grupos RP 02 a RP 13 vieram do sistema de documentos da Defesa Civil, com o DFD, o ETP e o TR em Word. Clique na linha para abrir o arquivo, trocar o link ou anexar outro. O ano separa os grupos.", `<button class="btn btn-primary" type="button" data-act="novo" data-tipo="registro">Novo grupo</button>`)}
    ${extra}
    <div class="filters">
      <input id="filtro-rp-texto" class="input" data-live="1" data-grupo="registros" data-filtro="texto" placeholder="Grupo, material ou objeto" value="${esc(f.texto)}">
      <label class="field"><span>Ano</span><select class="select" data-filtro="ano" data-grupo="registros">${optAno}</select></label>
      <div class="count">${lista.length} grupos</div>
    </div>
    ${foraDaLista}
    ${tabela(
      [{ label: "Ano" }, { label: "Grupo" }, { label: "Situação" }, { label: "Materiais", cls: "num" }, { label: "Documentos" }],
      rows,
      "Nenhum grupo neste ano. Use Novo grupo para incluir os materiais, o DFD, o ETP e o TR."
    )}`;
}

function viewBusca() {
  const q = ui.q.trim();
  if (q.length < 2) {
    return topo("Busca", "Digite ao menos duas letras na busca geral e pressione Enter.");
  }
  const emps = state.empenhos.filter((e) => tem([e.numero, e.fornecedor, e.cnpj, e.objeto, e.observacoes, ...sisesEmpenho(e)].join(" "), q)).slice(0, 30);
  const emendas = state.emendas.filter((e) => tem([e.numero, e.vereador, e.destino, e.ubs, e.objeto, e.ementa, e.statusExecucao].join(" "), q)).slice(0, 30);
  const procs = state.processos.filter((p) => tem([p.sis, p.objeto, p.descricao, p.observacao, p.setor].join(" "), q)).slice(0, 30);
  const itens = state.itens.filter((i) => tem([i.numeroEmpenho, i.descricao, i.destino, i.sis, i.emenda, i.fornecedor].join(" "), q)).slice(0, 30);
  const bloco = (titulo, rows, cols) => `<section class="panel" style="margin-bottom:14px"><h2>${titulo}</h2>${tabela(cols, rows)}</section>`;
  return `
    ${topo("Busca", `Resultados para “${esc(q)}”.`)}
    ${bloco("Empenhos", emps.map((e) => `<tr data-tipo="empenho" data-id="${esc(e.id)}">${td(linkDoc(e), "nowrap")}${td(esc(clip(e.fornecedor, 36)))}${td(brl(e.valor), "num")}${td(pill(e.status))}</tr>`), [{ label: "Empenho" }, { label: "Fornecedor" }, { label: "Valor", cls: "num" }, { label: "Status" }])}
    ${bloco("Emendas", emendas.map((e) => `<tr data-tipo="emenda" data-id="${esc(e.id)}">${td(linkDoc(e), "nowrap")}${td(esc(e.vereador || "—"))}${td(brl(e.valor), "num")}${td(esc(clip(e.destino || e.ubs, 40)))}</tr>`), [{ label: "Emenda" }, { label: "Autor" }, { label: "Valor", cls: "num" }, { label: "Destino" }])}
    ${bloco("Compras", procs.map((p) => `<tr data-tipo="processo" data-id="${esc(p.id)}">${td(esc(p.sis || "—"), "nowrap")}${td(esc(clip(p.objeto, 48)))}${td(esc(p.setor || "—"))}</tr>`), [{ label: "SIS" }, { label: "Objeto" }, { label: "Setor" }])}
    ${bloco("Itens", itens.map((i) => `<tr data-tipo="grupo" data-id="${esc(normEmp(i.numeroEmpenho))}">${td(esc(i.numeroEmpenho), "nowrap")}${td(esc(clip(i.descricao, 48)))}${td(esc(clip(i.destino, 32)))}${td(brl(i.valorTotal), "num")}</tr>`), [{ label: "Empenho" }, { label: "Descrição" }, { label: "Destino" }, { label: "Valor", cls: "num" }])}
    ${bloco("Registros de preços", state.registros.filter((r) => tem([r.nome, r.objeto, r.ano, ...(r.materiais || []).map((m) => m.item)].join(" "), q)).slice(0, 30).map((r) => `<tr data-tipo="registro" data-id="${esc(r.id)}">${td(esc(r.ano), "nowrap")}${td(esc(r.nome))}${td(esc(r.situacao || "—"))}${td(String((r.materiais || []).length), "num")}</tr>`), [{ label: "Ano" }, { label: "Grupo" }, { label: "Situação" }, { label: "Materiais", cls: "num" }])}
    ${bloco("Materiais", state.lotes.filter((l) => !ehLegenda(l) && tem([l.item, l.lote, l.descricao, l.destinos, l.emendas, l.sis].join(" "), q)).slice(0, 30).map((l) => {
      const n = emendasCitadas(l.emendas)[0];
      const em = n ? emendaPorNumero(n) : null;
      const abre = em ? ` data-tipo="emenda" data-id="${esc(em.id)}"` : "";
      return `<tr${abre}>${td(esc(l.item))}${td(esc(l.lote || "—"))}${td(esc(clip(l.destinos, 40)))}${td(esc(l.emendas || "—"))}</tr>`;
    }), [{ label: "Item" }, { label: "Lote" }, { label: "Destino" }, { label: "Emenda" }])}`;
}

function viewDados() {
  return `
    ${topo("Cópia dos dados", "O painel guarda as alterações neste navegador. A cópia em JSON leva os dados e também os documentos e as notas fiscais anexados.")}
    <div class="notice">Para abrir de novo, use o arquivo index.html desta pasta. Não precisa instalar nada.</div>
    <section class="mix">
      <div class="mix-card a"><div class="caption">Empenhos</div><div class="n">${state.empenhos.length}</div><div class="caption">Controle de Empenhos 2026</div></div>
      <div class="mix-card a"><div class="caption">Itens</div><div class="n">${state.itens.length}</div><div class="caption">Relação de empenhos</div></div>
      <div class="mix-card b"><div class="caption">Emendas</div><div class="n">${state.emendas.length}</div><div class="caption">Relação 2024 e impositivas</div></div>
      <div class="mix-card b"><div class="caption">Processos</div><div class="n">${state.processos.length}</div><div class="caption">Acompanhamento de compras</div></div>
      <div class="mix-card c"><div class="caption">Liberações</div><div class="n">${state.liberacoes.length}</div><div class="caption">Aba Empenhos do acompanhamento</div></div>
      <div class="mix-card c"><div class="caption">Materiais</div><div class="n">${state.lotes.filter((l) => !ehLegenda(l)).length}</div><div class="caption">Itens ligados às emendas</div></div>
      <div class="mix-card b"><div class="caption">Registros de preços</div><div class="n">${state.registros.length}</div><div class="caption">Grupos com DFD, ETP e TR</div></div>
      <div class="mix-card a"><div class="caption">Arquivos</div><div class="n">${anexos.length}</div><div class="caption">${anexos.filter((a) => a.tipo === "nota").length} notas fiscais</div></div>
    </section>
    <div class="top-actions" style="margin-bottom:16px">
      <button class="btn btn-primary" type="button" data-act="exportar">Baixar cópia JSON</button>
      <label class="btn">Restaurar de um arquivo<input id="arquivo-copia" type="file" accept="application/json,.json" hidden></label>
      <button class="btn btn-danger" type="button" data-act="reset">Recarregar as planilhas de origem</button>
    </div>
    <p class="caption">Recarregar substitui o que foi alterado nos registros pela leitura original das quatro planilhas. Os arquivos anexados permanecem. Baixe uma cópia antes, se quiser guardar o que mudou.</p>`;
}

function viewHtml() {
  switch (ui.view) {
    case "empenhos": return viewEmpenhos();
    case "emendas": return viewEmendas();
    case "compras": return viewCompras();
    case "materiais": return viewMateriais();
    case "registros": return viewRegistros();
    case "busca": return viewBusca();
    case "dados": return viewDados();
    default: return viewPainel();
  }
}

function preservarSituacao() {
  const form = document.querySelector("#view form[data-form='andamento']");
  if (!form || ui.aberto?.tipo !== "processo") return;
  const p = procPorId(ui.aberto.id);
  if (!p) return;
  const ler = (nome) => {
    const el = form.querySelector(`[name='${nome}']`);
    return el ? el.value : null;
  };
  if (ler("descricao") != null) p.descricao = ler("descricao");
  if (ler("sis") != null) p.sis = ler("sis");
  if (ler("data") != null) p.data = ler("data");
  if (ler("objeto") != null) p.objeto = ler("objeto");
}
function preservarRegistro() {
  const form = document.querySelector("#view form[data-form='registro']");
  if (!form || !form.dataset.id) return;
  const r = registroPorId(form.dataset.id);
  if (!r) return;
  ["ano", "nome", "objeto", "situacao", "dfd", "etp", "tr", "observacoes"].forEach((nome) => {
    const el = form.querySelector(`[name='${nome}']`);
    if (el) r[nome] = el.value;
  });
}
function valorCampo(form, nome) {
  const el = form.querySelector(`[name='${nome}']`);
  return el ? el.value : "";
}
function preservarMaterialRp() {
  const form = document.querySelector("#view form[data-form='material-rp']");
  if (!form || (form.dataset.id || "") !== (ui.editId || "")) return;
  ui._mat = {
    item: valorCampo(form, "item"),
    unidade: valorCampo(form, "unidade"),
    quantidade: valorCampo(form, "quantidade"),
    descricao: valorCampo(form, "descricao"),
  };
}
function render() {
  preservarSituacao();
  preservarRegistro();
  preservarMaterialRp();
  const active = document.activeElement;
  const activeId = active && active.id && active.closest("#view") ? active.id : "";
  const pos = activeId && active.selectionStart != null ? active.selectionStart : null;
  document.getElementById("view").innerHTML = viewHtml();
  document.querySelectorAll("#nav a").forEach((a) => a.classList.toggle("active", a.dataset.view === ui.view));
  if (activeId) {
    const next = document.getElementById(activeId);
    if (next) {
      next.focus();
      if (pos != null && next.setSelectionRange) {
        try { next.setSelectionRange(pos, pos); } catch { /* tipos date/number */ }
      }
    }
  }
  const marca = ui.aberto ? `${ui.view}:${ui.aberto.tipo}:${ui.aberto.id}` : "";
  if (marca && marca !== render._marca) {
    const linha = document.querySelector("#view tr.detail-row, #view article.detail");
    if (linha) linha.scrollIntoView({ block: "nearest" });
  }
  render._marca = marca;
  const editorLote = document.querySelector("#view form[data-form='lote'][data-id]");
  const marcaLote = editorLote && editorLote.dataset.id ? editorLote.dataset.id : "";
  if (marcaLote && marcaLote !== render._lote) editorLote.scrollIntoView({ block: "nearest" });
  render._lote = marcaLote;
}

function abrir(tipo, id) {
  if (tipo === "liberacao") {
    const l = state.liberacoes.find((x) => x.id === id);
    const p = l && procPorSis(l.sis);
    if (p) {
      abrir("processo", p.id);
      return;
    }
  }
  const destino = {
    empenho: "empenhos",
    grupo: "empenhos",
    emenda: "emendas",
    processo: "compras",
    registro: "registros",
  }[tipo] || ui.view;
  if (tipo === "grupo") ui.tab = "itens";
  if (tipo === "empenho" && ui.tab !== "arquivos") ui.tab = "controle";
  ui.aberto = { tipo, id };
  ui.form = null;
  ui.editId = null;
  ui._mat = null;
  ui.view = destino;
  const hash = `#${destino}`;
  if (location.hash !== hash) {
    ui._keep = true;
    location.hash = hash;
  } else {
    render();
  }
}

function fechar() {
  ui.aberto = null;
  ui.form = null;
  ui.editId = null;
  ui._mat = null;
  render();
}

function csvCelula(v) {
  return `"${String(v ?? "").replace(/"/g, '""')}"`;
}
function baixar(nome, texto, tipo) {
  const blob = new Blob([texto], { type: tipo });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = nome;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}
function exportarCsv(qual) {
  let linhas = [];
  if (qual === "empenhos") {
    linhas = [["Data", "Empenho", "Fornecedor", "CNPJ", "Objeto", "Valor", "Status", "Fonte", "Observações"]];
    empenhosFiltrados().forEach((e) => linhas.push([dataBR(e.data) || e.dataTexto, e.numero, e.fornecedor, e.cnpj, e.objeto, e.valor, e.status, fonteDe(e.observacoes), e.observacoes]));
  } else if (qual === "emendas") {
    linhas = [["Emenda", "Ano", "Vereador", "Partido", "Valor", "Destino", "Andamento", "Objeto"]];
    emendasFiltradas().forEach((e) => linhas.push([e.numero, e.ano, e.vereador, e.partido, e.valor, e.destino || e.ubs, e.statusExecucao, e.objeto]));
  } else if (qual === "compras") {
    linhas = [["Data", "SIS", "Objeto", "Setor", "Última verificação", "Descrição", "Observação"]];
    comprasFiltradas().forEach((p) => linhas.push([dataBR(p.data) || p.dataTexto, p.sis, p.objeto, p.setor, dataBR(p.ultimaVerificacao) || p.ultimaTexto, situacaoAtual(p), p.observacao]));
  }
  const csv = `\uFEFF${linhas.map((r) => r.map(csvCelula).join(";")).join("\r\n")}`;
  baixar(`painel-${qual}.csv`, csv, "text/csv;charset=utf-8");
}

function salvarEmpenho(fd) {
  const numero = fd.get("numero").trim();
  if (!numero) { avisar("Informe o número do empenho."); return; }
  const data = {
    numero,
    documento: fd.get("documento").trim(),
    ano: fd.get("ano") || anoEmpenho({ numero, data: fd.get("data") }),
    anoManual: true,
    data: fd.get("data") || "",
    dataTexto: "",
    fornecedor: fd.get("fornecedor").trim(),
    cnpj: fd.get("cnpj").trim(),
    objeto: fd.get("objeto").trim(),
    unidade: fd.get("unidade").trim(),
    valor: parseValor(fd.get("valor")),
    status: fd.get("status") || "",
    observacoes: fd.get("observacoes").trim(),
  };
  if (ui.editId) {
    const i = state.empenhos.findIndex((e) => e.id === ui.editId);
    if (i >= 0) state.empenhos[i] = { ...state.empenhos[i], ...data };
    ui.aberto = { tipo: "empenho", id: ui.editId };
  } else {
    data.id = uid();
    data.origem = "manual";
    state.empenhos.push(data);
    ui.aberto = { tipo: "empenho", id: data.id };
  }
  ui.form = null;
  ui.editId = null;
  save();
  render();
  avisar("Empenho salvo neste navegador.");
}

function salvarProcesso(fd) {
  const sis = normSis(fd.get("sis")) || fd.get("sis").trim();
  const objeto = fd.get("objeto").trim();
  if (!sis && !objeto) { avisar("Informe o SIS ou o objeto."); return; }
  const data = {
    id: uid(),
    sis,
    objeto,
    data: fd.get("data") || hojeISO(),
    dataTexto: "",
    setor: fd.get("setor") || "",
    descricao: fd.get("descricao").trim(),
    observacao: "",
    ultimaVerificacao: hojeISO(),
    ultimaTexto: "",
    dataTramitacao: fd.get("setor") ? hojeISO() : "",
    ano: (sis.match(/(\d{4})$/) || [])[1] || (fd.get("data") || hojeISO()).slice(0, 4),
  };
  state.processos.push(data);
  ui.form = null;
  ui.editId = null;
  ui.aberto = { tipo: "processo", id: data.id };
  save();
  render();
  avisar("Processo salvo neste navegador.");
}

function salvarAndamento(fd) {
  const p = ui.aberto && procPorId(ui.aberto.id);
  if (!p) return;
  const setor = fd.get("setor") || "";
  if (setor !== (p.setor || "")) p.dataTramitacao = hojeISO();
  p.setor = setor;
  p.descricao = String(fd.get("descricao") || "").trim();
  p.sis = normSis(fd.get("sis")) || String(fd.get("sis") || "").trim();
  p.data = fd.get("data") || "";
  p.dataTexto = "";
  p.objeto = String(fd.get("objeto") || "").trim();
  p.ano = (String(p.sis).match(/(\d{4})$/) || [])[1] || String(p.data || "").slice(0, 4);
  save();
  render();
  avisar("Processo atualizado neste navegador.");
}

function salvarFicha(fd) {
  const p = ui.aberto && procPorId(ui.aberto.id);
  if (!p) return;
  const sis = normSis(fd.get("sis")) || fd.get("sis").trim();
  p.sis = sis;
  p.data = fd.get("data") || "";
  p.dataTexto = "";
  p.objeto = fd.get("objeto").trim();
  p.ano = (sis.match(/(\d{4})$/) || [])[1] || (p.data || "").slice(0, 4);
  save();
  render();
  avisar("Ficha atualizada neste navegador.");
}

function salvarEmenda(fd) {
  const numero = normEmenda(fd.get("numero"));
  if (!numero) { avisar("Use o número no formato 033/2025."); return; }
  const data = {
    numero,
    documento: fd.get("documento").trim(),
    ano: numero.slice(-4),
    vereador: fd.get("vereador").trim(),
    partido: fd.get("partido").trim(),
    valor: parseValor(fd.get("valor")),
    destino: fd.get("destino").trim(),
    ubs: fd.get("ubs").trim(),
    natureza: fd.get("natureza").trim(),
    objeto: fd.get("objeto").trim(),
    ementa: fd.get("ementa").trim(),
    justificativa: fd.get("justificativa").trim(),
    definicao: fd.get("definicao").trim(),
    materiaisDefinidos: fd.get("materiaisDefinidos").trim(),
    statusExecucao: fd.get("statusExecucao").trim(),
    observacao: fd.get("observacao").trim(),
    memorando: fd.get("memorando").trim(),
    sis: normSis(fd.get("sis")),
  };
  if (ui.editId) {
    const i = state.emendas.findIndex((e) => e.id === ui.editId);
    if (i >= 0) state.emendas[i] = { ...state.emendas[i], ...data };
    ui.aberto = { tipo: "emenda", id: ui.editId };
  } else {
    if (emendaPorNumero(numero)) { avisar("Já existe emenda com esse número. Abra a ficha e atualize."); return; }
    data.id = uid();
    state.emendas.push(data);
    ui.aberto = { tipo: "emenda", id: data.id };
  }
  ui.form = null;
  ui.editId = null;
  save();
  render();
  avisar("Emenda salva neste navegador.");
}

function salvarItem(fd) {
  const grupo = ui.aberto?.tipo === "grupo" ? ui.aberto.id : "";
  const base = state.itens.find((i) => normEmp(i.numeroEmpenho) === grupo) || {};
  const data = {
    descricao: fd.get("descricao").trim(),
    quantidade: parseValor(fd.get("quantidade")),
    unidade: fd.get("unidade").trim(),
    valorTotal: parseValor(fd.get("valorTotal")),
    destino: fd.get("destino").trim(),
    previsaoEntrega: fd.get("previsaoEntrega") || "",
    previsaoTexto: "",
    sis: normSis(fd.get("sis")),
    emenda: normEmenda(fd.get("emenda")),
    observacoes: fd.get("observacoes").trim(),
  };
  if (!data.descricao && !data.destino) { avisar("Informe a descrição ou o destino."); return; }
  if (ui.editId) {
    const i = state.itens.findIndex((x) => x.id === ui.editId);
    if (i >= 0) state.itens[i] = { ...state.itens[i], ...data };
  } else {
    state.itens.push({
      ...data,
      id: uid(),
      numeroEmpenho: base.numeroEmpenho || grupo,
      data: base.data || "",
      dataTexto: base.dataTexto || "",
      tipo: base.tipo || "",
      autor: base.autor || "",
      solicitante: base.solicitante || "",
      fornecedor: base.fornecedor || "",
      valorUnitario: null,
      origem: "manual",
    });
  }
  ui.form = null;
  ui.editId = null;
  save();
  render();
  avisar("Item salvo neste navegador.");
}

function salvarRegistro(fd) {
  const nome = String(fd.get("nome") || "").trim();
  if (!nome) { avisar("Informe o nome do grupo."); return; }
  const ano = String(fd.get("ano") || "").trim();
  if (!anoPlausivel(ano)) { avisar("Selecione um ano válido."); return; }
  const links = {};
  for (let i = 0; i < PECAS_RP.length; i += 1) {
    const p = PECAS_RP[i];
    const lido = linkDocRp(fd.get(p.chave));
    if (lido == null) { avisar(`O ${p.rotulo} precisa ser um link https:// ou o arquivo já incluído no grupo.`); return; }
    links[p.chave] = lido;
  }
  const data = {
    ano,
    nome,
    objeto: String(fd.get("objeto") || "").trim(),
    situacao: SITUACOES_RP.includes(fd.get("situacao")) ? fd.get("situacao") : "Em elaboração",
    observacoes: String(fd.get("observacoes") || "").trim(),
    dfd: links.dfd,
    etp: links.etp,
    tr: links.tr,
  };
  const idAberto = ui.aberto && ui.aberto.tipo === "registro" ? ui.aberto.id : "";
  if (idAberto) {
    const i = state.registros.findIndex((r) => r.id === idAberto);
    if (i >= 0) state.registros[i] = Object.assign({}, state.registros[i], data);
  } else {
    const id = uid();
    state.registros.push(Object.assign({ id, materiais: [] }, data));
    ui.aberto = { tipo: "registro", id };
    ui.form = null;
  }
  save();
  render();
  avisar("Grupo salvo neste navegador.");
}

function salvarMaterialRp(fd) {
  preservarRegistro();
  const reg = ui.aberto && ui.aberto.tipo === "registro" ? registroPorId(ui.aberto.id) : null;
  if (!reg) { avisar("Abra o grupo antes de incluir o material."); return; }
  const item = String(fd.get("item") || "").trim();
  if (!item) { avisar("Informe o item."); return; }
  const data = {
    item,
    unidade: String(fd.get("unidade") || "").trim(),
    quantidade: parseValor(fd.get("quantidade")),
    descricao: String(fd.get("descricao") || "").trim(),
  };
  reg.materiais = reg.materiais || [];
  if (ui.editId) {
    const i = reg.materiais.findIndex((m) => m.id === ui.editId);
    if (i >= 0) reg.materiais[i] = Object.assign({}, reg.materiais[i], data);
  } else {
    reg.materiais.push(Object.assign({ id: uid() }, data));
  }
  ui.form = null;
  ui.editId = null;
  ui._mat = null;
  save();
  render();
  avisar("Material salvo neste navegador.");
}

function salvarLote(fd) {
  const item = fd.get("item").trim();
  if (!item) { avisar("Informe o item."); return; }
  const data = {
    lote: fd.get("lote").trim(),
    ano: fd.get("ano") || ui.filtros.materiais.ano || "2026",
    anoManual: true,
    item,
    sis: fd.get("sis").trim(),
    descricao: fd.get("descricao").trim(),
    destinos: fd.get("destinos").trim(),
    emendas: fd.get("emendas").trim(),
  };
  if (ui.editId) {
    const i = state.lotes.findIndex((l) => l.id === ui.editId);
    if (i >= 0) state.lotes[i] = { ...state.lotes[i], ...data };
  } else {
    data.id = uid();
    state.lotes.push(data);
  }
  ui.form = null;
  ui.editId = null;
  save();
  render();
  avisar("Material salvo neste navegador.");
}

function excluir(tipo, id) {
  if (tipo === "empenho") {
    const e = empenhoPorId(id);
    if (!e || !confirm(`Excluir o empenho ${e.numero}?`)) return;
    state.empenhos = state.empenhos.filter((x) => x.id !== id);
    ui.aberto = null;
    ui.form = null;
    save();
    removerAnexosDoEmpenho(id).catch(() => {}).finally(() => {
      render();
      avisar("Registro excluído neste navegador.");
    });
    return;
  } else if (tipo === "emenda") {
    const e = emendaPorId(id);
    if (!e || !confirm(`Excluir a emenda ${e.numero}?`)) return;
    state.emendas = state.emendas.filter((x) => x.id !== id);
  } else if (tipo === "lote") {
    const l = state.lotes.find((x) => x.id === id);
    if (!l || !confirm(`Excluir o material ${l.item}?`)) return;
    state.lotes = state.lotes.filter((x) => x.id !== id);
  } else if (tipo === "item") {
    const item = state.itens.find((x) => x.id === id);
    if (!item || !confirm("Excluir este destino?")) return;
    state.itens = state.itens.filter((x) => x.id !== id);
    if (ui.aberto?.tipo === "grupo" && !state.itens.some((x) => normEmp(x.numeroEmpenho) === ui.aberto.id)) ui.aberto = null;
    ui.form = null;
    ui.editId = null;
    save();
    render();
    avisar("Destino excluído neste navegador.");
    return;
  } else if (tipo === "registro") {
    const r = registroPorId(id);
    if (!r || !confirm(`Excluir o grupo ${r.nome}? Os materiais deste grupo também saem.`)) return;
    state.registros = state.registros.filter((x) => x.id !== id);
    ui.aberto = null;
    ui.form = null;
    ui.editId = null;
    ui._mat = null;
    save();
    removerAnexosDoRegistro(id).catch(() => {}).finally(() => {
      render();
      avisar("Grupo excluído neste navegador.");
    });
    return;
  } else if (tipo === "material-rp") {
    const reg = state.registros.find((r) => (r.materiais || []).some((m) => m.id === id));
    const m = reg && reg.materiais.find((x) => x.id === id);
    if (!m || !confirm(`Excluir o material ${m.item}?`)) return;
    reg.materiais = reg.materiais.filter((x) => x.id !== id);
    ui.form = null;
    ui.editId = null;
    ui._mat = null;
    save();
    render();
    avisar("Material excluído deste grupo.");
    return;
  } else if (tipo === "processo") {
    const p = procPorId(id);
    if (!p || !confirm(`Excluir o processo ${p.sis || p.objeto}?`)) return;
    state.processos = state.processos.filter((x) => x.id !== id);
  } else return;
  ui.aberto = null;
  ui.form = null;
  save();
  render();
  avisar("Registro excluído neste navegador.");
}

function onClick(ev) {
  const actEl = ev.target.closest("[data-act]");
  if (actEl) {
    const act = actEl.dataset.act;
    if (act === "open") abrir(actEl.dataset.tipo, actEl.dataset.id);
    else if (act === "goto") { location.hash = actEl.dataset.arg; }
    else if (act === "fechar") fechar();
    else if (act === "fechar-form") { ui.form = null; ui.editId = null; ui._mat = null; render(); }
    else if (act === "tab") { ui.tab = actEl.dataset.arg; ui.aberto = null; ui.form = null; ui._mat = null; render(); }
    else if (act === "novo") {
      ui.form = actEl.dataset.tipo === "item" ? "item" : actEl.dataset.tipo;
      ui.editId = null;
      ui._mat = null;
      if (actEl.dataset.tipo === "empenho") ui.tab = "controle";
      if (actEl.dataset.tipo !== "item" && actEl.dataset.tipo !== "material-rp") ui.aberto = null;
      render();
    } else if (act === "edit") {
      ui.form = actEl.dataset.tipo === "item" ? "item" : actEl.dataset.tipo;
      ui.editId = actEl.dataset.id;
      ui._mat = null;
      render();
    } else if (act === "del") excluir(actEl.dataset.tipo, actEl.dataset.id);
    else if (act === "verificar") {
      const p = procPorId(actEl.dataset.id);
      if (p) {
        const form = actEl.closest("form");
        if (form) p.descricao = String(new FormData(form).get("descricao") || "").trim();
        p.ultimaVerificacao = hojeISO();
        p.ultimaTexto = "";
        save();
        render();
        avisar("Verificação marcada para hoje.");
      }
    } else if (act === "csv") exportarCsv(actEl.dataset.arg);
    else if (act === "abrir-anexo") abrirAnexo(actEl.dataset.id);
    else if (act === "del-anexo") removerAnexo(actEl.dataset.id);
    else if (act === "exportar") exportarCopia();
    else if (act === "reset") {
      if (confirm("Isso substitui tudo o que foi alterado neste navegador pela leitura original das planilhas. Continuar?")) {
        const grupos = state.registros || [];
        state = sane(clone(window.SEED || {}));
        state.registros = grupos;
        save();
        fechar();
        avisar("Dados de origem recarregados.");
      }
    }
    return;
  }
  if (ev.target.closest("select, input, textarea, button, a, label, summary")) return;
  const tr = ev.target.closest("tr[data-tipo]");
  if (tr) abrir(tr.dataset.tipo, tr.dataset.id);
}

function onSubmit(ev) {
  const form = ev.target.closest("form[data-form]");
  if (!form) return;
  ev.preventDefault();
  const fd = new FormData(form);
  if (form.dataset.form === "empenho") salvarEmpenho(fd);
  else if (form.dataset.form === "processo") salvarProcesso(fd);
  else if (form.dataset.form === "andamento") salvarAndamento(fd);
  else if (form.dataset.form === "ficha") salvarFicha(fd);
  else if (form.dataset.form === "emenda") salvarEmenda(fd);
  else if (form.dataset.form === "item") salvarItem(fd);
  else if (form.dataset.form === "lote") salvarLote(fd);
  else if (form.dataset.form === "registro") salvarRegistro(fd);
  else if (form.dataset.form === "material-rp") salvarMaterialRp(fd);
}

function onChange(ev) {
  const el = ev.target;
  if (el.dataset.act === "status") {
    const e = empenhoPorId(el.dataset.id);
    if (!e) return;
    e.status = el.value;
    save();
    render();
    avisar("Status atualizado.");
    return;
  }
  if (el.dataset.act === "atalho") {
    const url = el.value.trim();
    if (url && !/^https?:\/\//i.test(url)) {
      avisar("O link precisa começar com https://");
      return;
    }
    state.atalhos = Object.assign({ documentosRelacao: "", pastaEmendas: "" }, state.atalhos || {});
    state.atalhos[el.dataset.chave] = url;
    save();
    render();
    avisar(url ? "Link guardado neste navegador." : "Link removido neste navegador.");
    return;
  }
  if (el.dataset.act === "setor-proc") {
    const p = procPorId(el.dataset.id);
    if (!p) return;
    const setor = el.value || "";
    if (setor !== (p.setor || "")) p.dataTramitacao = hojeISO();
    p.setor = setor;
    save();
    render();
    avisar("Setor atualizado.");
    return;
  }
  if (el.dataset.act === "verificacao") {
    const p = procPorId(el.dataset.id);
    if (!p) return;
    p.ultimaVerificacao = el.value || "";
    p.ultimaTexto = "";
    save();
    render();
    avisar("Verificação atualizada.");
    return;
  }
  if (el.dataset.anexo && el.files && el.files[0]) {
    const file = el.files[0];
    const empenhoId = el.dataset.id;
    const tipo = el.dataset.anexo;
    el.value = "";
    gravarAnexo(empenhoId, tipo, file);
    return;
  }
  if (el.id === "arquivo-copia" && el.files && el.files[0]) {
    const file = el.files[0];
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(String(reader.result || ""));
        if (!parsed.empenhos || !parsed.emendas || !parsed.processos) throw new Error("formato");
        const arquivos = Array.isArray(parsed.anexos) ? parsed.anexos : null;
        delete parsed.anexos;
        state = sane(parsed);
        save();
        const pronto = () => {
          fechar();
          avisar(arquivos ? "Cópia restaurada, com os arquivos." : "Cópia restaurada neste navegador.");
        };
        if (arquivos) restaurarAnexos(arquivos).then(pronto).catch(() => avisar("Os registros voltaram, mas os arquivos não puderam ser restaurados."));
        else pronto();
      } catch {
        avisar("Esse arquivo não é uma cópia deste painel.");
      }
    };
    reader.readAsText(file, "utf-8");
    return;
  }
  if (el.dataset.filtro && el.dataset.grupo) {
    ui.filtros[el.dataset.grupo][el.dataset.filtro] = el.value;
    render();
  }
}

function onInput(ev) {
  const el = ev.target;
  if (!el.dataset.live || !el.dataset.grupo) return;
  ui.filtros[el.dataset.grupo][el.dataset.filtro] = el.value;
  render();
}

function nomeView(raw) {
  const v = raw === "atas" ? "materiais" : raw;
  return VIEWS.includes(v) ? v : "painel";
}
function syncHash() {
  const bruto = (location.hash || "#painel").slice(1);
  const v = nomeView(bruto);
  if (ui._keep) {
    ui._keep = false;
    ui.view = v;
    render();
    return;
  }
  ui.view = v;
  ui.aberto = null;
  ui.form = null;
  ui.editId = null;
  render();
}

function init() {
  if (!window.SEED) {
    document.getElementById("view").innerHTML = `<p class="notice">Não encontrei a base importada das planilhas (js/seed.js).</p>`;
    return;
  }
  document.getElementById("view").addEventListener("click", onClick);
  document.getElementById("view").addEventListener("submit", onSubmit);
  document.getElementById("view").addEventListener("change", onChange);
  document.getElementById("view").addEventListener("input", onInput);
  document.getElementById("q").addEventListener("keydown", (ev) => {
    if (ev.key === "Enter") {
      ui.q = ev.target.value.trim();
      if (location.hash !== "#busca") location.hash = "busca";
      else render();
    }
  });
  window.addEventListener("hashchange", syncHash);
  document.addEventListener("keydown", (ev) => {
    if (ev.key === "Escape" && (ui.aberto || ui.form)) fechar();
  });
  listarAnexos().catch(() => {}).then(() => syncHash());
}

init();

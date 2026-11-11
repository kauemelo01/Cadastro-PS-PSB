"use strict";
const $ = id => document.getElementById(id);
const embedded = window.parent !== window;
const state = { db: null, dataset: new URLSearchParams(location.search).get("dataset"),
  snapshot: null, operations: [], editing: null, changes: null, syncing: false, booted: false };
const channel = typeof BroadcastChannel !== "undefined" ? new BroadcastChannel("ps-visitas") : null;
const textNode = (tag, value, className) => {
  const element = document.createElement(tag);
  element.textContent = value;
  if (className) element.className = className;
  return element;
};
function send(type, extra = {}) {
  if (embedded) window.parent.postMessage({ isStreamlitMessage: true, type, ...extra }, "*");
}
function resize() { send("streamlit:setFrameHeight", { height: document.body.scrollHeight + 24 }); }
function message(text, kind = "") {
  $("message").textContent = text; $("message").className = kind; $("message").hidden = false; resize();
}
function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open("pastoral-social-visitas", 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      db.createObjectStore("snapshots", { keyPath: "dataset" });
      const ops = db.createObjectStore("operations", { keyPath: "_sequence", autoIncrement: true });
      ops.createIndex("id", "id", { unique: true });
      ops.createIndex("dataset", "dataset");
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}
function transaction(stores, mode, callback) {
  return new Promise((resolve, reject) => {
    const tx = state.db.transaction(stores, mode);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error || new Error("Falha ao gravar no cache."));
    tx.onabort = () => reject(tx.error || new Error("Gravação do cache cancelada."));
    try { callback(tx); } catch (error) { tx.abort(); reject(error); }
  });
}
async function refresh() {
  if (!state.dataset) return;
  await transaction(["snapshots", "operations"], "readonly", tx => {
    const snapshot = tx.objectStore("snapshots").get(state.dataset);
    snapshot.onsuccess = () => { state.snapshot = snapshot.result || null; };
    const operations = tx.objectStore("operations").index("dataset").getAll(state.dataset);
    operations.onsuccess = () => { state.operations = operations.result.sort((a,b) => a._sequence - b._sequence); };
  });
  renderStatus();
  if (!state.editing) renderResults();
}
function notifyTabs() { if (channel) channel.postMessage({ dataset: state.dataset }); }
function displayTimestamp(iso) {
  return new Intl.DateTimeFormat("pt-BR", { timeZone: "America/Sao_Paulo", dateStyle: "short", timeStyle: "medium" }).format(new Date(iso));
}
function currentRows() {
  const rows = (state.snapshot?.records || []).map(row => ({ ...row }));
  for (const op of state.operations) {
    const matches = rows.filter(row => row.NUMERO === op.numero);
    let row = matches[0];
    if (!row) { row = { ...op.row, NUMERO: op.numero }; rows.push(row); }
    for (const [col, ch] of Object.entries(op.changes)) row[col] = ch.new;
    row[state.snapshot.audit_name] = op.editor;
    row[state.snapshot.audit_ts] = displayTimestamp(op.timestamp);
  }
  return rows;
}
function renderStatus() {
  const n = state.operations.length;
  $("connection").textContent = navigator.onLine ? "Cache pronto · aparelho com conexão" : "Sem internet · usando o cache deste aparelho";
  $("pending-count").textContent = `${n} ${n === 1 ? "alteração pendente" : "alterações pendentes"} neste aparelho.`;
  $("sync").disabled = !n || state.syncing || !navigator.onLine;
  $("sync").textContent = state.syncing ? "Enviando… Aguarde a confirmação" : "☁️ Enviar todas as alterações";
}
function renderResults() {
  const area = $("results"); area.replaceChildren();
  if (!state.snapshot) {
    area.append(textNode("p", "Nenhum cadastro em cache. Acesse uma vez com internet para preparar este aparelho.", "empty")); resize(); return;
  }
  let query = $("query").value.trim();
  if (/^\d+$/.test(query)) query = query.replace(/^0+(?=\d)/, "");
  const rows = currentRows();
  if (!query) {
    area.append(textNode("p", `Digite o número do registro para consultar e editar. ${rows.length} ${rows.length === 1 ? "cadastro disponível" : "cadastros disponíveis"} neste aparelho.`, "empty"));
  } else {
    const matches = rows.filter(row => (row.NUMERO || "").trim() === query);
    if (!matches.length) area.append(textNode("p", "Nenhum registro encontrado. Confira o número e tente de novo.", "empty"));
    for (const row of matches) {
      const card = document.createElement("article"); card.className = "card";
      card.append(textNode("h2", row.NOME || "—"), textNode("small", `Nº ${row.NUMERO} · CPF ${row.CPF || "—"}`));
      const details = document.createElement("dl");
      for (const [label, value] of [["Tipo", row.TIPO], ["CID 2026", row["CID 2026"]],
        ["Telefone", row.TELEFONE], ["Endereço", [row["RUA PRINCIPAL"], row["Nº CASA"], row.VIELA].filter(Boolean).join(", ")],
        ["Precisa cesta", row["PRECISA RECEBER CESTAS BÁSICAS DA PASTORAL SOCIAL? [AVALIAÇÃO]"]],
        ["Beneficiário", row["TIPO DE BENEFICIÁRIO [TIPO]"]]]) {
        details.append(textNode("dt", label), textNode("dd", value || "—"));
      }
      card.append(details);
      const editor = row[state.snapshot.audit_name], ts = row[state.snapshot.audit_ts];
      if (editor || ts) card.append(textNode("p", `Última edição: ${editor || ""} · ${ts || ""}`, "hint"));
      const button = textNode("button", "✏️ Editar Cadastro"); button.type = "button";
      button.addEventListener("click", () => editRow(row)); card.append(button); area.append(card);
    }
  }
  resize();
}
function toISODate(text) {
  const match = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(text);
  return match ? `${match[3]}-${match[2]}-${match[1]}` : "";
}
function fromISODate(text) { return text ? text.split("-").reverse().join("/") : ""; }
function addField(col, group) {
  const row = state.editing.row;
  if (!(col in row) || state.editing.controls.has(col)) return;
  const spec = state.snapshot.specs[col] || { label: col, kind: "text" };
  const initial = row[col] || "";
  const field = document.createElement("div"); field.className = "field";
  const label = textNode("label", spec.label); field.append(label);
  const kind = spec.kind;
  let input, getValue;
  if (kind === "multi") {
    input = document.createElement("div"); input.className = "checks";
    const tokens = initial.split(/[,;]/).map(x => x.trim()).filter(Boolean);
    for (const option of [...new Set([...(spec.options || []), ...tokens])]) {
      const choice = document.createElement("input"); choice.type = "checkbox"; choice.value = option; choice.checked = tokens.includes(option);
      const optionLabel = document.createElement("label"); optionLabel.append(choice, document.createTextNode(option)); input.append(optionLabel);
    }
    getValue = () => Array.from(input.querySelectorAll("input:checked"), x => x.value).join(", ");
  } else if (["select", "choice", "yesno"].includes(kind)) {
    input = document.createElement("select");
    const options = kind === "yesno" ? ["SIM", "NÃO"] : spec.options;
    for (const option of [...new Set(["", ...options, ...(initial ? [initial] : [])])]) {
      const element = textNode("option", option || "Selecione…"); element.value = option; input.append(element);
    }
    input.value = initial; getValue = () => input.value;
  } else {
    input = document.createElement(kind === "textarea" ? "textarea" : "input");
    if (kind === "date") {
      input.type = "date"; input.min = "1900-01-01";
      const today = new Date(); input.max = `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,"0")}-${String(today.getDate()).padStart(2,"0")}`;
      input.value = toISODate(initial); getValue = () => fromISODate(input.value);
      if (initial && !input.value) field.append(textNode("p", `Valor atual: ${initial}`, "hint"));
    } else {
      input.value = initial; input.placeholder = spec.ph || ""; getValue = () => input.value.trim();
    }
    if (kind === "locked") input.disabled = true;
  }
  input.id = `field-${state.editing.controls.size}`; label.htmlFor = input.id;
  let dirty = false; input.addEventListener("input", () => { dirty = true; }); input.addEventListener("change", () => { dirty = true; });
  state.editing.controls.set(col, { getValue, initial, spec, isDirty: () => dirty });
  field.append(input); group.append(field);
}
function editRow(row) {
  state.editing = { row: { ...row }, controls: new Map() };
  $("fields").replaceChildren(); $("edit-title").textContent = `${row.NOME || "—"} · Nº ${row.NUMERO}`;
  const shown = new Set();
  state.snapshot.sections.forEach(([title, cols], index) => {
    const group = document.createElement("details"); group.open = index === 0 || title.includes("⭐");
    if (title.includes("⭐")) group.className = "cestas";
    group.append(textNode("summary", title)); cols.forEach(col => { shown.add(col); addField(col, group); });
    $("fields").append(group);
  });
  const extras = Object.keys(row).filter(col => !shown.has(col) && ![state.snapshot.audit_name,state.snapshot.audit_ts].includes(col));
  if (extras.length) {
    const group = document.createElement("details"); group.append(textNode("summary", "📎 Outros campos"));
    extras.forEach(col => addField(col, group)); $("fields").append(group);
  }
  $("home").hidden = true; $("editor").hidden = false; resize(); window.scrollTo(0, 0);
}
function closeEditor() { state.editing = null; $("editor").hidden = true; $("home").hidden = false; renderResults(); }
function reviewChanges(event) {
  event.preventDefault();
  const changes = {};
  for (const [col, control] of state.editing.controls) {
    if (!control.isDirty() || control.spec.kind === "locked") continue;
    let value = control.getValue();
    if (["CPF", "CPF RESERVA 1", "CPF RESERVA 2"].includes(col)) {
      value = value.replace(/^(\d{3})\.?(\d{3})\.?(\d{3})-?(\d{2})$/, "$1.$2.$3-$4");
    }
    if (value !== control.initial) changes[col] = { old: control.initial, new: value, label: control.spec.label };
  }
  if (!Object.keys(changes).length) { message("Nenhuma alteração para salvar."); return; }
  if (changes.NOME && !changes.NOME.new) { message("O nome não pode ficar em branco.", "error"); return; }
  state.changes = changes; $("changes-list").replaceChildren();
  for (const change of Object.values(changes)) {
    const item = document.createElement("div"); item.className = "change";
    item.append(textNode("strong", change.label), textNode("span", change.old || "(vazio)", "old"),
      document.createTextNode(" → "), textNode("span", change.new || "(vazio)", "new"));
    $("changes-list").append(item);
  }
  $("confirm-dialog").showModal(); resize();
}
async function confirmSave(event) {
  event.preventDefault();
  const editor = $("editor-name").value.trim().replace(/\s+/g, " ");
  if (editor.length < 2) return;
  $("confirm-save").disabled = true;
  try {
    const autoSync = embedded && navigator.onLine && !state.operations.length && state.snapshot.write_enabled;
    const op = { id: crypto.randomUUID(), dataset: state.dataset, numero: state.editing.row.NUMERO,
      row: state.editing.row, changes: state.changes, editor, timestamp: new Date().toISOString() };
    await transaction(["operations"], "readwrite", tx => tx.objectStore("operations").add(op));
    $("confirm-dialog").close(); closeEditor(); await refresh(); notifyTabs();
    message("Alteração salva no cache deste aparelho. Use “Enviar todas as alterações” quando houver internet.", "success");
    if (autoSync) await requestSync();
  } catch (error) {
    $("confirm-dialog").close(); message(`Não foi possível salvar no aparelho: ${error.message}. A edição continua aberta.`, "error");
  } finally { $("confirm-save").disabled = false; }
}
function appRoot() {
  const prefix = location.pathname.split("/component/")[0];
  return location.origin + prefix + "/";
}
async function requestSync() {
  if (state.syncing) return;
  await refresh();
  if (!state.operations.length) return;
  const requestId = crypto.randomUUID();
  const inflight = { request_id: requestId, operations: state.operations.map(op => ({ ...op })) };
  await transaction(["snapshots"], "readwrite", tx => {
    const store = tx.objectStore("snapshots"), request = store.get(state.dataset);
    request.onsuccess = () => store.put({ ...request.result, inflight });
  });
  state.syncing = true; renderStatus();
  if (embedded) {
    send("streamlit:setComponentValue", { value: { type: "sync", ...inflight }, dataType: "json" });
    startSyncTimer();
  } else {
    // O clique autoriza o envio. O site lê o mesmo IndexedDB e transmite este lote.
    location.assign(appRoot());
  }
}
let syncTimer;
function startSyncTimer() {
  clearTimeout(syncTimer);
  syncTimer = setTimeout(() => {
    state.syncing = false; renderStatus();
    message("A confirmação do envio ainda não chegou. As pendências continuam salvas; você pode tentar novamente.");
  }, 45000);
}
async function consumeResponse(response) {
  if (!response || state.snapshot?.last_response === response.request_id) return;
  const acked = new Set(Object.keys(response.targets || {}));
  if (response.ok && !state.operations.some(op => acked.has(op.id))) return;
  await transaction(["snapshots", "operations"], "readwrite", tx => {
    const store = tx.objectStore("snapshots"), request = store.get(state.dataset);
    request.onsuccess = () => {
      const snapshot = request.result;
      if (!snapshot || snapshot.last_response === response.request_id) return;
      // Uma resposta atrasada não pode apagar um pedido mais recente.
      const inflight = snapshot.inflight;
      const aliases = {};
      if (response.ok) {
        for (const op of inflight?.operations || []) {
          if (response.targets[op.id]) aliases[op.numero] = response.targets[op.id];
        }
        const pending = tx.objectStore("operations").index("dataset").openCursor(state.dataset);
        pending.onsuccess = () => {
          const cursor = pending.result;
          if (!cursor) return;
          const op = cursor.value;
          if (acked.has(op.id)) cursor.delete();
          else if (aliases[op.numero]) cursor.update({ ...op, numero: aliases[op.numero], row: { ...op.row, NUMERO: aliases[op.numero] } });
          cursor.continue();
        };
        snapshot.records = response.records;
      }
      snapshot.last_response = response.request_id;
      if (inflight?.request_id === response.request_id) snapshot.inflight = null;
      store.put(snapshot);
    };
  });
  clearTimeout(syncTimer); state.syncing = false; await refresh(); notifyTabs();
  message(response.ok ? `${acked.size} ${acked.size === 1 ? "alteração enviada" : "alterações enviadas"}. Datas e horários originais preservados.` : response.error, response.ok ? "success" : "error");
}
let renderQueue = Promise.resolve();
async function onRender(args) {
  state.dataset = args.dataset;
  await refresh();
  const previous = state.snapshot;
  const snapshot = { ...previous, dataset: args.dataset, specs: args.specs, sections: args.sections,
    audit_name: args.audit_name, audit_ts: args.audit_ts, write_enabled: args.write_enabled };
  if (args.records !== null) snapshot.records = args.records;
  // Se a fonte estiver fora do ar, preserva a última cópia em disco.
  if (!snapshot.records) { message("Não há cadastros em cache. Conecte à internet e tente novamente.", "error"); return; }
  await transaction(["snapshots"], "readwrite", tx => {
    const store = tx.objectStore("snapshots"), request = store.get(state.dataset);
    request.onsuccess = () => store.put({ ...request.result, ...snapshot,
      inflight: request.result?.inflight, last_response: request.result?.last_response });
  });
  await refresh(); await consumeResponse(args.response);
  if (!state.booted) {
    state.booted = true;
    if (state.snapshot.inflight) {
      state.syncing = true; renderStatus();
      send("streamlit:setComponentValue", { value: { type: "sync", ...state.snapshot.inflight }, dataType: "json" });
      startSyncTimer();
    }
  }
  await prepareOffline(); resize();
}
let prepared = false;
async function prepareOffline() {
  if (prepared) return;
  if (!("serviceWorker" in navigator)) {
    $("offline-hint").textContent = "Este navegador permite editar enquanto a tela estiver aberta. Para reabrir offline, use um navegador com suporte a cache offline em HTTPS."; return;
  }
  try {
    await navigator.serviceWorker.register("./sw.js", { scope: "./" });
    await navigator.serviceWorker.ready;
    const url = new URL("./index.html", location.href); url.searchParams.set("dataset", state.dataset);
    $("offline-link").href = url.href; $("offline-link").hidden = !embedded;
    $("offline-hint").textContent = embedded
      ? "Abra a versão offline e salve seu endereço antes de sair para as visitas. Ela permite consultar, editar e reabrir a tela sem internet."
      : "Versão offline pronta. Salve este endereço nos favoritos para reabrir sem internet. Os dados e as pendências ficam neste navegador.";
    prepared = true;
  } catch (error) {
    $("offline-hint").textContent = `Não foi possível preparar a reabertura offline (${error.message}). Mantenha esta tela aberta para editar sem conexão.`;
  }
}
async function init() {
  $("standalone-header").hidden = embedded;
  $("query").addEventListener("input", renderResults);
  $("edit-form").addEventListener("submit", reviewChanges);
  $("confirm-form").addEventListener("submit", confirmSave);
  $("cancel-edit").addEventListener("click", closeEditor);
  $("back").addEventListener("click", () => $("confirm-dialog").close());
  $("sync").addEventListener("click", () => requestSync().catch(error => message(error.message, "error")));
  for (const name of ["online", "offline"]) window.addEventListener(name, renderStatus);
  if (channel) channel.onmessage = event => { if (event.data.dataset === state.dataset) refresh().catch(error => message(error.message, "error")); };
  new ResizeObserver(resize).observe(document.body);
  try {
    state.db = await openDatabase();
    if (embedded) {
      window.addEventListener("message", event => {
        if (event.source !== window.parent || event.data.type !== "streamlit:render") return;
        renderQueue = renderQueue.then(() => onRender(event.data.args)).catch(error => message(`Falha no cache: ${error.message}`, "error"));
      });
      send("streamlit:componentReady", { apiVersion: 1 });
    } else {
      await refresh(); await prepareOffline();
      if (state.snapshot?.inflight) message("Há um envio sem confirmação. Reconecte e toque em enviar para conferir; alterações já recebidas não serão duplicadas.");
    }
    if (navigator.storage?.persist) navigator.storage.persist().catch(() => {});
  } catch (error) {
    message(`O cache deste navegador está indisponível: ${error.message}. Habilite o armazenamento para usar o modo offline.`, "error");
  }
}
init();

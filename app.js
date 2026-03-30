// Exam Writer – SharePoint sync via Cloudflare Worker ROPC broker

let fileHandle = null;
let autosaveTimer = null;
let dirty = false;
let adminPinHash = localStorage.getItem("adminPinHash") || null;

// SharePoint state
let cfg = null;
let spAccessToken = null;
let spTokenExpiry = 0;
let spRefreshTimer = null;

const $ = (s) => document.querySelector(s);

/* ══ SERVICE WORKER ══ */
async function registerSW() {
  if ("serviceWorker" in navigator) {
    try { await navigator.serviceWorker.register("./service-worker.js", { scope: "./" }); } catch {}
  }
}

/* ══ UI HELPERS ══ */
function setDirty(v) { dirty = v; $("#dirtyDot").hidden = !v; }
function updateWordCount() {
  const t = $("#editor").value.trim();
  const w = t ? (t.match(/\b\w+\b/g)?.length ?? 0) : 0;
  $("#wordCount").textContent = `${w} word${w === 1 ? "" : "s"}`;
}
async function ensurePersistence() {
  if (navigator.storage?.persist) { try { await navigator.storage.persist(); } catch {} }
}

/* ══ THEME ══ */
function applyTheme(t) { document.documentElement.setAttribute("data-theme", t); localStorage.setItem("theme", t); }
function initTheme() { applyTheme(localStorage.getItem("theme") || "dark"); }
function toggleTheme() { applyTheme((localStorage.getItem("theme") || "dark") === "dark" ? "light" : "dark"); }

/* ══ CONFIG ══ */
function loadConfig() {
  try {
    const raw = localStorage.getItem("ewConfig");
    if (!raw) return false;
    cfg = JSON.parse(raw);
    const required = ["workerUrl", "workerSecret", "siteUrl", "rootFolder"];
    if (required.some(k => !cfg[k])) { cfg = null; return false; }
    // Strip leading document library name if someone included it in rootFolder.
    // e.g. "Shared Documents/CandidateWork" or "Documents/CandidateWork" → "CandidateWork"
    cfg.rootFolder = cfg.rootFolder
      .replace(/^shared\s+documents\//i, "")
      .replace(/^documents\//i, "")
      .replace(/^\/+|\/+$/g, ""); // also trim any leading/trailing slashes
    return true;
  } catch { cfg = null; return false; }
}
function isDeviceConfigured() { return !!localStorage.getItem("ewConfig"); }

/* ══ SP STATUS UI ══ */
function setSPStatus(state, label) {
  document.querySelectorAll(".sp-banner").forEach(el => {
    el.setAttribute("data-state", state);
    el.textContent = label;
  });
}
function updateAllSpUI() {
  const configured = !!cfg;
  document.querySelectorAll(".sp-config-status").forEach(el => {
    el.setAttribute("data-state", configured ? "ok" : "missing");
    el.textContent = configured ? `Config: ${cfg.username || cfg.workerUrl}` : "Config: not loaded";
  });
}

/* ══ TOKEN BROKER ══ */
async function acquireTokenROPC() {
  if (!cfg) return false;
  if (!cfg.workerUrl || !cfg.workerSecret) {
    console.error("Missing workerUrl or workerSecret");
    setSPStatus("error", "SP: worker not configured");
    return false;
  }
  try {
    setSPStatus("syncing", "SP: authenticating...");
    const res = await fetch(cfg.workerUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Worker-Secret": cfg.workerSecret },
    });
    const data = await res.json();
    if (!res.ok) {
      console.group("Token broker auth failed");
      console.error("HTTP:", res.status, data);
      console.groupEnd();
      throw new Error(data.error_description || data.error || `HTTP ${res.status}`);
    }
    spAccessToken = data.access_token;
    spTokenExpiry = Date.now() + (data.expires_in - 60) * 1000;
    scheduleTokenRefresh(data.expires_in - 60);
    console.info("SP auth ok:", cfg.username);
    setSPStatus("connected", `SP: ${cfg.username}`);
    updateAllSpUI();
    return true;
  } catch (e) {
    console.error("Worker fetch failed:", e);
    spAccessToken = null;
    setSPStatus("error", "SP: auth failed — see console");
    updateAllSpUI();
    return false;
  }
}
function scheduleTokenRefresh(inSeconds) {
  if (spRefreshTimer) clearTimeout(spRefreshTimer);
  spRefreshTimer = setTimeout(async () => { await acquireTokenROPC(); }, Math.max(inSeconds * 1000, 10_000));
}
async function getSpToken() {
  if (spAccessToken && Date.now() < spTokenExpiry) return spAccessToken;
  const ok = await acquireTokenROPC();
  return ok ? spAccessToken : null;
}

/* ══ DRIVE RESOLUTION ══ */
async function resolveSiteId(token, siteUrl) {
  const url = new URL(siteUrl);
  const endpoint = `https://graph.microsoft.com/v1.0/sites/${url.hostname}:${url.pathname.replace(/\/$/, "")}`;
  const res = await fetch(endpoint, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`resolveSiteId HTTP ${res.status}`);
  return (await res.json()).id;
}
async function resolveDriveId(token, siteId) {
  const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
    { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`resolveDriveId HTTP ${res.status}`);
  const drives = (await res.json()).value || [];
  console.log("Drives:", drives.map(d => d.name));
  const match = drives.find(d => ["documents","shared documents"].includes(d.name.toLowerCase())) || drives[0];
  if (!match) throw new Error("No drives found");
  console.log("Using drive:", match.name, match.id);
  return match.id;
}
async function ensureDriveResolved() {
  const token = await getSpToken();
  if (!token) throw new Error("Could not get SP token");
  let siteId = sessionStorage.getItem("ewSiteId");
  if (!siteId) { siteId = await resolveSiteId(token, cfg.siteUrl); sessionStorage.setItem("ewSiteId", siteId); }
  let driveId = sessionStorage.getItem("ewDriveId");
  if (!driveId) { driveId = await resolveDriveId(token, siteId); sessionStorage.setItem("ewDriveId", driveId); }
  return { token, siteId, driveId };
}

/* ══ SP SYNC ══ */
async function syncToSharePoint(text) {
  if (!cfg) return;
  const candidateId = ($("#candidateId").value || "unknown").replace(/\s+/g, "_");
  const filename = buildSuggestedName();
  const filePath = `${cfg.rootFolder}/${candidateId}/${filename}`;
  try {
    setSPStatus("syncing", "SP: saving...");
    const { token, driveId } = await ensureDriveResolved();
    const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`;
    const res = await fetch(url, {
      method: "PUT",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "text/plain" },
      body: text,
    });
    if (!res.ok) throw new Error(`Upload HTTP ${res.status}: ${await res.text()}`);
    setSPStatus("connected", `SP: saved ${new Date().toLocaleTimeString("en-GB")}`);
    console.info("SP upload ok:", filePath);
  } catch (e) {
    console.error("SP sync failed:", e);
    setSPStatus("error", "SP: save failed — see console");
  }
}

/* ══ CANDIDATE FILE RECOVERY ══ */
async function findCandidateFile(candidateId) {
  const { token, driveId } = await ensureDriveResolved();
  // Sanitise to match how syncToSharePoint saves the folder name
  const safeCandidateId = candidateId.trim().replace(/\s+/g, "_");
  const folder = `${cfg.rootFolder}/${safeCandidateId}`;
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${folder}:/children`;
  console.log("Listing:", folder);
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`findCandidateFile HTTP ${res.status}: ${await res.text()}`);
  const files = ((await res.json()).value || []).filter(f => f.name.endsWith(".txt"));
  if (!files.length) return null;
  files.sort((a, b) => new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime));
  const f = files[0];
  return { name: f.name, modified: f.lastModifiedDateTime, downloadUrl: f["@microsoft.graph.downloadUrl"] };
}
async function loadCandidateFileFromSP(downloadUrl) {
  const res = await fetch(downloadUrl);
  if (!res.ok) throw new Error(`Download HTTP ${res.status}`);
  return await res.text();
}

/* ══ AUTO-DETECT (invigilator tab unlock) ══ */
async function autoDetectCandidateFile() {
  if (!cfg) return;
  const candidateId = $("#candidateId")?.value?.trim();
  if (!candidateId) return;
  const notice = $("#spResumeNotice");
  const btn = $("#spResumeBtn");
  if (!notice) return;
  notice.hidden = true;
  if (btn) btn.onclick = null;
  try {
    const file = await findCandidateFile(candidateId);
    if (!file) return;
    const modTime = new Date(file.modified).toLocaleString("en-GB");
    notice.hidden = false;
    const fnEl = notice.querySelector(".sp-resume-filename");
    const tmEl = notice.querySelector(".sp-resume-time");
    if (fnEl) fnEl.textContent = file.name;
    if (tmEl) tmEl.textContent = modTime;
    if (btn) {
      btn.onclick = async () => {
        try {
          const text = await loadCandidateFileFromSP(file.downloadUrl);
          loadFromText(text);
          fileHandle = await window.showSaveFilePicker({
            suggestedName: file.name,
            types: [{ description: "Exam Document", accept: { "text/plain": [".txt"] } }],
          });
          await writeToFile(buildDocumentText());
          await afterSuccessfulSavePick();
          notice.hidden = true;
        } catch (e) { console.error("Resume failed:", e); alert("Could not load: " + e.message); }
      };
    }
  } catch (e) { console.error("autoDetect failed:", e); }
}

/* ══ MANUAL RECOVERY (admin panel) ══ */
async function recoverCandidateFile(inputId="recoverCandidateId", statusId="recoverStatus", closeDialog=true) {
  const input = $(`#${inputId}`);
  const status = $(`#${statusId}`);
  const candidateId = input?.value?.trim();
  if (!candidateId) { if (status) { status.textContent = "Enter a candidate ID first."; status.setAttribute("data-state","error"); } return; }
  if (!cfg) { if (status) { status.textContent = "SP not configured."; status.setAttribute("data-state","error"); } return; }
  if (status) { status.textContent = "Searching SharePoint…"; status.setAttribute("data-state","searching"); }
  try {
    const file = await findCandidateFile(candidateId);
    if (!file) { if (status) { status.textContent = `No file found for "${candidateId}".`; status.setAttribute("data-state","not-found"); } return; }
    const modTime = new Date(file.modified).toLocaleString("en-GB");
    if (status) { status.textContent = `Found: ${file.name} (${modTime}). Loading…`; status.setAttribute("data-state","found"); }
    const text = await loadCandidateFileFromSP(file.downloadUrl);
    loadFromText(text);
    try {
      fileHandle = await window.showSaveFilePicker({ suggestedName: file.name, types: [{ description: "Exam Document", accept: { "text/plain": [".txt"] } }] });
      await writeToFile(buildDocumentText());
      await afterSuccessfulSavePick();
    } catch {}
    if (status) { status.textContent = `Loaded: ${file.name}`; status.setAttribute("data-state","ok"); }
    if (closeDialog) $("#adminDialog")?.close();
  } catch (e) {
    console.error("recoverCandidateFile failed:", e);
    if (status) { status.textContent = `Error: ${e.message}`; status.setAttribute("data-state","error"); }
  }
}

/* ══ FILE HELPERS ══ */
function buildSuggestedName() {
  const c = ($("#centerNumber").value || "center").replace(/\s+/g,"_");
  const id = ($("#candidateId").value || "candidate").replace(/\s+/g,"_");
  const title = ($("#examTitle").value || "exam").replace(/\s+/g,"_");
  return `${c}-${id}-${title}-${new Date().toISOString().slice(0,10)}.txt`;
}
async function writeToFile(text) {
  if (!fileHandle) return;
  const w = await fileHandle.createWritable(); await w.write(text); await w.close(); setDirty(false);
}
async function writeOPFSBackup(text) {
  if (!navigator.storage?.getDirectory) return;
  const root = await navigator.storage.getDirectory();
  const fh = await root.getFileHandle("autosave-backup.txt",{create:true});
  const w = await fh.createWritable(); await w.write(text); await w.close();
}

/* ══ DOCUMENT TEXT ══ */
function buildDocumentText() {
  return `Center Number: ${$("#centerNumber").value||""}\n`+
    `Candidate ID: ${$("#candidateId").value||""}\n`+
    `Candidate Name: ${$("#candidateName").value||""}\n`+
    `Exam Title: ${$("#examTitle").value||""}\n`+
    `Saved: ${new Date().toLocaleString()}\n---\n\n`+
    $("#editor").value;
}

/* ══ OPEN / SAVE / NEW ══ */
async function openExisting() {
  try {
    const [h] = await window.showOpenFilePicker({types:[{description:"Text",accept:{"text/plain":[".txt"]}}]});
    if (!h) return;
    fileHandle = h;
    loadFromText(await (await h.getFile()).text());
    await afterSuccessfulSavePick();
  } catch {}
}
function loadFromText(text) {
  const sep = /\r?\n---\r?\n\r?\n?/;
  if (sep.test(text)) {
    const [hdr, body=""] = text.split(sep,2);
    const get = l => (hdr.match(new RegExp(`^${l}:\\s*(.*)$`,"mi"))||[])[1]||"";
    $("#centerNumber").value=get("Center Number");
    $("#candidateId").value=get("Candidate ID");
    $("#candidateName").value=get("Candidate Name");
    $("#examTitle").value=get("Exam Title");
    $("#editor").value=body;
  } else { $("#editor").value=text; }
  updateWordCount();
}
async function newDoc() { $("#editor").value=""; setDirty(false); fileHandle=null; updateWordCount(); await showSetupDialog(); }
async function saveAs() {
  try {
    fileHandle = await window.showSaveFilePicker({suggestedName:buildSuggestedName(),types:[{description:"Exam Document",accept:{"text/plain":[".txt"]}}]});
    await writeToFile(buildDocumentText()); await afterSuccessfulSavePick();
  } catch {}
}

/* ══ SETUP DIALOG ══ */
async function showSetupDialog() { const d=$("#setupDialog"); if(d&&!d.open) d.showModal(); }
async function hideSetupDialog() { const d=$("#setupDialog"); if(d?.open) d.close(); }
async function firstSaveFlow() {
  try {
    fileHandle = await window.showSaveFilePicker({suggestedName:buildSuggestedName(),types:[{description:"Exam Document",accept:{"text/plain":[".txt"]}}]});
    await writeToFile(""); await afterSuccessfulSavePick();
  } catch(e) { console.warn("firstSaveFlow failed",e); }
}
async function afterSuccessfulSavePick() {
  await hideSetupDialog(); $("#editor").disabled=false; $("#editor").focus(); if(!autosaveTimer) startAutosave();
}

/* ══ AUTOSAVE ══ */
function startAutosave() {
  if (autosaveTimer) clearInterval(autosaveTimer);
  autosaveTimer = setInterval(async () => {
    if (!$("#editor").value.trim()) return;
    try {
      $("#autosaveStatus").textContent="Autosave: saving…";
      const t=buildDocumentText();
      if(fileHandle) await writeToFile(t); else await writeOPFSBackup(t);
      if(cfg) await syncToSharePoint(t);
      $("#autosaveStatus").textContent="Autosave: up to date";
    } catch(e) { $("#autosaveStatus").textContent="Autosave: error"; console.error(e); }
  }, 30_000);
}

/* ══ PRINT ══ */
const LINES_PER_PAGE = 34;
function escapeHtml(str) { return str.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"); }
function buildPrintPages() {
  const meta = {
    center: $("#centerNumber").value||"", id: $("#candidateId").value||"",
    name: $("#candidateName").value||"", title: $("#examTitle").value||"",
    date: new Date().toLocaleDateString("en-GB"),
  };
  const allLines = $("#editor").value.split("\n");
  const chunks = [];
  for(let i=0;i<allLines.length;i+=LINES_PER_PAGE) chunks.push(allLines.slice(i,i+LINES_PER_PAGE).join("\n"));
  if(!chunks.length) chunks.push("");
  return chunks.map((body,idx) => {
    const div = document.createElement("div");
    div.className = "print-page";
    div.innerHTML = `
      <div class="print-page-header">
        <div class="print-page-ids">
          <div><strong>Center Number:</strong> ${escapeHtml(meta.center)}</div>
          <div><strong>Candidate ID:</strong> ${escapeHtml(meta.id)}</div>
          <div><strong>Candidate Name:</strong> ${escapeHtml(meta.name)}</div>
          <div><strong>Exam Title:</strong> ${escapeHtml(meta.title)}</div>
        </div>
        <div class="print-page-meta">
          <span>${escapeHtml(meta.date)}</span>
          <span>Page ${idx+1} of ${chunks.length}</span>
        </div>
        <hr>
      </div>
      <div class="print-page-body">${escapeHtml(body)}</div>`;
    return div;
  });
}
let _printContainer = null;
function printDoc() {
  if(_printContainer){_printContainer.remove();_printContainer=null;}
  const pages=buildPrintPages();
  _printContainer=document.createElement("div");
  _printContainer.id="printContainer";
  pages.forEach(p=>_printContainer.appendChild(p));
  document.body.appendChild(_printContainer);
  window.print();
}
window.addEventListener("afterprint",()=>{if(_printContainer){_printContainer.remove();_printContainer=null;}});

/* ══ FULLSCREEN ══ */
function toggleFullscreen() {
  if(!document.fullscreenElement) document.documentElement.requestFullscreen({navigationUI:"hide"}).catch(()=>{});
  else document.exitFullscreen().catch(()=>{});
}

/* ══ EDITOR HARDENING ══ */
function hardenEditor() {
  const ed=$("#editor"); const block=e=>{e.preventDefault();e.stopPropagation();};
  ed.addEventListener("paste",block); ed.addEventListener("drop",block); ed.addEventListener("contextmenu",block);
}

/* ══ ADMIN PIN ══ */
async function sha256Hex(str) {
  const buf=await crypto.subtle.digest("SHA-256",new TextEncoder().encode(str));
  return Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,"0")).join("");
}
async function verifyAdminPin(pin) { if(!adminPinHash) return pin==="0000"; return (await sha256Hex(pin))===adminPinHash; }
function openAdminDialog() { $("#adminDialog").showModal(); $("#adminPanel").hidden=true; $("#adminPin").value=""; }

/* ══ INVIGILATOR PIN GATE ══ */
let invigPinHash = localStorage.getItem("invigPinHash") || null;
async function verifyInvigPin(pin) { if(!invigPinHash) return pin==="1234"; return (await sha256Hex(pin))===invigPinHash; }
function lockInvigTab() {
  const gate=$("#invigPinGate"); const ctrl=$("#invigControls");
  if(gate) gate.hidden=false; if(ctrl) ctrl.hidden=true;
  const notice=$("#spResumeNotice"); if(notice) notice.hidden=true;
}
async function tryUnlockInvigTab(pin) {
  const ok=await verifyInvigPin(pin);
  if(!ok){alert("Incorrect invigilator PIN");return;}
  const gate=$("#invigPinGate"); const ctrl=$("#invigControls");
  if(gate) gate.hidden=true; if(ctrl) ctrl.hidden=false;
  await autoDetectCandidateFile();
}

/* ══ SETUP MODAL ══ */
function initSetupModal() {
  // Tabs
  document.querySelectorAll(".setup-tab-btn").forEach(btn=>{
    btn.addEventListener("click",()=>{
      const t=btn.dataset.tab;
      document.querySelectorAll(".setup-tab-btn").forEach(b=>b.classList.toggle("active",b.dataset.tab===t));
      document.querySelectorAll(".setup-tab-panel").forEach(p=>p.hidden=p.dataset.tab!==t);
    });
  });
  // Filename preview
  ["s_centerNumber","s_candidateId","s_examTitle"].forEach(id=>{
    const el=document.getElementById(id);
    if(el) el.addEventListener("input",()=>{ const p=$("#filenamePreview"); if(p) p.textContent=buildSuggestedName(); });
  });
  // Begin button
  const beginBtn=$("#beginExamBtn");
  if(beginBtn){
    beginBtn.addEventListener("click",async ev=>{
      ev.preventDefault();
      // Sync setup modal fields → main meta bar fields
      const map = [["s_centerNumber","centerNumber"],["s_candidateId","candidateId"],["s_candidateName","candidateName"],["s_examTitle","examTitle"]];
      map.forEach(([src,dst])=>{ const s=document.getElementById(src),d=document.getElementById(dst); if(s&&d) d.value=s.value; });
      if(!$("#centerNumber").value.trim()||!$("#candidateId").value.trim()||!$("#candidateName").value.trim()||!$("#examTitle").value.trim()){
        alert("Please fill in all fields before beginning."); return;
      }
      await firstSaveFlow();
    });
  }
  // Invigilator PIN
  const invigUnlockBtn=$("#invigUnlockBtn");
  if(invigUnlockBtn) invigUnlockBtn.addEventListener("click",async()=>await tryUnlockInvigTab($("#invigPinInput")?.value||""));
  lockInvigTab();
}

/* ══ ADMIN TRIGGERS ══ */
function initAdminTriggers() {
  let taps=0,timer;
  $("#brandHotspot").addEventListener("click",()=>{
    taps++;clearTimeout(timer);timer=setTimeout(()=>{taps=0;},1500);
    if(taps>=5){taps=0;openAdminDialog();}
  });
  document.addEventListener("keydown",e=>{
    if(e.ctrlKey&&e.shiftKey&&e.key.toLowerCase()==="e"){e.preventDefault();openAdminDialog();}
    if(e.ctrlKey&&e.altKey&&e.key.toLowerCase()==="k"){e.preventDefault();openAdminDialog();}
  });
  $("#adminCancel").addEventListener("click",()=>$("#adminDialog").close());
  $("#adminUnlock").addEventListener("click",async()=>{
    const ok=await verifyAdminPin($("#adminPin").value);
    if(!ok){alert("Incorrect PIN");return;}
    $("#adminPanel").hidden=false;
  });
  $("#forceSaveBtn").addEventListener("click",async()=>{
    const t=buildDocumentText();
    if(fileHandle) await writeToFile(t); else await writeOPFSBackup(t);
    if(cfg) await syncToSharePoint(t);
  });
  const recoverBtn=$("#recoverBtn");
  if(recoverBtn){
    recoverBtn.addEventListener("click",async()=>{
      const status=$("#recoverStatus");
      if(status){status.textContent="";status.removeAttribute("data-state");}
      await recoverCandidateFile("recoverCandidateId","recoverStatus",true);
    });
  }
  const invigRecoverBtn=$("#invigRecoverBtn");
  if(invigRecoverBtn){
    invigRecoverBtn.addEventListener("click",async()=>{
      const status=$("#invigRecoverStatus");
      if(status){status.textContent="";status.removeAttribute("data-state");}
      await recoverCandidateFile("invigRecoverCandidateId","invigRecoverStatus",false);
    });
  }
  const nextBtn=$("#nextCandidateBtn");
  if(nextBtn){
    nextBtn.addEventListener("click",async()=>{
      if(!confirm("Start next candidate? Current session will be cleared.")) return;
      $("#editor").value="";$("#candidateId").value="";$("#candidateName").value="";
      fileHandle=null;setDirty(false);updateWordCount();
      $("#adminDialog").close();await showSetupDialog();
    });
  }
  $("#exitAppBtn").addEventListener("click",()=>{try{window.close();}catch{}});
}

/* ══ DEVICE SETUP MODAL ══ */
function initDeviceSetup() {
  const STORAGE_KEY = "ewConfig";
  const REQUIRED = ["clientId","tenantId","username","password","siteUrl","rootFolder","workerUrl","workerSecret"];
  let parsedConfig = null;

  const dialog   = $("#deviceSetupDialog");
  const openBtn  = $("#openDeviceSetupBtn");
  const closeBtn = $("#dsCloseBtn");
  const dropzone = $("#dsDropzone");
  const fileInput= $("#dsFileInput");
  const preview  = $("#dsConfigPreview");
  const saveBtn  = $("#dsSaveBtn");
  const result   = $("#dsResult");
  const clearBtn = $("#dsClearBtn");
  const banner   = $("#dsConfiguredBanner");

  function dsStepState(id, state) {
    const el = document.getElementById(id);
    el.className = "ds-step" + (state ? " " + state : "");
  }

  function resetDialog() {
    parsedConfig = null;
    dsStepState("dsStep1", "active");
    dsStepState("dsStep2", "");
    dropzone.innerHTML = "Click to select <code>examwriter-config.json</code> from USB";
    dropzone.classList.remove("loaded","dragover");
    preview.innerHTML = "";
    preview.classList.remove("visible");
    saveBtn.disabled = true;
    saveBtn.textContent = "Save configuration";
    result.className = "ds-result";
    result.innerHTML = "";
    fileInput.value = "";
    banner.hidden = !localStorage.getItem(STORAGE_KEY);
  }

  openBtn?.addEventListener("click", () => { resetDialog(); dialog.showModal(); });
  closeBtn?.addEventListener("click", () => dialog.close());

  dropzone.addEventListener("click",   () => fileInput.click());
  dropzone.addEventListener("keydown", e => { if (e.key==="Enter"||e.key===" ") fileInput.click(); });
  dropzone.addEventListener("dragover",  e => { e.preventDefault(); dropzone.classList.add("dragover"); });
  dropzone.addEventListener("dragleave", () => dropzone.classList.remove("dragover"));
  dropzone.addEventListener("drop", e => {
    e.preventDefault(); dropzone.classList.remove("dragover");
    if (e.dataTransfer.files[0]) readFile(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener("change", () => { if (fileInput.files[0]) readFile(fileInput.files[0]); });

  function readFile(file) {
    const r = new FileReader();
    r.onload = e => { try { processConfig(JSON.parse(e.target.result)); } catch(err) { showParseError(err.message); } };
    r.readAsText(file);
  }

  function processConfig(c) {
    const missing = REQUIRED.filter(k => !c[k]);
    if (missing.length) {
      preview.innerHTML = `<span class="bad">✕ Missing required fields: ${missing.join(", ")}</span>`;
      preview.classList.add("visible");
      dsStepState("dsStep1","error");
      parsedConfig = null; saveBtn.disabled = true;
      return;
    }
    let host = c.siteUrl;
    try { host = new URL(c.siteUrl).hostname; } catch {}
    preview.innerHTML =
      `<span class="ok">✓ clientId</span>   ${c.clientId.slice(0,8)}…\n` +
      `<span class="ok">✓ tenantId</span>   ${c.tenantId.slice(0,8)}…\n` +
      `<span class="ok">✓ username</span>   ${c.username}\n` +
      `<span class="ok">✓ password</span>   ${"•".repeat(12)}\n` +
      `<span class="ok">✓ siteUrl</span>    ${host}\n` +
      `<span class="ok">✓ rootFolder</span> ${c.rootFolder}`;
    preview.classList.add("visible");
    dropzone.innerHTML = `✓ ${c.username} → ${host}`;
    dropzone.classList.add("loaded");
    dsStepState("dsStep1","done");
    dsStepState("dsStep2","active");
    parsedConfig = c; saveBtn.disabled = false;
  }

  function showParseError(msg) {
    preview.innerHTML = `<span class="bad">✕ Could not read file: ${msg}</span>`;
    preview.classList.add("visible");
    dsStepState("dsStep1","error");
    parsedConfig = null; saveBtn.disabled = true;
  }

  saveBtn.addEventListener("click", () => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(parsedConfig));
      dsStepState("dsStep2","done");
      result.className = "ds-result success";
      result.innerHTML = "✓ Configuration saved.<br><strong>Remove the USB drive.</strong>";
      saveBtn.textContent = "Saved ✓"; saveBtn.disabled = true;
      banner.hidden = false;
      if (loadConfig()) { setSPStatus("syncing","SP: connecting..."); acquireTokenROPC(); }
      updateAllSpUI();
    } catch(e) {
      result.className = "ds-result error";
      result.textContent = `Save failed: ${e.message}`;
    }
  });

  clearBtn.addEventListener("click", () => {
    if (!confirm("Clear the stored configuration from this device?")) return;
    localStorage.removeItem(STORAGE_KEY);
    cfg = null;
    setSPStatus("idle","SP: not configured");
    updateAllSpUI();
    resetDialog();
  });
}

/* ══ SHORTCUTS ══ */
function bindShortcuts() {
  document.addEventListener("keydown",e=>{
    if(e.ctrlKey&&e.key.toLowerCase()==="n"){e.preventDefault();newDoc();}
    if(e.ctrlKey&&e.key.toLowerCase()==="o"){e.preventDefault();openExisting();}
    if(e.ctrlKey&&e.shiftKey&&e.key.toLowerCase()==="s"){e.preventDefault();saveAs();}
  });
}

/* ══ DEBUG (remove before deploy) ══ */
window.debugListDrive=async(path="")=>{
  try{
    const{token,driveId}=await ensureDriveResolved();
    const url=path
      ?`https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${path}:/children`
      :`https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`;
    const res=await fetch(url,{headers:{Authorization:`Bearer ${token}`}});
    const data=await res.json();
    console.table((data.value||[]).map(f=>({name:f.name,type:f.folder?"folder":"file",modified:f.lastModifiedDateTime})));
  }catch(e){console.error(e);}
};

/* ══ INIT ══ */
window.addEventListener("DOMContentLoaded",async()=>{
  registerSW();ensurePersistence();initTheme();

  $("#newBtn").addEventListener("click",newDoc);
  $("#openBtn").addEventListener("click",openExisting);
  $("#saveAsBtn").addEventListener("click",saveAs);
  $("#printBtn").addEventListener("click",printDoc);
  $("#fullscreenBtn").addEventListener("click",toggleFullscreen);
  $("#themeBtn").addEventListener("click",toggleTheme);
  $("#syncBtn")?.addEventListener("click",async()=>{const t=buildDocumentText();if(fileHandle)await writeToFile(t);if(cfg)await syncToSharePoint(t);});

  bindShortcuts();
  hardenEditor();
  $("#editor").addEventListener("input",()=>{setDirty(true);updateWordCount();});
  updateWordCount();

  initSetupModal();
  initAdminTriggers();
  initDeviceSetup();

  const configured=loadConfig();
  if(configured){setSPStatus("syncing","SP: connecting...");await acquireTokenROPC();}
  else{setSPStatus("idle",isDeviceConfigured()?"SP: config error":"SP: not configured");}
  updateAllSpUI();

  await showSetupDialog();

  window.addEventListener("beforeunload",e=>{if(dirty){e.preventDefault();e.returnValue="";}});
  startAutosave();
});

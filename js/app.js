/**
 * app.js — Attendance app behaviour
 * ---------------------------------------------------------------
 * Depends on: js/config.js (must load first)
 *
 * Sections inside this file:
 *  1. Shared state + date helpers (sheet format: "11 August 2026")
 *  2. API helpers (Apps Script GET calls + local roster cache)
 *  3. Session / banner / toast
 *  4. Tabs + Register form
 *  5. Find Me (local filter — fast after first list load)
 *  6. Moderator PIN (keyboard ONLY while Moderator tab + PIN gate open)
 *  7. Moderator sheet table (edit name/company/control no/date/size)
 *  8. Boot
 *
 * IMPORTANT FIX: PIN digit capture must NOT run on Register / Find Me,
 * otherwise typing a control number steals keys for the PIN.
 */


let pinBuffer = "";
let pinUnlocked = false;
let sessionConfig = { date: "", dateWords: "", location: "Sasol Club", facilitator: "" };
let todayNumeric = "";
let todayWords = "";
let cachedRecords = [];
let rosterCache = []; // fast local Find Me (no Apps Script search round-trip)
let rosterLoadedAt = 0;
let modFilter = "all";
let modDateFilter = ""; // "" = all days; otherwise "11 August 2026"
let rosterPromise = null;
let syncTimer = null;

/** Control number / ID as plain text (keep leading zeros) */
function controlKey(emp){
  return String(emp == null ? "" : emp).replace(/^\u200B+/, "").trim();
}

function duplicateControlCounts(list){
  const counts = {};
  (list || []).forEach(r => {
    const k = controlKey(r.employeeNumber);
    if (!k) return;
    counts[k] = (counts[k] || 0) + 1;
  });
  return counts;
}

function renderDupBanner(list){
  const el = document.getElementById("dupBanner");
  if (!el) return;
  const counts = duplicateControlCounts(list);
  const dups = Object.keys(counts).filter(k => counts[k] > 1);
  if (!dups.length) {
    el.style.display = "none";
    el.innerHTML = "";
    return;
  }
  el.style.display = "block";
  el.innerHTML = `<b>Duplicate check:</b> ${dups.length} control number/ID value${dups.length === 1 ? "" : "s"} appear more than once `
    + `(${dups.slice(0, 6).map(d => `<b>${esc(d)}</b>×${counts[d]}`).join(", ")}${dups.length > 6 ? ", …" : ""}). `
    + `Flagged in yellow below — <b>not deleted</b>. Fix on the sheet or edit/delete the extra row.`;
}

/** Re-pull from spreadsheet so UI matches sheet (after delete/save/sheet edits) */
async function reloadFromSheet(opts){
  const silent = !!(opts && opts.silent);
  try {
    rosterLoadedAt = 0;
    const list = await ensureRoster(true);
    cachedRecords = list.slice();
    if (pinUnlocked) {
      fillModDateFilter();
      renderModList();
    }
    return list;
  } catch (e) {
    if (!silent) toast("Could not refresh from sheet");
    return cachedRecords;
  }
}

function startSheetSync(){
  stopSheetSync();
  syncTimer = setInterval(() => {
    if (!pinUnlocked) return;
    if (document.hidden) return;
    reloadFromSheet({ silent: true }).catch(() => {});
  }, SYNC_POLL_MS);
}

function stopSheetSync(){
  if (syncTimer) { clearInterval(syncTimer); syncTimer = null; }
}

/** Sort key for training date (newest first). Unknown/blank dates go last. */
function dateSortKey(value){
  const words = toSheetDate(value);
  const m = String(words || "").match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);
  if (!m) return 0;
  const mo = MONTHS.findIndex(x => x.toLowerCase() === m[2].toLowerCase());
  if (mo < 0) return 0;
  return Date.UTC(Number(m[3]), mo, Number(m[1]));
}

/** Unique training dates from roster, newest first */
function uniqueTrainingDates(list){
  const seen = new Set();
  const dates = [];
  (list || []).forEach(r => {
    const d = toSheetDate(r.date);
    if (!d || seen.has(d)) return;
    seen.add(d);
    dates.push(d);
  });
  dates.sort((a, b) => dateSortKey(b) - dateSortKey(a));
  return dates;
}

function fillModDateFilter(){
  const sel = document.getElementById("modDateFilter");
  if (!sel) return;
  const dates = uniqueTrainingDates(cachedRecords);
  const prev = modDateFilter;
  sel.innerHTML = `<option value="">All days</option>` +
    dates.map(d => `<option value="${esc(d)}">${esc(d)}</option>`).join("");
  if (prev && dates.includes(prev)) {
    sel.value = prev;
    modDateFilter = prev;
  } else {
    sel.value = "";
    modDateFilter = "";
  }
  updateModDateHint();
}

function updateModDateHint(){
  const hint = document.getElementById("modDateHint");
  if (!hint) return;
  if (!modDateFilter) {
    hint.textContent = "Showing all days (newest first). Pick a day to narrow the list.";
  } else {
    const n = filteredRecords().length;
    hint.textContent = `Showing ${n} attendee${n === 1 ? "" : "s"} for ${modDateFilter}. Search only looks inside this day.`;
  }
}

function pad2(n){ return String(n).padStart(2,"0"); }

function computeToday(){
  const now = new Date();
  const d = now.getDate();
  const m = now.getMonth();
  const y = now.getFullYear();
  todayNumeric = `${pad2(d)}/${pad2(m+1)}/${y}`;
  todayWords = `${d} ${MONTHS[m]} ${y}`;
  return { numeric: todayNumeric, words: todayWords };
}

function wordsFromNumeric(numeric){
  const m = String(numeric || "").trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return "";
  const d = Number(m[1]);
  const mo = Number(m[2]);
  const y = Number(m[3]);
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return "";
  return `${d} ${MONTHS[mo - 1]} ${y}`;
}

/** Sheet date = "11 August 2026" */
function toSheetDate(value){
  const s = String(value || "").trim();
  if (!s) return "";
  // Already words: 11 August 2026
  const words = s.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);
  if (words) {
    const mo = MONTHS.findIndex(m => m.toLowerCase() === words[2].toLowerCase());
    if (mo >= 0) return `${Number(words[1])} ${MONTHS[mo]} ${words[3]}`;
  }
  // From DD/MM/YYYY or combined "11/08/2026 (11 August 2026)"
  const num = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (num) {
    const d = Number(num[1]);
    const mo = Number(num[2]);
    const y = num[3];
    if (mo >= 1 && mo <= 12) return `${d} ${MONTHS[mo - 1]} ${y}`;
  }
  // Extract trailing words in parentheses
  const paren = s.match(/\(([^)]+)\)\s*$/);
  if (paren) return toSheetDate(paren[1]);
  return s;
}

function sessionSheetDate(){
  return toSheetDate(sessionConfig.dateWords || sessionConfig.date || todayWords) || todayWords;
}

/** True when sheet cell is already "11 August 2026" (not 11/08/2026 (...)) */
function isWordsDate(value){
  return /^(\d{1,2})\s+[A-Za-z]+\s+(\d{4})$/.test(String(value || "").trim());
}

function needsDateFix(value){
  const s = String(value || "").trim();
  if (!s) return true; // blank
  return !isWordsDate(s);
}

function loadLocalSession(){
  try {
    const raw = localStorage.getItem(SESSION_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) { return null; }
}

function saveLocalSession(){
  localStorage.setItem(SESSION_KEY, JSON.stringify({
    date: sessionConfig.date,
    dateWords: sessionConfig.dateWords,
    location: sessionConfig.location,
    facilitator: sessionConfig.facilitator
  }));
}

function syncDateUI(){
  const sheetDate = sessionSheetDate();
  document.getElementById("dateWords").textContent = sheetDate || "â€”";
  document.getElementById("dateNumeric").textContent = todayNumeric;
  document.getElementById("cfgDate").value = todayNumeric;
  document.getElementById("cfgDateWords").value = sheetDate || "";
  document.getElementById("cfgDateHint").textContent =
    `Sheet date format: ${sheetDate || todayWords}`;
}

function applyTodayDate(){
  computeToday();
  sessionConfig.date = todayNumeric;
  sessionConfig.dateWords = todayWords;
  syncDateUI();
  saveLocalSession();
}

function readDateFromInputs(){
  let words = document.getElementById("cfgDateWords").value.trim();
  words = toSheetDate(words) || todayWords;
  sessionConfig.dateWords = words;
  sessionConfig.date = todayNumeric;
  syncDateUI();
}

function toast(msg){
  const t = document.getElementById("toast");
  document.getElementById("toastMsg").textContent = msg;
  t.classList.add("show");
  setTimeout(() => t.classList.remove("show"), 2400);
}

function esc(s){
  return (s||"").toString().replace(/[&<>"']/g, c => ({
    "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"
  }[c]));
}

function hasSize(r){
  const s = (r.size || "").toString().trim().toUpperCase();
  return s === "S" || s === "M" || s === "L";
}

function sizeBadge(size){
  if (hasSize({ size })) return `<span class="size-tag">${esc(size)}</span>`;
  return `<span class="size-tag pending">Pending</span>`;
}

function setApiStatus(ok, msg){
  const el = document.getElementById("apiStatus");
  el.className = "status-bar show " + (ok ? "ok" : "bad");
  el.textContent = msg;
}

/** GET-based API â€” reliable with GitHub Pages + Apps Script */
async function apiCall(params){
  const url = API_URL + "?" + new URLSearchParams(params).toString();
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), 25000);
  try {
    const res = await fetch(url, { method: "GET", signal: ctrl.signal, cache: "no-store" });
    const text = await res.text();
    let data;
    try { data = JSON.parse(text); }
    catch (e) { throw new Error("Bad response from server. Redeploy Apps Script as Web App (Anyone)."); }
    return data;
  } finally {
    clearTimeout(timer);
  }
}

/** One list fetch, reused for Find Me + moderator (fast local search) */
async function ensureRoster(force){
  const fresh = Date.now() - rosterLoadedAt < ROSTER_TTL_MS;
  if (!force && rosterCache.length && fresh) return rosterCache;
  if (rosterPromise) return rosterPromise;
  rosterPromise = (async () => {
    try {
      const r = await apiCall({ action: "list" });
      rosterCache = r.success ? (r.data || []) : [];
      rosterLoadedAt = Date.now();
      return rosterCache;
    } finally {
      rosterPromise = null;
    }
  })();
  return rosterPromise;
}

function findByLast4(digits, list){
  const q = String(digits || "").trim();
  return (list || []).filter(r => String(r.employeeNumber || "").slice(-4) === q);
}

document.querySelectorAll(".tab-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
    document.querySelectorAll("main section").forEach(s => s.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById("tab-" + btn.dataset.tab).classList.add("active");
    // Only focus PIN when opening Moderator — never on Register/Find Me
    if (btn.dataset.tab === "mod" && !pinUnlocked) {
      setTimeout(() => {
        const pinType = document.getElementById("pinType");
        const gate = document.getElementById("pinGate");
        if (pinType && gate && gate.style.display !== "none") pinType.focus();
      }, 50);
    }
  });
});

function renderSessionMeta(){
  const loc = sessionConfig.location || "Location not set";
  const fac = sessionConfig.facilitator;
  document.getElementById("sessionMeta").innerHTML =
    `<span class="meta-pill">${esc(loc)}</span>` +
    (fac ? `<span class="meta-pill">${esc(fac)}</span>` : "");
}

async function loadSessionConfig(){
  computeToday();
  const local = loadLocalSession();
  if (local && local.date) {
    sessionConfig.date = local.date;
    sessionConfig.dateWords = local.dateWords || wordsFromNumeric(local.date) || todayWords;
    sessionConfig.location = local.location || "Sasol Club";
    sessionConfig.facilitator = local.facilitator || "";
  } else {
    sessionConfig.date = todayNumeric;
    sessionConfig.dateWords = todayWords;
  }

  try {
    const r = await apiCall({ action: "config" });
    if (r.success) {
      sessionConfig.location = r.data.location || sessionConfig.location || "Sasol Club";
      sessionConfig.facilitator = r.data.facilitator || sessionConfig.facilitator || "";
      // Only fall back to server today if moderator never typed a date
      if (!local || !local.dateWords) {
        sessionConfig.date = todayNumeric;
        sessionConfig.dateWords = r.data.dateWords || todayWords;
      }
      setApiStatus(true, "Connected to register sheet");
    } else {
      setApiStatus(false, "API error: " + (r.error || "unknown"));
    }
  } catch (e) {
    setApiStatus(false, "Cannot reach Apps Script â€” redeploy Web App");
  }

  document.getElementById("cfgLocation").value = sessionConfig.location || "Sasol Club";
  document.getElementById("cfgFacilitator").value = sessionConfig.facilitator || "";
  syncDateUI();
  renderSessionMeta();
}

document.getElementById("useTodayBtn").addEventListener("click", () => {
  applyTodayDate();
  toast("Date set to today");
});

document.getElementById("applyDateBlankBtn").addEventListener("click", async () => {
  readDateFromInputs();
  const sheetDate = sessionSheetDate(); // e.g. 11 August 2026
  // Blank OR old format like 11/08/2026 (11 August 2026)
  const targets = cachedRecords.filter(r => needsDateFix(r.date));
  if (!targets.length) { toast("All dates already look correct"); return; }
  let ok = 0;
  for (const r of targets) {
    const next = String(r.date || "").trim() ? (toSheetDate(r.date) || sheetDate) : sheetDate;
    try {
      const res = await apiCall({
        action: "update",
        employeeNumber: r.employeeNumber,
        date: next
      });
      if (res.success) {
        r.date = next;
        ok++;
      }
    } catch (e) {}
  }
  rosterCache = cachedRecords.slice();
  rosterLoadedAt = Date.now();
  toast(`Set "${sheetDate}" style on ${ok} row(s)`);
  renderModList();
});

document.getElementById("saveCfgBtn").addEventListener("click", async () => {
  readDateFromInputs();
  sessionConfig.location = document.getElementById("cfgLocation").value.trim() || "Sasol Club";
  sessionConfig.facilitator = document.getElementById("cfgFacilitator").value.trim();
  saveLocalSession();
  try {
    await apiCall({
      action: "setConfig",
      location: sessionConfig.location,
      facilitator: sessionConfig.facilitator
    });
    renderSessionMeta();
    toast("Session details saved");
  } catch (e) {
    toast("Saved locally â€” sheet location save failed");
  }
});

document.getElementById("regForm").addEventListener("submit", async (ev) => {
  ev.preventDefault();
  computeToday();

  // Combine initials (uppercase, strip dots/spaces) + surname
  const initialsRaw = document.getElementById("regInitials").value.trim().replace(/[\s.]+/g, "").toUpperCase();
  const surname = document.getElementById("regSurname").value.trim();
  const name = initialsRaw && surname ? `${initialsRaw} ${surname}` : (initialsRaw || surname);

  // Control number / ID kept as text (leading zeros allowed)
  const empRaw = controlKey(document.getElementById("regEmp").value);
  const company = document.getElementById("regCompany").value.trim();
  const resultDiv = document.getElementById("regResult");
  resultDiv.innerHTML = "";

  if (!initialsRaw || !surname) {
    resultDiv.innerHTML = `<div class="result err"><div class="mark">!</div><div><div class="rtitle">Missing name</div><div class="rbody">Enter your initials (e.g. <b>NM</b>) and your surname separately.</div></div></div>`;
    return;
  }
  if (!empRaw || !company) {
    resultDiv.innerHTML = `<div class="result err"><div class="mark">!</div><div><div class="rtitle">Missing information</div><div class="rbody">Fill in control number / ID and company.</div></div></div>`;
    return;
  }

  const submitBtn = document.getElementById("submitBtn");
  submitBtn.disabled = true;
  submitBtn.textContent = "Submitting…";
  resultDiv.innerHTML = `<div class="spinner"></div>`;

  try {
    // Fresh roster check so we catch duplicates even if cache was stale
    try {
      const list = await ensureRoster(true);
      const hit = list.find(r => controlKey(r.employeeNumber) === empRaw);
      if (hit) {
        const sized = hasSize(hit) ? `Size on file: <b>${esc(hit.size)}</b>.` : `Size still pending assignment.`;
        resultDiv.innerHTML = `<div class="result warn"><div class="mark">!</div><div><div class="rtitle">Already on the register</div><div class="rbody">Control number / ID <b>${esc(empRaw)}</b> belongs to <b>${esc(hit.name)}</b>. ${sized}<br>Duplicates are not allowed.</div></div></div>`;
        submitBtn.disabled = false;
        submitBtn.textContent = "Sign Me In";
        return;
      }
    } catch (e) { /* server will still enforce */ }

    const r = await apiCall({
      action: "register",
      name,
      employeeNumber: empRaw,
      company,
      size: ""
    });
    if (!r.success) {
      resultDiv.innerHTML = `<div class="result err"><div class="mark">!</div><div><div class="rtitle">Something went wrong</div><div class="rbody">${esc(r.error||"Please try again.")}</div></div></div>`;
    } else if (r.duplicate) {
      const sized = hasSize(r.data) ? `Size on file: <b>${esc(r.data.size)}</b>.` : `Size still pending assignment.`;
      resultDiv.innerHTML = `<div class="result warn"><div class="mark">!</div><div><div class="rtitle">Already signed in</div><div class="rbody">Control number / ID <b>${esc(empRaw)}</b> is already used by <b>${esc(r.data.name)}</b>. ${sized}</div></div></div>`;
    } else {
      const stamped = sessionSheetDate() || todayWords;
      try {
        await apiCall({ action: "update", employeeNumber: empRaw, date: stamped, row: r.data && r.data.row ? String(r.data.row) : "" });
      } catch (e) {}
      resultDiv.innerHTML = `<div class="result ok"><div class="mark">✓</div><div><div class="rtitle">You're signed in</div><div class="rbody"><b>${esc(name)}</b> · ${esc(company)}<br>Date: <b>${esc(stamped)}</b><br>A moderator will assign your respirator size.</div></div></div>`;
      document.getElementById("regInitials").value = "";
      document.getElementById("regSurname").value = "";
      document.getElementById("regEmp").value = "";
      document.getElementById("regCompany").value = "";
      await reloadFromSheet({ silent: true });
      toast("Saved to spreadsheet");
    }
  } catch (e) {
    resultDiv.innerHTML = `<div class="result err"><div class="mark">!</div><div><div class="rtitle">Connection error</div><div class="rbody">${esc(e.message || "Check internet / redeploy Apps Script.")}</div></div></div>`;
  }
  submitBtn.disabled = false;
  submitBtn.textContent = "Sign Me In";
});

async function doFind(){
  const digits = document.getElementById("findInput").value.trim();
  const out = document.getElementById("findResults");
  if (digits.length !== 4) {
    out.innerHTML = `<div class="empty">Enter exactly 4 digits</div>`;
    return;
  }
  out.innerHTML = '<div class="spinner"></div>';
  try {
    // Always refresh so deletes / sheet edits show without a hard page reload
    const list = await ensureRoster(true);
    const records = findByLast4(digits, list);
    if (records.length === 0) {
      out.innerHTML = `<div class="empty">No match found. Use Register if you haven't signed in yet.</div>`;
      return;
    }
    out.innerHTML = records.map(rr => `
      <div class="entry">
        <div>
          <div class="name">${esc(rr.name)}</div>
          <div class="meta">${esc(rr.company)} · Control ••${esc(String(rr.employeeNumber).slice(-4))}<br>${esc(toSheetDate(rr.date) || "Date pending")}</div>
        </div>
        ${sizeBadge(rr.size)}
      </div>
    `).join("");
  } catch (e) {
    out.innerHTML = `<div class="empty">Connection error — try again.</div>`;
  }
}
document.getElementById("findBtn").addEventListener("click", doFind);
document.getElementById("findInput").addEventListener("input", () => {
  const digits = document.getElementById("findInput").value.trim();
  if (digits.length === 4) doFind();
});
document.getElementById("findInput").addEventListener("keydown", e => { if (e.key === "Enter") doFind(); });

function renderPinDots(){
  document.querySelectorAll("#pinDots .dot").forEach((d, i) => {
    d.classList.toggle("filled", i < pinBuffer.length);
  });
  const pinType = document.getElementById("pinType");
  if (pinType && document.activeElement !== pinType) pinType.value = pinBuffer;
}

function tryUnlockPin(){
  if (pinBuffer.length !== 4 || pinUnlocked) return;
  if (pinBuffer === MOD_PIN) {
    pinUnlocked = true;
    document.getElementById("pinGate").style.display = "none";
    document.getElementById("modDashboard").style.display = "block";
    startSheetSync();
    loadModDashboard();
  } else {
    toast("Incorrect PIN");
    pinBuffer = "";
    renderPinDots();
  }
}

function handlePinKey(k){
  if (pinUnlocked) return;
  if (k === "clear") pinBuffer = "";
  else if (k === "back") pinBuffer = pinBuffer.slice(0, -1);
  else if (/^\d$/.test(k) && pinBuffer.length < 4) pinBuffer += k;
  else return;
  renderPinDots();
  if (pinBuffer.length === 4) setTimeout(tryUnlockPin, 80);
}

document.getElementById("pinPad").addEventListener("click", e => {
  const btn = e.target.closest("button");
  if (!btn) return;
  handlePinKey(btn.dataset.k);
});

document.getElementById("pinType").addEventListener("input", e => {
  const digits = e.target.value.replace(/\D/g, "").slice(0, 4);
  pinBuffer = digits;
  e.target.value = digits;
  renderPinDots();
  if (pinBuffer.length === 4) setTimeout(tryUnlockPin, 80);
});

/**
 * PIN keyboard capture — ONLY when Moderator tab is active AND PIN gate is showing.
 * Never steal digits from Register / Find Me (control number typing).
 */
document.addEventListener("keydown", e => {
  if (pinUnlocked) return;

  const modTab = document.getElementById("tab-mod");
  if (!modTab || !modTab.classList.contains("active")) return;

  const gate = document.getElementById("pinGate");
  if (!gate || gate.style.display === "none") return;

  const ae = document.activeElement;
  const typingElsewhere =
    ae &&
    ae !== document.getElementById("pinType") &&
    (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA" || ae.tagName === "SELECT" || ae.isContentEditable);
  if (typingElsewhere) return;

  if (e.key >= "0" && e.key <= "9") { e.preventDefault(); handlePinKey(e.key); }
  else if (e.key === "Backspace") { e.preventDefault(); handlePinKey("back"); }
  else if (e.key === "Escape" || e.key === "Delete") { e.preventDefault(); handlePinKey("clear"); }
});

function filteredRecords(){
  const q = document.getElementById("modSearch").value.trim().toLowerCase();
  const day = modDateFilter;
  const list = cachedRecords.filter(r => {
    if (day) {
      const rd = toSheetDate(r.date);
      if (rd !== day) return false;
    }
    if (modFilter === "pending" && hasSize(r)) return false;
    if (modFilter === "sized" && !hasSize(r)) return false;
    if (!q) return true;
    const name = (r.name || "").toLowerCase();
    const emp = String(r.employeeNumber || "");
    const company = (r.company || "").toLowerCase();
    return name.includes(q) || company.includes(q) || emp.includes(q) || emp.slice(-4) === q;
  });
  // Sort by training date (newest first), then name
  list.sort((a, b) => {
    const dk = dateSortKey(b.date) - dateSortKey(a.date);
    if (dk !== 0) return dk;
    return String(a.name || "").localeCompare(String(b.name || ""), undefined, { sensitivity: "base" });
  });
  return list;
}

function renderModList(){
  const listEl = document.getElementById("modList");
  const records = filteredRecords();
  // Stats reflect the current day filter (All days = full roster)
  const scope = modDateFilter
    ? cachedRecords.filter(r => toSheetDate(r.date) === modDateFilter)
    : cachedRecords;
  const pending = scope.filter(r => !hasSize(r)).length;
  const sized = scope.length - pending;
  document.getElementById("statTotal").textContent = scope.length;
  document.getElementById("statPending").textContent = pending;
  document.getElementById("statSized").textContent = sized;
  updateModDateHint();
  renderDupBanner(cachedRecords);

  const dupCounts = duplicateControlCounts(cachedRecords);

  if (records.length === 0) {
    listEl.innerHTML = `<div class="empty">${modDateFilter ? "No attendees for this day" : "No matching attendees"}</div>`;
    return;
  }

  listEl.innerHTML = `
  <table class="sheet-table">
    <thead>
      <tr>
        <th class="col-name">Attendee Name and Surname</th>
        <th class="col-emp">Control No / ID</th>
        <th class="col-co">Company</th>
        <th class="col-date">Date of Training</th>
        <th class="col-size">Size</th>
        <th class="col-act">Actions</th>
      </tr>
    </thead>
    <tbody>
      ${records.map(r => {
        const emp = controlKey(r.employeeNumber);
        const size = (r.size || "").toString().toUpperCase();
        const isDup = (dupCounts[emp] || 0) > 1;
        const rowId = r.row || "";
        return `
        <tr data-emp="${esc(emp)}" data-row="${esc(String(rowId))}" class="${isDup ? "dup-row" : ""}">
          <td class="col-name"><input type="text" class="edit-name" value="${esc(r.name)}">${isDup ? `<span class="dup-tag">Duplicate</span>` : ""}</td>
          <td class="col-emp"><input type="text" class="edit-emp" value="${esc(emp)}" inputmode="text" autocomplete="off"></td>
          <td class="col-co"><input type="text" class="edit-company" value="${esc(r.company)}"></td>
          <td class="col-date"><input type="text" class="edit-date" value="${esc(toSheetDate(r.date) || "")}" placeholder="${esc(sessionSheetDate())}"></td>
          <td class="col-size">
            <select class="edit-size">
              <option value="" ${!size ? "selected" : ""}>—</option>
              <option value="S" ${size==="S"?"selected":""}>S</option>
              <option value="M" ${size==="M"?"selected":""}>M</option>
              <option value="L" ${size==="L"?"selected":""}>L</option>
            </select>
          </td>
          <td class="col-act">
            <div class="sheet-actions">
              <button type="button" class="save-row">Save</button>
              <button type="button" class="del-row">Delete</button>
            </div>
          </td>
        </tr>`;
      }).join("")}
    </tbody>
  </table>`;

  listEl.querySelectorAll("tr[data-emp]").forEach(rowEl => {
    const emp = controlKey(rowEl.dataset.emp);
    const sheetRow = rowEl.dataset.row || "";

    rowEl.querySelector(".edit-size").addEventListener("change", async (e) => {
      const size = e.target.value;
      if (!size) return;
      try {
        const payload = { action: "update", employeeNumber: emp, size };
        if (sheetRow) payload.row = sheetRow;
        const r = await apiCall(payload);
        if (!r.success) { toast(r.error || "Update failed"); return; }
        toast("Size saved");
        await reloadFromSheet({ silent: true });
      } catch (err) { toast("Connection error"); }
    });

    rowEl.querySelector(".save-row").addEventListener("click", async () => {
      const name = rowEl.querySelector(".edit-name").value.trim();
      const company = rowEl.querySelector(".edit-company").value.trim();
      const newEmp = controlKey(rowEl.querySelector(".edit-emp").value);
      const date = toSheetDate(rowEl.querySelector(".edit-date").value.trim());
      const size = rowEl.querySelector(".edit-size").value;
      if (!name || !company) { toast("Name and company required"); return; }
      if (!newEmp) { toast("Control number / ID required"); return; }
      if (newEmp !== emp) {
        const clash = cachedRecords.find(x => controlKey(x.employeeNumber) === newEmp && String(x.row || "") !== String(sheetRow));
        if (clash) {
          toast("That control number / ID already exists — duplicate not allowed");
          return;
        }
      }
      try {
        const payload = { action: "update", employeeNumber: emp, name, company };
        if (sheetRow) payload.row = sheetRow;
        if (newEmp !== emp) payload.newEmployeeNumber = newEmp;
        if (date) payload.date = date;
        if (size) payload.size = size;
        const r = await apiCall(payload);
        if (!r.success) { toast(r.error || "Update failed"); return; }
        toast("Row saved to sheet");
        await reloadFromSheet({ silent: true });
      } catch (err) { toast("Connection error"); }
    });

    rowEl.querySelector(".del-row").addEventListener("click", async () => {
      if (!confirm("Remove this attendee from the spreadsheet?")) return;
      try {
        const payload = { action: "delete", employeeNumber: emp };
        if (sheetRow) payload.row = sheetRow;
        const r = await apiCall(payload);
        if (!r.success) { toast(r.error || "Delete failed"); return; }
        toast("Deleted from sheet");
        // Always re-pull so list / row numbers / duplicates match the spreadsheet
        await reloadFromSheet({ silent: true });
      } catch (err) { toast("Connection error"); }
    });
  });
}

async function loadModDashboard(){
  const listEl = document.getElementById("modList");
  listEl.innerHTML = '<div class="spinner"></div>';
  try {
    await reloadFromSheet({ silent: true });
  } catch (e) {
    cachedRecords = [];
    toast("Could not load sheet data");
    fillModDateFilter();
    renderModList();
  }
}

document.getElementById("refreshBtn").addEventListener("click", loadModDashboard);
document.getElementById("modSearch").addEventListener("input", renderModList);
document.getElementById("modDateFilter").addEventListener("change", () => {
  modDateFilter = document.getElementById("modDateFilter").value;
  renderModList();
});
document.getElementById("modDateTodayBtn").addEventListener("click", () => {
  computeToday();
  const today = todayWords;
  const sel = document.getElementById("modDateFilter");
  // Prefer today's label if it exists in the list; otherwise still filter to today
  if (![...sel.options].some(o => o.value === today)) {
    const opt = document.createElement("option");
    opt.value = today;
    opt.textContent = today + " (no rows yet)";
    sel.appendChild(opt);
  }
  sel.value = today;
  modDateFilter = today;
  renderModList();
});
document.getElementById("modDateClearBtn").addEventListener("click", () => {
  document.getElementById("modDateFilter").value = "";
  modDateFilter = "";
  renderModList();
});
document.querySelectorAll(".filt").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".filt").forEach(b => b.classList.remove("on"));
    btn.classList.add("on");
    modFilter = btn.dataset.filt;
    renderModList();
  });
});

document.getElementById("exportBtn").addEventListener("click", () => {
  readDateFromInputs();
  const fallback = sessionSheetDate();
  const rows = [["Attendee Name and Surname","Sasol Control Number","Company","Date of Training","Respirator Size Required (S/M/L)"]];
  // Export what the moderator is currently looking at (day + size filters + search)
  filteredRecords().forEach(r => rows.push([r.name, r.employeeNumber, r.company, toSheetDate(r.date) || fallback, r.size || ""]));
  const csv = rows.map(row => row.map(cell => `"${(cell||"").toString().replace(/"/g,'""')}"`).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  const stamp = (modDateFilter || fallback).replace(/\s+/g, "-");
  a.download = `respirator-fitment-attendance-${stamp}.csv`;
  a.click();
  URL.revokeObjectURL(url);
  toast(modDateFilter ? `CSV for ${modDateFilter}` : "CSV downloaded");
});

document.getElementById("genQrBtn").addEventListener("click", () => {
  let url = document.getElementById("linkInput").value.trim() || location.href.split("#")[0];
  document.getElementById("linkInput").value = url;
  const qrImg = document.getElementById("qrImg");
  qrImg.src = `https://api.qrserver.com/v1/create-qr-code/?size=340x340&data=${encodeURIComponent(url)}`;
  qrImg.style.display = "inline-block";
});

computeToday();
loadSessionConfig();
// Prefetch roster so Find Me feels instant (does not affect Register typing)
ensureRoster(false).catch(() => {});

// When you come back to the tab / window, pull latest spreadsheet data
document.addEventListener("visibilitychange", () => {
  if (document.hidden) return;
  reloadFromSheet({ silent: true }).catch(() => {});
});
window.addEventListener("focus", () => {
  reloadFromSheet({ silent: true }).catch(() => {});
});


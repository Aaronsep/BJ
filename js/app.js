import { parseTimeToMinutes, hhmm } from "./time.js";
import { solveOptimalWithFixed } from "./solver.js";
import { createSheetsDb } from "./sheetsDb.js";

const SHEET_URL =
  "https://docs.google.com/spreadsheets/d/123IQNZ6SQ4KcRuWnNgQjdemcdCERkvnz9Qd0Qu3e9tA/gviz/tq?tqx=out:json";

const HISTORY_STORAGE_KEY = "solve-history-v1";
const HISTORY_TTL_MS = 3 * 24 * 60 * 60 * 1000;
const HISTORY_MAX_ITEMS = 150;
const DRAFT_STORAGE_KEY = "solve-draft-v1";
const DRAFT_TTL_MS = 3 * 24 * 60 * 60 * 1000;
const DRAFT_SAVE_DELAY_MS = 450;
const UI_THEME_KEY = "ui-theme-v1";

function formatLocalDateTime(timestamp) {
  if (!timestamp) return "";
  return new Date(timestamp).toLocaleString("en-US", { hour12: false });
}

function makeId(prefix = "id") {
  const random = Math.random().toString(36).slice(2, 8);
  return `${prefix}-${Date.now().toString(36)}-${random}`;
}

function safeReadJson(key) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function safeWriteJson(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // Keep app working even if storage quota is exceeded.
  }
}

export function initApp() {
  const $ = (id) => document.getElementById(id);

  const elList = $("list");
  const elCount = $("countInfo");
  const elErr = $("errBox");
  const elOut = $("out");
  const elStatus = $("status");
  const elDbInfo = $("dbInfo");
  const elHistoryInfo = $("historyInfo");

  const elAddJobForm = $("addJobForm");
  const elNameIn = $("nameIn");
  const elDurIn = $("durIn");
  const elFixIn = $("fixIn");
  const elNameSuggestions = $("nameSuggestions");
  const elK = $("k");
  const elQuant = $("quant");
  const elRefreshDbBtn = $("refreshDbBtn");
  const elHistorySelect = $("historySelect");
  const elLoadHistoryBtn = $("loadHistoryBtn");
  const elDeleteHistoryBtn = $("deleteHistoryBtn");
  const elClearHistoryBtn = $("clearHistoryBtn");
  const elThemeToggleBtn = $("themeToggleBtn");
  const elThemeToggleText = $("themeToggleText");
  const elConfirmModal = $("confirmModal");
  const elConfirmTitle = $("confirmTitle");
  const elConfirmMessage = $("confirmMessage");
  const elConfirmCancelBtn = $("confirmCancelBtn");
  const elConfirmOkBtn = $("confirmOkBtn");

  let jobs = [];
  let jobId = 1;
  let solveHistory = [];
  let editingCheckpointId = null;
  let draftSaveTimer = null;
  let addressCatalog = [];
  let suggestionItems = [];
  let suggestionActiveIndex = -1;
  let latestPlanSnapshot = null;
  let confirmResolver = null;
  let confirmPreviouslyFocused = null;

  const sheetsDb = createSheetsDb({
    spreadsheetUrl: SHEET_URL,
    ttlMs: 10 * 60 * 1000,
    onStateChange: onDbStateChange,
  });

  function getErrorMessage(error, fallback = "Something went wrong. Please try again.") {
    const message = error && error.message ? String(error.message).trim() : "";
    return message || fallback;
  }

  function setStatus(message = "", tone = "neutral") {
    elStatus.classList.remove("statusOk", "statusBad");
    elStatus.textContent = String(message || "");
    if (tone === "ok") elStatus.classList.add("statusOk");
    if (tone === "bad") elStatus.classList.add("statusBad");
  }

  function showError(msg) {
    elErr.style.display = "block";
    elErr.textContent = String(msg || "Something went wrong.");
  }

  function clearError() {
    elErr.style.display = "none";
    elErr.textContent = "";
  }

  function clearPlanOutput() {
    elOut.innerHTML = "";
    delete elOut.dataset.copyText;
  }

  function sanitizePlanSnapshot(raw) {
    if (!raw || typeof raw !== "object") return null;
    const teamCount = Math.max(1, parseInt(raw.teamCount, 10) || 1);
    const jobCount = Math.max(0, parseInt(raw.jobCount, 10) || 0);
    const minMinutes = Math.max(0, parseInt(raw.minMinutes, 10) || 0);
    const maxMinutes = Math.max(0, parseInt(raw.maxMinutes, 10) || 0);
    const gapMinutes = Math.max(0, parseInt(raw.gapMinutes, 10) || 0);
    const teamsRaw = Array.isArray(raw.teams) ? raw.teams : [];

    const teams = teamsRaw
      .map((team, idx) => {
        const totalMinutes = Math.max(0, parseInt(team && team.totalMinutes, 10) || 0);
        const jobsRaw = Array.isArray(team && team.jobs) ? team.jobs : [];
        const jobsClean = jobsRaw
          .map((job) => ({
            name: String((job && job.name) || "").trim(),
            minutes: Math.max(0, parseInt(job && job.minutes, 10) || 0),
          }))
          .filter((job) => job.name && job.minutes > 0);
        return {
          teamIndex: idx + 1,
          totalMinutes,
          jobs: jobsClean,
        };
      })
      .slice(0, teamCount);

    if (!teams.length) return null;

    return {
      teamCount,
      jobCount,
      minMinutes,
      maxMinutes,
      gapMinutes,
      teams,
    };
  }

  function buildPlanSnapshot(result, teamCount, jobCount) {
    const teams = Array.isArray(result.teams) ? result.teams : [];
    const totals = teams.map((team) => Math.max(0, parseInt(team.totalMinutes, 10) || 0));
    const minMinutes = totals.length ? Math.min(...totals) : 0;
    const maxMinutes = totals.length ? Math.max(...totals) : 0;

    return sanitizePlanSnapshot({
      teamCount,
      jobCount,
      minMinutes,
      maxMinutes,
      gapMinutes: maxMinutes - minMinutes,
      teams: teams.map((team, idx) => ({
        teamIndex: idx + 1,
        totalMinutes: team.totalMinutes,
        jobs: (Array.isArray(team.jobs) ? team.jobs : []).map((job) => ({
          name: job.name,
          minutes: job.minutes,
        })),
      })),
    });
  }

  function buildPlanText(snapshot) {
    if (!snapshot) return "";
    let out = "";
    out += `Plan summary\n`;
    out += `Teams: ${snapshot.teamCount}\n`;
    out += `Jobs: ${snapshot.jobCount}\n`;
    out += `Longest team time: ${snapshot.maxMinutes} min (${hhmm(snapshot.maxMinutes)})\n`;
    out += `Shortest team time: ${snapshot.minMinutes} min (${hhmm(snapshot.minMinutes)})\n`;
    out += `Gap: ${snapshot.gapMinutes} min (${hhmm(snapshot.gapMinutes)})\n`;
    out += `\n`;

    snapshot.teams.forEach((team, idx) => {
      out += `Team ${idx + 1} | Total: ${team.totalMinutes} min (${hhmm(team.totalMinutes)})\n`;
      team.jobs.forEach((job) => {
        out += `  - ${job.name} - ${job.minutes} min (${hhmm(job.minutes)})\n`;
      });
      out += "\n";
    });

    return out.trimEnd();
  }

  function renderPlanView(snapshot) {
    if (!snapshot) {
      clearPlanOutput();
      return;
    }

    const planText = buildPlanText(snapshot);
    elOut.innerHTML = "";
    elOut.dataset.copyText = planText;

    const view = document.createElement("div");
    view.className = "planView";

    const summary = document.createElement("section");
    summary.className = "planSummary";

    const stats = document.createElement("div");
    stats.className = "planStats";

    const statItems = [
      { label: "Teams", value: String(snapshot.teamCount) },
      { label: "Jobs", value: String(snapshot.jobCount) },
      { label: "Gap", value: `${snapshot.gapMinutes}m` },
      { label: "Range", value: `${snapshot.minMinutes}-${snapshot.maxMinutes}m` },
    ];

    for (const stat of statItems) {
      const statCard = document.createElement("div");
      statCard.className = "planStat";
      const statLabel = document.createElement("div");
      statLabel.className = "planStatLabel";
      statLabel.textContent = stat.label;
      const statValue = document.createElement("div");
      statValue.className = "planStatValue";
      statValue.textContent = stat.value;
      statCard.appendChild(statLabel);
      statCard.appendChild(statValue);
      stats.appendChild(statCard);
    }

    summary.appendChild(stats);
    view.appendChild(summary);

    const teamsWrap = document.createElement("section");
    teamsWrap.className = "planTeams";

    snapshot.teams.forEach((team, idx) => {
      const teamCard = document.createElement("article");
      teamCard.className = "planTeam";

      const teamHead = document.createElement("div");
      teamHead.className = "planTeamHead";

      const teamName = document.createElement("div");
      teamName.className = "planTeamName";
      teamName.textContent = `Team ${idx + 1}`;
      const teamTotal = document.createElement("div");
      teamTotal.className = "planTeamTotal";
      teamTotal.textContent = `${team.totalMinutes} min (${hhmm(team.totalMinutes)})`;
      teamHead.appendChild(teamName);
      teamHead.appendChild(teamTotal);

      const list = document.createElement("ul");
      list.className = "planJobList";

      if (!team.jobs.length) {
        const empty = document.createElement("li");
        empty.className = "planEmpty";
        empty.textContent = "No jobs assigned.";
        list.appendChild(empty);
      } else {
        team.jobs.forEach((job) => {
          const item = document.createElement("li");
          item.className = "planJobItem";

          const jobName = document.createElement("span");
          jobName.className = "planJobName";
          jobName.textContent = job.name;
          jobName.title = job.name;
          item.title = job.name;

          const jobTime = document.createElement("span");
          jobTime.className = "planJobTime";
          jobTime.textContent = `${job.minutes} min (${hhmm(job.minutes)})`;

          item.appendChild(jobName);
          item.appendChild(jobTime);
          list.appendChild(item);
        });
      }

      teamCard.appendChild(teamHead);
      teamCard.appendChild(list);
      teamsWrap.appendChild(teamCard);
    });

    view.appendChild(teamsWrap);
    elOut.appendChild(view);
  }

  function renderPlanTextFallback(text) {
    if (!String(text || "").trim()) {
      clearPlanOutput();
      return;
    }
    elOut.innerHTML = "";
    const plain = document.createElement("div");
    plain.className = "planPlainText";
    plain.textContent = text;
    elOut.appendChild(plain);
    elOut.dataset.copyText = text;
  }

  function isConfirmModalOpen() {
    return !!(elConfirmModal && elConfirmModal.classList.contains("isOpen"));
  }

  function closeConfirmModal(confirmed) {
    if (!isConfirmModalOpen()) return;

    elConfirmModal.classList.remove("isOpen");
    elConfirmModal.setAttribute("aria-hidden", "true");
    document.body.classList.remove("modalOpen");

    const resolve = confirmResolver;
    confirmResolver = null;
    if (resolve) resolve(Boolean(confirmed));

    if (confirmPreviouslyFocused && typeof confirmPreviouslyFocused.focus === "function") {
      confirmPreviouslyFocused.focus();
    }
    confirmPreviouslyFocused = null;
  }

  function openConfirmModal({
    title = "Confirm action",
    message = "Are you sure?",
    confirmText = "Delete",
    cancelText = "Cancel",
    tone = "danger",
  } = {}) {
    if (!elConfirmModal || !elConfirmTitle || !elConfirmMessage || !elConfirmCancelBtn || !elConfirmOkBtn) {
      return Promise.resolve(false);
    }

    if (confirmResolver) {
      closeConfirmModal(false);
    }

    confirmPreviouslyFocused = document.activeElement;
    elConfirmTitle.textContent = String(title);
    elConfirmMessage.textContent = String(message);
    elConfirmCancelBtn.textContent = String(cancelText);
    elConfirmOkBtn.textContent = String(confirmText);
    elConfirmOkBtn.classList.remove("bad", "primary");
    if (tone === "danger") {
      elConfirmOkBtn.classList.add("bad");
    } else {
      elConfirmOkBtn.classList.add("primary");
    }

    elConfirmModal.setAttribute("aria-hidden", "false");
    elConfirmModal.classList.add("isOpen");
    document.body.classList.add("modalOpen");

    requestAnimationFrame(() => {
      elConfirmCancelBtn.focus();
    });

    return new Promise((resolve) => {
      confirmResolver = resolve;
    });
  }

  function detectSystemTheme() {
    if (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches) {
      return "dark";
    }
    return "light";
  }

  function getSavedTheme() {
    try {
      const stored = localStorage.getItem(UI_THEME_KEY);
      if (stored === "dark" || stored === "light") return stored;
    } catch {
      // Ignore localStorage errors and fallback to system theme.
    }
    return null;
  }

  function setTheme(theme, persist = true) {
    const nextTheme = theme === "dark" ? "dark" : "light";
    const root = document.documentElement;
    root.classList.add("themeSwitching");
    root.setAttribute("data-theme", nextTheme);
    if (elThemeToggleText) {
      elThemeToggleText.textContent = nextTheme === "dark" ? "☀" : "☾";
    }
    if (elThemeToggleBtn) {
      elThemeToggleBtn.setAttribute(
        "aria-label",
        nextTheme === "dark" ? "Switch to light mode" : "Switch to dark mode",
      );
      elThemeToggleBtn.title = nextTheme === "dark" ? "Light mode" : "Dark mode";
    }
    if (persist) {
      try {
        localStorage.setItem(UI_THEME_KEY, nextTheme);
      } catch {
        // Ignore localStorage errors.
      }
    }
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        root.classList.remove("themeSwitching");
      });
    });
  }

  function initTheme() {
    const saved = getSavedTheme();
    const initial = document.documentElement.getAttribute("data-theme") || saved || detectSystemTheme();
    setTheme(initial, !!saved);
  }

  function toggleTheme() {
    const current = document.documentElement.getAttribute("data-theme") === "dark" ? "dark" : "light";
    setTheme(current === "dark" ? "light" : "dark", true);
  }

  function buildFixOptions(teamCount) {
    const opts = [{ value: 0, text: "Any team" }];
    for (let i = 1; i <= teamCount; i++) {
      opts.push({ value: i, text: `Team ${i}` });
    }
    return opts;
  }

  function fillSelect(selectEl, teamCount, value) {
    selectEl.innerHTML = "";
    for (const optionData of buildFixOptions(teamCount)) {
      const option = document.createElement("option");
      option.value = String(optionData.value);
      option.textContent = optionData.text;
      if (String(value ?? 0) === String(optionData.value)) option.selected = true;
      selectEl.appendChild(option);
    }
  }

  function currentTeamCount() {
    return Math.max(1, parseInt(elK.value, 10) || 1);
  }

  function currentQuant() {
    return Math.max(1, parseInt(elQuant.value, 10) || 1);
  }

  function normalizeSearch(value) {
    return String(value ?? "")
      .toLowerCase()
      .trim();
  }

  function syncAllFixSelects() {
    const teamCount = currentTeamCount();
    fillSelect(elFixIn, teamCount, parseInt(elFixIn.value, 10) || 0);

    for (const job of jobs) {
      if (job._fixSelectEl) {
        const current = parseInt(job._fixSelectEl.value, 10) || 0;
        fillSelect(job._fixSelectEl, teamCount, Math.min(current, teamCount));
        job.fixedTeam = parseInt(job._fixSelectEl.value, 10) || 0;
      } else {
        job.fixedTeam = Math.min(job.fixedTeam || 0, teamCount);
      }
    }
  }

  function normalizeJobForSave(job) {
    return {
      name: String(job.name ?? "").trim(),
      durRaw: String(job.durRaw ?? "").trim(),
      fixedTeam: Math.max(0, parseInt(job.fixedTeam, 10) || 0),
    };
  }

  function sanitizeDraft(raw, cutoff) {
    if (!raw || typeof raw !== "object") return null;
    const createdAt = Number(raw.createdAt);
    if (!Number.isFinite(createdAt) || createdAt < cutoff) return null;

    const jobsRaw = Array.isArray(raw.jobs) ? raw.jobs : [];
    const savedJobs = jobsRaw.map(normalizeJobForSave).filter((job) => job.name || job.durRaw || job.fixedTeam > 0);

    return {
      createdAt,
      jobs: savedJobs,
      k: Math.max(1, parseInt(raw.k, 10) || 1),
      quant: Math.max(1, parseInt(raw.quant, 10) || 1),
      editingCheckpointId: raw.editingCheckpointId ? String(raw.editingCheckpointId) : null,
      pendingName: String(raw.pendingName || ""),
      pendingDur: String(raw.pendingDur || ""),
      pendingFix: Math.max(0, parseInt(raw.pendingFix, 10) || 0),
    };
  }

  function buildDraftSnapshot() {
    return {
      createdAt: Date.now(),
      jobs: jobs.map(normalizeJobForSave).filter((job) => job.name || job.durRaw || job.fixedTeam > 0),
      k: currentTeamCount(),
      quant: currentQuant(),
      editingCheckpointId,
      pendingName: String(elNameIn.value || ""),
      pendingDur: String(elDurIn.value || ""),
      pendingFix: Math.max(0, parseInt(elFixIn.value, 10) || 0),
    };
  }

  function persistDraftNow() {
    safeWriteJson(DRAFT_STORAGE_KEY, buildDraftSnapshot());
  }

  function scheduleDraftSave() {
    if (draftSaveTimer) clearTimeout(draftSaveTimer);
    draftSaveTimer = setTimeout(() => {
      draftSaveTimer = null;
      persistDraftNow();
    }, DRAFT_SAVE_DELAY_MS);
  }

  function clearDraft() {
    if (draftSaveTimer) {
      clearTimeout(draftSaveTimer);
      draftSaveTimer = null;
    }
    try {
      localStorage.removeItem(DRAFT_STORAGE_KEY);
    } catch {
      // Ignore localStorage removal errors.
    }
  }

  function restoreDraftFromStorage() {
    const raw = safeReadJson(DRAFT_STORAGE_KEY);
    const cutoff = Date.now() - DRAFT_TTL_MS;
    const draft = sanitizeDraft(raw, cutoff);
    if (!draft) {
      clearDraft();
      return;
    }

    editingCheckpointId = draft.editingCheckpointId;
    jobId = 1;
    jobs = draft.jobs.map((job) => ({
      id: jobId++,
      name: job.name,
      durRaw: job.durRaw,
      fixedTeam: job.fixedTeam,
    }));

    elK.value = String(draft.k);
    elQuant.value = String(draft.quant);
    elNameIn.value = draft.pendingName;
    elDurIn.value = draft.pendingDur;
    syncAllFixSelects();
    elFixIn.value = String(Math.min(draft.pendingFix, currentTeamCount()));
    render();
    renderHistory();

    if (draft.jobs.length || draft.pendingName || draft.pendingDur) {
      setStatus(`Draft restored (${formatLocalDateTime(draft.createdAt)}).`, "ok");
    }
  }

  function sanitizeHistoryEntry(entry, cutoff) {
    if (!entry || typeof entry !== "object") return null;
    const createdAt = Number(entry.createdAt);
    if (!Number.isFinite(createdAt) || createdAt < cutoff) return null;

    const rawJobs = Array.isArray(entry.jobs) ? entry.jobs : [];
    const savedJobs = rawJobs
      .map(normalizeJobForSave)
      .filter((job) => job.name || job.durRaw || job.fixedTeam > 0);
    if (!savedJobs.length) return null;

    return {
      id: String(entry.id || makeId("checkpoint")),
      createdAt,
      label: String(entry.label || "Saved"),
      jobs: savedJobs,
      k: Math.max(1, parseInt(entry.k, 10) || 1),
      quant: Math.max(1, parseInt(entry.quant, 10) || 1),
      outputText: String(entry.outputText || ""),
      statusText: String(entry.statusText || ""),
      planSnapshot: sanitizePlanSnapshot(entry.planSnapshot),
    };
  }

  function cleanupHistory(entries) {
    const cutoff = Date.now() - HISTORY_TTL_MS;
    const clean = [];
    for (const entry of entries) {
      const sanitized = sanitizeHistoryEntry(entry, cutoff);
      if (sanitized) clean.push(sanitized);
    }
    clean.sort((a, b) => b.createdAt - a.createdAt);
    return clean.slice(0, HISTORY_MAX_ITEMS);
  }

  function persistHistory() {
    safeWriteJson(HISTORY_STORAGE_KEY, solveHistory);
  }

  function updateHistoryActionButtons() {
    const hasHistory = solveHistory.length > 0;
    const selectedId = getSelectedCheckpointId();

    if (!hasHistory) {
      elLoadHistoryBtn.disabled = true;
      if (elDeleteHistoryBtn) elDeleteHistoryBtn.disabled = true;
      elClearHistoryBtn.disabled = true;
      return;
    }

    elLoadHistoryBtn.disabled = !selectedId;
    if (elDeleteHistoryBtn) elDeleteHistoryBtn.disabled = !selectedId;
    elClearHistoryBtn.disabled = false;
  }

  function renderHistory() {
    const previousSelectedId = String(elHistorySelect.value || "");
    elHistorySelect.innerHTML = "";

    if (!solveHistory.length) {
      const option = document.createElement("option");
      option.value = "";
      option.textContent = "No saved checks";
      elHistorySelect.appendChild(option);
      elHistorySelect.disabled = true;
      updateHistoryActionButtons();
      elHistoryInfo.textContent = "Saved checks: 0 (kept for 3 days).";
      return;
    }

    elHistorySelect.disabled = false;

    const detachedOption = document.createElement("option");
    detachedOption.value = "";
    detachedOption.textContent = "Start a new check";
    elHistorySelect.appendChild(detachedOption);

    for (const entry of solveHistory) {
      const option = document.createElement("option");
      option.value = entry.id;
      option.textContent =
        `${formatLocalDateTime(entry.createdAt)} | ` +
        `${entry.jobs.length} jobs | ${entry.k} teams | ${entry.label}`;
      elHistorySelect.appendChild(option);
    }

    if (editingCheckpointId && solveHistory.some((entry) => entry.id === editingCheckpointId)) {
      elHistorySelect.value = editingCheckpointId;
    } else if (previousSelectedId && solveHistory.some((entry) => entry.id === previousSelectedId)) {
      elHistorySelect.value = previousSelectedId;
    } else {
      elHistorySelect.value = "";
    }

    updateHistoryActionButtons();
    elHistoryInfo.textContent = `Saved checks: ${solveHistory.length} (kept for 3 days).`;
  }

  function loadHistoryFromStorage() {
    const raw = safeReadJson(HISTORY_STORAGE_KEY);
    const entries = Array.isArray(raw) ? raw : [];
    const cleaned = cleanupHistory(entries);
    solveHistory = cleaned;
    if (!solveHistory.some((entry) => entry.id === editingCheckpointId)) {
      editingCheckpointId = null;
    }
    persistHistory();
    renderHistory();
  }

  function createCheckpoint(label, id) {
    return {
      id: id || makeId("checkpoint"),
      createdAt: Date.now(),
      label,
      jobs: jobs.map(normalizeJobForSave),
      k: currentTeamCount(),
      quant: currentQuant(),
      outputText: elOut.dataset.copyText || elOut.textContent || "",
      statusText: elStatus.textContent || "",
      planSnapshot: latestPlanSnapshot || null,
    };
  }

  function saveCheckpoint(label, preferredId = null, allowCreate = true) {
    const checkpointId = preferredId || editingCheckpointId || (allowCreate ? makeId("checkpoint") : null);
    if (!checkpointId) return;

    const checkpoint = createCheckpoint(label, checkpointId);
    const idx = solveHistory.findIndex((entry) => entry.id === checkpointId);

    if (idx >= 0) {
      solveHistory[idx] = checkpoint;
    } else if (allowCreate) {
      solveHistory.unshift(checkpoint);
    } else {
      return;
    }

    solveHistory = cleanupHistory(solveHistory);
    editingCheckpointId = checkpoint.id;
    persistHistory();
    renderHistory();
  }

  function getSelectedCheckpointId() {
    const selectedId = String(elHistorySelect.value || "");
    if (!selectedId) return null;
    return solveHistory.some((entry) => entry.id === selectedId) ? selectedId : null;
  }

  function saveOrCreateCheckpointProgress(label = "In progress") {
    const selectedId = getSelectedCheckpointId();
    if (selectedId) {
      saveCheckpoint(label, selectedId, false);
      return;
    }

    if (jobs.length === 0) return;
    saveCheckpoint(label);
  }

  function restoreCheckpoint(entry) {
    editingCheckpointId = entry.id;
    jobId = 1;
    jobs = entry.jobs.map((job) => ({
      id: jobId++,
      name: job.name,
      durRaw: job.durRaw,
      fixedTeam: job.fixedTeam,
    }));

    elK.value = String(entry.k);
    elQuant.value = String(entry.quant);
    syncAllFixSelects();
    render();

    latestPlanSnapshot = entry.planSnapshot || null;
    if (latestPlanSnapshot) {
      renderPlanView(latestPlanSnapshot);
    } else {
      renderPlanTextFallback(entry.outputText || "");
    }
    setStatus(`Check opened (${formatLocalDateTime(entry.createdAt)}).`, "ok");
    clearError();
    renderHistory();
    scheduleDraftSave();
  }

  function loadSelectedCheckpoint() {
    const id = String(elHistorySelect.value || "");
    if (!id) return;
    const selected = solveHistory.find((entry) => entry.id === id);
    if (!selected) return;
    restoreCheckpoint(selected);
  }

  async function clearHistory() {
    if (!solveHistory.length) return;
    const ok = await openConfirmModal({
      title: "Clear all saved checks?",
      message: "This will permanently remove all saved checks from the last 3 days.",
      confirmText: "Clear all",
      cancelText: "Cancel",
      tone: "danger",
    });
    if (!ok) return;

    solveHistory = [];
    editingCheckpointId = null;
    try {
      localStorage.removeItem(HISTORY_STORAGE_KEY);
    } catch {
      // Ignore localStorage removal errors.
    }
    renderHistory();
    setStatus("All checks cleared.", "ok");
  }

  async function deleteSelectedCheckpoint() {
    const selectedId = getSelectedCheckpointId();
    if (!selectedId) return;

    const entry = solveHistory.find((x) => x.id === selectedId);
    if (!entry) return;

    const ok = await openConfirmModal({
      title: "Delete this saved check?",
      message: `${formatLocalDateTime(entry.createdAt)} | ${entry.jobs.length} jobs`,
      confirmText: "Delete",
      cancelText: "Cancel",
      tone: "danger",
    });
    if (!ok) return;

    solveHistory = solveHistory.filter((x) => x.id !== selectedId);
    if (editingCheckpointId === selectedId) {
      editingCheckpointId = null;
    }

    persistHistory();
    renderHistory();
    scheduleDraftSave();
    setStatus("Check deleted.", "ok");
  }

  function render() {
    elList.innerHTML = "";
    const teamCount = currentTeamCount();

    jobs.forEach((job) => {
      const row = document.createElement("div");
      row.className = "rowItem";

      const nameWrap = document.createElement("div");
      nameWrap.className = "rowNameAuto";

      const name = document.createElement("input");
      name.type = "text";
      name.value = job.name;
      name.placeholder = "Address";
      name.setAttribute("aria-autocomplete", "list");
      name.setAttribute("aria-expanded", "false");

      const rowSuggestions = document.createElement("div");
      rowSuggestions.className = "suggestions hidden";
      rowSuggestions.setAttribute("role", "listbox");
      const rowSuggestionId = `row-suggestions-${job.id}`;
      rowSuggestions.id = rowSuggestionId;
      name.setAttribute("aria-controls", rowSuggestionId);

      let rowSuggestionItems = [];
      let rowSuggestionActiveIndex = -1;

      function hideRowSuggestions() {
        rowSuggestionItems = [];
        rowSuggestionActiveIndex = -1;
        rowSuggestions.innerHTML = "";
        rowSuggestions.classList.add("hidden");
        name.setAttribute("aria-expanded", "false");
      }

      function setRowSuggestionActiveIndex(index) {
        const buttons = rowSuggestions.querySelectorAll(".suggestionItem");
        buttons.forEach((btn, i) => {
          btn.classList.toggle("active", i === index);
        });
        rowSuggestionActiveIndex = index;
      }

      function applyRowDurationFromAddress(force = false) {
        const rowData = sheetsDb.findByAddress(name.value);
        if (!rowData || !rowData.scheduledTime) return false;
        if (!force && String(dur.value || "").trim()) return false;
        dur.value = rowData.scheduledTime;
        job.durRaw = dur.value;
        return true;
      }

      function applyRowSuggestionByIndex(index) {
        const address = rowSuggestionItems[index];
        if (!address) return false;
        name.value = address;
        job.name = address;
        hideRowSuggestions();
        applyRowDurationFromAddress(true);
        scheduleDraftSave();
        return true;
      }

      function renderRowSuggestions(matches) {
        rowSuggestions.innerHTML = "";
        rowSuggestionItems = matches.slice();
        rowSuggestionActiveIndex = -1;

        for (const address of matches) {
          const btn = document.createElement("button");
          btn.type = "button";
          btn.className = "suggestionItem";
          btn.textContent = address;
          btn.title = address;
          btn.addEventListener("mousedown", (event) => {
            event.preventDefault();
          });
          btn.addEventListener("click", () => {
            const idx = rowSuggestionItems.indexOf(address);
            applyRowSuggestionByIndex(idx);
          });
          rowSuggestions.appendChild(btn);
        }

        rowSuggestions.classList.remove("hidden");
        name.setAttribute("aria-expanded", "true");
      }

      function refreshRowSuggestions() {
        const query = normalizeSearch(name.value);
        if (!query) {
          hideRowSuggestions();
          return;
        }

        const matches = addressCatalog
          .filter((address) => normalizeSearch(address).includes(query))
          .slice(0, 8);

        if (!matches.length) {
          hideRowSuggestions();
          return;
        }

        renderRowSuggestions(matches);
      }

      name.addEventListener("input", () => {
        job.name = name.value;
        applyRowDurationFromAddress(false);
        refreshRowSuggestions();
        scheduleDraftSave();
      });
      name.addEventListener("change", () => {
        applyRowDurationFromAddress(false);
        hideRowSuggestions();
      });
      name.addEventListener("focus", () => {
        refreshRowSuggestions();
      });
      name.addEventListener("blur", () => {
        setTimeout(() => {
          hideRowSuggestions();
        }, 120);
      });

      const dur = document.createElement("input");
      dur.type = "text";
      dur.value = job.durRaw;
      dur.placeholder = "Duration (1.5 or 2:25)";
      dur.addEventListener("input", () => {
        job.durRaw = dur.value;
        scheduleDraftSave();
      });
      name.addEventListener("keydown", (event) => {
        const isOpen = !rowSuggestions.classList.contains("hidden");

        if (event.key === "ArrowDown" && isOpen) {
          event.preventDefault();
          const next = Math.min(rowSuggestionActiveIndex + 1, rowSuggestionItems.length - 1);
          setRowSuggestionActiveIndex(next);
          return;
        }

        if (event.key === "ArrowUp" && isOpen) {
          event.preventDefault();
          const next = Math.max(rowSuggestionActiveIndex - 1, 0);
          setRowSuggestionActiveIndex(next);
          return;
        }

        if (event.key === "Escape" && isOpen) {
          event.preventDefault();
          hideRowSuggestions();
          return;
        }

        if (event.key === "Enter") {
          if (isOpen && rowSuggestionItems.length) {
            event.preventDefault();
            const idx = rowSuggestionActiveIndex >= 0 ? rowSuggestionActiveIndex : 0;
            if (applyRowSuggestionByIndex(idx)) return;
          }
          dur.focus();
        }
      });

      const fix = document.createElement("select");
      job._fixSelectEl = fix;
      fillSelect(fix, teamCount, Math.min(job.fixedTeam || 0, teamCount));
      fix.addEventListener("change", () => {
        job.fixedTeam = parseInt(fix.value, 10) || 0;
        scheduleDraftSave();
      });

      const del = document.createElement("button");
      del.className = "btn icon bad del";
      del.type = "button";
      del.textContent = "✕";
      del.title = "Delete";
      del.setAttribute("aria-label", `Delete ${job.name || "job"}`);
      del.addEventListener("click", () => {
        jobs = jobs.filter((x) => x.id !== job.id);
        render();
        scheduleDraftSave();
      });

      nameWrap.appendChild(name);
      nameWrap.appendChild(rowSuggestions);
      row.appendChild(nameWrap);
      row.appendChild(dur);
      row.appendChild(fix);
      row.appendChild(del);
      elList.appendChild(row);
    });

    elCount.textContent = `Jobs: ${jobs.length} | Teams: ${teamCount}`;
  }

  function addJob(name, durRaw, fixedTeam) {
    clearError();
    const parsedName = String(name ?? "").trim();
    const parsedDur = String(durRaw ?? "").trim();
    const parsedFixedTeam = parseInt(fixedTeam, 10) || 0;

    if (!parsedName) throw new Error("Enter an address first");
    parseTimeToMinutes(parsedDur);

    jobs.push({
      id: jobId++,
      name: parsedName,
      durRaw: parsedDur,
      fixedTeam: parsedFixedTeam,
    });

    render();
    scheduleDraftSave();
  }

  function fillDurationFromAddress() {
    const row = sheetsDb.findByAddress(elNameIn.value);
    if (!row || !row.scheduledTime) return false;
    elDurIn.value = row.scheduledTime;
    elDurIn.dataset.autoFilled = "1";
    return true;
  }

  function hideNameSuggestions() {
    suggestionItems = [];
    suggestionActiveIndex = -1;
    elNameSuggestions.innerHTML = "";
    elNameSuggestions.classList.add("hidden");
    elNameIn.setAttribute("aria-expanded", "false");
  }

  function setSuggestionActiveIndex(index) {
    const buttons = elNameSuggestions.querySelectorAll(".suggestionItem");
    buttons.forEach((btn, i) => {
      btn.classList.toggle("active", i === index);
    });
    suggestionActiveIndex = index;
  }

  function applySuggestionByIndex(index) {
    const address = suggestionItems[index];
    if (!address) return false;
    elNameIn.value = address;
    hideNameSuggestions();
    fillDurationFromAddress();
    scheduleDraftSave();
    return true;
  }

  function renderNameSuggestions(matches) {
    elNameSuggestions.innerHTML = "";
    suggestionItems = matches.slice();
    suggestionActiveIndex = -1;

    for (const address of matches) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "suggestionItem";
      btn.textContent = address;
      btn.title = address;
      btn.addEventListener("mousedown", (event) => {
        event.preventDefault();
      });
      btn.addEventListener("click", () => {
        const idx = suggestionItems.indexOf(address);
        applySuggestionByIndex(idx);
      });
      elNameSuggestions.appendChild(btn);
    }

    elNameSuggestions.classList.remove("hidden");
    elNameIn.setAttribute("aria-expanded", "true");
  }

  function refreshNameSuggestions() {
    const query = normalizeSearch(elNameIn.value);
    if (!query) {
      hideNameSuggestions();
      return;
    }

    const matches = addressCatalog
      .filter((address) => normalizeSearch(address).includes(query))
      .slice(0, 8);

    if (!matches.length) {
      hideNameSuggestions();
      return;
    }

    renderNameSuggestions(matches);
  }

  function addFromInputs() {
    try {
      if (!String(elDurIn.value || "").trim()) fillDurationFromAddress();

      addJob(elNameIn.value, elDurIn.value, elFixIn.value);
      elNameIn.value = "";
      elDurIn.value = "";
      elFixIn.value = "0";
      elDurIn.dataset.autoFilled = "0";
      hideNameSuggestions();
      elNameIn.focus();
      scheduleDraftSave();
      saveOrCreateCheckpointProgress("Updated");
    } catch (error) {
      showError(getErrorMessage(error));
    }
  }

  function renderAddressSuggestions(records) {
    addressCatalog = records.map((row) => String(row.address || "").trim()).filter(Boolean);
    if (elNameIn.value.trim()) {
      refreshNameSuggestions();
    } else {
      hideNameSuggestions();
    }
  }

  function onDbStateChange(state) {
    renderAddressSuggestions(state.records);

    if (state.loading) {
      elDbInfo.className = "mini";
      elDbInfo.textContent = "Refreshing address list...";
    } else if (state.lastError) {
      elDbInfo.className = "mini badTxt";
      elDbInfo.textContent = "Address list is temporarily unavailable. Retrying soon.";
    } else {
      elDbInfo.className = "mini";
      elDbInfo.textContent = "Address list is ready.";
    }

    if (!state.loading && !state.lastError) {
      fillDurationFromAddress();
    }
  }

  async function refreshDbNow() {
    clearError();
    await sheetsDb.refreshNow();
    const current = sheetsDb.getState();
    if (current.lastError) {
      showError("We could not refresh the address list right now. Please try again.");
      return;
    }
    fillDurationFromAddress();
    setStatus("Address list refreshed.", "ok");
  }

  function solve() {
    clearError();
    clearPlanOutput();
    latestPlanSnapshot = null;

    try {
      const teamCount = parseInt(elK.value, 10);
      const quant = parseInt(elQuant.value, 10);
      if (!Number.isFinite(teamCount) || teamCount < 1) throw new Error("Please choose at least 1 team");
      if (!Number.isFinite(quant) || quant < 1) throw new Error("Rounding must be 1 minute or more");
      if (jobs.length === 0) throw new Error("Add at least one job");

      syncAllFixSelects();

      const jobsMinutes = jobs.map((job, idx) => {
        const name = String(job.name ?? "").trim();
        if (!name) throw new Error(`Job ${idx + 1}: address is missing`);
        const minutes = parseTimeToMinutes(job.durRaw);
        const fixedTeam = Math.min(Math.max(parseInt(job.fixedTeam, 10) || 0, 0), teamCount);
        return { name, minutes, fixedTeam };
      });

      const fixedCount = jobsMinutes.filter((x) => x.fixedTeam > 0).length;
      if (jobsMinutes.length > 90 && quant === 1) {
        setStatus("Creating plan. This can take a moment for larger lists.");
      } else {
        setStatus(
          fixedCount
          ? `Creating plan with ${fixedCount} assigned jobs...`
          : "Creating plan...",
        );
      }

      setTimeout(() => {
        try {
          const result = solveOptimalWithFixed(jobsMinutes, teamCount, quant);
          const snapshot = buildPlanSnapshot(result, teamCount, jobsMinutes.length);
          latestPlanSnapshot = snapshot;
          renderPlanView(snapshot);
          setStatus("Plan ready.", "ok");
          saveCheckpoint("Planned");
        } catch (error) {
          latestPlanSnapshot = null;
          clearPlanOutput();
          setStatus("We could not create a plan.", "bad");
          showError(getErrorMessage(error));
          saveCheckpoint("Plan failed");
        }
      }, 10);
    } catch (error) {
      setStatus("Please review your inputs.", "bad");
      showError(getErrorMessage(error));
      if (jobs.length > 0) {
        saveCheckpoint("Needs review");
      }
    }
  }

  async function copyResult() {
    try {
      const text = elOut.dataset.copyText || elOut.textContent || "";
      if (!text.trim()) throw new Error("Create a plan first");
      await navigator.clipboard.writeText(text);
      setStatus("Plan copied.", "ok");
    } catch (error) {
      showError(getErrorMessage(error));
    }
  }

  $("addBtn").addEventListener("click", addFromInputs);
  if (elAddJobForm) {
    elAddJobForm.addEventListener("submit", (event) => {
      event.preventDefault();
      addFromInputs();
    });
  }
  $("clearBtn").addEventListener("click", () => {
    jobs = [];
    editingCheckpointId = null;
    latestPlanSnapshot = null;
    if (!elHistorySelect.disabled) {
      elHistorySelect.value = "";
    }
    render();
    clearPlanOutput();
    setStatus("");
    clearError();
    hideNameSuggestions();
    clearDraft();
  });
  $("demoBtn").addEventListener("click", () => {
    editingCheckpointId = null;
    latestPlanSnapshot = null;
    jobs = [
      { id: jobId++, name: "4567 Dynasty Road", durRaw: "1.5", fixedTeam: 1 },
      { id: jobId++, name: "Kitchen wiring", durRaw: "2:15", fixedTeam: 0 },
      { id: jobId++, name: "Bedroom painting", durRaw: "3.1", fixedTeam: 2 },
      { id: jobId++, name: "Patio cleanup", durRaw: "0:45", fixedTeam: 0 },
      { id: jobId++, name: "Light fixture install", durRaw: "1:20", fixedTeam: 0 },
      { id: jobId++, name: "Plumbing check", durRaw: "2.1", fixedTeam: 0 },
    ];
    render();
    clearPlanOutput();
    clearError();
    hideNameSuggestions();
    scheduleDraftSave();
  });
  $("solveBtn").addEventListener("click", solve);
  $("copyBtn").addEventListener("click", copyResult);

  if (elThemeToggleBtn) {
    elThemeToggleBtn.addEventListener("click", toggleTheme);
  }

  elRefreshDbBtn.addEventListener("click", () => {
    void refreshDbNow();
  });

  elLoadHistoryBtn.addEventListener("click", loadSelectedCheckpoint);
  if (elDeleteHistoryBtn) {
    elDeleteHistoryBtn.addEventListener("click", deleteSelectedCheckpoint);
  }
  elClearHistoryBtn.addEventListener("click", clearHistory);
  elHistorySelect.addEventListener("change", () => {
    updateHistoryActionButtons();
  });

  elK.addEventListener("change", () => {
    syncAllFixSelects();
    render();
    scheduleDraftSave();
  });
  elQuant.addEventListener("change", scheduleDraftSave);
  elDurIn.addEventListener("input", () => {
    elDurIn.dataset.autoFilled = "0";
    scheduleDraftSave();
  });
  elDurIn.addEventListener("keydown", (event) => {
    if (event.key === "Enter") addFromInputs();
  });
  elNameIn.addEventListener("input", () => {
    fillDurationFromAddress();
    refreshNameSuggestions();
    scheduleDraftSave();
  });
  elNameIn.addEventListener("change", () => {
    fillDurationFromAddress();
    hideNameSuggestions();
    scheduleDraftSave();
  });
  elNameIn.addEventListener("focus", () => {
    refreshNameSuggestions();
  });
  elNameIn.addEventListener("blur", () => {
    setTimeout(() => {
      hideNameSuggestions();
    }, 120);
  });
  elNameIn.addEventListener("keydown", (event) => {
    const isOpen = !elNameSuggestions.classList.contains("hidden");

    if (event.key === "ArrowDown" && isOpen) {
      event.preventDefault();
      const next = Math.min(suggestionActiveIndex + 1, suggestionItems.length - 1);
      setSuggestionActiveIndex(next);
      return;
    }

    if (event.key === "ArrowUp" && isOpen) {
      event.preventDefault();
      const next = Math.max(suggestionActiveIndex - 1, 0);
      setSuggestionActiveIndex(next);
      return;
    }

    if (event.key === "Escape" && isOpen) {
      event.preventDefault();
      hideNameSuggestions();
      return;
    }

    if (event.key === "Enter") {
      if (isOpen && suggestionItems.length) {
        event.preventDefault();
        const idx = suggestionActiveIndex >= 0 ? suggestionActiveIndex : 0;
        if (applySuggestionByIndex(idx)) return;
      }
      elDurIn.focus();
    }
  });
  elFixIn.addEventListener("change", scheduleDraftSave);

  document.addEventListener("click", (event) => {
    if (!elNameSuggestions.contains(event.target) && event.target !== elNameIn) {
      hideNameSuggestions();
    }
  });

  if (elConfirmCancelBtn) {
    elConfirmCancelBtn.addEventListener("click", () => {
      closeConfirmModal(false);
    });
  }

  if (elConfirmOkBtn) {
    elConfirmOkBtn.addEventListener("click", () => {
      closeConfirmModal(true);
    });
  }

  if (elConfirmModal) {
    elConfirmModal.addEventListener("click", (event) => {
      if (event.target === elConfirmModal) {
        closeConfirmModal(false);
      }
    });
  }

  document.addEventListener("keydown", (event) => {
    if (!isConfirmModalOpen()) return;

    if (event.key === "Escape") {
      event.preventDefault();
      closeConfirmModal(false);
      return;
    }

    if (event.key !== "Tab" || !elConfirmModal) return;

    const focusable = Array.from(
      elConfirmModal.querySelectorAll(
        'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])',
      ),
    ).filter((el) => !el.hasAttribute("disabled"));

    if (!focusable.length) return;

    const first = focusable[0];
    const last = focusable[focusable.length - 1];
    const active = document.activeElement;

    if (event.shiftKey && active === first) {
      event.preventDefault();
      last.focus();
      return;
    }

    if (!event.shiftKey && active === last) {
      event.preventDefault();
      first.focus();
    }
  });

  initTheme();
  fillSelect(elFixIn, 2, 0);
  jobs = [];
  render();
  loadHistoryFromStorage();
  restoreDraftFromStorage();

  void (async () => {
    await sheetsDb.init();
    const current = sheetsDb.getState();
    if (current.lastError) {
      showError("We could not sync the address list right now. Please try again.");
    }
  })();

  window.addEventListener("beforeunload", () => {
    if (draftSaveTimer) {
      clearTimeout(draftSaveTimer);
      draftSaveTimer = null;
      persistDraftNow();
    }
    sheetsDb.destroy();
  });
}

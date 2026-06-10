const FALLBACK_PROVIDERS = {
  qwen: {
    label: "Qwen",
    siteLabel: "Qwen Studio",
    url: "https://chat.qwen.ai/",
    responseLabel: "Ответ Qwen",
    pastePlaceholder: "Вставь сюда полный ответ Qwen",
    maxUploadFilesPerBatch: 5,
    limitSummary: "Qwen web: до 5 файлов на один батч.",
  },
  gemini: {
    label: "Gemini",
    siteLabel: "Google AI Studio",
    url: "https://aistudio.google.com/prompts/new_chat",
    responseLabel: "Ответ Gemini",
    pastePlaceholder: "Вставь сюда полный ответ Gemini",
    maxUploadFilesPerBatch: 0,
    limitSummary: "Google AI Studio: file-count лимит для этого workflow не задаем.",
  },
  kimi: {
    label: "Kimi",
    siteLabel: "Kimi",
    url: "https://www.kimi.com/",
    responseLabel: "Ответ Kimi",
    pastePlaceholder: "Вставь сюда полный ответ Kimi",
    maxUploadFilesPerBatch: 50,
    limitSummary: "Kimi web: практический лимит 50 файлов на батч.",
  },
  mimo: {
    label: "MiMo",
    siteLabel: "Xiaomi MiMo",
    url: "https://aistudio.xiaomimimo.com/#/c",
    responseLabel: "Ответ MiMo",
    pastePlaceholder: "Вставь сюда полный ответ MiMo",
    maxUploadFilesPerBatch: 5,
    limitSummary: "Xiaomi MiMo: безопасные 5 файлов на батч.",
  },
  deepseek: {
    label: "DeepSeek",
    siteLabel: "DeepSeek",
    url: "https://chat.deepseek.com/",
    responseLabel: "Ответ DeepSeek",
    pastePlaceholder: "Вставь сюда полный ответ DeepSeek",
    maxUploadFilesPerBatch: 20,
    limitSummary: "DeepSeek web: консервативные 20 файлов на батч.",
  },
};

let PROVIDERS = { ...FALLBACK_PROVIDERS };

const state = {
  mode: "quarterly",
  provider: normalizeProvider(localStorage.getItem("cordiant-provider")),
  status: null,
  selectedBatch: null,
  selectedBatchData: null,
  lastOutputPath: null,
  preparePollTimer: null,
};

const el = {
  modeQuarterly: must("#mode-quarterly"),
  modeAnnual: must("#mode-annual"),
  packageState: must("#package-state"),
  packageDetail: must("#package-detail"),
  jsonState: must("#json-state"),
  jsonDetail: must("#json-detail"),
  excelState: must("#excel-state"),
  excelDetail: must("#excel-detail"),
  providerName: must("#provider-name"),
  currentModeLabel: must("#current-mode-label"),
  progressFill: must("#progress-fill"),
  prepareBtn: must("#prepare-btn"),
  revealPackageBtn: must("#reveal-package-btn"),
  providerSelect: must("#provider-select"),
  providerLimitNote: must("#provider-limit-note"),
  openProviderSidebar: must("#open-provider-sidebar"),
  openProviderPrompt: must("#open-provider-prompt"),
  refreshBtn: must("#refresh-btn"),
  applyBtn: must("#apply-btn"),
  workflowNote: must("#workflow-note"),
  screenTitle: must("#screen-title"),
  screenSubtitle: must("#screen-subtitle"),
  providerStepLabel: must("#provider-step-label"),
  stepSources: must("#step-sources"),
  stepJson: must("#step-json"),
  stepExcel: must("#step-excel"),
  batchCount: must("#batch-count"),
  batchList: must("#batch-list"),
  batchTitle: must("#batch-title"),
  batchMeta: must("#batch-meta"),
  batchStatus: must("#batch-status"),
  promptBox: must("#prompt-box"),
  startBatchBtn: must("#start-batch-btn"),
  jsonBox: must("#json-box"),
  jsonSectionTitle: must("#json-section-title"),
  saveJsonBtn: must("#save-json-btn"),
  jsonFeedback: must("#json-feedback"),
  outputText: must("#output-text"),
  revealOutputBtn: must("#reveal-output-btn"),
  toast: must("#toast"),
};

function must(selector) {
  const node = document.querySelector(selector);
  if (!node) {
    throw new Error(`Missing dashboard element: ${selector}`);
  }
  return node;
}

function normalizeProvider(provider) {
  return PROVIDERS[provider] ? provider : "qwen";
}

function providerInfo() {
  return PROVIDERS[state.provider] || PROVIDERS.qwen;
}

function normalizeProviderProfile(profile) {
  if (!profile) return null;
  return {
    label: profile.label,
    siteLabel: profile.site_label || profile.siteLabel || profile.label,
    url: profile.url,
    responseLabel: profile.response_label || profile.responseLabel || `Ответ ${profile.label}`,
    pastePlaceholder: profile.paste_placeholder || profile.pastePlaceholder || `Вставь сюда полный ответ ${profile.label}`,
    maxUploadFilesPerBatch: Number(profile.max_upload_files_per_batch ?? profile.maxUploadFilesPerBatch ?? 0),
    limitSummary: profile.limit_summary || profile.limitSummary || "",
    limitConfidence: profile.limit_confidence || profile.limitConfidence || "",
  };
}

function syncProvidersFromStatus(status) {
  if (!status?.providers) return;
  const next = {};
  Object.entries(status.providers).forEach(([key, profile]) => {
    next[key] = normalizeProviderProfile(profile);
  });
  PROVIDERS = { ...FALLBACK_PROVIDERS, ...next };
  if (!PROVIDERS[state.provider]) {
    state.provider = "qwen";
    localStorage.setItem("cordiant-provider", state.provider);
  }
  renderProvider();
}

function modeInfo(status = state.status) {
  return status?.modes?.[state.mode] || {};
}

function modeLabel() {
  return state.mode === "quarterly" ? "Квартал" : "Год";
}

function prepareJob() {
  return modeInfo().prepare_job || null;
}

function isPreparing() {
  return prepareJob()?.state === "running";
}

function batchesForCurrentMode(status = state.status) {
  return status?.modes?.[state.mode]?.batches || [];
}

function selectedBatchStillExists(status = state.status) {
  return Boolean(state.selectedBatch && batchesForCurrentMode(status).some((batch) => batch.id === state.selectedBatch));
}

function nextUnsavedBatchId(status = state.status) {
  return batchesForCurrentMode(status).find((batch) => !batch.json_saved)?.id || null;
}

async function api(path, options = {}) {
  const response = await fetch(path, {
    headers: { "Content-Type": "application/json" },
    ...options,
  });
  const type = response.headers.get("content-type") || "";
  const payload = type.includes("application/json")
    ? await response.json()
    : { ok: false, error: await response.text() };
  if (!payload.ok) {
    throw new Error(payload.error || payload.stderr || `Операция не выполнена (${response.status})`);
  }
  return payload;
}

function showToast(message) {
  el.toast.textContent = message;
  el.toast.classList.add("visible");
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => el.toast.classList.remove("visible"), 3200);
}

function setBusy(button, busy, text) {
  if (!button) return;
  if (busy) {
    button.dataset.originalText = button.textContent;
    button.textContent = text;
    button.disabled = true;
    button.setAttribute("aria-busy", "true");
    return;
  }
  button.textContent = button.dataset.originalText || button.textContent;
  button.disabled = false;
  button.removeAttribute("aria-busy");
}

function setJsonFeedback(message = "", tone = "neutral") {
  el.jsonFeedback.textContent = message;
  el.jsonFeedback.className = `inline-status ${tone}`;
}

function setMode(mode) {
  state.mode = mode;
  state.selectedBatch = null;
  state.selectedBatchData = null;
  setJsonFeedback();
  el.modeQuarterly.classList.toggle("active", mode === "quarterly");
  el.modeAnnual.classList.toggle("active", mode === "annual");
  el.modeQuarterly.setAttribute("aria-selected", String(mode === "quarterly"));
  el.modeAnnual.setAttribute("aria-selected", String(mode === "annual"));
  el.screenTitle.textContent = mode === "quarterly" ? "Квартальное обновление" : "Годовая сверка";
  el.screenSubtitle.textContent =
    mode === "quarterly"
      ? "Подготовь источники за последний самостоятельный квартал и собери один лист «Свод»."
      : "Подготовь источники за последний полный год и собери один лист «Свод».";
  render();
  autoSelectNextBatch().catch((error) => showToast(error.message));
}

async function setProvider(provider) {
  state.provider = normalizeProvider(provider);
  localStorage.setItem("cordiant-provider", state.provider);
  state.status = null;
  state.selectedBatch = null;
  state.selectedBatchData = null;
  renderProvider();
  render();
  await refreshStatus();
}

function renderProvider() {
  const provider = providerInfo();
  el.providerSelect.value = state.provider;
  el.providerName.textContent = provider.label;
  el.providerStepLabel.textContent = provider.label;
  el.openProviderSidebar.href = provider.url;
  el.openProviderSidebar.textContent = `Открыть ${provider.label}`;
  el.openProviderSidebar.title = provider.siteLabel;
  el.openProviderPrompt.href = provider.url;
  el.openProviderPrompt.title = `Открыть ${provider.siteLabel}`;
  el.jsonSectionTitle.textContent = provider.responseLabel;
  el.jsonBox.placeholder = provider.pastePlaceholder;
  const limitText = provider.maxUploadFilesPerBatch
    ? `Лимит: ${provider.maxUploadFilesPerBatch} файлов на батч.`
    : "Лимит файлов: не режем, контролируем объем входа.";
  el.providerLimitNote.textContent = provider.limitSummary || limitText;
}

function syncPreparePolling() {
  if (isPreparing() && !state.preparePollTimer) {
    state.preparePollTimer = setInterval(() => {
      refreshStatus().catch((error) => showToast(error.message));
    }, 2000);
  }
  if (!isPreparing() && state.preparePollTimer) {
    clearInterval(state.preparePollTimer);
    state.preparePollTimer = null;
  }
}

function render() {
  const info = modeInfo();
  const preparing = isPreparing();
  const batches = preparing ? [] : info.batches || [];
  const jsonCount = preparing ? 0 : info.json_count || 0;
  const batchCount = batches.length || 0;
  const complete = batchCount > 0 && jsonCount === batchCount;
  const job = prepareJob();
  const nextBatch = batches.find((batch) => !batch.json_saved);
  const progress = batchCount ? Math.round((jsonCount / batchCount) * 100) : 0;
  const provider = providerInfo();
  const packageProvider = info.package_provider || {};

  el.currentModeLabel.textContent = modeLabel();
  el.packageState.textContent = preparing ? "готовлю" : info.package_exists ? "готов" : "нет";
  el.packageDetail.textContent = preparing ? "скачиваю источники" : info.package_path ? shortPath(info.package_path) : "пакет не создан";
  el.jsonState.textContent = `${jsonCount}/${batchCount}`;
  el.jsonDetail.textContent = batchCount ? `${progress}% закрыто` : "нет батчей";
  el.excelState.textContent = complete ? "готов" : "закрыт";
  el.excelDetail.textContent = complete ? "можно собрать" : batchCount ? "нужны все JSON" : "нужен пакет";
  el.progressFill.style.width = `${complete ? 100 : progress}%`;

  el.stepSources.textContent = preparing
    ? "скачиваю источники"
    : info.package_exists
      ? `${batchCount} батчей${packageProvider.label ? ` · ${packageProvider.label}` : ""}`
      : "не подготовлены";
  el.stepJson.textContent = `${jsonCount} сохранено`;
  el.stepExcel.textContent = preparing ? "ждёт пакет" : complete ? "готов к финалу" : batchCount ? "нужны все JSON" : "нужен пакет";
  el.batchCount.textContent = String(batchCount);

  el.prepareBtn.disabled = preparing;
  el.prepareBtn.textContent = preparing ? "Готовлю источники..." : "Подготовить источники";
  el.prepareBtn.toggleAttribute("aria-busy", preparing);
  el.revealPackageBtn.disabled = preparing || !info.package_path;
  el.applyBtn.disabled = preparing || !complete;
  el.applyBtn.title = complete ? "" : "Финальная сборка доступна только после JSON для всех батчей";

  renderWorkflow({ preparing, job, info, nextBatch, complete, provider });
  renderBatchList({ preparing, batches });
  renderBatchDetail();
  syncPreparePolling();
}

function renderWorkflow({ preparing, job, info, nextBatch, complete, provider }) {
  if (preparing) {
    el.workflowNote.textContent = "Пакет готовится: источники скачиваются, промпты и upload-папки собираются. Статус обновится автоматически.";
  } else if (job?.state === "error") {
    el.workflowNote.textContent = `Подготовка не завершилась: ${job.message || "ошибка"}`;
  } else if (info.provider_mismatch) {
    const packageProvider = info.package_provider?.label || "другого провайдера";
    el.workflowNote.textContent = `Текущий пакет подготовлен для ${packageProvider}. Для ${provider.label} нажми «Подготовить источники», чтобы пересобрать батчи под его лимит файлов.`;
  } else if (!info.package_exists) {
    el.workflowNote.textContent = "Нажми «Подготовить источники». После этого очередь батчей появится здесь.";
  } else if (nextBatch) {
    el.workflowNote.textContent = `Следующий батч: ${nextBatch.id}. Нажми «Начать батч», загрузи открытую папку в ${provider.label}, вставь промпт и сохрани ответ.`;
  } else if (complete) {
    el.workflowNote.textContent = "Все ответы сохранены. Собери итоговый Excel: на выходе будет один лист «Свод».";
  } else {
    el.workflowNote.textContent = "Пакет есть, но очередь пока пустая. Обнови статус или подготовь источники заново.";
  }
}

function renderBatchList({ preparing, batches }) {
  el.batchList.innerHTML = "";
  if (preparing) {
    el.batchList.append(emptyBatch("Пакет готовится", "Батчи появятся автоматически после подготовки источников."));
    return;
  }
  if (!batches.length) {
    el.batchList.append(emptyBatch("Нет батчей", "Сначала подготовь источники для выбранного режима."));
    return;
  }

  const fragment = document.createDocumentFragment();
  batches.forEach((batch) => {
    const item = document.createElement("button");
    item.type = "button";
    item.className = `batch-item${state.selectedBatch === batch.id ? " active" : ""}${batch.provider_limit_exceeded ? " limit-warning" : ""}`;
    item.setAttribute("aria-pressed", String(state.selectedBatch === batch.id));
    const review = batch.json_summary ? ` · review ${batch.json_summary.review}` : "";
    const saved = batch.json_saved ? "JSON сохранён" : "нет JSON";
    const limitText = batch.provider_limit ? ` / лимит ${batch.provider_limit}` : "";
    item.innerHTML = `
      <span class="batch-topline">
        <strong>${escapeHtml(batch.id)}</strong>
        <em class="${batch.json_saved ? "saved" : "pending"}">${saved}</em>
      </span>
      <span class="batch-companies">${escapeHtml((batch.companies || []).join(", "))}</span>
      <small>${batch.upload_count || 0}${limitText} файлов · ${batch.downloaded_sources || 0} источников${review}</small>
    `;
    item.addEventListener("click", () => selectBatch(batch.id));
    fragment.append(item);
  });
  el.batchList.append(fragment);
}

function emptyBatch(title, text) {
  const node = document.createElement("div");
  node.className = "empty-state";
  node.innerHTML = `<strong>${escapeHtml(title)}</strong><span>${escapeHtml(text)}</span>`;
  return node;
}

function renderBatchDetail() {
  if (!state.selectedBatchData) {
    el.batchTitle.textContent = "Выбери батч";
    el.batchMeta.textContent = "";
    el.batchStatus.textContent = "нет JSON";
    el.batchStatus.className = "status-pill neutral";
    el.promptBox.value = "";
    el.jsonBox.value = "";
    el.startBatchBtn.disabled = true;
    return;
  }

  const batch = state.selectedBatchData.batch;
  const summary = batch.json_summary;
  const summaryText = summary
    ? ` · ${summary.non_null}/${summary.facts} значений, review ${summary.review}`
    : "";
  const providerLimitText = batch.provider_limit
    ? ` · лимит ${batch.provider_limit} файлов`
    : "";
  el.batchTitle.textContent = batch.id;
  el.batchMeta.textContent = `${(batch.companies || []).join(", ")} · ${batch.upload_count || 0} файлов${providerLimitText} · ${batch.prompt_chars || 0} знаков промпта${summaryText}`;
  el.batchStatus.textContent = batch.json_saved ? "JSON сохранён" : "нет JSON";
  if (batch.provider_limit_exceeded) {
    el.batchStatus.textContent = "выше лимита";
  }
  el.batchStatus.className = `status-pill ${batch.provider_limit_exceeded ? "warning" : batch.json_saved ? "ok" : "neutral"}`;
  el.startBatchBtn.disabled = !state.selectedBatchData.upload_folder_path || !state.selectedBatchData.prompt;
}

async function refreshStatus({ autoSelect = true } = {}) {
  state.status = await api("/api/status", {
    method: "POST",
    body: JSON.stringify({ provider: state.provider }),
  });
  syncProvidersFromStatus(state.status);
  if (!selectedBatchStillExists()) {
    state.selectedBatch = null;
    state.selectedBatchData = null;
  }
  if (autoSelect && !state.selectedBatch && !isPreparing()) {
    const nextBatchId = nextUnsavedBatchId();
    if (nextBatchId) {
      await selectBatch(nextBatchId, { refresh: false });
      return;
    }
  }
  render();
}

async function selectBatch(batchId, { refresh = true } = {}) {
  const payload = await api("/api/batch", {
    method: "POST",
    body: JSON.stringify({ mode: state.mode, provider: state.provider, batch: batchId }),
  });
  state.selectedBatch = batchId;
  state.selectedBatchData = payload;
  el.promptBox.value = payload.prompt;
  el.jsonBox.value = "";
  setJsonFeedback();
  if (refresh) {
    state.status = await api("/api/status", {
      method: "POST",
      body: JSON.stringify({ provider: state.provider }),
    });
    syncProvidersFromStatus(state.status);
  }
  render();
}

async function autoSelectNextBatch() {
  if (!state.status || isPreparing() || state.selectedBatch) return;
  const nextBatchId = nextUnsavedBatchId();
  if (nextBatchId) {
    await selectBatch(nextBatchId, { refresh: false });
  }
}

async function prepareSources() {
  try {
    const payload = await api("/api/prepare", {
      method: "POST",
      body: JSON.stringify({ mode: state.mode, provider: state.provider }),
    });
    state.status = payload;
    state.selectedBatch = null;
    state.selectedBatchData = null;
    showToast("Подготовка источников запущена");
    render();
  } catch (error) {
    showToast(error.message);
  }
}

async function saveJson() {
  if (!state.selectedBatch) {
    showToast("Сначала выбери батч");
    setJsonFeedback("Сначала выбери батч в очереди.", "error");
    return;
  }
  const text = el.jsonBox.value.trim();
  if (!text) {
    const provider = providerInfo();
    showToast(`Вставь ответ ${provider.label}`);
    setJsonFeedback(`Вставь полный ответ ${provider.label} в это поле.`, "error");
    return;
  }
  const savedBatch = state.selectedBatch;
  setBusy(el.saveJsonBtn, true, "Сохраняю...");
  setJsonFeedback("Проверяю и сохраняю ответ...", "neutral");
  try {
    const payload = await api("/api/save-json", {
      method: "POST",
      body: JSON.stringify({ mode: state.mode, provider: state.provider, batch: savedBatch, text }),
    });
    state.status = payload;
    el.jsonBox.value = "";
    const nextBatchId = nextUnsavedBatchId(payload);
    if (nextBatchId) {
      showToast("Ответ сохранён. Открываю следующий батч.");
      setJsonFeedback(`Сохранено для ${savedBatch}. Открываю ${nextBatchId}.`, "success");
      await selectBatch(nextBatchId, { refresh: false });
    } else {
      state.selectedBatch = null;
      state.selectedBatchData = null;
      render();
      setJsonFeedback("Все ответы сохранены. Можно собрать итоговый Excel.", "success");
      showToast("Все ответы сохранены");
    }
  } catch (error) {
    showToast(error.message);
    setJsonFeedback(error.message, "error");
  } finally {
    setBusy(el.saveJsonBtn, false);
  }
}

async function applyExcel() {
  setBusy(el.applyBtn, true, "Собираю...");
  try {
    const payload = await api("/api/apply", {
      method: "POST",
      body: JSON.stringify({ mode: state.mode, provider: state.provider }),
    });
    state.lastOutputPath = payload.output_path;
    el.outputText.textContent = payload.output_path || "Итоговый Excel создан.";
    el.revealOutputBtn.disabled = !payload.output_path;
    showToast("Итоговый лист Excel собран");
    await refreshStatus();
  } catch (error) {
    showToast(error.message);
  } finally {
    setBusy(el.applyBtn, false);
  }
}

async function reveal(path) {
  if (!path) return;
  try {
    await api("/api/reveal", {
      method: "POST",
      body: JSON.stringify({ path }),
    });
  } catch (error) {
    showToast(error.message);
  }
}

async function startBatch() {
  if (!state.selectedBatchData) {
    showToast("Сначала выбери батч");
    return;
  }
  setBusy(el.startBatchBtn, true, "Готовлю...");
  try {
    await api("/api/start-batch", {
      method: "POST",
      body: JSON.stringify({ mode: state.mode, provider: state.provider, batch: state.selectedBatch }),
    });
    showToast(`Промпт скопирован. Загрузи файлы в ${providerInfo().label} из открытой папки.`);
  } catch (error) {
    showToast(error.message);
  } finally {
    setBusy(el.startBatchBtn, false);
  }
}

function shortPath(path) {
  if (!path) return "";
  const parts = String(path).split("/");
  return parts.slice(-2).join("/");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

el.modeQuarterly.addEventListener("click", () => setMode("quarterly"));
el.modeAnnual.addEventListener("click", () => setMode("annual"));
el.providerSelect.addEventListener("change", (event) => {
  setProvider(event.target.value).catch((error) => showToast(error.message));
});
el.refreshBtn.addEventListener("click", () => refreshStatus().catch((error) => showToast(error.message)));
el.prepareBtn.addEventListener("click", prepareSources);
el.applyBtn.addEventListener("click", applyExcel);
el.revealPackageBtn.addEventListener("click", () => reveal(modeInfo().package_path));
el.startBatchBtn.addEventListener("click", startBatch);
el.revealOutputBtn.addEventListener("click", () => reveal(state.lastOutputPath));
el.saveJsonBtn.addEventListener("click", saveJson);

document.querySelector(".mode-toggle").addEventListener("keydown", (event) => {
  if (event.key !== "ArrowLeft" && event.key !== "ArrowRight") return;
  event.preventDefault();
  const nextMode = state.mode === "quarterly" ? "annual" : "quarterly";
  setMode(nextMode);
  (nextMode === "quarterly" ? el.modeQuarterly : el.modeAnnual).focus();
});

renderProvider();
refreshStatus().catch((error) => showToast(error.message));

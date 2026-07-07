const form = document.querySelector("#uploadForm");
const fileInput = document.querySelector("#drawingFile");
const fileMeta = document.querySelector("#fileMeta");
const knowledgeSelect = document.querySelector("#knowledgeSelect");
const knowledgeFile = document.querySelector("#knowledgeFile");
const knowledgeMeta = document.querySelector("#knowledgeMeta");
const uploadKnowledgeButton = document.querySelector("#uploadKnowledgeButton");
const submitButton = document.querySelector("#submitButton");
const serviceState = document.querySelector("#serviceState");
const jobBadge = document.querySelector("#jobBadge");
const progressBar = document.querySelector("#progressBar");
const stageList = document.querySelector("#stageList");
const downloads = document.querySelector("#downloads");
const errorBox = document.querySelector("#errorBox");

const fields = {
  jobId: document.querySelector("#jobId"),
  stage: document.querySelector("#stage"),
  knowledgeFileName: document.querySelector("#knowledgeFileName"),
  connectionRows: document.querySelector("#connectionRows"),
  sheetMaterials: document.querySelector("#sheetMaterials"),
};

let pollTimer = null;

fileInput.addEventListener("change", () => {
  const file = fileInput.files[0];
  fileMeta.textContent = file ? `${file.name} · ${formatBytes(file.size)}` : "未选择文件";
});

knowledgeFile.addEventListener("change", () => {
  const file = knowledgeFile.files[0];
  knowledgeMeta.textContent = file ? `${file.name} · ${formatBytes(file.size)}` : "上传新的 .xlsx 治理文件";
});

knowledgeSelect.addEventListener("change", () => {
  fields.knowledgeFileName.textContent = currentKnowledgeName() || "-";
});

uploadKnowledgeButton.addEventListener("click", async () => {
  const file = knowledgeFile.files[0];
  if (!file) {
    showError("请选择要上传的治理文件。");
    return;
  }
  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    showError("治理文件必须是 .xlsx。");
    return;
  }

  clearError();
  uploadKnowledgeButton.disabled = true;
  uploadKnowledgeButton.textContent = "上传中";

  const data = new FormData();
  data.append("file", file);
  try {
    const response = await fetch("/api/knowledge-files", { method: "POST", body: data });
    if (!response.ok) {
      throw new Error(await response.text());
    }
    const payload = await response.json();
    await loadKnowledgeFiles(payload.item?.id);
    knowledgeFile.value = "";
    knowledgeMeta.textContent = "上传新的 .xlsx 治理文件";
  } catch (error) {
    showError(error.message);
  } finally {
    uploadKnowledgeButton.disabled = false;
    uploadKnowledgeButton.textContent = "上传并使用";
  }
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const file = fileInput.files[0];
  if (!file) {
    showError("请选择 PDF 图纸或 DOCX 技术要求文件。");
    return;
  }
  const name = file.name.toLowerCase();
  if (!name.endsWith(".pdf") && !name.endsWith(".docx")) {
    showError("目前只支持 PDF 图纸或 DOCX 技术要求文件。");
    return;
  }

  clearError();
  setBusy(true);
  setProgress(8);
  downloads.innerHTML = "";

  const data = new FormData();
  data.append("file", file);
  data.append("use_llm_extract", document.querySelector("#useLlm").checked ? "true" : "false");
  data.append("knowledge_file_id", knowledgeSelect.value || "__default__");

  try {
    const response = await fetch("/api/jobs", { method: "POST", body: data });
    if (!response.ok) {
      throw new Error(await response.text());
    }
    const payload = await response.json();
    fields.jobId.textContent = payload.job_id;
    pollJob(payload.job_id);
  } catch (error) {
    setBusy(false);
    setProgress(0);
    showError(error.message);
  }
});

async function pollJob(jobId) {
  window.clearTimeout(pollTimer);
  try {
    const response = await fetch(`/api/jobs/${jobId}`);
    if (!response.ok) {
      throw new Error(await response.text());
    }
    const job = await response.json();
    renderJob(job);
    if (job.status === "queued" || job.status === "running") {
      pollTimer = window.setTimeout(() => pollJob(jobId), 1800);
    } else {
      setBusy(false);
    }
  } catch (error) {
    setBusy(false);
    showError(error.message);
  }
}

function renderJob(job) {
  jobBadge.textContent = statusText(job.status);
  jobBadge.className = `badge ${job.status || ""}`;
  fields.jobId.textContent = job.job_id || "-";
  fields.stage.textContent = job.stage || "-";
  fields.knowledgeFileName.textContent = job.knowledge_file_name || currentKnowledgeName() || "-";

  const summary = job.summary || {};
  fields.connectionRows.textContent = valueOrDash(summary.connection_rows);
  fields.sheetMaterials.textContent = valueOrDash(summary.sheet_materials);

  if (job.status === "queued") setProgress(12);
  if (job.status === "running") setProgress(progressForStage(job.stage));
  if (job.status === "done") setProgress(100);
  if (job.status === "failed") setProgress(100);
  renderStages(job.stage, job.status);

  if (job.status === "failed") {
    showError(job.error || "任务失败");
  } else {
    clearError();
  }

  renderDownloads(job.downloads || {});
}

function renderDownloads(items) {
  downloads.innerHTML = "";
  const labels = {
    output: "下载接线表",
    requirements: "图纸需求 JSON",
    recommendations: "推荐结果 JSON",
    ocr_json: "OCR JSON",
    ocr_text: "OCR 文本",
    docx_tables: "DOCX 表格 JSON",
  };
  for (const [key, href] of Object.entries(items)) {
    const link = document.createElement("a");
    link.href = href;
    link.textContent = labels[key] || key;
    if (key !== "output") link.className = "secondary";
    downloads.appendChild(link);
  }
}

function progressForStage(stage) {
  if (!stage) return 18;
  if (stage.includes("OCR") || stage.includes("DOCX")) return 30;
  if (stage.includes("解析")) return 50;
  if (stage.includes("抽取")) return 62;
  if (stage.includes("匹配")) return 76;
  if (stage.includes("回填")) return 90;
  return 45;
}

function setBusy(isBusy) {
  submitButton.disabled = isBusy;
  submitButton.textContent = isBusy ? "处理中" : "开始生成";
}

function setProgress(value) {
  progressBar.style.width = `${Math.max(0, Math.min(100, value))}%`;
}

function renderStages(stage, status) {
  if (!stageList) return;
  const steps = ["read", "parse", "extract", "match", "fill"];
  let active = -1;
  if (stage?.includes("OCR") || stage?.includes("DOCX")) active = 0;
  if (stage?.includes("解析") && !stage?.includes("DOCX")) active = 1;
  if (stage?.includes("抽取")) active = 2;
  if (stage?.includes("匹配")) active = 3;
  if (stage?.includes("回填")) active = 4;
  if (status === "done") active = steps.length;
  if (status === "failed") active = Math.max(active, 0);

  for (const item of stageList.querySelectorAll("li")) {
    const index = steps.indexOf(item.dataset.step);
    item.className = "";
    if (active === steps.length || index < active) item.classList.add("done");
    if (index === active && status !== "done") item.classList.add("active");
  }
}

function showError(message) {
  errorBox.hidden = false;
  errorBox.textContent = message;
}

function clearError() {
  errorBox.hidden = true;
  errorBox.textContent = "";
}

function statusText(status) {
  const map = {
    queued: "排队中",
    running: "处理中",
    done: "已完成",
    failed: "失败",
  };
  return map[status] || "未开始";
}

function valueOrDash(value) {
  return value === undefined || value === null || value === "" ? "-" : String(value);
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

async function checkHealth() {
  try {
    await fetch("/api/health");
    serviceState.textContent = "服务正常";
    serviceState.className = "service-state ok";
  } catch {
    serviceState.textContent = "服务不可用";
    serviceState.className = "service-state";
  }
}

async function loadKnowledgeFiles(selectedId) {
  const response = await fetch("/api/knowledge-files");
  if (!response.ok) {
    throw new Error(await response.text());
  }
  const payload = await response.json();
  const items = payload.items || [];
  knowledgeSelect.innerHTML = "";
  for (const item of items) {
    const option = document.createElement("option");
    option.value = item.id;
    option.dataset.name = item.name;
    const prefix = item.is_default ? "默认" : "上传";
    option.textContent = `${prefix} · ${item.name} · ${formatBytes(item.size || 0)}`;
    knowledgeSelect.appendChild(option);
  }
  if (selectedId && items.some((item) => item.id === selectedId)) {
    knowledgeSelect.value = selectedId;
  }
  fields.knowledgeFileName.textContent = currentKnowledgeName() || "-";
}

function currentKnowledgeName() {
  const selected = knowledgeSelect.options[knowledgeSelect.selectedIndex];
  if (!selected) return "";
  return selected.dataset.name || selected.textContent;
}

async function init() {
  await checkHealth();
  try {
    await loadKnowledgeFiles();
  } catch (error) {
    showError(error.message);
  }
}

init();

const form = document.querySelector("#uploadForm");
const fileInput = document.querySelector("#fileInput");
const fileName = document.querySelector("#fileName");
const button = document.querySelector("#generateButton");
const dropZone = document.querySelector(".drop-zone");
const result = document.querySelector("#result");
const emptyState = document.querySelector("#emptyState");

function setFile(file) {
  if (!file) return;
  const dt = new DataTransfer();
  dt.items.add(file);
  fileInput.files = dt.files;
  fileName.textContent = file.name;
  button.disabled = false;
  showEmpty();
}

function showEmpty() {
  result.classList.add("hidden");
  result.innerHTML = "";
  emptyState.classList.remove("hidden");
}

function listBlock(title, items, type) {
  if (!items || items.length === 0) return "";
  return `
    <div class="notice ${type}">
      <strong>${title}</strong>
      <ul>${items.map((item) => `<li>${item}</li>`).join("")}</ul>
    </div>
  `;
}

function showResult(payload) {
  emptyState.classList.add("hidden");
  result.classList.remove("hidden");

  if (payload.ok) {
    result.innerHTML = `
      <div class="summary">已生成 ${payload.count} 份 PDF 工资单。</div>
      <a class="download-link" href="${payload.downloadUrl}">下载 ${payload.fileName}</a>
      ${listBlock("生成提示", payload.warnings, "warning")}
    `;
  } else {
    result.innerHTML = `
      ${listBlock("生成失败", payload.errors || ["未知错误"], "error")}
      ${listBlock("生成提示", payload.warnings, "warning")}
    `;
  }
}

fileInput.addEventListener("change", () => {
  const file = fileInput.files[0];
  if (file) {
    fileName.textContent = file.name;
    button.disabled = false;
    showEmpty();
  }
});

dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("dragging");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("dragging");
});

dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("dragging");
  setFile(event.dataTransfer.files[0]);
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!fileInput.files[0]) return;

  button.disabled = true;
  button.textContent = "生成中...";

  const data = new FormData();
  data.append("file", fileInput.files[0]);

  try {
    const response = await fetch("/api/generate", { method: "POST", body: data });
    const payload = await response.json();
    showResult(payload);
  } catch (error) {
    showResult({ ok: false, errors: [`请求失败：${error.message}`], warnings: [] });
  } finally {
    button.disabled = false;
    button.textContent = "生成 ZIP";
  }
});

import "./app.css";
import { parseText, renderDocxBlob, validateDocument } from "./core/index.js";
import type { DocumentNode, DocumentSpec, TableNode, TemplateId } from "./core/types.js";

const SAMPLE_TEXT = `关于项目现场检查情况的报告

一、检查概况
本次检查于 2026 年 8 月 12 日进行，参与人员包括项目负责人、现场工程师和客户代表。
姓名：张三
项目：移动办公文档实验
日期：2026 年 8 月 12 日

二、发现事项
| 事项 | 负责人 | 截止日期 |
| --- | --- | --- |
| 补充现场照片 | 李四 | 2026 年 8 月 14 日 |
| 确认整改结果 | 王五 | 2026 年 8 月 16 日 |

三、下一步安排
相关负责人应按期完成事项，并在移动端生成最终报告后提交。`;

interface BeforeInstallPromptEvent extends Event {
  prompt: () => Promise<void>;
  userChoice: Promise<{ outcome: "accepted" | "dismissed" }>;
}

const app = document.querySelector<HTMLDivElement>("#app");
if (!app) throw new Error("找不到应用挂载节点。");

app.innerHTML = `
  <header class="topbar">
    <a class="wordmark" href="#main">文成</a>
    <button class="install-button" id="install-button" type="button" hidden>安装</button>
  </header>
  <main class="app-shell" id="main">
    <section class="intro" aria-labelledby="page-title">
      <h1 id="page-title">粘贴内容，生成 Word</h1>
      <p>自动识别标题、正文和简单表格，再按固定办公模板排版。无需逐段调整格式。</p>
      <span class="privacy-note">内容仅在当前设备处理</span>
    </section>

    <div class="workbench">
      <section class="panel input-panel" aria-labelledby="input-heading">
        <header class="panel__header">
          <h2 id="input-heading">原始内容</h2>
        </header>
        <div class="panel__body">
          <div class="field" id="title-field">
            <label for="document-title">文档标题（可选）</label>
            <input id="document-title" type="text" maxlength="80" autocomplete="off" placeholder="留空时从正文自动识别" aria-describedby="title-help" />
            <p class="field__help" id="title-help">手动填写会覆盖自动识别结果。</p>
          </div>

          <div class="field" id="text-field">
            <label for="source-text">文本内容</label>
            <textarea id="source-text" aria-required="true" aria-describedby="text-help" placeholder="示例：粘贴带有“一、工作概况”、字段：值，或 Markdown 表格的完整文本。"></textarea>
            <p class="field__help" id="text-help">支持中文标题、制表符、Markdown 表格和连续“字段：值”。</p>
          </div>

          <div class="input-actions" aria-label="文本操作">
            <button class="text-button" id="paste-button" type="button">粘贴剪贴板</button>
            <button class="text-button" id="sample-button" type="button">载入示例</button>
            <button class="text-button" id="clear-button" type="button">清空</button>
          </div>

          <fieldset>
            <legend>办公模板</legend>
            <div class="template-options">
              <label class="template-option">
                <input type="radio" name="template" value="formal" checked />
                <span><strong>正式公文</strong><span>宋体正文、紧凑行距、蓝色分级标题</span></span>
              </label>
              <label class="template-option">
                <input type="radio" name="template" value="report" />
                <span><strong>工作报告</strong><span>微软雅黑正文、舒展行距、适合说明材料</span></span>
              </label>
            </div>
          </fieldset>
        </div>
      </section>

      <section class="panel preview-panel" aria-labelledby="preview-heading">
        <header class="panel__header">
          <h2 id="preview-heading">识别结果</h2>
        </header>
        <div class="summary" id="summary" aria-label="文档结构统计"></div>
        <div id="notice" class="notice" hidden></div>
        <article class="document-preview" id="document-preview" aria-live="polite"></article>
      </section>
    </div>
  </main>

  <footer class="foot-line"><p>文成 · 本地生成可编辑 DOCX · 首版支持两类办公模板</p></footer>

  <aside class="generate-bar" aria-label="生成文档">
    <div class="generate-bar__copy">
      <strong>准备生成 Word</strong>
      <span id="generate-hint">请先粘贴需要整理的内容</span>
    </div>
    <button class="primary-button" id="generate-button" type="button" disabled>生成 Word</button>
  </aside>
`;

const titleInput = requireElement<HTMLInputElement>("document-title");
const sourceInput = requireElement<HTMLTextAreaElement>("source-text");
const textField = requireElement<HTMLDivElement>("text-field");
const preview = requireElement<HTMLElement>("document-preview");
const summary = requireElement<HTMLDivElement>("summary");
const notice = requireElement<HTMLDivElement>("notice");
const generateButton = requireElement<HTMLButtonElement>("generate-button");
const generateHint = requireElement<HTMLSpanElement>("generate-hint");
const pasteButton = requireElement<HTMLButtonElement>("paste-button");
const sampleButton = requireElement<HTMLButtonElement>("sample-button");
const clearButton = requireElement<HTMLButtonElement>("clear-button");
const installButton = requireElement<HTMLButtonElement>("install-button");

let currentSpec: DocumentSpec | null = null;
let installPrompt: BeforeInstallPromptEvent | null = null;
let previewTimer: number | undefined;

function requireElement<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) throw new Error(`缺少界面元素：${id}`);
  return element as T;
}

function escapeHtml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function selectedTemplate(): TemplateId {
  const selected = document.querySelector<HTMLInputElement>('input[name="template"]:checked');
  return selected?.value === "report" ? "report" : "formal";
}

function tableMarkup(node: TableNode): string {
  const headers = node.hasHeader ? node.rows[0] ?? [] : node.style === "key-value" ? ["字段", "内容"] : [];
  const rows = node.hasHeader ? node.rows.slice(1) : node.rows;
  const columnCount = Math.max(node.rows[0]?.length ?? 1, 1);
  const headerMarkup = node.hasHeader
    ? `<div class="preview-table__row preview-table__row--header" style="--columns:${columnCount}">${headers
        .map((cell) => `<div class="preview-table__cell">${escapeHtml(cell)}</div>`)
        .join("")}</div>`
    : "";

  const bodyMarkup = rows
    .map(
      (row) => `<div class="preview-table__row" style="--columns:${columnCount}">${row
        .map((cell, index) => {
          const label = headers[index] || `第 ${index + 1} 列`;
          return `<div class="preview-table__cell"><span class="preview-table__label">${escapeHtml(label)}</span><span>${escapeHtml(cell)}</span></div>`;
        })
        .join("")}</div>`,
    )
    .join("");

  return `<div class="preview-table" role="table">${headerMarkup}${bodyMarkup}</div>`;
}

function nodeMarkup(node: DocumentNode): string {
  if (node.type === "heading") {
    return `<h3 class="preview-heading" data-level="${node.level}">${escapeHtml(node.text)}</h3>`;
  }
  if (node.type === "table") return tableMarkup(node);
  return `<p class="preview-paragraph">${escapeHtml(node.text)}</p>`;
}

function emptyPreview(): void {
  currentSpec = null;
  summary.innerHTML = ["标题", "正文", "表格", "表格行"]
    .map((label) => `<div class="summary__item"><span class="summary__value">0</span><span class="summary__label">${label}</span></div>`)
    .join("");
  preview.innerHTML = '<div class="preview-empty"><p>粘贴一段完整文本，识别后的文档结构会显示在这里。</p></div>';
  notice.hidden = true;
  generateButton.disabled = true;
  generateHint.textContent = "请先粘贴需要整理的内容";
  textField.dataset.state = "default";
  sourceInput.removeAttribute("aria-invalid");
}

function updatePreview(): void {
  const text = sourceInput.value.trim();
  if (!text) {
    emptyPreview();
    return;
  }

  const spec = parseText({
    text,
    title: titleInput.value.trim() || undefined,
    template: selectedTemplate(),
  });
  const validation = validateDocument(spec);
  currentSpec = spec;

  summary.innerHTML = [
    [validation.stats.headings, "标题"],
    [validation.stats.paragraphs, "正文"],
    [validation.stats.tables, "表格"],
    [validation.stats.tableRows, "表格行"],
  ]
    .map(([value, label]) => `<div class="summary__item"><span class="summary__value">${value}</span><span class="summary__label">${label}</span></div>`)
    .join("");

  preview.innerHTML = `<h2 class="document-title">${escapeHtml(spec.title)}</h2>${spec.nodes.map(nodeMarkup).join("")}`;
  generateButton.disabled = !validation.ok;
  generateHint.textContent = validation.ok
    ? `已识别 ${spec.nodes.length} 个内容块，可直接生成`
    : "识别结果存在错误，请检查输入";

  if (validation.errors.length > 0 || validation.warnings.length > 0) {
    notice.hidden = false;
    notice.dataset.tone = validation.errors.length > 0 ? "error" : "default";
    notice.textContent = [...validation.errors, ...validation.warnings].join(" ");
  } else {
    notice.hidden = true;
  }

  textField.dataset.state = validation.ok ? "success" : "error";
  if (validation.ok) {
    sourceInput.removeAttribute("aria-invalid");
  } else {
    sourceInput.setAttribute("aria-invalid", "true");
  }
}

function schedulePreview(): void {
  window.clearTimeout(previewTimer);
  previewTimer = window.setTimeout(updatePreview, 160);
}

function setButtonState(state: "default" | "loading" | "success" | "error", label: string): void {
  generateButton.dataset.state = state;
  generateButton.textContent = label;
}

function setUtilityButtonState(
  button: HTMLButtonElement,
  state: "default" | "loading" | "success" | "error",
  label: string,
  resetLabel?: string,
): void {
  button.dataset.state = state;
  button.textContent = label;
  if (resetLabel) {
    window.setTimeout(() => {
      button.dataset.state = "default";
      button.textContent = resetLabel;
    }, 1200);
  }
}

async function generateDocument(): Promise<void> {
  updatePreview();
  if (!currentSpec || generateButton.disabled) return;

  generateButton.disabled = true;
  setButtonState("loading", "正在生成");
  notice.hidden = true;
  const startedAt = performance.now();

  try {
    const blob = await renderDocxBlob(currentSpec);
    const remainingLoadingTime = 300 - (performance.now() - startedAt);
    if (remainingLoadingTime > 0) {
      await new Promise((resolve) => window.setTimeout(resolve, remainingLoadingTime));
    }
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `${safeFilename(currentSpec.title)}.docx`;
    document.body.append(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);

    setButtonState("success", "已生成 Word");
    generateHint.textContent = "文件已下载，可用 Word 或 WPS 打开";
    notice.hidden = false;
    notice.dataset.tone = "success";
    notice.textContent = "文档已在当前设备生成并开始下载。";
    window.setTimeout(() => setButtonState("default", "再次生成"), 1800);
  } catch (error) {
    setButtonState("error", "生成失败");
    generateHint.textContent = "未能生成文件，请保留内容后重试";
    notice.hidden = false;
    notice.dataset.tone = "error";
    notice.textContent = error instanceof Error ? `生成失败：${error.message}` : "生成失败：发生未知错误，请重试。";
  } finally {
    generateButton.disabled = false;
  }
}

function safeFilename(title: string): string {
  const cleaned = title.replace(/[\\/:*?"<>|]/g, "-").trim();
  return cleaned || "未命名文档";
}

sourceInput.addEventListener("input", schedulePreview);
titleInput.addEventListener("input", schedulePreview);
document.querySelectorAll<HTMLInputElement>('input[name="template"]').forEach((input) => input.addEventListener("change", updatePreview));

sourceInput.addEventListener("blur", () => {
  if (!sourceInput.value.trim()) {
    textField.dataset.state = "error";
    sourceInput.setAttribute("aria-invalid", "true");
    requireElement<HTMLParagraphElement>("text-help").textContent = "文本内容为空。请粘贴需要整理的完整内容。";
  }
});

sourceInput.addEventListener("input", () => {
  if (sourceInput.value.trim()) {
    requireElement<HTMLParagraphElement>("text-help").textContent = "支持中文标题、制表符、Markdown 表格和连续“字段：值”。";
  }
});

pasteButton.addEventListener("click", async () => {
  pasteButton.disabled = true;
  setUtilityButtonState(pasteButton, "loading", "读取中");
  try {
    const text = await navigator.clipboard.readText();
    if (!text.trim()) throw new Error("剪贴板中没有可用文本。请复制内容后重试。");
    sourceInput.value = text;
    updatePreview();
    sourceInput.focus({ preventScroll: true });
    setUtilityButtonState(pasteButton, "success", "已粘贴", "粘贴剪贴板");
  } catch (error) {
    notice.hidden = false;
    notice.dataset.tone = "error";
    notice.textContent = error instanceof Error ? error.message : "无法读取剪贴板。请长按输入框后选择“粘贴”。";
    sourceInput.focus({ preventScroll: true });
    setUtilityButtonState(pasteButton, "error", "无法粘贴", "粘贴剪贴板");
  } finally {
    pasteButton.disabled = false;
  }
});

sampleButton.addEventListener("click", () => {
  sourceInput.value = SAMPLE_TEXT;
  titleInput.value = "";
  updatePreview();
  setUtilityButtonState(sampleButton, "success", "已载入", "载入示例");
});

clearButton.addEventListener("click", () => {
  sourceInput.value = "";
  titleInput.value = "";
  emptyPreview();
  sourceInput.focus({ preventScroll: true });
  setUtilityButtonState(clearButton, "success", "已清空", "清空");
});

generateButton.addEventListener("click", generateDocument);

window.addEventListener("beforeinstallprompt", (event) => {
  event.preventDefault();
  installPrompt = event as BeforeInstallPromptEvent;
  installButton.hidden = false;
});

installButton.addEventListener("click", async () => {
  if (!installPrompt) return;
  installButton.disabled = true;
  setUtilityButtonState(installButton, "loading", "安装中");
  await installPrompt.prompt();
  const choice = await installPrompt.userChoice;
  setUtilityButtonState(
    installButton,
    choice.outcome === "accepted" ? "success" : "default",
    choice.outcome === "accepted" ? "已安装" : "安装",
  );
  installPrompt = null;
  installButton.hidden = true;
  installButton.disabled = false;
});

if ("serviceWorker" in navigator && import.meta.env.PROD) {
  window.addEventListener("load", () => {
    void navigator.serviceWorker.register(`${import.meta.env.BASE_URL}sw.js`);
  });
}

if (new URLSearchParams(window.location.search).get("demo") === "1") {
  sourceInput.value = SAMPLE_TEXT;
  updatePreview();
} else {
  emptyPreview();
}

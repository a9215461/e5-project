/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
import { formatMessage, templates, formatByTemplateId, generateSampleBlock, joinTemplatesWithSeparator, getTemplateById, templatesAsJSON } from "./utils";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    // hide sideload message and show app
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // primary action: 插入文本
    document.getElementById("run").onclick = run;

    // preview action: 在状态区显示将要插入的文本（不写入文档）
    const previewBtn = document.getElementById("preview");
    if (previewBtn) {
      previewBtn.onclick = preview;
    }

    // 模板下拉与插入模板按钮
    const tplSelect = document.getElementById("template-select");
    if (tplSelect) {
      // 填充模板选项
      templates.forEach((t) => {
        const opt = document.createElement("option");
        opt.value = t.id;
        opt.textContent = t.title;
        tplSelect.appendChild(opt);
      });
      // 当选择变化时，自动更新预览
      tplSelect.onchange = () => {
        updatePreviewFromSelect();
      };
    }

    const insertTplBtn = document.getElementById("insert-template");
    if (insertTplBtn) {
      insertTplBtn.onclick = insertTemplate;
    }
    const insertSampleBtn = document.getElementById("insert-sample");
    if (insertSampleBtn) {
      insertSampleBtn.onclick = insertSample;
    }
    const insertIndexBtn = document.getElementById("insert-templates-index");
    if (insertIndexBtn) {
      insertIndexBtn.onclick = insertTemplatesIndex;
    }
    const copyBtn = document.getElementById("copy-template");
    if (copyBtn) {
      copyBtn.onclick = copyCurrentTemplate;
    }
    const exportBtn = document.getElementById("export-templates");
    if (exportBtn) {
      exportBtn.onclick = exportTemplatesJson;
    }
  }
});

/**
 * 将格式化后的文本写入当前选区
 */
export async function run() {
  const statusEl = document.getElementById("status");
  const options = { coercionType: Office.CoercionType.Text };
  const content = formatMessage("Hello from e5-project");

  try {
    statusEl.textContent = "正在写入文本...";
    await Office.context.document.setSelectedDataAsync(content, options);
    statusEl.textContent = `已插入: "${content}"`;
  } catch (err) {
    console.error(err);
    statusEl.textContent = "插入失败: " + (err && err.message ? err.message : String(err));
  }
}

/** 导出当前模板集合为 JSON 文件并触发下载 */
export function exportTemplatesJson() {
  const statusEl = document.getElementById("status");
  try {
    const json = templatesAsJSON();
    const blob = new Blob([json], { type: "application/json;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "e5-project-templates.json";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    statusEl.textContent = "已生成模板导出文件。";
  } catch (err) {
    console.error(err);
    statusEl.textContent = "导出模板失败：" + (err && err.message ? err.message : String(err));
  }
}

/** 插入所有模板的索引/说明到当前选区，便于快速把模板参考写入文档 */
export async function insertTemplatesIndex() {
  const statusEl = document.getElementById("status");
  const options = { coercionType: Office.CoercionType.Text };
  const content = joinTemplatesWithSeparator("\n\n---\n\n");

  try {
    statusEl.textContent = "正在插入模板索引...";
    await Office.context.document.setSelectedDataAsync(content, options);
    statusEl.textContent = `已插入模板索引（共 ${templates.length} 个模板）`;
  } catch (err) {
    console.error(err);
    statusEl.textContent = "插入模板索引失败: " + (err && err.message ? err.message : String(err));
  }
}

/** 插入由 generateSampleBlock 生成的多行示例文本 */
export async function insertSample() {
  const statusEl = document.getElementById("status");
  const options = { coercionType: Office.CoercionType.Text };
  // 生成 20 行示例文本
  const content = generateSampleBlock(20);

  try {
    statusEl.textContent = "正在插入多行示例...";
    await Office.context.document.setSelectedDataAsync(content, options);
    statusEl.textContent = `已插入多行示例（共 20 行）`;
  } catch (err) {
    console.error(err);
    statusEl.textContent = "插入多行示例失败: " + (err && err.message ? err.message : String(err));
  }
}

/**
 * 在状态区预览要插入的文本（不修改文档）
 */
export function preview() {
  const statusEl = document.getElementById("status");
  // 如果存在模板下拉，则使用选中的模板来预览；否则使用简单的预览
  const tplSelect = document.getElementById("template-select");
  let previewText;
  if (tplSelect && tplSelect.value) {
    previewText = formatByTemplateId(tplSelect.value);
  } else {
    previewText = formatMessage("预览: Hello from e5-project");
  }
  statusEl.textContent = `预览文本: "${previewText}"`;
}

/** 将所选模板插入到当前选区 */
export async function insertTemplate() {
  const statusEl = document.getElementById("status");
  const tplSelect = document.getElementById("template-select");
  const options = { coercionType: Office.CoercionType.Text };
  const id = tplSelect && tplSelect.value ? tplSelect.value : "simple";
  const content = formatByTemplateId(id);

  try {
    statusEl.textContent = "正在插入模板...";
    await Office.context.document.setSelectedDataAsync(content, options);
    statusEl.textContent = `已插入模板: "${content}"`;
  } catch (err) {
    console.error(err);
    statusEl.textContent = "模板插入失败: " + (err && err.message ? err.message : String(err));
  }
}

function updatePreviewFromSelect() {
  const previewBtn = document.getElementById("preview");
  // 复用 preview 的逻辑，直接触发 preview 更新状态
  if (previewBtn) preview();
}

/** 复制当前选中的模板文本到剪贴板 */
export async function copyCurrentTemplate() {
  const statusEl = document.getElementById("status");
  const tplSelect = document.getElementById("template-select");
  if (!tplSelect || !tplSelect.value) {
    statusEl.textContent = "请先选择一个模板再复制。";
    return;
  }
  const tpl = getTemplateById(tplSelect.value);
  if (!tpl) {
    statusEl.textContent = "未找到所选模板。";
    return;
  }
  const text = formatMessage(tpl.text);
  // 尝试使用 navigator.clipboard，回退到旧方法
  try {
    if (navigator && navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(text);
    } else {
      // 创建临时 textarea 回退
      const ta = document.createElement('textarea');
      ta.value = text;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      document.body.removeChild(ta);
    }
    statusEl.textContent = "已复制模板到剪贴板。";
    statusEl.classList.add('copy-success');
    setTimeout(() => statusEl.classList.remove('copy-success'), 2500);
  } catch (err) {
    console.error(err);
    statusEl.textContent = "复制失败：" + (err && err.message ? err.message : String(err));
  }
}

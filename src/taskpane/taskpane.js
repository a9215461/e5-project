/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
import { formatMessage } from "./utils";

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

/**
 * 在状态区预览要插入的文本（不修改文档）
 */
export function preview() {
  const statusEl = document.getElementById("status");
  const previewText = formatMessage("预览: Hello from e5-project");
  statusEl.textContent = `预览文本: "${previewText}"`;
}

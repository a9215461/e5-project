/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  const statusEl = document.getElementById("status");
  const options = { coercionType: Office.CoercionType.Text };
  const timestamp = new Date().toLocaleString();
  const content = `Hello from e5-project @ ${timestamp}`;

  try {
    statusEl.textContent = "正在写入文本...";
    await Office.context.document.setSelectedDataAsync(content, options);
    statusEl.textContent = `已插入: "${content}"`;
  } catch (err) {
    console.error(err);
    statusEl.textContent = "插入失败: " + err.message;
  }
}

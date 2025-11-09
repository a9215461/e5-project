/*
 * 小工具：格式化要插入的文本
 */

export function formatMessage(prefix = "") {
  const now = new Date();
  // 使用 yyyy-MM-dd HH:mm:ss 格式（本地时间）
  const pad = (v) => String(v).padStart(2, "0");
  const ts = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
  if (prefix) {
    return `${prefix} — ${ts}`;
  }
  return `Created at ${ts}`;
}

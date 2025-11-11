/*
 * 小工具：时间格式与模板集合
 * 目的：提供复用的时间戳格式化函数和示例模板以供任务窗格使用。
 */

/** 把 Date 对象格式化为 yyyy-MM-dd HH:mm:ss */
export function formatTimestamp(d = new Date()) {
  const pad = (v) => String(v).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

/**
 * 返回针对插入操作的最终字符串。
 * 如果提供了 template（字符串，含可替换占位符 {timestamp}），则替换占位符。
 */
export function formatMessage(template = "Created at {timestamp}") {
  const ts = formatTimestamp(new Date());
  return template.replace(/\{timestamp\}/g, ts);
}

/** 示例模板集合：可以在任务窗格中扩展或本地化 */
export const templates = [
  { id: "simple", title: "简单消息", text: "示例文本 — {timestamp}" },
  { id: "note", title: "备忘条目", text: "[备忘] 请注意：{timestamp} — 重要事项" },
  { id: "signature", title: "带签名的模板", text: "此处为自动插入内容。\n生成时间：{timestamp}\n— e5-project" },
  { id: "custom", title: "自定义（空白）", text: "自定义文本 — {timestamp}" }
];

/** 根据模板 id 查找并格式化，如果找不到则返回默认格式化文本 */
export function formatByTemplateId(id) {
  const t = templates.find((x) => x.id === id);
  if (!t) return formatMessage();
  return formatMessage(t.text);
}

/**
 * 返回当前可用模板的简短信息（id 与 title），方便在 UI 或测试中展示
 */
export function listTemplates() {
  return templates.map((t) => ({ id: t.id, title: t.title }));
}

// 额外添加一个简洁的横幅模板，便于测试插入多行文本的效果
templates.push({ id: "banner", title: "横幅模板", text: "=== 系统通知：{timestamp} ===" });

/**
 * 生成一个多行的示例块，包含行号与时间戳，便于测试插入多行文本行为。
 * 返回字符串，默认生成 20 行变体，每行以行号开头。
 */
export function generateSampleBlock(lines = 20) {
  const now = formatTimestamp(new Date());
  const header = `示例块 — 生成于 ${now}`;
  const parts = [header, ""]; // 空行分隔头部
  for (let i = 1; i <= lines; i++) {
    parts.push(`${i.toString().padStart(2, "0")}. 这是第 ${i} 行示例文本。`);
  }
  parts.push("\n--- 结束 ---");
  return parts.join("\n");
}

/**
 * 小工具：把模板列表转换成一段带索引的文本，方便在文档中直接查看可用模板
 */
export function templatesAsText() {
  return templates.map((t, idx) => `${idx + 1}. ${t.title} (id=${t.id})`).join("\n");
}

/**
 * 返回模板的统计信息：总数、最长标题长度等，便于在 UI 中展示统计数据。
 */
export function templatesInfo() {
  const count = templates.length;
  const longest = templates.reduce((acc, t) => Math.max(acc, t.title.length), 0);
  const ids = templates.map((t) => t.id);
  return { count, longestTitleLength: longest, ids };
}

/**
 * 将所有模板合并为一个带分隔符的文本块，便于一次性插入文档作为参考。
 */
export function joinTemplatesWithSeparator(sep = "\n---\n") {
  return templates.map((t) => `【${t.title}】\n${t.text}`).join(sep);
}




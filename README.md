# e5-project PowerPoint 加载项

一个使用 Office JavaScript API 的简易 PowerPoint 任务窗格 (Task Pane) 加载项示例。点击界面中的“插入文本”按钮，会在当前选中位置写入带时间戳的字符串，帮助你快速验证加载项的运行与权限配置。

## ✨ 功能概览

- 任务窗格 UI（HTML/CSS/JS），自定义按钮与状态区域。
- 使用 `Office.context.document.setSelectedDataAsync` 向选区写入文本。
- Webpack 构建与调试脚本，支持开发/生产模式。
- `manifest.xml` 描述加载项元数据与按钮入口。

## 🗂 项目结构（节选）

```
manifest.xml            # 加载项清单
package.json            # NPM 脚本与依赖
webpack.config.js       # 打包配置
src/
    taskpane/
        taskpane.html       # 任务窗格界面
        taskpane.css        # 任务窗格样式
        taskpane.js         # 业务逻辑与 Office API 调用
    commands/             # 加载项命令（功能入口示例）
assets/                 # 图标与静态资源
```

## 🚀 快速开始

### 1. 安装依赖
```powershell
npm install
```

### 2. 开发调试（桌面版 PowerPoint）
使用 VS Code 任务或直接执行：
```powershell
npm run start -- desktop --app powerpoint
```
脚本会启动本地服务器并自动侧载 (sideload) 加载项。首次运行若出现证书或权限提示，请按照终端指引操作。

### 3. 构建
开发构建：
```powershell
npm run build:dev
```
生产构建：
```powershell
npm run build
```

### 4. 停止调试
```powershell
npm run stop
```

## 模板功能（新增）

在任务窗格中你现在可以看到以下新按钮：

- 选择模板（下拉框）
- 插入模板：将所选模板格式化后插入当前选区
- 预览文本：在状态区查看将要插入的内容（不写入文档）
- 插入多行示例：插入一个默认 20 行的示例文本块，便于测试换行与多行插入
- 插入模板索引：将所有模板的标题与内容合并为一段参考文本并插入到文档

复制模板功能：

- 在任务窗格中点击“复制模板”可将当前选择的模板文本复制到系统剪贴板，便于粘贴到其他文档或聊天窗口。
- 如果你的浏览器/宿主支持 `navigator.clipboard`，会使用现代 API；否则会回退到兼容性方法。

使用步骤：选择模板 -> 点击“插入模板”或点击“预览文本”查看效果；点击“复制模板”将模板内容复制到剪贴板。要插入所有模板说明，可使用“插入模板索引”。

调试提示：

- 复制到剪贴板需要在安全上下文（HTTPS）或受信任的 WebView 环境中执行；如果复制失败，请检查宿主环境的剪贴板权限。
- 插入多行内容（例如“插入多行示例”）在 PowerPoint 中会把整段文本插入当前选中的文本框中，注意选区类型。

导出模板（JSON）：

- 在任务窗格中点击“导出模板(JSON)”会把当前模板集合导出为一个名为 `e5-project-templates.json` 的文件，便于备份或跨项目导入。
- 导出的 JSON 文件是一个包含 `templates` 数组的漂亮（pretty-printed）JSON，包含每个模板的 `id`、`title` 和 `text` 字段。
- 下载会在浏览器/宿主环境中触发文件保存对话框；如果保存失败，请检查宿主对 blob/url 下载的支持。

插入编号列表：

- 新增“插入编号列表”按钮，用于在任务窗格中插入一段带序号的多行文本（默认 40 行），便于在 PowerPoint 中测试多行插入、换行与样式表现。
- 如需修改数量，可在 `src/taskpane/taskpane.js` 中调整 `generateNumberedList` 的参数（默认在 UI 中是 40）。

插入模板表格与示例日志：

- 新增“插入模板表格”按钮，会把当前模板集合格式化为一个 Markdown 风格的表格并插入到当前选区，表格包含序号、模板 id、模板标题及示例文本（被截断以便展示）。
- 插入动作同时会在表格后附加若干示例日志条目（默认 10 条），便于测试批量文本插入以及查阅日志的场景。
- 该功能使用 `src/taskpane/utils.js` 中的 `templatesAsMarkdownTable()` 与 `generateLogEntries()` 方法生成文本；如需调整样式或样例长度，可在相应文件中修改参数。




## 🧪 试用步骤
1. 打开 PowerPoint，确保加载项已侧载（任务窗格可见）。
2. 在幻灯片上建立一个文本占位选区（比如点击一个文本框）。
3. 在任务窗格点击“插入文本”。
4. 观察选区与状态栏文字是否更新为带时间戳的内容。

## 🛠 常用脚本速览

| 命令 | 作用 |
|------|------|
| `npm run start` | 启动调试并侧载清单 |
| `npm run stop` | 停止调试并清理侧载 |
| `npm run build:dev` | 开发模式打包（含 SourceMap） |
| `npm run build` | 生产模式打包 |
| `npm run watch` | 监听源码增量构建 |
| `npm run lint` | 代码规范检查 |
| `npm run lint:fix` | 自动修复可处理的问题 |

## 🔧 常见问题 (FAQ)

**Q: 任务窗格没有显示？**  
请确认 PowerPoint 已正确侧载，或重新运行 `npm run start` 并关闭所有旧的 PowerPoint 进程后再试。

**Q: 写入文本失败？**  
确保当前焦点在可写入的文本形状/占位符中；某些对象（图片等）不可直接写入文本。

**Q: 如何修改插入内容？**  
编辑 `src/taskpane/taskpane.js` 中 `run()` 函数的 `content` 变量即可。

## 📄 Manifest 清单提示
修改 `manifest.xml` 后请重新侧载（停止再启动），并可执行：
```powershell
npm run validate
```
验证清单格式与必填字段。

## 🧩 后续可扩展方向
- 添加自定义功能按钮与 Ribbon 分组。
- 与后端服务（Graph API 等）集成，动态获取内容。
- 增加单元测试与更严谨的错误处理。

## 📜 许可证
本项目采用 MIT License，详见仓库 LICENSE（若缺失可自行添加）。

## 🙌 贡献 & 反馈
欢迎 Fork 与提 Issue 改进项目。若需了解更多 Office 加载项开发，可参考官方文档：
https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins

---
更新说明：近期对任务窗格 UI 做了本地化与交互增强，并重写了本 README 以聚焦当前项目实际功能。
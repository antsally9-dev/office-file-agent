# Office File Agent

一个用于练习 Agent 工具调用、同时可以完成移动办公最小闭环的项目。输入普通文本和简单表格，输出可编辑的 Word `.docx`。

在线试用：[https://antsally9-dev.github.io/office-file-agent/](https://antsally9-dev.github.io/office-file-agent/)

## 当前能力

- 自动识别文档标题和三级中文/Markdown 标题；
- 识别制表符、Markdown 竖线表格；
- 识别连续的“字段：值”信息表；
- 提供 `formal` 和 `report` 两个确定性模板；
- 在渲染前检查空文档、异常列数和超宽表格；
- 支持确定性 CLI，也支持通过 OpenAI Responses API 让模型规划工具调用。

## 当前状态

- 移动优先 PWA、文本解析、校验、DOCX 渲染和 Agent CLI 已完成；
- 自动化测试覆盖标题、正文、两类表格和异常行补齐；
- 示例文件已用本机 Word 实际打开、导出并逐页检查；
- PWA 已在 320、375、414、768 和 1280px 浏览器视口完成检查；
- 浏览器端已实际下载并验证 DOCX 文件。

## 本地运行

```powershell
pnpm install
pnpm test
pnpm dev
```

开发服务器会监听局域网地址。同一 Wi-Fi 下，可在手机浏览器中打开终端显示的 Network 地址。

生成可部署的静态版本：

```powershell
pnpm build
pnpm preview
```

`dist` 目录可以部署到任意支持 HTTPS 的静态托管服务。HTTPS 上线后，剪贴板读取、PWA 安装和离线缓存才能获得完整的浏览器支持。

## Agent CLI

```powershell
pnpm agent -- --input examples/sample.txt --output outputs/sample.docx --trace
```

模型规划模式是可选的：

```powershell
$env:OPENAI_API_KEY="你的 API Key"
pnpm agent -- --input examples/sample.txt --output outputs/model-sample.docx --model "你有权限使用的模型名" --trace
```

没有 API Key 时，确定性 Agent 仍会依次执行 `parse_text → validate_document → render_docx`。

生成后的示例位于 `outputs/sample.docx`。

## 设计边界

首版不做复杂合并单元格、图片表格 OCR、多人协作、自由式富文本编辑或企业模板管理。模型只负责任务规划；DOCX 格式始终由确定性代码生成。

模型工具循环按照 [OpenAI 官方 Function calling 文档](https://developers.openai.com/api/docs/guides/function-calling) 实现。

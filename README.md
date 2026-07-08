# AI 驱动跨境电商表格处理工具

## 基础目录结构

```text
.
├── app
│   ├── api
│   │   └── generate-excel
│   │       └── route.ts
│   ├── globals.css
│   ├── layout.tsx
│   └── page.tsx
├── .env.example
├── next.config.mjs
├── package.json
├── postcss.config.mjs
├── tailwind.config.ts
└── tsconfig.json
```

## 工作流

1. 前端上传源 Excel、CSV、图片，并输入附加规则。
2. 前端解析源表格为文本、图片为 Data URL、模板首个非空行作为表头。
3. `/api/generate-excel` 只接收预处理数据、图片、用户规则、目标表头和模型名，不再上传完整模板文件。
4. 服务端通过 OpenAI-compatible 中转格式调用云雾中转模型。
5. 服务端解析 AI 返回的 JSON 数组并返回给浏览器。
6. 浏览器用本地选择的目标模板和 `exceljs` 写入数据并触发 `.xlsx` 下载。

这个流程可以避免 Vercel API Route 因大模板文件触发 `HTTP 413 Payload Too Large`。

## 环境变量

复制 `.env.example` 为 `.env.local`，填入实际云雾中转 Key。前台可在 `gemini-3.5-flash` 和 `gemini-3.1-pro-preview` 之间切换，服务端会对白名单模型做二次校验。

```bash
AI_BASE_URL=https://yunwu.ai
AI_API_KEY=your_key
AI_MODEL=gemini-3.5-flash
```

如果 `AI_BASE_URL` 只填写域名，服务端会自动补齐为 OpenAI-compatible 的 `/v1/chat/completions` 请求地址。

## 启动

```bash
npm install
npm run dev
```

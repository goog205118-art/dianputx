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
3. `/api/generate-excel` 接收预处理数据和模板文件。
4. 服务端通过 OpenAI-compatible 中转格式调用 Gemini 1.5 Pro。
5. 服务端解析 AI 返回的 JSON 数组，用 `exceljs` 写入模板并返回 `.xlsx` 下载流。

## 环境变量

复制 `.env.example` 为 `.env.local`，填入实际中转站或 Gemini Key。

```bash
AI_BASE_URL=https://generativelanguage.googleapis.com/v1beta/openai
AI_API_KEY=your_key
AI_MODEL=gemini-1.5-pro
```

## 启动

```bash
npm install
npm run dev
```

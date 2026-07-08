import ExcelJS from "exceljs";

export const runtime = "nodejs";
export const maxDuration = 120;

type GeneratePayload = {
  sourceText?: string;
  images?: string[];
  userPrompt?: string;
  targetHeaders?: string[];
  model?: string;
};

type AiRow = Record<string, unknown>;

type ValidatedPayload = {
  sourceText: string;
  images: string[];
  userPrompt: string;
  targetHeaders: string[];
  model: SupportedModel;
};

const supportedModels = ["gemini-3.5-flash", "gemini-3.1-pro-preview"] as const;
type SupportedModel = (typeof supportedModels)[number];

const jsonResponse = (status: number, error: string) =>
  Response.json(
    {
      error
    },
    {
      status
    }
  );

function buildSystemPrompt(targetHeaders: string[]) {
  return [
    "你是一个跨境电商财务数据整理专家。",
    "请从用户提供的源文本和截图中提取数据，并严格按照目标 Excel 表头输出 JSON 数组。",
    `目标表头：${JSON.stringify(targetHeaders)}`,
    "输出要求：",
    "1. 只输出 JSON 数组，不要输出 Markdown、解释或额外文字。",
    "2. 数组中的每个对象代表一行数据，对象 key 必须来自目标表头。",
    "3. 如果某个字段不确定或缺失，请填入空字符串。",
    "4. 必须遵循用户的附加处理规则。",
    "5. 不要臆造源文本和截图中不存在的关键业务事实。"
  ].join("\n");
}

function normalizeBaseUrl(value?: string) {
  const normalized = (value || "https://yunwu.ai").replace(/\/+$/, "");

  try {
    const parsed = new URL(normalized);
    const path = parsed.pathname.replace(/\/+$/, "");
    const hasApiVersion = /\/(v\d+|v\d+beta)(\/|$)/i.test(path);

    if (!path || path === "/") {
      return `${normalized}/v1`;
    }

    return hasApiVersion ? normalized : `${normalized}/v1`;
  } catch {
    return normalized;
  }
}

function getTextFromAiMessage(content: unknown) {
  if (typeof content === "string") {
    return content;
  }

  if (Array.isArray(content)) {
    return content
      .map((part) => {
        if (part && typeof part === "object" && "text" in part) {
          return String((part as { text?: unknown }).text ?? "");
        }
        return "";
      })
      .join("");
  }

  return "";
}

function extractJsonArray(text: string): AiRow[] {
  const withoutFence = text
    .replace(/^```(?:json)?/i, "")
    .replace(/```$/i, "")
    .trim();

  const start = withoutFence.indexOf("[");
  const end = withoutFence.lastIndexOf("]");

  if (start === -1 || end === -1 || end <= start) {
    throw new Error("AI 返回内容不是 JSON 数组。");
  }

  const parsed = JSON.parse(withoutFence.slice(start, end + 1));

  if (!Array.isArray(parsed)) {
    throw new Error("AI 返回 JSON 顶层结构不是数组。");
  }

  return parsed.map((row) => (row && typeof row === "object" && !Array.isArray(row) ? (row as AiRow) : {}));
}

async function callAi(payload: ValidatedPayload) {
  const apiKey = process.env.AI_API_KEY;
  const baseUrl = normalizeBaseUrl(process.env.AI_BASE_URL);
  const model = payload.model;

  if (!apiKey) {
    throw new Error("服务端缺少 AI_API_KEY 环境变量。");
  }

  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), 90_000);

  try {
    const userContent: Array<
      | {
          type: "text";
          text: string;
        }
      | {
          type: "image_url";
          image_url: {
            url: string;
          };
        }
    > = [
      {
        type: "text",
        text: [
          `用户附加规则：${payload.userPrompt || "无"}`,
          "",
          "源表格文本：",
          payload.sourceText || "无源表格文本，仅根据图片和用户规则处理。"
        ].join("\n")
      },
      ...payload.images.map((url) => ({
        type: "image_url" as const,
        image_url: {
          url
        }
      }))
    ];

    const response = await fetch(`${baseUrl}/chat/completions`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        model,
        temperature: 0.1,
        response_format: {
          type: "json_object"
        },
        messages: [
          {
            role: "system",
            content: buildSystemPrompt(payload.targetHeaders)
          },
          {
            role: "user",
            content: userContent
          }
        ]
      }),
      signal: controller.signal
    });

    if (!response.ok) {
      const detail = await response.text().catch(() => "");
      throw new Error(`AI 接口请求失败：HTTP ${response.status}${detail ? `，${detail.slice(0, 500)}` : ""}`);
    }

    const body = (await response.json()) as {
      choices?: Array<{
        message?: {
          content?: unknown;
        };
      }>;
    };

    const content = getTextFromAiMessage(body.choices?.[0]?.message?.content);

    if (!content.trim()) {
      throw new Error("AI 未返回可解析内容。");
    }

    return extractJsonArray(content);
  } catch (error) {
    if (error instanceof DOMException && error.name === "AbortError") {
      throw new Error("AI 接口请求超时。");
    }

    throw error;
  } finally {
    clearTimeout(timeoutId);
  }
}

function getCellText(cell: ExcelJS.Cell) {
  const value = cell.value;

  if (value == null) {
    return "";
  }

  if (typeof value === "object") {
    if (value instanceof Date) {
      return value.toISOString();
    }

    if ("text" in value && typeof value.text === "string") {
      return value.text;
    }

    if ("result" in value) {
      return String(value.result ?? "");
    }

    if ("richText" in value && Array.isArray(value.richText)) {
      return value.richText.map((part: { text?: string }) => part.text ?? "").join("");
    }
  }

  return String(value);
}

function findHeaderRow(worksheet: ExcelJS.Worksheet, targetHeaders: string[]) {
  const normalizedTargets = new Set(targetHeaders.map((header) => header.trim()).filter(Boolean));
  let bestRow = 1;
  let bestScore = 0;

  worksheet.eachRow((row, rowNumber) => {
    let score = 0;
    row.eachCell((cell) => {
      if (normalizedTargets.has(getCellText(cell).trim())) {
        score += 1;
      }
    });

    if (score > bestScore) {
      bestScore = score;
      bestRow = rowNumber;
    }
  });

  return bestRow;
}

function buildHeaderMap(worksheet: ExcelJS.Worksheet, headerRowNumber: number, targetHeaders: string[]) {
  const headerRow = worksheet.getRow(headerRowNumber);
  const map = new Map<string, number>();

  headerRow.eachCell((cell, colNumber) => {
    const header = getCellText(cell).trim();
    if (header) {
      map.set(header, colNumber);
    }
  });

  let nextColumn = Math.max(headerRow.cellCount, worksheet.columnCount) + 1;

  targetHeaders.forEach((header) => {
    if (!map.has(header)) {
      const cell = headerRow.getCell(nextColumn);
      cell.value = header;
      map.set(header, nextColumn);
      nextColumn += 1;
    }
  });

  headerRow.commit();
  return map;
}

function toExcelValue(value: unknown): ExcelJS.CellValue {
  if (value == null) {
    return "";
  }

  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return value;
  }

  if (value instanceof Date) {
    return value;
  }

  return JSON.stringify(value);
}

function copyRowStyle(sourceRow: ExcelJS.Row, targetRow: ExcelJS.Row, columns: number[]) {
  columns.forEach((columnNumber) => {
    const sourceCell = sourceRow.getCell(columnNumber);
    const targetCell = targetRow.getCell(columnNumber);

    if (sourceCell.style && Object.keys(sourceCell.style).length > 0) {
      targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
    }
  });
}

async function fillWorkbook(templateFile: File, targetHeaders: string[], rows: AiRow[]) {
  const workbook = new ExcelJS.Workbook();
  const templateBuffer = Buffer.from(await templateFile.arrayBuffer());
  await workbook.xlsx.load(templateBuffer as unknown as ArrayBuffer);

  const worksheet = workbook.worksheets[0];

  if (!worksheet) {
    throw new Error("目标模板中没有可写入的工作表。");
  }

  const headerRowNumber = findHeaderRow(worksheet, targetHeaders);
  const headerMap = buildHeaderMap(worksheet, headerRowNumber, targetHeaders);
  const columns = targetHeaders.map((header) => headerMap.get(header)).filter((column): column is number => Boolean(column));
  const styleSourceRow = worksheet.getRow(headerRowNumber + 1);
  let nextRowNumber = Math.max(worksheet.actualRowCount, headerRowNumber) + 1;

  rows.forEach((item) => {
    const row = worksheet.getRow(nextRowNumber);
    copyRowStyle(styleSourceRow, row, columns);

    targetHeaders.forEach((header) => {
      const columnNumber = headerMap.get(header);
      if (columnNumber) {
        row.getCell(columnNumber).value = toExcelValue(item[header]);
      }
    });

    row.commit();
    nextRowNumber += 1;
  });

  return workbook.xlsx.writeBuffer();
}

function resolveModel(value?: string): SupportedModel {
  if (supportedModels.includes(value as SupportedModel)) {
    return value as SupportedModel;
  }

  const fallback = process.env.AI_MODEL;

  if (supportedModels.includes(fallback as SupportedModel)) {
    return fallback as SupportedModel;
  }

  return "gemini-3.5-flash";
}

function validatePayload(payload: GeneratePayload): ValidatedPayload {
  const sourceText = payload.sourceText ?? "";
  const images = Array.isArray(payload.images) ? payload.images : [];
  const userPrompt = payload.userPrompt ?? "";
  const targetHeaders = Array.isArray(payload.targetHeaders)
    ? payload.targetHeaders.map((header) => String(header).trim()).filter(Boolean)
    : [];

  if (!sourceText.trim() && images.length === 0) {
    throw new Error("缺少源数据文本或图片。");
  }

  if (targetHeaders.length === 0) {
    throw new Error("缺少目标模板表头。");
  }

  return {
    sourceText,
    images,
    userPrompt,
    targetHeaders,
    model: resolveModel(payload.model)
  };
}

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const payloadFile = formData.get("payload");
    const templateFile = formData.get("template");

    if (!(payloadFile instanceof File)) {
      return jsonResponse(400, "请求中缺少 payload。");
    }

    if (!(templateFile instanceof File)) {
      return jsonResponse(400, "请求中缺少目标 Excel 模板。");
    }

    const rawPayload = JSON.parse(await payloadFile.text()) as GeneratePayload;
    const payload = validatePayload(rawPayload);
    const aiRows = await callAi(payload);
    const outputBuffer = await fillWorkbook(templateFile, payload.targetHeaders, aiRows);
    const fileName = `ai-filled-${new Date().toISOString().slice(0, 10)}.xlsx`;

    return new Response(outputBuffer, {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${fileName}"`,
        "Cache-Control": "no-store"
      }
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "服务端处理失败。";
    return jsonResponse(500, message);
  }
}

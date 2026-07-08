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
    const rawPayload = (await request.json().catch(() => null)) as GeneratePayload | null;

    if (!rawPayload) {
      return jsonResponse(400, "请求 JSON 格式无效。");
    }

    const payload = validatePayload(rawPayload);
    const aiRows = await callAi(payload);

    return Response.json(
      {
        rows: aiRows
      },
      {
        status: 200,
        headers: {
          "Cache-Control": "no-store"
        }
      }
    );
  } catch (error) {
    console.error(error);
    const message = error instanceof Error ? error.message : "服务端处理失败。";
    return jsonResponse(500, message);
  }
}

"use client";

import { useCallback, useMemo, useState } from "react";
import { useDropzone } from "react-dropzone";
import * as XLSX from "xlsx";
import {
  AlertCircle,
  CheckCircle2,
  Download,
  FileSpreadsheet,
  ImageIcon,
  Loader2,
  Play,
  Plus,
  Trash2,
  UploadCloud
} from "lucide-react";

type DropAreaProps = {
  title: string;
  description: string;
  files: File[];
  accept: Record<string, string[]>;
  multiple?: boolean;
  disabled?: boolean;
  onDrop: (files: File[]) => void;
  onRemove: (index: number) => void;
};

type GeneratePayload = {
  sourceText: string;
  images: string[];
  userPrompt: string;
  targetHeaders: string[];
  model: AiModel;
};

const modelOptions = [
  {
    id: "gemini-3.5-flash",
    label: "Flash",
    description: "快速处理"
  },
  {
    id: "gemini-3.1-pro-preview",
    label: "Pro Preview",
    description: "复杂识别"
  }
] as const;

type AiModel = (typeof modelOptions)[number]["id"];

const sourceAccept = {
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
  "text/csv": [".csv"],
  "image/png": [".png"],
  "image/jpeg": [".jpg", ".jpeg"]
};

const templateAccept = {
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"]
};

function formatBytes(bytes: number) {
  if (bytes === 0) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  const exponent = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${(bytes / 1024 ** exponent).toFixed(exponent === 0 ? 0 : 1)} ${units[exponent]}`;
}

function getFileIcon(file: File) {
  if (file.type.startsWith("image/")) {
    return <ImageIcon className="h-4 w-4 text-teal-700" aria-hidden="true" />;
  }

  return <FileSpreadsheet className="h-4 w-4 text-amber-700" aria-hidden="true" />;
}

function DropArea({
  title,
  description,
  files,
  accept,
  multiple = true,
  disabled,
  onDrop,
  onRemove
}: DropAreaProps) {
  const handleDrop = useCallback(
    (acceptedFiles: File[]) => {
      if (acceptedFiles.length > 0) {
        onDrop(acceptedFiles);
      }
    },
    [onDrop]
  );

  const { getRootProps, getInputProps, isDragActive, fileRejections } = useDropzone({
    accept,
    multiple,
    disabled,
    onDrop: handleDrop
  });

  return (
    <section className="rounded-lg border border-gray-200 bg-white p-5 shadow-soft">
      <div className="mb-4 flex items-start justify-between gap-4">
        <div>
          <h2 className="text-base font-semibold text-gray-950">{title}</h2>
          <p className="mt-1 text-sm text-gray-500">{description}</p>
        </div>
        <Plus className="mt-0.5 h-5 w-5 text-gray-400" aria-hidden="true" />
      </div>

      <div
        {...getRootProps()}
        className={[
          "flex min-h-44 cursor-pointer flex-col items-center justify-center rounded-lg border border-dashed px-6 py-8 text-center transition",
          isDragActive
            ? "border-teal-500 bg-teal-50 text-teal-900"
            : "border-gray-300 bg-gray-50 text-gray-700 hover:border-teal-400 hover:bg-teal-50/60",
          disabled ? "pointer-events-none opacity-60" : ""
        ].join(" ")}
      >
        <input {...getInputProps()} />
        <UploadCloud className="mb-3 h-8 w-8 text-teal-700" aria-hidden="true" />
        <p className="text-sm font-medium">{isDragActive ? "释放文件以上传" : "拖拽文件到这里，或点击选择"}</p>
        <p className="mt-2 text-xs text-gray-500">{Object.values(accept).flat().join(", ")}</p>
      </div>

      {fileRejections.length > 0 ? (
        <div className="mt-3 flex items-center gap-2 rounded-lg bg-red-50 px-3 py-2 text-sm text-red-700">
          <AlertCircle className="h-4 w-4" aria-hidden="true" />
          有文件类型不符合当前上传区域要求。
        </div>
      ) : null}

      {files.length > 0 ? (
        <ul className="mt-4 space-y-2">
          {files.map((file, index) => (
            <li
              key={`${file.name}-${file.lastModified}-${index}`}
              className="flex min-h-12 items-center justify-between gap-3 rounded-lg border border-gray-200 bg-gray-50 px-3"
            >
              <div className="flex min-w-0 items-center gap-3">
                {getFileIcon(file)}
                <div className="min-w-0">
                  <p className="truncate text-sm font-medium text-gray-900">{file.name}</p>
                  <p className="text-xs text-gray-500">{formatBytes(file.size)}</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => onRemove(index)}
                className="inline-flex h-9 w-9 items-center justify-center rounded-md text-gray-500 transition hover:bg-white hover:text-red-600"
                aria-label={`移除 ${file.name}`}
                disabled={disabled}
              >
                <Trash2 className="h-4 w-4" aria-hidden="true" />
              </button>
            </li>
          ))}
        </ul>
      ) : null}
    </section>
  );
}

function sheetRowsToText(workbook: XLSX.WorkBook, fileName: string) {
  const parts: string[] = [`文件：${fileName}`];

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json<Array<string | number | boolean | Date | null>>(sheet, {
      header: 1,
      defval: "",
      raw: false
    });

    parts.push(`工作表：${sheetName}`);
    rows.forEach((row) => {
      const line = row.map((cell) => String(cell ?? "").trim()).join("\t");
      if (line.trim()) {
        parts.push(line);
      }
    });
  });

  return parts.join("\n");
}

async function parseSourceFile(file: File) {
  if (file.name.toLowerCase().endsWith(".csv") || file.type === "text/csv") {
    return `文件：${file.name}\n${await file.text()}`;
  }

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  return sheetRowsToText(workbook, file.name);
}

async function extractTargetHeaders(file: File) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const firstSheetName = workbook.SheetNames[0];

  if (!firstSheetName) {
    return [];
  }

  const rows = XLSX.utils.sheet_to_json<Array<string | number | boolean | Date | null>>(
    workbook.Sheets[firstSheetName],
    {
      header: 1,
      defval: "",
      raw: false
    }
  );

  const headerRow = rows.find((row) => row.some((cell) => String(cell ?? "").trim().length > 0));
  return (headerRow ?? [])
    .map((cell) => String(cell ?? "").trim())
    .filter(Boolean);
}

function fileToDataUrl(file: File) {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(new Error(`图片读取失败：${file.name}`));
    reader.readAsDataURL(file);
  });
}

function downloadBlob(blob: Blob, fileName: string) {
  const href = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = href;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(href);
}

export default function Home() {
  const [sourceFiles, setSourceFiles] = useState<File[]>([]);
  const [templateFiles, setTemplateFiles] = useState<File[]>([]);
  const [userPrompt, setUserPrompt] = useState("");
  const [selectedModel, setSelectedModel] = useState<AiModel>("gemini-3.5-flash");
  const [isProcessing, setIsProcessing] = useState(false);
  const [status, setStatus] = useState("等待上传文件");
  const [error, setError] = useState<string | null>(null);

  const templateFile = templateFiles[0];
  const canSubmit = sourceFiles.length > 0 && Boolean(templateFile) && !isProcessing;

  const sourceSummary = useMemo(() => {
    const sheetCount = sourceFiles.filter((file) => /\.(xlsx|csv)$/i.test(file.name)).length;
    const imageCount = sourceFiles.filter((file) => file.type.startsWith("image/")).length;
    return { sheetCount, imageCount };
  }, [sourceFiles]);

  const handleGenerate = async () => {
    if (isProcessing) return;

    setError(null);

    if (sourceFiles.length === 0) {
      setError("请先上传源数据文件。");
      return;
    }

    if (!templateFile) {
      setError("请上传目标 Excel 模板。");
      return;
    }

    setIsProcessing(true);
    setStatus("正在解析源数据和模板表头");

    const controller = new AbortController();
    const timeoutId = window.setTimeout(() => controller.abort(), 120_000);

    try {
      const sourceTextBlocks: string[] = [];
      const images: string[] = [];

      for (const file of sourceFiles) {
        if (file.type.startsWith("image/")) {
          images.push(await fileToDataUrl(file));
        } else {
          sourceTextBlocks.push(await parseSourceFile(file));
        }
      }

      const targetHeaders = await extractTargetHeaders(templateFile);

      if (targetHeaders.length === 0) {
        throw new Error("目标模板未识别到表头，请确认第一张工作表包含字段名称。");
      }

      const payload: GeneratePayload = {
        sourceText: sourceTextBlocks.join("\n\n---\n\n"),
        images,
        userPrompt,
        targetHeaders,
        model: selectedModel
      };

      const formData = new FormData();
      formData.append("payload", new Blob([JSON.stringify(payload)], { type: "application/json" }));
      formData.append("template", templateFile);

      setStatus(`正在请求 ${selectedModel} 生成结构化数据`);

      const response = await fetch("/api/generate-excel", {
        method: "POST",
        body: formData,
        signal: controller.signal
      });

      if (!response.ok) {
        const message = await response
          .json()
          .then((body: { error?: string }) => body.error)
          .catch(() => null);
        throw new Error(message || `处理失败，HTTP ${response.status}`);
      }

      setStatus("正在下载生成后的 Excel");
      const blob = await response.blob();
      downloadBlob(blob, `ai-filled-${new Date().toISOString().slice(0, 10)}.xlsx`);
      setStatus("处理完成，文件已开始下载");
    } catch (caughtError) {
      const message =
        caughtError instanceof DOMException && caughtError.name === "AbortError"
          ? "请求超时，请减少输入文件体积后重试。"
          : caughtError instanceof Error
            ? caughtError.message
            : "处理失败，请稍后重试。";
      setError(message);
      setStatus("处理失败");
    } finally {
      window.clearTimeout(timeoutId);
      setIsProcessing(false);
    }
  };

  return (
    <main className="mx-auto flex min-h-screen w-full max-w-7xl flex-col px-4 py-6 sm:px-6 lg:px-8">
      <header className="mb-6 flex flex-col justify-between gap-4 border-b border-gray-200 pb-5 lg:flex-row lg:items-end">
        <div>
          <p className="mb-2 text-sm font-semibold text-teal-700">内部运营工具</p>
          <h1 className="text-3xl font-semibold tracking-normal text-gray-950 sm:text-4xl">AI 驱动跨境电商表格处理工作台</h1>
          <p className="mt-3 max-w-3xl text-sm leading-6 text-gray-600">
            上传源 Excel、CSV 或商品截图，再上传目标模板，系统会按模板表头整理数据并导出填好的 Excel。
          </p>
        </div>

        <div className="grid grid-cols-3 gap-2 rounded-lg border border-gray-200 bg-white p-2 text-center shadow-soft sm:min-w-80">
          <div className="rounded-md bg-gray-50 px-3 py-2">
            <p className="text-lg font-semibold text-gray-950">{sourceSummary.sheetCount}</p>
            <p className="text-xs text-gray-500">表格</p>
          </div>
          <div className="rounded-md bg-gray-50 px-3 py-2">
            <p className="text-lg font-semibold text-gray-950">{sourceSummary.imageCount}</p>
            <p className="text-xs text-gray-500">图片</p>
          </div>
          <div className="rounded-md bg-gray-50 px-3 py-2">
            <p className="text-lg font-semibold text-gray-950">{templateFile ? 1 : 0}</p>
            <p className="text-xs text-gray-500">模板</p>
          </div>
        </div>
      </header>

      <div className="grid flex-1 gap-5 lg:grid-cols-[minmax(0,1fr)_minmax(360px,0.8fr)]">
        <div className="space-y-5">
          <DropArea
            title="源数据"
            description="支持 Excel、CSV、商品截图，可一次上传多个文件。"
            files={sourceFiles}
            accept={sourceAccept}
            disabled={isProcessing}
            onDrop={(files) => setSourceFiles((current) => [...current, ...files])}
            onRemove={(index) => setSourceFiles((current) => current.filter((_, currentIndex) => currentIndex !== index))}
          />

          <section className="rounded-lg border border-gray-200 bg-white p-5 shadow-soft">
            <label htmlFor="prompt" className="text-base font-semibold text-gray-950">
              附加处理规则
            </label>
            <textarea
              id="prompt"
              value={userPrompt}
              disabled={isProcessing}
              onChange={(event) => setUserPrompt(event.target.value)}
              placeholder="例如：金额保留两位小数；SKU 以截图中的款式编号为准；缺失字段留空。"
              className="mt-4 min-h-40 w-full resize-y rounded-lg border border-gray-300 bg-gray-50 px-4 py-3 text-sm leading-6 text-gray-900 outline-none transition placeholder:text-gray-400 focus:border-teal-500 focus:bg-white focus:ring-4 focus:ring-teal-100 disabled:opacity-60"
            />
          </section>
        </div>

        <aside className="space-y-5">
          <DropArea
            title="目标 Excel 模板"
            description="仅支持 .xlsx，系统会按第一张工作表的表头写入数据。"
            files={templateFiles}
            accept={templateAccept}
            multiple={false}
            disabled={isProcessing}
            onDrop={(files) => setTemplateFiles(files.slice(0, 1))}
            onRemove={() => setTemplateFiles([])}
          />

          <section className="rounded-lg border border-gray-200 bg-white p-5 shadow-soft">
            <div className="mb-4">
              <h2 className="text-base font-semibold text-gray-950">模型切换</h2>
              <p className="mt-1 text-sm text-gray-500">按任务复杂度选择云雾中转模型。</p>
            </div>

            <div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-1 xl:grid-cols-2" role="radiogroup" aria-label="选择 AI 模型">
              {modelOptions.map((option) => {
                const isSelected = selectedModel === option.id;

                return (
                  <button
                    key={option.id}
                    type="button"
                    role="radio"
                    aria-checked={isSelected}
                    disabled={isProcessing}
                    onClick={() => setSelectedModel(option.id)}
                    className={[
                      "min-h-16 rounded-lg border px-4 py-3 text-left transition disabled:cursor-not-allowed disabled:opacity-60",
                      isSelected
                        ? "border-teal-600 bg-teal-50 text-teal-950 ring-2 ring-teal-100"
                        : "border-gray-200 bg-gray-50 text-gray-700 hover:border-teal-300 hover:bg-white"
                    ].join(" ")}
                  >
                    <span className="block text-sm font-semibold">{option.label}</span>
                    <span className="mt-1 block text-xs text-gray-500">{option.id}</span>
                    <span className="mt-1 block text-xs text-gray-500">{option.description}</span>
                  </button>
                );
              })}
            </div>
          </section>

          <section className="rounded-lg border border-gray-200 bg-white p-5 shadow-soft">
            <div className="flex items-start gap-3">
              {error ? (
                <AlertCircle className="mt-0.5 h-5 w-5 text-red-600" aria-hidden="true" />
              ) : status.includes("完成") ? (
                <CheckCircle2 className="mt-0.5 h-5 w-5 text-teal-700" aria-hidden="true" />
              ) : (
                <Download className="mt-0.5 h-5 w-5 text-gray-500" aria-hidden="true" />
              )}
              <div>
                <h2 className="text-base font-semibold text-gray-950">导出状态</h2>
                <p className={["mt-1 text-sm", error ? "text-red-700" : "text-gray-500"].join(" ")}>
                  {error ?? status}
                </p>
              </div>
            </div>

            <button
              type="button"
              onClick={handleGenerate}
              disabled={!canSubmit}
              className="mt-5 inline-flex h-12 w-full items-center justify-center gap-2 rounded-lg bg-gray-950 px-5 text-sm font-semibold text-white transition hover:bg-teal-800 disabled:cursor-not-allowed disabled:bg-gray-300 disabled:text-gray-500"
            >
              {isProcessing ? (
                <Loader2 className="h-4 w-4 animate-spin" aria-hidden="true" />
              ) : (
                <Play className="h-4 w-4" aria-hidden="true" />
              )}
              {isProcessing ? "处理中" : "开始处理并导出"}
            </button>
          </section>
        </aside>
      </div>
    </main>
  );
}

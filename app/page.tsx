"use client";

import { useCallback, useMemo, useState } from "react";
import ExcelJS from "exceljs";
import { useDropzone } from "react-dropzone";
import * as XLSX from "xlsx";
import {
  AlertCircle,
  ArrowRight,
  Bot,
  Braces,
  CheckCircle2,
  ChevronDown,
  Code2,
  Download,
  FileSpreadsheet,
  Home as HomeIcon,
  ImageIcon,
  Library,
  Loader2,
  Palette,
  Play,
  Plus,
  Sparkles,
  Sun,
  Trash2,
  Type,
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

type AiRow = Record<string, unknown>;

type GenerateResponse = {
  rows?: AiRow[];
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

const sidebarItems = [
  { label: "工作台", icon: HomeIcon, active: true },
  { label: "字段", icon: Type },
  { label: "规则", icon: Palette },
  { label: "模板库", icon: Library },
  { label: "AI 就绪", icon: Bot },
  { label: "接口", icon: Braces },
  { label: "安装", icon: Code2 }
];

function formatBytes(bytes: number) {
  if (bytes === 0) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  const exponent = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${(bytes / 1024 ** exponent).toFixed(exponent === 0 ? 0 : 1)} ${units[exponent]}`;
}

function getFileIcon(file: File) {
  if (file.type.startsWith("image/")) {
    return <ImageIcon className="h-4 w-4 text-[#1559e8]" aria-hidden="true" />;
  }

  return <FileSpreadsheet className="h-4 w-4 text-[#e36d43]" aria-hidden="true" />;
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
    <section className="rounded-lg border border-[#e5e8f0] bg-white p-5 shadow-panel">
      <div className="mb-4 flex items-start justify-between gap-4">
        <div>
          <h2 className="text-base font-bold text-[#0d1017]">{title}</h2>
          <p className="mt-1 text-sm leading-5 text-[#667085]">{description}</p>
        </div>
        <Plus className="mt-0.5 h-5 w-5 text-[#1559e8]" aria-hidden="true" />
      </div>

      <div
        {...getRootProps()}
        className={[
          "flex min-h-44 cursor-pointer flex-col items-center justify-center rounded-lg border border-dashed px-6 py-8 text-center transition",
          isDragActive
            ? "border-[#1559e8] bg-[#edf4ff] text-[#0f48c7]"
            : "border-[#d7deea] bg-[#fbfcff] text-[#1b2533] hover:border-[#1559e8] hover:bg-[#f4f8ff]",
          disabled ? "pointer-events-none opacity-60" : ""
        ].join(" ")}
      >
        <input {...getInputProps()} />
        <div className="mb-3 flex h-12 w-12 items-center justify-center rounded-full bg-white shadow-float">
          <UploadCloud className="h-6 w-6 text-[#1559e8]" aria-hidden="true" />
        </div>
        <p className="text-sm font-semibold">{isDragActive ? "释放文件以上传" : "拖拽文件到这里，或点击选择"}</p>
        <p className="mt-2 text-xs text-[#667085]">{Object.values(accept).flat().join(", ")}</p>
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
              className="flex min-h-12 items-center justify-between gap-3 rounded-lg border border-[#e5e8f0] bg-[#fbfcff] px-3"
            >
              <div className="flex min-w-0 items-center gap-3">
                {getFileIcon(file)}
                <div className="min-w-0">
                  <p className="truncate text-sm font-semibold text-[#111318]">{file.name}</p>
                  <p className="text-xs text-[#667085]">{formatBytes(file.size)}</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => onRemove(index)}
                className="inline-flex h-9 w-9 items-center justify-center rounded-md text-[#667085] transition hover:bg-white hover:text-red-600"
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
  const templateBuffer = await templateFile.arrayBuffer();
  await workbook.xlsx.load(templateBuffer);

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

      setStatus(`正在请求 ${selectedModel} 生成结构化数据`);

      const response = await fetch("/api/generate-excel", {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify(payload),
        signal: controller.signal
      });

      if (!response.ok) {
        const message = await response
          .json()
          .then((body: { error?: string }) => body.error)
          .catch(() => null);
        throw new Error(message || `处理失败，HTTP ${response.status}`);
      }

      const body = (await response.json()) as GenerateResponse;
      const rows = Array.isArray(body.rows) ? body.rows : null;

      if (!rows) {
        throw new Error("服务端返回数据格式异常。");
      }

      setStatus("正在写入模板并导出 Excel");
      const outputBuffer = await fillWorkbook(templateFile, targetHeaders, rows);
      const blob = new Blob([outputBuffer as BlobPart], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      });
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
    <main className="min-h-screen overflow-hidden bg-[#fbfcff] text-[#111318]">
      <header className="sticky top-0 z-30 flex h-16 items-center justify-between border-b border-[#e5e8f0] bg-white/95 px-4 backdrop-blur sm:px-7">
        <div className="flex items-center gap-9">
          <div className="flex items-center gap-3">
            <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-[#1559e8] text-white shadow-blue">
              <Sparkles className="h-4 w-4" aria-hidden="true" />
            </div>
            <span className="text-xl font-extrabold tracking-normal text-[#101828]">DianpuTX</span>
          </div>

          <nav className="hidden items-center gap-9 text-sm font-semibold text-[#111318] md:flex" aria-label="主导航">
            <button type="button" className="inline-flex items-center gap-1.5 transition hover:text-[#1559e8]">
              工具
              <ChevronDown className="h-4 w-4" aria-hidden="true" />
            </button>
            <button type="button" className="transition hover:text-[#1559e8]">
              处理流程
            </button>
            <button type="button" className="transition hover:text-[#1559e8]">
              模板库
            </button>
          </nav>
        </div>

        <div className="flex items-center gap-3">
          <span className="hidden text-sm font-semibold text-[#111318] sm:inline">Workspace</span>
          <button
            type="button"
            className="inline-flex h-10 items-center justify-center rounded-lg bg-[#1559e8] px-4 text-sm font-bold text-white shadow-blue transition hover:bg-[#0f48c7]"
          >
            Excel 工具
          </button>
        </div>
      </header>

      <div className="flex min-h-[calc(100vh-4rem)]">
        <aside className="hidden w-[84px] shrink-0 flex-col border-r border-[#e5e8f0] bg-white md:flex">
          <nav className="flex flex-1 flex-col items-center py-5" aria-label="工具导航">
            <div className="space-y-1">
              {sidebarItems.map((item) => {
                const Icon = item.icon;
                return (
                  <button
                    key={item.label}
                    type="button"
                    className={[
                      "flex h-[84px] w-[84px] flex-col items-center justify-center gap-2 border-l-4 text-xs font-medium transition",
                      item.active
                        ? "border-[#1559e8] bg-[#edf4ff] text-[#1559e8]"
                        : "border-transparent text-[#475467] hover:bg-[#f4f7fb] hover:text-[#1559e8]"
                    ].join(" ")}
                  >
                    <Icon className="h-5 w-5" aria-hidden="true" />
                    {item.label}
                  </button>
                );
              })}
            </div>
            <div className="mt-auto flex flex-col items-center gap-4 pb-3 text-[#475467]">
              <Sun className="h-5 w-5" aria-hidden="true" />
              <span className="text-xs font-bold">v1.1</span>
            </div>
          </nav>
        </aside>

        <section className="relative flex-1 px-4 py-8 sm:px-8 lg:px-12">
          <div className="pointer-events-none absolute inset-0 overflow-hidden">
            <div className="absolute left-1/2 top-8 h-[780px] w-[780px] -translate-x-1/2 rounded-full border border-[#d9e4ff]" />
            <div className="absolute left-1/2 top-[-80px] h-[1040px] w-[1040px] -translate-x-1/2 rounded-full border border-[#e6edff]" />
            <div className="absolute left-1/2 top-[-190px] h-[1260px] w-[1260px] -translate-x-1/2 rounded-full border border-[#eef3ff]" />
          </div>

          <div className="relative mx-auto max-w-7xl">
            <div className="mx-auto mb-8 max-w-4xl text-center">
              <div className="mb-5 inline-flex items-center gap-2 rounded-full border border-[#d9e4ff] bg-white px-4 py-2 text-sm font-bold text-[#1559e8] shadow-float">
                <Sparkles className="h-4 w-4" aria-hidden="true" />
                AI 表格处理工作台
              </div>
              <h1 className="text-4xl font-extrabold leading-tight tracking-normal text-[#111318] sm:text-5xl lg:text-6xl">
                跨境店铺表格处理。
                <span className="block text-[#1559e8]">Fast. Clean. Ready.</span>
              </h1>
              <p className="mx-auto mt-5 max-w-2xl text-base leading-7 text-[#344054]">
                上传源数据和目标模板，按你的规则生成结构化结果，并直接导出可交付 Excel。
              </p>

              <div className="mt-6 grid gap-2 text-center sm:grid-cols-3">
                <div className="rounded-lg border border-[#e5e8f0] bg-white px-4 py-3 shadow-panel">
                  <p className="text-2xl font-extrabold text-[#111318]">{sourceSummary.sheetCount}</p>
                  <p className="text-xs font-semibold text-[#667085]">表格</p>
                </div>
                <div className="rounded-lg border border-[#e5e8f0] bg-white px-4 py-3 shadow-panel">
                  <p className="text-2xl font-extrabold text-[#111318]">{sourceSummary.imageCount}</p>
                  <p className="text-xs font-semibold text-[#667085]">图片</p>
                </div>
                <div className="rounded-lg border border-[#e5e8f0] bg-white px-4 py-3 shadow-panel">
                  <p className="text-2xl font-extrabold text-[#111318]">{templateFile ? 1 : 0}</p>
                  <p className="text-xs font-semibold text-[#667085]">模板</p>
                </div>
              </div>
            </div>

            <div className="grid gap-5 lg:grid-cols-[minmax(0,1fr)_420px]">
              <div className="space-y-5">
                <div className="grid gap-5 xl:grid-cols-2">
                  <DropArea
                    title="源数据"
                    description="支持 Excel、CSV、商品截图，可一次上传多个文件。"
                    files={sourceFiles}
                    accept={sourceAccept}
                    disabled={isProcessing}
                    onDrop={(files) => setSourceFiles((current) => [...current, ...files])}
                    onRemove={(index) => setSourceFiles((current) => current.filter((_, currentIndex) => currentIndex !== index))}
                  />

                  <DropArea
                    title="目标 Excel 模板"
                    description="仅支持 .xlsx，按第一张工作表的表头写入数据。"
                    files={templateFiles}
                    accept={templateAccept}
                    multiple={false}
                    disabled={isProcessing}
                    onDrop={(files) => setTemplateFiles(files.slice(0, 1))}
                    onRemove={() => setTemplateFiles([])}
                  />
                </div>

                <section className="rounded-lg border border-[#e5e8f0] bg-white p-5 shadow-panel">
                  <label htmlFor="prompt" className="text-base font-bold text-[#0d1017]">
                    附加处理规则
                  </label>
                  <textarea
                    id="prompt"
                    value={userPrompt}
                    disabled={isProcessing}
                    onChange={(event) => setUserPrompt(event.target.value)}
                    placeholder="例如：金额保留两位小数；SKU 以截图中的款式编号为准；缺失字段留空。"
                    className="mt-4 min-h-44 w-full resize-y rounded-lg border border-[#d7deea] bg-[#fbfcff] px-4 py-3 text-sm leading-6 text-[#111318] outline-none transition placeholder:text-[#98a2b3] focus:border-[#1559e8] focus:bg-white focus:ring-4 focus:ring-[#d9e4ff] disabled:opacity-60"
                  />
                </section>
              </div>

              <aside className="space-y-5">
                <section className="rounded-lg border border-[#e5e8f0] bg-white p-5 shadow-panel">
                  <div className="mb-4">
                    <h2 className="text-base font-bold text-[#0d1017]">模型切换</h2>
                    <p className="mt-1 text-sm leading-5 text-[#667085]">按任务复杂度选择云雾中转模型。</p>
                  </div>

                  <div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-1" role="radiogroup" aria-label="选择 AI 模型">
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
                            "min-h-20 rounded-lg border px-4 py-3 text-left transition disabled:cursor-not-allowed disabled:opacity-60",
                            isSelected
                              ? "border-[#1559e8] bg-[#edf4ff] text-[#0f48c7] ring-2 ring-[#d9e4ff]"
                              : "border-[#e5e8f0] bg-white text-[#344054] hover:border-[#1559e8] hover:bg-[#f4f8ff]"
                          ].join(" ")}
                        >
                          <span className="block text-sm font-extrabold">{option.label}</span>
                          <span className="mt-1 block text-xs font-medium text-[#667085]">{option.id}</span>
                          <span className="mt-1 block text-xs font-semibold text-[#344054]">{option.description}</span>
                        </button>
                      );
                    })}
                  </div>
                </section>

                <section className="rounded-lg border border-[#e5e8f0] bg-white p-5 shadow-panel">
                  <div className="flex items-start gap-3">
                    <div
                      className={[
                        "flex h-9 w-9 shrink-0 items-center justify-center rounded-full",
                        error ? "bg-red-50 text-red-600" : status.includes("完成") ? "bg-emerald-50 text-emerald-600" : "bg-[#edf4ff] text-[#1559e8]"
                      ].join(" ")}
                    >
                      {error ? (
                        <AlertCircle className="h-5 w-5" aria-hidden="true" />
                      ) : status.includes("完成") ? (
                        <CheckCircle2 className="h-5 w-5" aria-hidden="true" />
                      ) : (
                        <Download className="h-5 w-5" aria-hidden="true" />
                      )}
                    </div>
                    <div className="min-w-0">
                      <h2 className="text-base font-bold text-[#0d1017]">导出状态</h2>
                      <p className={["mt-1 text-sm leading-5", error ? "text-red-700" : "text-[#667085]"].join(" ")}>
                        {error ?? status}
                      </p>
                    </div>
                  </div>

                  <button
                    type="button"
                    onClick={handleGenerate}
                    disabled={!canSubmit}
                    className="mt-5 inline-flex h-12 w-full items-center justify-center gap-2 rounded-lg bg-[#1559e8] px-5 text-sm font-extrabold text-white shadow-blue transition hover:bg-[#0f48c7] disabled:cursor-not-allowed disabled:bg-[#d0d5dd] disabled:text-[#667085] disabled:shadow-none"
                  >
                    {isProcessing ? (
                      <Loader2 className="h-4 w-4 animate-spin" aria-hidden="true" />
                    ) : (
                      <Play className="h-4 w-4" aria-hidden="true" />
                    )}
                    {isProcessing ? "处理中" : "开始处理并导出"}
                    {!isProcessing ? <ArrowRight className="h-4 w-4" aria-hidden="true" /> : null}
                  </button>
                </section>
              </aside>
            </div>
          </div>
        </section>
      </div>
    </main>
  );
}

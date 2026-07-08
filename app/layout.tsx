import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "AI 表格处理工作台",
  description: "内部跨境电商 Excel 数据整理工具"
};

export default function RootLayout({
  children
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}

import type { DocumentSpec, ValidationResult } from "./types.js";

export function validateDocument(spec: DocumentSpec): ValidationResult {
  const errors: string[] = [];
  const warnings: string[] = [];
  const headings = spec.nodes.filter((node) => node.type === "heading").length;
  const paragraphs = spec.nodes.filter((node) => node.type === "paragraph").length;
  const tables = spec.nodes.filter((node) => node.type === "table");
  const tableRows = tables.reduce((total, table) => total + table.rows.length, 0);

  if (!spec.title.trim()) errors.push("文档标题为空。");
  if (spec.nodes.length === 0) errors.push("没有可生成的正文内容。");

  for (const [tableIndex, table] of tables.entries()) {
    if (table.rows.length === 0) {
      errors.push(`第 ${tableIndex + 1} 个表格没有数据行。`);
      continue;
    }
    const widths = new Set(table.rows.map((row) => row.length));
    if (widths.size > 1) {
      errors.push(`第 ${tableIndex + 1} 个表格的列数不一致。`);
    }
    if ((table.rows[0]?.length ?? 0) > 8) {
      warnings.push(`第 ${tableIndex + 1} 个表格超过 8 列，手机和 A4 页面可能较拥挤。`);
    }
  }

  if (headings === 0 && paragraphs + tables.length > 4) {
    warnings.push("未识别出分级标题，请确认原文是否需要章节结构。");
  }
  if (spec.title === "未命名文档") {
    warnings.push("未自动识别标题，已使用“未命名文档”。");
  }

  return {
    ok: errors.length === 0,
    errors,
    warnings,
    stats: { headings, paragraphs, tables: tables.length, tableRows },
  };
}

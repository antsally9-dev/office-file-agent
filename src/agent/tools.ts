import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { parseText, renderDocxBlob, validateDocument } from "../core/index.js";
import type {
  DocumentSpec,
  TemplateId,
  ValidationResult,
} from "../core/index.js";

export interface AgentContext {
  sourceText: string;
  title?: string;
  template: TemplateId;
  outputPath: string;
  spec?: DocumentSpec;
  validation?: ValidationResult;
  trace: AgentTraceEntry[];
}

export interface AgentTraceEntry {
  tool: "parse_text" | "validate_document" | "render_docx";
  at: string;
  result: Record<string, unknown>;
}

function trace(
  context: AgentContext,
  tool: AgentTraceEntry["tool"],
  result: Record<string, unknown>,
): void {
  context.trace.push({ tool, at: new Date().toISOString(), result });
}

export async function parseTextTool(
  context: AgentContext,
  args: { title?: string; template?: TemplateId } = {},
): Promise<Record<string, unknown>> {
  context.spec = parseText({
    text: context.sourceText,
    title: args.title ?? context.title,
    template: args.template ?? context.template,
  });
  const result = {
    title: context.spec.title,
    template: context.spec.template,
    nodes: context.spec.nodes.length,
    tables: context.spec.nodes.filter((node) => node.type === "table").length,
  };
  trace(context, "parse_text", result);
  return result;
}

export async function validateDocumentTool(
  context: AgentContext,
): Promise<Record<string, unknown>> {
  if (!context.spec) throw new Error("请先调用 parse_text。 ");
  context.validation = validateDocument(context.spec);
  const result = {
    ok: context.validation.ok,
    errors: context.validation.errors,
    warnings: context.validation.warnings,
    stats: context.validation.stats,
  };
  trace(context, "validate_document", result);
  return result;
}

export async function renderDocxTool(
  context: AgentContext,
  args: { outputPath?: string } = {},
): Promise<Record<string, unknown>> {
  if (!context.spec) throw new Error("请先调用 parse_text。 ");
  const validation = context.validation ?? validateDocument(context.spec);
  if (!validation.ok) {
    throw new Error(`文档校验未通过：${validation.errors.join("；")}`);
  }

  const outputPath = resolve(args.outputPath ?? context.outputPath);
  const blob = await renderDocxBlob(context.spec);
  const bytes = new Uint8Array(await blob.arrayBuffer());
  await mkdir(dirname(outputPath), { recursive: true });
  await writeFile(outputPath, bytes);
  context.outputPath = outputPath;

  const result = { outputPath, bytes: bytes.byteLength };
  trace(context, "render_docx", result);
  return result;
}

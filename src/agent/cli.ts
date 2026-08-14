import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { OfficeFileAgent } from "./office-agent.js";
import type { TemplateId } from "../core/index.js";

function option(name: string): string | undefined {
  const index = process.argv.indexOf(`--${name}`);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

function hasFlag(name: string): boolean {
  return process.argv.includes(`--${name}`);
}

function usage(): never {
  console.error(`
用法：
  pnpm agent -- --input examples/sample.txt --output outputs/sample.docx [选项]

选项：
  --title <标题>              指定文档标题
  --template formal|report   选择模板，默认 formal
  --model <模型名>           启用 OpenAI Responses API 工具规划器
  --trace                    输出工具调用轨迹
`);
  process.exit(1);
}

const inputPath = option("input");
const outputPath = option("output");
if (!inputPath || !outputPath) usage();

const templateValue = option("template") ?? "formal";
if (templateValue !== "formal" && templateValue !== "report") {
  throw new Error("--template 只能是 formal 或 report。");
}

const text = await readFile(resolve(inputPath), "utf8");
const result = await new OfficeFileAgent().run({
  text,
  title: option("title"),
  template: templateValue as TemplateId,
  outputPath: resolve(outputPath),
  model: option("model") ?? process.env.OPENAI_MODEL,
});

console.log(`已生成：${result.outputPath}`);
console.log(
  `模式：${result.mode}；段落 ${result.validation.stats.paragraphs}；标题 ${result.validation.stats.headings}；表格 ${result.validation.stats.tables}`,
);
if (result.validation.warnings.length > 0) {
  console.log(`提示：${result.validation.warnings.join("；")}`);
}
if (hasFlag("trace")) {
  console.log(JSON.stringify(result.trace, null, 2));
}

import OpenAI from "openai";
import type { ResponseInput } from "openai/resources/responses/responses";
import type { TemplateId, ValidationResult } from "../core/index.js";
import {
  parseTextTool,
  renderDocxTool,
  validateDocumentTool,
  type AgentContext,
  type AgentTraceEntry,
} from "./tools.js";

export interface AgentRequest {
  text: string;
  title?: string;
  template?: TemplateId;
  outputPath: string;
  model?: string;
}

export interface AgentResult {
  outputPath: string;
  validation: ValidationResult;
  trace: AgentTraceEntry[];
  mode: "deterministic" | "model-planned";
  finalMessage?: string;
}

const TOOL_DEFINITIONS = [
  {
    type: "function" as const,
    name: "parse_text",
    description:
      "Parse the provided source text into headings, paragraphs, and simple tables. Call this first.",
    strict: true,
    parameters: {
      type: "object",
      properties: {
        title: { type: ["string", "null"] },
        template: { type: ["string", "null"], enum: ["formal", "report", null] },
      },
      required: ["title", "template"],
      additionalProperties: false,
    },
  },
  {
    type: "function" as const,
    name: "validate_document",
    description: "Validate the parsed document structure before rendering.",
    strict: true,
    parameters: {
      type: "object",
      properties: {},
      required: [],
      additionalProperties: false,
    },
  },
  {
    type: "function" as const,
    name: "render_docx",
    description: "Render a validated document to the requested DOCX output path.",
    strict: true,
    parameters: {
      type: "object",
      properties: {},
      required: [],
      additionalProperties: false,
    },
  },
];

function createContext(request: AgentRequest): AgentContext {
  return {
    sourceText: request.text,
    title: request.title,
    template: request.template ?? "formal",
    outputPath: request.outputPath,
    trace: [],
  };
}

async function executeTool(
  context: AgentContext,
  name: string,
  rawArguments: string,
): Promise<Record<string, unknown>> {
  const args = JSON.parse(rawArguments) as Record<string, unknown>;
  if (name === "parse_text") {
    const template =
      args.template === "formal" || args.template === "report"
        ? args.template
        : undefined;
    return parseTextTool(context, {
      title: typeof args.title === "string" ? args.title : undefined,
      template,
    });
  }
  if (name === "validate_document") return validateDocumentTool(context);
  if (name === "render_docx") return renderDocxTool(context);
  throw new Error(`未知工具：${name}`);
}

export class OfficeFileAgent {
  async run(request: AgentRequest): Promise<AgentResult> {
    return request.model
      ? this.runWithModel(request)
      : this.runDeterministically(request);
  }

  private async runDeterministically(request: AgentRequest): Promise<AgentResult> {
    const context = createContext(request);
    await parseTextTool(context);
    await validateDocumentTool(context);
    await renderDocxTool(context);

    if (!context.validation) throw new Error("内部错误：缺少校验结果。");
    return {
      outputPath: context.outputPath,
      validation: context.validation,
      trace: context.trace,
      mode: "deterministic",
    };
  }

  private async runWithModel(request: AgentRequest): Promise<AgentResult> {
    if (!process.env.OPENAI_API_KEY) {
      throw new Error("使用模型规划器需要设置 OPENAI_API_KEY。");
    }

    const context = createContext(request);
    const openai = new OpenAI();
    const input: ResponseInput = [
      {
        role: "user",
        content: [
          {
            type: "input_text",
            text: [
              "请把以下内容生成 Word。必须依次调用 parse_text、validate_document、render_docx。",
              `建议标题：${request.title ?? "由工具自动判断"}`,
              `模板：${request.template ?? "formal"}`,
              "源文本：",
              request.text,
            ].join("\n"),
          },
        ],
      },
    ];

    let finalMessage = "";
    for (let turn = 0; turn < 8; turn += 1) {
      const response = await openai.responses.create({
        model: request.model,
        instructions:
          "你是办公文件生成 Agent。不要改写事实或补充源文本中不存在的数据。通过工具完成文件生成；校验失败时解释错误，不得绕过校验。",
        input,
        tools: TOOL_DEFINITIONS,
      });
      input.push(...response.output);
      finalMessage = response.output_text;

      const calls = response.output.filter((item) => item.type === "function_call");
      if (calls.length === 0) break;

      for (const call of calls) {
        const result = await executeTool(context, call.name, call.arguments);
        input.push({
          type: "function_call_output",
          call_id: call.call_id,
          output: JSON.stringify(result),
        });
      }
    }

    if (!context.validation || !context.outputPath || !context.spec) {
      throw new Error("模型没有完成完整的解析、校验和生成工具链。");
    }
    return {
      outputPath: context.outputPath,
      validation: context.validation,
      trace: context.trace,
      mode: "model-planned",
      finalMessage,
    };
  }
}

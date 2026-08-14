import { describe, expect, it } from "vitest";
import { parseText, validateDocument } from "../src/core/index.js";

describe("parseText", () => {
  it("识别标题、中文层级、键值表和制表符表格", () => {
    const spec = parseText({
      text: [
        "现场报告",
        "",
        "一、基本信息",
        "姓名：张三",
        "日期：2026年8月12日",
        "",
        "事项\t负责人",
        "补充材料\t李四",
      ].join("\n"),
    });

    expect(spec.title).toBe("现场报告");
    expect(spec.nodes.filter((node) => node.type === "heading")).toHaveLength(1);
    const tables = spec.nodes.filter((node) => node.type === "table");
    expect(tables).toHaveLength(2);
    expect(tables[0]?.style).toBe("key-value");
    expect(tables[1]?.rows).toEqual([
      ["事项", "负责人"],
      ["补充材料", "李四"],
    ]);
  });

  it("为不规则表格补齐空单元格", () => {
    const spec = parseText({
      title: "测试",
      text: "名称|负责人|日期\n事项一|张三\n事项二|李四|8月12日",
    });
    const table = spec.nodes.find((node) => node.type === "table");
    expect(table?.rows[1]).toEqual(["事项一", "张三", ""]);
    expect(validateDocument(spec).ok).toBe(true);
  });

  it("忽略 Markdown 表格两端空列和分隔行", () => {
    const spec = parseText({
      title: "测试",
      text: "| 事项 | 负责人 |\n| --- | --- |\n| 补充材料 | 张三 |",
    });
    const table = spec.nodes.find((node) => node.type === "table");
    expect(table?.rows).toEqual([
      ["事项", "负责人"],
      ["补充材料", "张三"],
    ]);
  });
});

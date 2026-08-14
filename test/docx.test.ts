import { describe, expect, it } from "vitest";
import { parseText, renderDocxBlob } from "../src/core/index.js";

describe("renderDocxBlob", () => {
  it("生成有效的 OOXML zip 文件", async () => {
    const spec = parseText({
      title: "测试文档",
      text: "一、概况\n正文内容。\n字段：值\n日期：2026年8月12日",
    });
    const blob = await renderDocxBlob(spec);
    const bytes = new Uint8Array(await blob.arrayBuffer());
    expect(bytes.byteLength).toBeGreaterThan(1000);
    expect(String.fromCharCode(bytes[0] ?? 0, bytes[1] ?? 0)).toBe("PK");
  });
});

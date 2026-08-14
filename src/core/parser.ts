import type {
  DocumentNode,
  DocumentSpec,
  ParseInput,
  TableNode,
} from "./types.js";

const MARKDOWN_HEADING = /^(#{1,3})\s+(.+)$/;
const HEADING_LEVEL_1 = /^[一二三四五六七八九十]+、\s*(.+)$/;
const HEADING_LEVEL_2 = /^（[一二三四五六七八九十]+）\s*(.+)$/;
const HEADING_LEVEL_3 = /^\d+[.、]\s*(.+)$/;
const KEY_VALUE = /^([^：:\t|]{1,24})[：:]\s*(.*)$/;

function cleanCell(value: string): string {
  return value.trim().replace(/^\|\s*|\s*\|$/g, "").trim();
}

function splitDelimitedLine(line: string, delimiter: "\t" | "|"): string[] {
  const cells = line.split(delimiter).map(cleanCell);
  if (delimiter === "|" && cells[0] === "") cells.shift();
  if (delimiter === "|" && cells.at(-1) === "") cells.pop();
  return cells;
}

function delimiterFor(line: string): "\t" | "|" | null {
  if (line.includes("\t")) return "\t";
  const pipes = (line.match(/\|/g) ?? []).length;
  return pipes >= 2 ? "|" : null;
}

function normalizeRows(rows: string[][]): string[][] {
  const width = Math.max(...rows.map((row) => row.length));
  return rows.map((row) => [
    ...row,
    ...Array.from({ length: width - row.length }, () => ""),
  ]);
}

function looksLikeMarkdownDivider(row: string[]): boolean {
  return row.every((cell) => /^:?-{3,}:?$/.test(cell));
}

function headingFromLine(line: string): DocumentNode | null {
  const markdown = line.match(MARKDOWN_HEADING);
  if (markdown) {
    const level = Math.min(markdown[1]?.length ?? 1, 3) as 1 | 2 | 3;
    return { type: "heading", level, text: markdown[2]?.trim() ?? "" };
  }

  if (HEADING_LEVEL_1.test(line)) {
    return { type: "heading", level: 1, text: line };
  }
  if (HEADING_LEVEL_2.test(line)) {
    return { type: "heading", level: 2, text: line };
  }
  if (HEADING_LEVEL_3.test(line) && line.length <= 48) {
    return { type: "heading", level: 3, text: line };
  }
  return null;
}

function inferTitle(lines: string[], provided?: string): {
  title: string;
  consumedLine: number | null;
} {
  if (provided?.trim()) return { title: provided.trim(), consumedLine: null };

  const firstIndex = lines.findIndex((line) => line.trim().length > 0);
  const first = firstIndex >= 0 ? lines[firstIndex]?.trim() ?? "" : "";
  const isCandidate =
    first.length > 0 &&
    first.length <= 40 &&
    !delimiterFor(first) &&
    !KEY_VALUE.test(first) &&
    !/[。！？；]$/.test(first);

  return isCandidate
    ? { title: first.replace(/^#{1,3}\s+/, ""), consumedLine: firstIndex }
    : { title: "未命名文档", consumedLine: null };
}

export function parseText(input: ParseInput): DocumentSpec {
  const normalized = input.text.replace(/\r\n?/g, "\n").trim();
  const lines = normalized.length > 0 ? normalized.split("\n") : [];
  const { title, consumedLine } = inferTitle(lines, input.title);
  const nodes: DocumentNode[] = [];

  for (let index = 0; index < lines.length; ) {
    if (index === consumedLine) {
      index += 1;
      continue;
    }

    const line = lines[index]?.trim() ?? "";
    if (!line) {
      index += 1;
      continue;
    }

    const delimiter = delimiterFor(line);
    if (delimiter) {
      const rows: string[][] = [];
      let cursor = index;
      while (cursor < lines.length) {
        const candidate = lines[cursor]?.trim() ?? "";
        if (!candidate.includes(delimiter)) break;
        rows.push(splitDelimitedLine(candidate, delimiter));
        cursor += 1;
      }

      const consistentRows = rows.filter((row) => row.length >= 2);
      if (consistentRows.length >= 2) {
        const withoutDivider = consistentRows.filter(
          (row) => !looksLikeMarkdownDivider(row),
        );
        nodes.push({
          type: "table",
          style: "grid",
          rows: normalizeRows(withoutDivider),
          hasHeader: withoutDivider.length > 1,
        });
        index = cursor;
        continue;
      }
    }

    const keyValueRows: string[][] = [];
    let keyValueCursor = index;
    while (keyValueCursor < lines.length) {
      const candidate = lines[keyValueCursor]?.trim() ?? "";
      const match = candidate.match(KEY_VALUE);
      if (!match) break;
      keyValueRows.push([match[1]?.trim() ?? "", match[2]?.trim() ?? ""]);
      keyValueCursor += 1;
    }
    if (keyValueRows.length >= 2) {
      const table: TableNode = {
        type: "table",
        style: "key-value",
        rows: keyValueRows,
        hasHeader: false,
      };
      nodes.push(table);
      index = keyValueCursor;
      continue;
    }

    const heading = headingFromLine(line);
    if (heading) {
      nodes.push(heading);
    } else {
      nodes.push({ type: "paragraph", text: line });
    }
    index += 1;
  }

  return {
    title,
    template: input.template ?? "formal",
    nodes,
    metadata: {
      parsedAt: new Date().toISOString(),
      sourceCharacters: input.text.length,
    },
  };
}

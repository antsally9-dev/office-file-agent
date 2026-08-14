export type TemplateId = "formal" | "report";

export interface HeadingNode {
  type: "heading";
  level: 1 | 2 | 3;
  text: string;
}

export interface ParagraphNode {
  type: "paragraph";
  text: string;
}

export interface TableNode {
  type: "table";
  style: "grid" | "key-value";
  rows: string[][];
  hasHeader: boolean;
}

export type DocumentNode = HeadingNode | ParagraphNode | TableNode;

export interface DocumentSpec {
  title: string;
  template: TemplateId;
  nodes: DocumentNode[];
  metadata: {
    parsedAt: string;
    sourceCharacters: number;
  };
}

export interface ParseInput {
  text: string;
  title?: string;
  template?: TemplateId;
}

export interface ValidationResult {
  ok: boolean;
  errors: string[];
  warnings: string[];
  stats: {
    headings: number;
    paragraphs: number;
    tables: number;
    tableRows: number;
  };
}

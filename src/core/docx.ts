import {
  AlignmentType,
  BorderStyle,
  Document,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  VerticalAlignTable,
  WidthType,
} from "docx";
import type { DocumentNode, DocumentSpec, TableNode, TemplateId } from "./types.js";

interface TemplateStyle {
  bodyFont: string;
  headingFont: string;
  titleSize: number;
  bodySize: number;
  accentFill: string;
  headerText: string;
  lineSpacing: number;
  bodyAfter: number;
  bodyAlignment: (typeof AlignmentType)[keyof typeof AlignmentType];
  titleAfter: number;
  headings: Record<
    1 | 2 | 3,
    { size: number; color: string; before: number; after: number }
  >;
}

const TEMPLATES: Record<TemplateId, TemplateStyle> = {
  formal: {
    bodyFont: "SimSun",
    headingFont: "SimHei",
    titleSize: 46,
    bodySize: 22,
    accentFill: "F2F4F7",
    headerText: "202124",
    lineSpacing: 264,
    bodyAfter: 120,
    bodyAlignment: AlignmentType.LEFT,
    titleAfter: 320,
    headings: {
      1: { size: 32, color: "2E74B5", before: 320, after: 160 },
      2: { size: 26, color: "2E74B5", before: 240, after: 120 },
      3: { size: 24, color: "1F4D78", before: 160, after: 80 },
    },
  },
  report: {
    bodyFont: "Microsoft YaHei",
    headingFont: "Microsoft YaHei",
    titleSize: 48,
    bodySize: 22,
    accentFill: "F4F6F9",
    headerText: "17365D",
    lineSpacing: 320,
    bodyAfter: 160,
    bodyAlignment: AlignmentType.JUSTIFIED,
    titleAfter: 360,
    headings: {
      1: { size: 32, color: "2E74B5", before: 360, after: 200 },
      2: { size: 26, color: "2E74B5", before: 240, after: 120 },
      3: { size: 24, color: "1F4D78", before: 160, after: 80 },
    },
  },
};

const CONTENT_WIDTH_DXA = 9360;
const TABLE_INDENT_DXA = 120;

function columnWidthsFor(node: TableNode): number[] {
  const columnCount = node.rows[0]?.length ?? 1;
  if (node.style === "key-value" && columnCount === 2) {
    return [2700, 6660];
  }

  const weights = Array.from({ length: columnCount }, (_, columnIndex) => {
    const longest = Math.max(
      ...node.rows.map((row) => Array.from(row[columnIndex] ?? "").length),
      1,
    );
    return Math.max(2, Math.min(longest, 14));
  });
  const totalWeight = weights.reduce((total, weight) => total + weight, 0);
  const widths = weights.map((weight) =>
    Math.floor((CONTENT_WIDTH_DXA * weight) / totalWeight),
  );
  const assigned = widths.reduce((total, width) => total + width, 0);
  widths[widths.length - 1] = (widths.at(-1) ?? 0) + CONTENT_WIDTH_DXA - assigned;
  return widths;
}

function cellParagraph(text: string, style: TemplateStyle, header = false): Paragraph {
  return new Paragraph({
    alignment: header ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { before: 80, after: 80 },
    children: [
      new TextRun({
        text,
        bold: header,
        size: style.bodySize,
        color: header ? style.headerText : "202124",
        font: header ? style.headingFont : style.bodyFont,
      }),
    ],
  });
}

function tableFromNode(node: TableNode, style: TemplateStyle): Table {
  const border = {
    style: BorderStyle.SINGLE,
    size: 6,
    color: "B7BDC5",
  };

  const columnWidths = columnWidthsFor(node);
  return new Table({
    width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
    indent: { size: TABLE_INDENT_DXA, type: WidthType.DXA },
    columnWidths,
    layout: TableLayoutType.FIXED,
    margins: {
      marginUnitType: WidthType.DXA,
      top: 80,
      bottom: 80,
      left: 120,
      right: 120,
    },
    borders: {
      top: border,
      bottom: border,
      left: border,
      right: border,
      insideHorizontal: border,
      insideVertical: border,
    },
    rows: node.rows.map(
      (row, rowIndex) =>
        new TableRow({
          tableHeader: node.hasHeader && rowIndex === 0,
          cantSplit: true,
          children: row.map(
            (cell, cellIndex) =>
              new TableCell({
                verticalAlign: VerticalAlignTable.CENTER,
                margins: {
                  marginUnitType: WidthType.DXA,
                  top: 80,
                  bottom: 80,
                  left: 120,
                  right: 120,
                },
                shading:
                  node.hasHeader && rowIndex === 0
                    ? { fill: style.accentFill, type: ShadingType.CLEAR }
                    : node.style === "key-value" && cellIndex === 0
                      ? { fill: "F4F5F6", type: ShadingType.CLEAR }
                      : undefined,
                width: {
                  size: columnWidths[cellIndex] ?? 0,
                  type: WidthType.DXA,
                },
                children: [
                  cellParagraph(
                    cell,
                    style,
                    (node.hasHeader && rowIndex === 0) ||
                      (node.style === "key-value" && cellIndex === 0),
                  ),
                ],
              }),
          ),
        }),
    ),
  });
}

function paragraphsFromNode(node: DocumentNode, style: TemplateStyle): (Paragraph | Table)[] {
  if (node.type === "table") {
    return [
      tableFromNode(node, style),
      new Paragraph({ spacing: { after: 120 }, children: [] }),
    ];
  }

  if (node.type === "heading") {
    return [
      new Paragraph({
        style: `OfficeHeading${node.level}`,
        children: [new TextRun(node.text)],
      }),
    ];
  }

  return [
    new Paragraph({
      style: "OfficeBody",
      children: [new TextRun(node.text)],
    }),
  ];
}

export function createDocxDocument(spec: DocumentSpec): Document {
  const style = TEMPLATES[spec.template];
  const children = spec.nodes.flatMap((node) => paragraphsFromNode(node, style));

  return new Document({
    creator: "Office File Agent",
    title: spec.title,
    description: "Generated from structured text by Office File Agent",
    styles: {
      default: {
        document: {
          run: { font: style.bodyFont, size: style.bodySize, color: "202124" },
          paragraph: { spacing: { after: style.bodyAfter, line: style.lineSpacing } },
        },
      },
      paragraphStyles: [
        {
          id: "OfficeTitle",
          name: "Office Title",
          basedOn: "Normal",
          next: "OfficeBody",
          quickFormat: true,
          run: {
            font: style.headingFont,
            size: style.titleSize,
            bold: true,
            color: "202124",
          },
          paragraph: {
            alignment: AlignmentType.CENTER,
            keepNext: true,
            spacing: { before: 0, after: style.titleAfter },
          },
        },
        {
          id: "OfficeBody",
          name: "Office Body",
          basedOn: "Normal",
          next: "OfficeBody",
          quickFormat: true,
          run: { font: style.bodyFont, size: style.bodySize, color: "202124" },
          paragraph: {
            alignment: style.bodyAlignment,
            spacing: { after: style.bodyAfter, line: style.lineSpacing },
          },
        },
        ...([1, 2, 3] as const).map((level) => ({
          id: `OfficeHeading${level}`,
          name: `Office Heading ${level}`,
          basedOn: `Heading${level}`,
          next: "OfficeBody",
          quickFormat: true,
          run: {
            font: style.headingFont,
            size: style.headings[level].size,
            bold: true,
            color: style.headings[level].color,
          },
          paragraph: {
            keepNext: true,
            outlineLevel: level - 1,
            spacing: {
              before: style.headings[level].before,
              after: style.headings[level].after,
            },
          },
        })),
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 12240, height: 15840 },
            margin: {
              top: 1440,
              right: 1440,
              bottom: 1440,
              left: 1440,
              header: 708,
              footer: 708,
            },
          },
        },
        children: [
          new Paragraph({
            style: "OfficeTitle",
            children: [new TextRun(spec.title)],
          }),
          ...children,
        ],
      },
    ],
  });
}

export async function renderDocxBlob(spec: DocumentSpec): Promise<Blob> {
  return Packer.toBlob(createDocxDocument(spec));
}

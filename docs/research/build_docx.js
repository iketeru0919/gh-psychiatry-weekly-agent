const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  Header, Footer, PageNumber, LevelFormat, convertInchesToTwip,
} = require("docx");

const SRC = process.argv[2];
const OUT = process.argv[3];

const FONT = "Yu Gothic";
const NAVY = "27436E";
const INK = "151C26";
const GREY = "4A5666";
const RULE = "D6DBE3";
const HEADBG = "EFF1F5";

// page: A4 portrait, 1in margins -> content width
const CONTENT_W = 9026; // twips

const md = fs.readFileSync(SRC, "utf8").split("\n");

// ---- inline **bold** -> TextRun[] ----
function runs(text, opts = {}) {
  text = text.replace(/\[([^\]]+)\]\(([^)]+)\)/g, "$1（$2）");
  const out = [];
  const parts = text.split(/(\*\*[^*]+\*\*|`[^`]+`)/g);
  for (const p of parts) {
    if (!p) continue;
    if (p.startsWith("**") && p.endsWith("**")) {
      out.push(new TextRun({ text: p.slice(2, -2), bold: true, font: FONT, ...opts }));
    } else if (p.startsWith("`") && p.endsWith("`")) {
      out.push(new TextRun({ text: p.slice(1, -1), font: "Consolas", ...opts }));
    } else {
      out.push(new TextRun({ text: p, font: FONT, ...opts }));
    }
  }
  return out.length ? out : [new TextRun({ text: "", font: FONT })];
}

function cell(text, { head = false, width } = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: head ? { type: ShadingType.CLEAR, fill: HEADBG, color: "auto" } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      spacing: { before: 0, after: 0, line: 264 },
      children: runs(text.trim(), { size: 18, bold: head, color: head ? GREY : INK }),
    })],
  });
}

const children = [];

function push(p) { children.push(p); }

function para(text, o = {}) {
  push(new Paragraph({
    spacing: { before: o.before ?? 0, after: o.after ?? 120, line: 300 },
    indent: o.indent,
    children: runs(text, { size: o.size ?? 20, color: o.color ?? INK }),
  }));
}

let i = 0;
let olInstance = 0;
while (i < md.length) {
  const line = md[i];
  const t = line.trim();

  // horizontal rule
  if (/^---+$/.test(t)) {
    push(new Paragraph({
      spacing: { before: 60, after: 200 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 1 } },
      children: [new TextRun({ text: "", font: FONT })],
    }));
    i++; continue;
  }

  // headings
  const h = /^(#{1,4})\s+(.*)$/.exec(t);
  if (h) {
    const lvl = h[1].length;
    const txt = h[2];
    if (lvl === 1) {
      push(new Paragraph({
        spacing: { before: 0, after: 160 },
        children: runs(txt, { size: 34, bold: true, color: INK }),
      }));
    } else if (lvl === 2) {
      push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 400, after: 160 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 4 } },
        children: runs(txt, { size: 26, bold: true, color: NAVY }),
      }));
    } else if (lvl === 3) {
      push(new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 300, after: 120 },
        children: runs(txt, { size: 22, bold: true, color: INK }),
      }));
    } else {
      push(new Paragraph({
        heading: HeadingLevel.HEADING_3,
        spacing: { before: 220, after: 100 },
        children: runs(txt, { size: 20, bold: true, color: GREY }),
      }));
    }
    i++; continue;
  }

  // fenced code
  if (t.startsWith("```")) {
    i++;
    const buf = [];
    while (i < md.length && !md[i].trim().startsWith("```")) { buf.push(md[i]); i++; }
    i++;
    push(new Paragraph({
      spacing: { before: 80, after: 200, line: 280 },
      shading: { type: ShadingType.CLEAR, fill: HEADBG, color: "auto" },
      border: {
        top: { style: BorderStyle.SINGLE, size: 4, color: RULE, space: 6 },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE, space: 6 },
        left: { style: BorderStyle.SINGLE, size: 4, color: RULE, space: 6 },
        right: { style: BorderStyle.SINGLE, size: 4, color: RULE, space: 6 },
      },
      children: [new TextRun({ text: buf.join("  "), font: "Consolas", size: 19 })],
    }));
    continue;
  }

  // table
  if (t.startsWith("|") && i + 1 < md.length && /^\|[\s:|-]+\|$/.test(md[i + 1].trim())) {
    const rowsRaw = [];
    while (i < md.length && md[i].trim().startsWith("|")) { rowsRaw.push(md[i].trim()); i++; }
    const cells = rowsRaw.map(r => r.replace(/^\||\|$/g, "").split("|"));
    const header = cells[0];
    const body = cells.slice(2);
    const n = header.length;
    const w = Math.floor(CONTENT_W / n);
    const widths = Array(n).fill(w);
    widths[n - 1] = CONTENT_W - w * (n - 1);
    const rows = [new TableRow({
      tableHeader: true,
      children: header.map((c, k) => cell(c, { head: true, width: widths[k] })),
    })];
    for (const r of body) {
      const rr = r.slice(0, n);
      while (rr.length < n) rr.push("");
      rows.push(new TableRow({ children: rr.map((c, k) => cell(c, { width: widths[k] })) }));
    }
    push(new Table({
      columnWidths: widths,
      width: { size: CONTENT_W, type: WidthType.DXA },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 4, color: RULE },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE },
        left: { style: BorderStyle.SINGLE, size: 4, color: RULE },
        right: { style: BorderStyle.SINGLE, size: 4, color: RULE },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: RULE },
        insideVertical: { style: BorderStyle.SINGLE, size: 4, color: RULE },
      },
      rows,
    }));
    push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "", font: FONT, size: 10 })] }));
    continue;
  }

  // blockquote
  if (t.startsWith(">")) {
    const buf = [];
    while (i < md.length && md[i].trim().startsWith(">")) {
      buf.push(md[i].trim().replace(/^>\s?/, "")); i++;
    }
    push(new Paragraph({
      spacing: { before: 100, after: 200, line: 300 },
      indent: { left: 240 },
      border: { left: { style: BorderStyle.SINGLE, size: 18, color: NAVY, space: 10 } },
      children: runs(buf.join(" "), { size: 19, color: GREY }),
    }));
    continue;
  }

  // bullet list
  if (/^[-*]\s+/.test(t)) {
    push(new Paragraph({
      numbering: { reference: "bullets", level: 0 },
      spacing: { after: 60, line: 300 },
      children: runs(t.replace(/^[-*]\s+/, ""), { size: 20 }),
    }));
    i++; continue;
  }

  // ordered list — a new numbering instance whenever the source restarts at 1,
  // so lists interleaved with continuation lines still number 1,2,3.
  const ol = /^(\d+)\.\s+(.*)$/.exec(t);
  if (ol) {
    if (Number(ol[1]) === 1 || olInstance === 0) olInstance++;
    push(new Paragraph({
      numbering: { reference: "numbers", level: 0, instance: olInstance },
      spacing: { after: 60, line: 300 },
      children: runs(ol[2], { size: 20 }),
    }));
    i++; continue;
  }

  // blank
  if (t === "") { i++; continue; }

  // italic-only trailing note
  if (/^\*[^*].*\*$/.test(t)) {
    push(new Paragraph({
      spacing: { before: 200, after: 120 },
      children: [new TextRun({ text: t.slice(1, -1), font: FONT, size: 18, italics: true, color: GREY })],
    }));
    i++; continue;
  }

  para(t.replace(/<br>/g, " "), { size: 20 });
  i++;
}

const doc = new Document({
  creator: "Claude Code",
  title: "就労継続支援B型 新モデル構築のための 学術データ・研究・実態データ集",
  description: "GH連携型 × ひきこもり・在宅就労型（2026年8月時点）",
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 360, hanging: 200 } } },
        }],
      },
      {
        reference: "numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 400, hanging: 240 } } },
        }],
      },
    ],
  },
  styles: {
    default: {
      document: { run: { font: FONT, size: 20, color: INK } },
    },
  },
  sections: [{
    properties: {
      page: {
        margin: {
          top: convertInchesToTwip(0.9), bottom: convertInchesToTwip(0.9),
          left: convertInchesToTwip(0.85), right: convertInchesToTwip(0.85),
        },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 80 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE, space: 4 } },
          children: [new TextRun({
            text: "就労継続支援B型 新モデル エビデンス集　／　2026年8月",
            font: FONT, size: 16, color: GREY,
          })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: GREY })],
        })],
      }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("wrote", OUT, b.length, "bytes"); });

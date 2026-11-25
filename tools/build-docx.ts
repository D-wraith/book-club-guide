// Builds exports/Book-Club-Guide-v<version>.docx from the docs/ pages, in menu order.
// No dependencies: writes the WordprocessingML parts and zips them (stored, CRC-32 via Bun).
// Usage: bun build-docx.ts [version]   (default: 0.10)
import { mkdirSync, readFileSync, writeFileSync } from "fs";
import { join } from "path";

const ROOT = join(import.meta.dir, "..");
const VERSION = process.argv[2] ?? "0.10";
const PAGES = ["index.md", "about.md", "meetings.md", "books.md", "join.md", "recaps/index.md", "contributing.md"];
const TABLE_W = 9360; // 6.5 in of US Letter text width, in DXA

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const links: string[] = []; // external URLs, one relationship each
const run = (t: string, rPr = "") => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}<w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;

/** **bold**, *italic*, [text](url). Links to other guide pages become plain text in the Word file. */
function inline(text: string, bold = false): string {
  const base = bold ? "<w:b/>" : "";
  let out = "", last = 0;
  for (const m of text.matchAll(/(\*\*[^*]+\*\*|\*[^*]+\*|\[[^\]]+\]\([^)]+\))/g)) {
    if (m.index! > last) out += run(text.slice(last, m.index), base);
    const t = m[0];
    if (t.startsWith("**")) out += run(t.slice(2, -2), "<w:b/>");
    else if (t.startsWith("[")) {
      const [, label, url] = t.match(/\[([^\]]+)\]\(([^)]+)\)/)!;
      if (/^https?:/.test(url)) { links.push(url); out += `<w:hyperlink r:id="rLink${links.length}">${run(label, '<w:rStyle w:val="Hyperlink"/>')}</w:hyperlink>`; }
      else out += run(label, base);
    } else out += run(t.slice(1, -1), `${base}<w:i/>`);
    last = m.index! + t.length;
  }
  if (last < text.length) out += run(text.slice(last), base);
  return out;
}
const para = (content: string, pPr = "") => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}${content}</w:p>`;

function table(rows: string[][]): string {
  const cols = rows[0].length;
  const weight = Array.from({ length: cols }, (_, c) => Math.max(8, Math.min(60, ...rows.map((r) => (r[c] ?? "").length))));
  const sum = weight.reduce((a, b) => a + b, 0);
  const w = weight.map((x) => Math.floor((x / sum) * TABLE_W));
  w[cols - 1] += TABLE_W - w.reduce((a, b) => a + b, 0);
  const b = '<w:top w:val="single" w:sz="4" w:color="BFBFBF"/><w:left w:val="single" w:sz="4" w:color="BFBFBF"/><w:bottom w:val="single" w:sz="4" w:color="BFBFBF"/><w:right w:val="single" w:sz="4" w:color="BFBFBF"/><w:insideH w:val="single" w:sz="4" w:color="BFBFBF"/><w:insideV w:val="single" w:sz="4" w:color="BFBFBF"/>';
  const trs = rows.map((r, i) => `<w:tr>${i === 0 ? "<w:trPr><w:tblHeader/></w:trPr>" : ""}${r.map((cell, c) =>
    `<w:tc><w:tcPr><w:tcW w:w="${w[c]}" w:type="dxa"/>${i === 0 ? '<w:shd w:val="clear" w:color="auto" w:fill="E8EEF7"/>' : ""}</w:tcPr>${para(inline(cell, i === 0), '<w:spacing w:after="0"/>')}</w:tc>`).join("")}</w:tr>`).join("");
  return `<w:tbl><w:tblPr><w:tblW w:w="${TABLE_W}" w:type="dxa"/><w:tblBorders>${b}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${w.map((x) => `<w:gridCol w:w="${x}"/>`).join("")}</w:tblGrid>${trs}</w:tbl>${para("")}`;
}

const REPO_URL = "https://github.com/Scrub-Jay-Collective/bookclub";
let body = "";
PAGES.forEach((page, p) => {
  const lines = readFileSync(join(ROOT, "docs", page), "utf8").replace(/^---[\s\S]*?---\n/, "").split("\n");
  if (p > 0) body += para('<w:r><w:br w:type="page"/></w:r>');
  let rows: string[][] = [], skipNav = false;
  const flush = () => { if (rows.length) { body += table(rows); rows = []; } };
  for (const raw of lines) {
    // Site-only HTML wrappers (the home page's Next meeting box) have no place in the Word file;
    // the Liquid repo URL becomes the real one.
    const line = raw.trimEnd().replaceAll("{{ site.github.repository_url }}", REPO_URL);
    if (/^<\/?div/.test(line)) continue;
    if (line.startsWith("|")) {
      const cells = line.split("|").slice(1, -1).map((c) => c.trim());
      if (!cells.every((c) => /^:?-+:?$/.test(c))) rows.push(cells);
      continue;
    }
    flush();
    if (!line.trim()) continue;
    // The home page's link list is site navigation; the Word file has its own page order.
    if (page === "index.md" && line.startsWith("## Find your way around")) { skipNav = true; continue; }
    if (skipNav) continue;
    if (line.startsWith("# ")) body += para(inline(line.slice(2)), `<w:pStyle w:val="${p === 0 ? "Title" : "Heading1"}"/>`);
    else if (line.startsWith("## ")) body += para(inline(line.slice(3)), '<w:pStyle w:val="Heading2"/>');
    else if (line.startsWith("### ")) body += para(inline(line.slice(4)), '<w:pStyle w:val="Heading3"/>');
    else if (/^\s*- /.test(line)) body += para(inline(line.replace(/^\s*- /, "")), '<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');
    else if (/^\d+\. /.test(line)) body += para(inline(line.replace(/^\d+\. /, "")), '<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>');
    else body += para(inline(line));
  }
  flush();
});

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}<w:sectPr><w:footerReference w:type="default" r:id="rFooter"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`;
const heading = (id: string, name: string, size: number, before: number, after: number, outline: number) =>
  `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="${before}" w:after="${after}"/>${outline >= 0 ? `<w:outlineLvl w:val="${outline}"/>` : ""}</w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:color w:val="1F3864"/><w:sz w:val="${size}"/></w:rPr></w:style>`;
const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`
  + `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>`
  + heading("Title", "Title", 48, 0, 160, -1) + heading("Heading1", "heading 1", 36, 0, 160, 0) + heading("Heading2", "heading 2", 28, 240, 100, 1) + heading("Heading3", "heading 3", 24, 200, 80, 2)
  + `<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="60"/><w:ind w:left="720"/></w:pPr></w:style>`
  + `<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>`
  + `<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/><w:spacing w:after="0"/></w:pPr><w:rPr><w:color w:val="7F7F7F"/><w:sz w:val="18"/></w:rPr></w:style></w:styles>`;
const level = (fmt: string, text: string) => `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>${fmt === "bullet" ? '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>' : ""}</w:lvl>`;
const numbering = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${NS}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/>${level("bullet", "")}</w:abstractNum><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="singleLevel"/>${level("decimal", "%1.")}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
const footer = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${NS}><w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr>${run(`Book Club Guide · v${VERSION} · page `)}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`;
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rStyles" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rNumbering" Type="${R}/numbering" Target="numbering.xml"/><Relationship Id="rFooter" Type="${R}/footer" Target="footer1.xml"/>${links.map((u, i) => `<Relationship Id="rLink${i + 1}" Type="${R}/hyperlink" Target="${esc(u)}" TargetMode="External"/>`).join("")}</Relationships>`;
const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
const core = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Book Club Guide</dc:title><dc:description>Book Club Guide v${esc(VERSION)}</dc:description><dc:creator>Book Club founders</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`;
const types = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`;
const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="${R}/officeDocument" Target="word/document.xml"/><Relationship Id="rCore" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`;

/** Minimal ZIP writer: stored entries, no folder entries (Word refuses some files that have them). */
function zip(files: [string, string][]): Uint8Array {
  const enc = new TextEncoder(); const chunks: Uint8Array[] = []; const central: Uint8Array[] = []; let offset = 0;
  const u16 = (v: number) => [v & 255, (v >> 8) & 255], u32 = (v: number) => [v & 255, (v >> 8) & 255, (v >> 16) & 255, (v >>> 24) & 255];
  for (const [name, text] of files) {
    const data = enc.encode(text), nm = enc.encode(name), crc = Bun.hash.crc32(data) >>> 0;
    const common = [...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21), ...u32(crc), ...u32(data.length), ...u32(data.length), ...u16(nm.length), ...u16(0)];
    const local = new Uint8Array([...u32(0x04034b50), ...common, ...nm]);
    chunks.push(local, data);
    central.push(new Uint8Array([...u32(0x02014b50), ...u16(20), ...common, ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(offset), ...nm]));
    offset += local.length + data.length;
  }
  const cdSize = central.reduce((a, c) => a + c.length, 0);
  const end = new Uint8Array([...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(files.length), ...u16(files.length), ...u32(cdSize), ...u32(offset), ...u16(0)]);
  const all = [...chunks, ...central, end]; const out = new Uint8Array(all.reduce((a, c) => a + c.length, 0)); let p = 0;
  for (const c of all) { out.set(c, p); p += c.length; }
  return out;
}

mkdirSync(join(ROOT, "exports"), { recursive: true });
const out = join(ROOT, "exports", `Book-Club-Guide-v${VERSION}.docx`);
writeFileSync(out, zip([["[Content_Types].xml", types], ["_rels/.rels", rootRels], ["word/document.xml", document], ["word/_rels/document.xml.rels", docRels], ["word/styles.xml", styles], ["word/numbering.xml", numbering], ["word/footer1.xml", footer], ["docProps/core.xml", core]]));
console.log(`wrote ${out} (${links.length} links)`);

// Builds exports/Book-Club-Edit-Log-v<version>.xlsx from data/*.csv and private/*.csv.
// Usage: bun build-xlsx.ts [version]   (default: 0.10)
import { existsSync, mkdirSync, readFileSync } from "fs";
import { join } from "path";
import ExcelJS from "exceljs";

const ROOT = join(import.meta.dir, "..");
const VERSION = process.argv[2] ?? "0.10";

/** RFC 4180 CSV: quoted fields, doubled quotes, commas and newlines inside quotes. */
function parseCsv(text: string): string[][] {
  const rows: string[][] = []; let row: string[] = []; let f = ""; let q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) { if (ch === '"' && text[i + 1] === '"') { f += '"'; i++; } else if (ch === '"') q = false; else f += ch; }
    else if (ch === '"') q = true;
    else if (ch === ",") { row.push(f); f = ""; }
    else if (ch === "\n") { row.push(f); rows.push(row); row = []; f = ""; }
    else if (ch !== "\r") f += ch;
  }
  if (f || row.length) { row.push(f); rows.push(row); }
  return rows;
}
const load = (rel: string) => (existsSync(join(ROOT, rel)) ? parseCsv(readFileSync(join(ROOT, rel), "utf8")) : []);

const wb = new ExcelJS.Workbook();
wb.creator = "Book Club founders";
wb.title = `Book Club Edit Log v${VERSION}`;

function sheet(name: string, rows: string[][], widths: number[], opts: { status?: number; options?: string[]; tab?: string } = {}) {
  const ws = wb.addWorksheet(name, { views: [{ state: "frozen", ySplit: 1 }], properties: { tabColor: opts.tab ? { argb: opts.tab } : undefined } });
  ws.columns = widths.map((w) => ({ width: w }));
  rows.forEach((r) => ws.addRow(r));
  const head = ws.getRow(1);
  head.font = { bold: true, color: { argb: "FF1F3864" } };
  head.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8EEF7" } };
  ws.eachRow((row) => row.eachCell((c) => { c.alignment = { vertical: "top", wrapText: true }; }));
  if (rows.length) ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: rows[0].length } };
  if (opts.status) {
    // A drop-down on the status column, with room for new rows.
    for (let r = 2; r <= 300; r++) ws.getCell(r, opts.status).dataValidation = { type: "list", allowBlank: true, formulae: [`"${opts.options!.join(",")}"`] };
  }
  return ws;
}

const readMe = sheet("Read me", [
  ["Book Club Edit Log", ""],
  ["Version", `${VERSION}`],
  ["What this is", "The log of every suggested change to the Book Club Guide, plus the organizers' tasks and meeting records. The Guide (Google Doc, later GitHub Pages) is the published text; this workbook is where changes are proposed and decided."],
  ["Edit Log", "One row per suggested change. Status is Proposed, Accepted or Declined. Accepted rows are applied to the Guide and added to Revision history. IDs stay stable: E = edit, H = housekeeping, R = restructure."],
  ["Actions", "Organizer tasks with owner and due date. Internal: never published."],
  ["Books", "Book suggestions B01–B10. The order is set by the organizers."],
  ["Meeting records", "One row per meeting. Recap drafts stay private until someone has reviewed them for personal information; leaving personal details out entirely is best. Only reviewed recaps are published."],
  ["Revision history", "One row per Guide version."],
  ["Privacy", "No phone numbers, emails or home addresses in this workbook."],
], [22, 110]);
readMe.getRow(1).font = { bold: true, size: 14, color: { argb: "FF1F3864" } };
readMe.getColumn(1).font = { bold: true };

sheet("Edit Log", load("private/edit-log.csv"), [12, 7, 26, 45, 45, 22, 12, 18, 40], { status: 7, options: ["Proposed", "Accepted", "Declined"], tab: "FF4472C4" });
sheet("Actions", load("private/actions.csv"), [7, 55, 20, 16, 22, 30], { tab: "FFED7D31" });
sheet("Books", load("data/books.csv"), [7, 38, 20, 32, 50], { tab: "FF70AD47" });
sheet("Meeting records", load("private/meeting-records.csv"), [12, 22, 11, 14, 12, 12, 30, 18, 11, 50], { status: 9, options: ["No", "Yes"], tab: "FFA5A5A5" });
sheet("Revision history", load("data/revision-history.csv"), [11, 14, 90]);

mkdirSync(join(ROOT, "exports"), { recursive: true });
const out = join(ROOT, "exports", `Book-Club-Edit-Log-v${VERSION}.xlsx`);
await wb.xlsx.writeFile(out);
console.log(`wrote ${out}`);

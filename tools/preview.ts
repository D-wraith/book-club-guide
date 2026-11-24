// Local preview of the GitHub Pages site: docs/*.md → exports/preview/*.html with the same menu.
// An approximation of the minima theme, for checking content and navigation before Pages is on.
// Usage: bun preview.ts
import { mkdirSync, readFileSync, writeFileSync } from "fs";
import { join } from "path";

const ROOT = join(import.meta.dir, "..");
const OUT = join(ROOT, "exports", "preview");
const MENU: [string, string][] = [["about.md", "About"], ["meetings.md", "How meetings work"], ["books.md", "Books"], ["join.md", "Join us"], ["recaps/index.md", "Recaps"]];
const PAGES = ["index.md", ...MENU.map((m) => m[0]), "contributing.md"];
const html = (p: string) => p.replace(/\.md$/, ".html").replace("/", "-");
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const inline = (s: string) => esc(s)
  .replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>").replace(/\*([^*]+)\*/g, "<em>$1</em>")
  .replace(/\[([^\]]+)\]\(([^)]+)\)/g, (_, t, u) => `<a href="${/^https?:/.test(u) ? u : html(u)}">${t}</a>`);

function render(md: string): string {
  let out = "", list = "", rows: string[][] = [];
  const close = () => { if (list) { out += `</${list}>`; list = ""; } if (rows.length) { out += `<table><thead><tr>${rows[0].map((c) => `<th>${inline(c)}</th>`).join("")}</tr></thead><tbody>${rows.slice(1).map((r) => `<tr>${r.map((c) => `<td>${inline(c)}</td>`).join("")}</tr>`).join("")}</tbody></table>`; rows = []; } };
  for (const line of md.replace(/^---[\s\S]*?---\n/, "").split("\n")) {
    if (line.startsWith("|")) { const c = line.split("|").slice(1, -1).map((x) => x.trim()); if (!c.every((x) => /^:?-+:?$/.test(x))) rows.push(c); continue; }
    const li = line.match(/^\s*- (.*)/), ol = line.match(/^\d+\. (.*)/);
    if (li || ol) { const t = li ? "ul" : "ol"; if (rows.length || list !== t) { close(); out += `<${t}>`; list = t; } out += `<li>${inline((li ?? ol)![1])}</li>`; continue; }
    close();
    const h = line.match(/^(#{1,3}) (.*)/);
    if (h) out += `<h${h[1].length}>${inline(h[2])}</h${h[1].length}>`;
    else if (line.trim()) out += `<p>${inline(line)}</p>`;
  }
  close();
  return out;
}

const css = `:root{--bg:#fdfdfd;--fg:#111;--muted:#828282;--line:#e8e8e8;--link:#2a7ae2;--head:#f5f8fc}
@media (prefers-color-scheme:dark){:root{--bg:#181818;--fg:#e6e6e6;--muted:#9a9a9a;--line:#333;--link:#79b8ff;--head:#222a35}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:16px/1.6 -apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif}
header{border-top:5px solid #424242;border-bottom:1px solid var(--line)}.wrap{max-width:800px;margin:0 auto;padding:0 16px}
header .wrap{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:center;gap:8px;min-height:56px}
.site{font-size:24px;font-weight:300;color:var(--fg);text-decoration:none}nav a{color:var(--fg);text-decoration:none;margin-left:16px;white-space:nowrap}nav a.on{font-weight:600}
main{padding:24px 0 48px}h1{font-size:2rem;line-height:1.2}h2{font-size:1.5rem;margin-top:2rem}h3{font-size:1.15rem}a{color:var(--link)}
table{border-collapse:collapse;width:100%;margin:1rem 0;display:block;overflow-x:auto}th,td{border:1px solid var(--line);padding:8px 10px;text-align:left;vertical-align:top}th{background:var(--head)}
footer{border-top:1px solid var(--line);color:var(--muted);font-size:14px;padding:16px 0}.banner{background:#fff4ce;color:#5c4400;padding:8px 16px;text-align:center;font-size:14px}
@media (max-width:600px){nav{width:100%}nav a{display:inline-block;margin:0 12px 8px 0}}`;

mkdirSync(OUT, { recursive: true });
for (const page of PAGES) {
  const body = render(readFileSync(join(ROOT, "docs", page), "utf8"));
  const nav = MENU.map(([p, t]) => `<a href="${html(p)}"${p === page ? ' class="on"' : ""}>${t}</a>`).join("");
  writeFileSync(join(OUT, html(page)), `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Book Club Guide</title><style>${css}</style></head><body><div class="banner">Local preview: the real site goes live on GitHub Pages when the repo is made public.</div><header><div class="wrap"><a class="site" href="index.html">Book Club Guide</a><nav>${nav}</nav></div></header><main><div class="wrap">${body}</div></main><footer><div class="wrap">Book Club Guide · <a href="contributing.html">How this guide stays current</a></div></footer></body></html>`);
}
console.log(`wrote ${PAGES.length} pages to ${OUT}`);

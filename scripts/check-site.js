const fs = require("fs");
const path = require("path");
const root = path.resolve(__dirname, "..");
const expectedMacroCount = 21;
const htmlFiles = [];
function walk(dir) {
  for (const item of fs.readdirSync(dir, { withFileTypes: true })) {
    if (item.name === ".git" || item.name === ".vs") continue;
    const p = path.join(dir, item.name);
    if (item.isDirectory()) walk(p);
    else if (p.endsWith(".html")) htmlFiles.push(p);
  }
}
function existsHref(file, href) {
  if (/^(https?:|mailto:|#)/.test(href)) return true;
  const clean = href.split("#")[0];
  if (!clean) return true;
  const target = clean.startsWith("/")
    ? path.join(root, clean.slice(1))
    : path.resolve(path.dirname(file), clean);
  if (fs.existsSync(target)) return true;
  if (fs.existsSync(target + ".html")) return true;
  if (fs.existsSync(path.join(target, "index.html"))) return true;
  return false;
}
walk(root);
let failures = [];
for (const file of htmlFiles) {
  const html = fs.readFileSync(file, "utf8");
  for (const match of html.matchAll(/(?:href|src)="([^"]+)"/g)) {
    if (!existsHref(file, match[1])) failures.push(path.relative(root, file) + " -> " + match[1]);
  }
}
const macroSources = fs.readdirSync(path.join(root, "macros", "src")).filter((f) => f.endsWith(".bas"));
if (macroSources.length !== expectedMacroCount) failures.push("Expected " + expectedMacroCount + " .bas files, found " + macroSources.length);
if (failures.length) { console.error(failures.join("\n")); process.exit(1); }
console.log("Checked " + htmlFiles.length + " HTML files and " + macroSources.length + " macro sources.");

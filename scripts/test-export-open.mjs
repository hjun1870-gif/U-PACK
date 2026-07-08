/**
 * 데모 워크북 export → xlsx 저장 → (Windows) Excel 실행
 */
import { createRequire } from "module";
import fs from "fs";
import path from "path";
import vm from "vm";
import { fileURLToPath } from "url";
import { execSync } from "child_process";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(__dirname, "..");
const outPath = path.join(root, "재고조사표_테스트_export.xlsx");

const require = createRequire(import.meta.url);
const XLSX = require("xlsx-js-style");

const lines = fs.readFileSync(path.join(root, "app.js"), "utf8").split("\n");
const appCode = lines.slice(0, 910).concat(lines.slice(1034)).join("\n");

const mockEl = () => ({
  style: {},
  classList: { add() {}, remove() {}, toggle() {}, contains: () => false },
  appendChild() {},
  addEventListener() {},
  value: "",
  textContent: "",
});
const sb = {
  XLSX,
  console,
  localStorage: { getItem: () => null, setItem() {} },
  document: {
    getElementById: () => mockEl(),
    querySelector: () => mockEl(),
    querySelectorAll: () => [],
    createElement: () => mockEl(),
    head: { appendChild() {} },
    body: { appendChild() {} },
  },
  supabaseClient: null,
  setTimeout,
  clearTimeout,
  URL: { createObjectURL: () => "", revokeObjectURL() {} },
  Blob,
  Uint8Array,
  ArrayBuffer,
  fetch: async () => ({ ok: false }),
  alert: () => {},
  navigator: { onLine: true },
  showToast: () => {},
  indexedDB: null,
  crypto: globalThis.crypto,
  performance: { now: () => Date.now() },
};
sb.window = sb;
sb.addEventListener = () => {};

vm.createContext(sb);
vm.runInContext(appCode, sb, { filename: "app.js" });

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, sb.createMockSheet("B2"), "B2");
XLSX.utils.book_append_sheet(wb, sb.createMockSheet("B1"), "B1");
XLSX.utils.book_append_sheet(wb, sb.createMockSheet("B3"), "B3");

const packSheet = {
  A1: { t: "s", v: "입수량 · 박스단위 기준표" },
  A2: { t: "s", v: "제품" },
  B2: { t: "s", v: "입수량" },
  C2: { t: "s", v: "박스단위" },
  A3: { t: "s", v: "M300" },
  B3: { t: "n", v: 210 },
  C3: { t: "n", v: 20 },
  A4: { t: "s", v: "K500" },
  B4: { t: "n", v: 120 },
  C4: { t: "n", v: 20 },
  "!ref": "A1:C4",
};
XLSX.utils.book_append_sheet(wb, packSheet, "입수량 박스단위");

sb.ensureTotalSheet(wb);
sb.aggregateTotalSheet(wb);
sb.applyExcelStyles(wb);

const results = [];
for (const name of wb.SheetNames) {
  const sheet = wb.Sheets[name];
  if (name.includes("입수량")) {
    results.push(`${name}: cols=${sheet["!cols"]?.length}, merges=${sheet["!merges"]?.length ?? 0}`);
    continue;
  }
  if (name.toUpperCase() === "TOTAL") {
    results.push(`${name}: styled`);
    continue;
  }
  const ok = sb.hasWarehouseLayout(sheet);
  const merges = sheet["!merges"]?.length ?? 0;
  const sample = sheet.B8?.s?.fill?.fgColor?.rgb;
  results.push(`${name}: layout=${ok}, merges=${merges}, B8fill=${sample}`);
}

const wbout = XLSX.write(wb, { bookType: "xlsx", type: "buffer", cellStyles: true });
fs.writeFileSync(outPath, wbout);

console.log("Saved:", outPath);
console.log("Size:", fs.statSync(outPath).size, "bytes");
results.forEach((r) => console.log(" ", r));

try {
  execSync(`start "" "${outPath}"`, { shell: true, stdio: "inherit" });
  console.log("Opened in default Excel app.");
} catch (e) {
  console.error("Could not open Excel:", e.message);
  process.exit(1);
}

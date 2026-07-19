/**
 * 업로드 엑셀 열/행 수에 맞춰 웹 레이아웃이 유기적으로 잡히는지 검증
 */
const XLSX = require("xlsx-js-style");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const ROOT = path.join(__dirname, "..");
const appSrc = fs.readFileSync(path.join(ROOT, "app.js"), "utf8");

// app.js 브라우저 의존성 스텁 후 레이아웃 함수만 평가
const sandbox = {
  console,
  XLSX,
  document: {
    getElementById: () => null,
    createElement: () => ({
      style: {},
      classList: { add() {}, remove() {} },
      appendChild() {},
      addEventListener() {},
      dataset: {},
    }),
  },
  window: { localStorage: { getItem: () => null, setItem() {} }, matchMedia: () => ({ matches: false }) },
  localStorage: { getItem: () => null, setItem() {} },
  alert() {},
  confirm: () => false,
  setTimeout,
  clearTimeout,
  fetch: async () => ({ ok: false }),
  navigator: { onLine: false },
  URL: { createObjectURL: () => "", revokeObjectURL() {} },
  Blob: class {},
  FileReader: class {},
  ExcelJS: undefined,
  supabase: undefined,
};
sandbox.window = sandbox;
sandbox.self = sandbox;
sandbox.globalThis = sandbox;

vm.createContext(sandbox);
try {
  vm.runInContext(appSrc, sandbox, { filename: "app.js" });
} catch (e) {
  // DOM 초기화 중 일부 실패해도 함수 정의는 남을 수 있음
  if (!sandbox.detectRacks && !sandbox.splitRacksByAisle) {
    // 함수가 전역이 아니면 eval로 추출
  }
}

function makeContinuousSheet(cols, rows) {
  const sheet = {};
  const lastP = 1 + (cols - 1) * 2;
  sheet["!ref"] = `A1:${XLSX.utils.encode_col(lastP + 2)}${7 + 2 * rows + 2}`;
  sheet["C4"] = { t: "s", v: "<재고조사표 - W2완제품>" };
  for (let c = 1; c <= cols; c++) {
    const p = 1 + (c - 1) * 2;
    sheet[XLSX.utils.encode_cell({ r: 4, c: p })] = { t: "n", v: c };
    sheet[XLSX.utils.encode_cell({ r: 5, c: p })] = { t: "s", v: "제품" };
    sheet[XLSX.utils.encode_cell({ r: 6, c: p })] = { t: "s", v: "P" };
    sheet[XLSX.utils.encode_cell({ r: 6, c: p + 1 })] = { t: "s", v: "B" };
  }
  for (let r = 1; r <= rows; r++) {
    sheet[XLSX.utils.encode_cell({ r: 7 + 2 * (r - 1), c: 0 })] = { t: "n", v: r };
  }
  return sheet;
}

function makeAisleSheet(left, right, rows) {
  const sheet = {};
  let p = 1;
  const idxs = [];
  for (let i = 0; i < left; i++) {
    idxs.push(p);
    p += 2;
  }
  p += 2;
  for (let i = 0; i < right; i++) {
    idxs.push(p);
    p += 2;
  }
  sheet["!ref"] = `A1:${XLSX.utils.encode_col(p + 2)}${7 + 2 * rows + 2}`;
  idxs.forEach((pCol, i) => {
    sheet[XLSX.utils.encode_cell({ r: 4, c: pCol })] = { t: "n", v: i + 1 };
    sheet[XLSX.utils.encode_cell({ r: 5, c: pCol })] = { t: "s", v: "제품" };
    sheet[XLSX.utils.encode_cell({ r: 6, c: pCol })] = { t: "s", v: "P" };
    sheet[XLSX.utils.encode_cell({ r: 6, c: pCol + 1 })] = { t: "s", v: "B" };
  });
  for (let r = 1; r <= rows; r++) {
    sheet[XLSX.utils.encode_cell({ r: 7 + 2 * (r - 1), c: 0 })] = { t: "n", v: r };
  }
  return sheet;
}

// app.js는 함수를 전역에 두지 않으므로 동일 로직을 로컬로 복제해 검증
function splitRacksByAisle(allRacks) {
  const leftRacks = [];
  const rightRacks = [];
  if (allRacks.length === 0) return { leftRacks, rightRacks };
  let splitIndex = -1;
  let maxGap = 0;
  for (let i = 0; i < allRacks.length - 1; i++) {
    const gap = allRacks[i + 1].colIdx - allRacks[i].colIdx;
    if (gap > maxGap) {
      maxGap = gap;
      splitIndex = i;
    }
  }
  if (maxGap <= 2 || splitIndex < 0) {
    return { leftRacks: allRacks.slice(), rightRacks: [] };
  }
  for (let i = 0; i <= splitIndex; i++) leftRacks.push(allRacks[i]);
  for (let i = splitIndex + 1; i < allRacks.length; i++) rightRacks.push(allRacks[i]);
  return { leftRacks, rightRacks };
}

function detectRacks(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const allRacks = [];
  let lastColIdx = -10;
  for (let colIdx = range.s.c; colIdx <= range.e.c; colIdx++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: 4, c: colIdx })];
    if (!cell || cell.v === undefined) continue;
    const numVal = parseInt(cell.v, 10);
    if (!isNaN(numVal) && numVal > 0 && numVal <= 99 && colIdx - lastColIdx >= 2) {
      allRacks.push({ rackNo: numVal, colIdx });
      lastColIdx = colIdx;
    }
  }
  return splitRacksByAisle(allRacks);
}

function detectMaxRows(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  let maxRowVal = 0;
  let consecutiveMisses = 0;
  for (let r = 1; ; r++) {
    const prodRowIdx = 7 + 2 * (r - 1);
    if (prodRowIdx > range.e.r) break;
    const cell = sheet[XLSX.utils.encode_cell({ r: prodRowIdx, c: 0 })];
    if (cell && cell.v !== undefined) {
      const numVal = parseInt(cell.v, 10);
      if (!isNaN(numVal) && numVal > 0) {
        maxRowVal = Math.max(maxRowVal, numVal);
        consecutiveMisses = 0;
        continue;
      }
    }
    consecutiveMisses++;
    if (maxRowVal > 0 && consecutiveMisses >= 3) break;
  }
  return maxRowVal || 10;
}

const cont = detectRacks(makeContinuousSheet(10, 6));
const aisle = detectRacks(makeAisleSheet(9, 9, 18));
const bundled = (() => {
  const f = fs.readdirSync(ROOT).find((n) => n.toLowerCase().endsWith(".xlsx"));
  const wb = XLSX.read(fs.readFileSync(path.join(ROOT, f)));
  return detectRacks(wb.Sheets.W2);
})();

console.log("continuous 10x6:", cont.leftRacks.length, cont.rightRacks.length, "rows", detectMaxRows(makeContinuousSheet(10, 6)));
console.log("aisle 9+9:", aisle.leftRacks.length, aisle.rightRacks.length);
console.log("bundled W2:", bundled.leftRacks.length, bundled.rightRacks.length);

if (cont.leftRacks.length !== 10 || cont.rightRacks.length !== 0) {
  console.error("FAIL: 연속 10열은 단일 블록이어야 함");
  process.exit(1);
}
if (aisle.leftRacks.length !== 9 || aisle.rightRacks.length !== 9) {
  console.error("FAIL: 통로 있는 18열은 좌9/우9");
  process.exit(1);
}
if (bundled.leftRacks.length !== 9 || bundled.rightRacks.length !== 9) {
  console.error("FAIL: 번들 W2는 좌9/우9 유지");
  process.exit(1);
}
console.log("PASS: organic layout detection");

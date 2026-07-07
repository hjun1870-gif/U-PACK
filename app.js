/**
 * app.js
 * 창고 재고조사표 시각화 및 편집 애플리케이션의 핵심 비즈니스 로직
 * - SheetJS (xlsx.js) 라이브러리를 활용하여 로컬 엑셀 파일을 파싱하고 갱신
 * - 시트별 동적 컬럼 매핑으로 구조 변경 유연성 제공
 * - 실시간 재고 검색, 수정 및 다운로드 기능 제공
 */

// 애플리케이션 상태 관리 전역 변수
let currentWorkbook = null; // 로드된 전체 엑셀 워크북 객체
let currentSheetName = "";  // 현재 화면에 렌더링 중인 시트 이름
let currentSheetData = null; // 현재 시트의 파싱된 엑셀 데이터 객체
let selectedCellInfo = null; // 현재 편집 모달이 열려 있는 셀의 정보

// Supabase 연동 상태 변수
let supabaseClient = null;
let isOnline = false;
let productsMetadataCache = {}; // 제품 메타데이터 캐시 (TOTAL 시트 갱신용)
let isUserLocalWorkbook = false; // 사용자가 업로드한 엑셀이 우선 (Supabase 템플릿 덮어쓰기 방지)
let supabaseLoadToken = 0; // 진행 중인 Supabase 로드 취소용 토큰
let isDbImportInProgress = false; // DB 일괄 덮어쓰기 중 realtime/서버로드 차단

/** 모바일·태블릿 뷰포트 여부 */
function isMobileViewport() {
  return window.matchMedia("(max-width: 768px)").matches;
}

/** 모바일에서 가로 스크롤이 자연스럽게 되도록 격자 최소 너비 적용 */
function applyResponsiveGridLayout(totalCols) {
  if (!gridBoard) return;
  if (isMobileViewport()) {
    const minWidth = Math.max(920, 80 + totalCols * 58 + 50);
    gridBoard.style.width = `${minWidth}px`;
    gridBoard.style.maxWidth = "none";
  } else {
    gridBoard.style.width = "";
  }
}

/**
 * 랙 셀 터치/클릭 — 스크롤과 탭을 구분해 모바일에서도 편집 모달이 열리게 함
 */
function bindRackCellTap(cell, onActivate) {
  let touchStartX = 0;
  let touchStartY = 0;
  let touchMoved = false;
  let lastTouchEnd = 0;

  cell.addEventListener(
    "touchstart",
    (e) => {
      touchMoved = false;
      const t = e.touches[0];
      touchStartX = t.clientX;
      touchStartY = t.clientY;
    },
    { passive: true }
  );

  cell.addEventListener(
    "touchmove",
    (e) => {
      const t = e.touches[0];
      if (
        Math.abs(t.clientX - touchStartX) > 12 ||
        Math.abs(t.clientY - touchStartY) > 12
      ) {
        touchMoved = true;
      }
    },
    { passive: true }
  );

  cell.addEventListener(
    "touchend",
    (e) => {
      if (touchMoved) return;
      e.preventDefault();
      lastTouchEnd = Date.now();
      onActivate();
    },
    { passive: false }
  );

  cell.addEventListener("click", () => {
    if (Date.now() - lastTouchEnd < 400) return;
    onActivate();
  });
}

// DOM 요소 캐싱
const fileInput = document.getElementById("fileInput");
const demoBtn = document.getElementById("demoBtn");
const sheetSelect = document.getElementById("sheetSelect");
const searchInput = document.getElementById("searchInput");
const searchResults = document.getElementById("searchResults");     // 결과 목록 패널
const searchResultsWrap = document.getElementById("searchResultsWrap"); // 래퍼: 이것을 show/hide하여 창고 도면을 밀어냄
const exportBtn = document.getElementById("exportBtn");
const gridBoard = document.getElementById("gridBoard");
const emptyGuide = document.getElementById("emptyGuide");

// 요약 통계 요소
const totalProductsEl = document.getElementById("totalProducts");
const totalPalletsEl = document.getElementById("totalPallets");
const totalBoxesEl = document.getElementById("totalBoxes");
const layoutInfoEl = document.getElementById("layoutInfo");

// 모달 관련 요소
const editModal = document.getElementById("editModal");
const closeModal = document.getElementById("closeModal");
const modalCellTitle = document.getElementById("modalCellTitle");
const inputProduct = document.getElementById("inputProduct");
const inputPallet = document.getElementById("inputPallet");
const inputBox = document.getElementById("inputBox");
const cancelBtn = document.getElementById("cancelBtn");
const saveBtn = document.getElementById("saveBtn");

// DB 설정 모달 캐싱
const dbConfigBtn = document.getElementById("dbConfigBtn");
const dbConfigModal = document.getElementById("dbConfigModal");
const closeDbModal = document.getElementById("closeDbModal");
const inputDbUrl = document.getElementById("inputDbUrl");
const inputDbKey = document.getElementById("inputDbKey");
const disconnectDbBtn = document.getElementById("disconnectDbBtn");
const saveDbConfigBtn = document.getElementById("saveDbConfigBtn");
const syncStatusEl = document.getElementById("syncStatus");
const mobileScrollHint = document.getElementById("mobileScrollHint");
const autoSyncOnUploadInput = document.getElementById("autoSyncOnUpload");
const cloudSyncBtn = document.getElementById("cloudSyncBtn");

/** 엑셀 업로드 시 Supabase 자동 동기화 여부 (기본값: 켜짐) */
function isAutoSyncOnUploadEnabled() {
  const stored = localStorage.getItem("auto_sync_on_upload");
  return stored === null || stored === "true";
}

const WORKBOOK_IDB_NAME = "upack-workbook-cache";
const WORKBOOK_IDB_STORE = "files";
const WORKBOOK_IDB_KEY = "current";
const WORKBOOK_STORAGE_BUCKET = "workbooks";
const WORKBOOK_STORAGE_PATH = "latest.xlsx";
const WORKBOOK_SNAPSHOT_ID = "latest";
const WORKBOOK_ARCHIVE_SHEET = "__WORKBOOK__";
const WORKBOOK_LAYOUT_SHEET = "__LAYOUT__";
const WORKBOOK_CHUNK_SIZE = 20000;
const SYSTEM_SHEET_NAMES = [WORKBOOK_ARCHIVE_SHEET, WORKBOOK_LAYOUT_SHEET];
const LAST_SHEET_KEY = "last_sheet_name";

/** TOTAL 시트 열 정의 (0-based) */
const TOTAL_HEADER_ROW = 4;
const TOTAL_DATA_START_ROW = 5;
const TOTAL_COL = {
  PRODUCT: 0,
  IPSUR: 1,
  BOX_UNIT: 2,
  PALLET: 3,
  BOX: 4,
  JAN: 5,
  STOCK: 6,
  CHULGO: 7,
  SUM: 8,
  PANMAE: 9,
  DIFF: 10,
  LOCATION: 11,
};

function uint8ArrayToBase64(bytes) {
  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

function base64ToUint8Array(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

function workbookToBytes(workbook) {
  return XLSX.write(workbook, { bookType: "xlsx", type: "array", cellStyles: true });
}

/** DB 아카이브용 — 용량 축소 (cellStyles 제외) */
function workbookToArchiveBytes(workbook) {
  return XLSX.write(workbook, { bookType: "xlsx", type: "array", cellStyles: false });
}

function openWorkbookIdb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(WORKBOOK_IDB_NAME, 1);
    req.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(WORKBOOK_IDB_STORE)) {
        db.createObjectStore(WORKBOOK_IDB_STORE);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function saveWorkbookToLocalCache(workbook) {
  if (!workbook) return;
  try {
    const bytes = workbookToBytes(workbook);
    const db = await openWorkbookIdb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(WORKBOOK_IDB_STORE, "readwrite");
      tx.objectStore(WORKBOOK_IDB_STORE).put(bytes, WORKBOOK_IDB_KEY);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
    db.close();
    localStorage.setItem("workbook_cache_active", "true");
  } catch (e) {
    console.warn("로컬 워크북 캐시 저장 실패:", e);
  }
}

async function loadWorkbookBytesFromLocalCache() {
  try {
    if (localStorage.getItem("workbook_cache_active") !== "true") return null;
    const db = await openWorkbookIdb();
    const bytes = await new Promise((resolve, reject) => {
      const tx = db.transaction(WORKBOOK_IDB_STORE, "readonly");
      const req = tx.objectStore(WORKBOOK_IDB_STORE).get(WORKBOOK_IDB_KEY);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
    db.close();
    return bytes || null;
  } catch (e) {
    console.warn("로컬 워크북 캐시 로드 실패:", e);
    return null;
  }
}

async function uploadWorkbookToStorage(workbook) {
  if (!supabaseClient || !workbook) return false;
  try {
    const bytes = workbookToBytes(workbook);
    const blob = new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const { error } = await supabaseClient.storage
      .from(WORKBOOK_STORAGE_BUCKET)
      .upload(WORKBOOK_STORAGE_PATH, blob, { upsert: true, contentType: blob.type });
    if (error) throw error;
    return true;
  } catch (e) {
    console.warn("Supabase Storage 워크북 업로드 실패 (로컬 캐시는 유지됩니다):", e);
    return false;
  }
}

async function downloadWorkbookBytesFromStorage() {
  if (!supabaseClient) return null;
  try {
    const { data, error } = await supabaseClient.storage
      .from(WORKBOOK_STORAGE_BUCKET)
      .download(WORKBOOK_STORAGE_PATH);
    if (error || !data) return null;
    return new Uint8Array(await data.arrayBuffer());
  } catch (e) {
    console.warn("Supabase Storage 워크북 다운로드 실패:", e);
    return null;
  }
}

async function uploadWorkbookToSnapshotTable(workbook) {
  if (!supabaseClient || !workbook) return false;
  try {
    const bytes = workbookToArchiveBytes(workbook);
    if (!bytes || bytes.length < 100) throw new Error("엑셀 변환 결과가 비어 있습니다.");
    const base64 = uint8ArrayToBase64(bytes);
    const { error } = await supabaseClient.from("workbook_snapshot").upsert(
      {
        id: WORKBOOK_SNAPSHOT_ID,
        file_base64: base64,
        updated_at: new Date().toISOString(),
      },
      { onConflict: "id" }
    );
    if (error) throw error;
    return true;
  } catch (e) {
    console.warn("workbook_snapshot 테이블 저장 실패 (마이그레이션 003 미적용 시 정상):", e);
    return false;
  }
}

async function downloadWorkbookBytesFromSnapshotTable() {
  if (!supabaseClient) return null;
  try {
    const { data, error } = await supabaseClient
      .from("workbook_snapshot")
      .select("file_base64")
      .eq("id", WORKBOOK_SNAPSHOT_ID)
      .maybeSingle();
    if (error || !data?.file_base64) return null;
    return base64ToUint8Array(data.file_base64);
  } catch (e) {
    console.warn("workbook_snapshot 테이블 로드 실패:", e);
    return null;
  }
}

async function uploadWorkbookToDbSnapshot(workbook) {
  if (!supabaseClient || !workbook) return false;
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      const bytes = workbookToArchiveBytes(workbook);
      if (!bytes || bytes.length < 100) throw new Error("엑셀 변환 결과가 비어 있습니다.");

      const base64 = uint8ArrayToBase64(bytes);
      const chunks = [];
      for (let i = 0; i < base64.length; i += WORKBOOK_CHUNK_SIZE) {
        chunks.push(base64.slice(i, i + WORKBOOK_CHUNK_SIZE));
      }

      const { error: delErr } = await supabaseClient
        .from("rack_inventory")
        .delete()
        .eq("sheet_name", WORKBOOK_ARCHIVE_SHEET);
      if (delErr) throw delErr;

      const rows = chunks.map((chunk, index) => ({
        sheet_name: WORKBOOK_ARCHIVE_SHEET,
        rack_row: index + 1,
        rack_col: 0,
        product_name: chunk,
        pallet_qty: chunks.length,
        box_qty: base64.length,
      }));

      for (const row of rows) {
        const { error } = await supabaseClient.from("rack_inventory").insert(row);
        if (error) throw error;
      }

      const { count, error: countErr } = await supabaseClient
        .from("rack_inventory")
        .select("id", { count: "exact", head: true })
        .eq("sheet_name", WORKBOOK_ARCHIVE_SHEET);
      if (countErr || count !== rows.length) {
        throw new Error(`청크 검증 실패 (${count}/${rows.length})`);
      }
      return true;
    } catch (e) {
      console.warn(`Supabase DB 워크북 저장 실패 (시도 ${attempt}/3):`, e);
      if (attempt === 3) return false;
      await new Promise((r) => setTimeout(r, 800 * attempt));
    }
  }
  return false;
}

async function downloadWorkbookBytesFromDbSnapshot() {
  if (!supabaseClient) return null;
  try {
    const { data, error } = await supabaseClient
      .from("rack_inventory")
      .select("rack_row, product_name")
      .eq("sheet_name", WORKBOOK_ARCHIVE_SHEET)
      .order("rack_row", { ascending: true });
    if (error || !data?.length) return null;
    const base64 = data.map((row) => row.product_name).join("");
    return base64ToUint8Array(base64);
  } catch (e) {
    console.warn("Supabase DB 워크북 로드 실패:", e);
    return null;
  }
}

/** 업로드한 엑셀을 브라우저 캐시 + Supabase에 보관 (새로고침·모바일 복원용) */
async function persistUploadedWorkbook(workbook) {
  if (!workbook) return false;
  await saveWorkbookToLocalCache(workbook);
  if (!isOnline) return true;
  if (supabaseClient) {
    try {
      await saveSheetLayoutsToDb(workbook);
    } catch (e) {
      console.warn("도면 레이아웃 DB 저장 실패:", e);
    }
  }
  const results = await Promise.allSettled([
    uploadWorkbookToSnapshotTable(workbook),
    uploadWorkbookToStorage(workbook),
    uploadWorkbookToDbSnapshot(workbook),
  ]);
  return results.some((r) => r.status === "fulfilled" && r.value === true);
}

function getInventoryRacks(dbRacks) {
  return (dbRacks || []).filter((rack) => !SYSTEM_SHEET_NAMES.includes(rack.sheet_name));
}

async function fetchLayoutRowsFromDb() {
  if (!supabaseClient) return [];
  try {
    const { data, error } = await supabaseClient
      .from("rack_inventory")
      .select("rack_row, rack_col, product_name, pallet_qty, box_qty")
      .eq("sheet_name", WORKBOOK_LAYOUT_SHEET)
      .order("rack_row", { ascending: true });
    if (error) throw error;
    return data || [];
  } catch (e) {
    console.warn("도면 레이아웃 조회 실패:", e);
    return [];
  }
}

async function saveSheetLayoutsToDb(workbook) {
  if (!supabaseClient || !workbook) return false;
  const rows = getWarehouseSheetNames(workbook).map((sheetName, idx) => {
    const sheet = workbook.Sheets[sheetName];
    const { leftRacks, rightRacks } = detectRacks(sheet);
    return {
      sheet_name: WORKBOOK_LAYOUT_SHEET,
      rack_row: idx + 1,
      rack_col: leftRacks.length,
      product_name: sheetName,
      pallet_qty: rightRacks.length,
      box_qty: detectMaxRows(sheet),
    };
  });
  const { error: delErr } = await supabaseClient
    .from("rack_inventory")
    .delete()
    .eq("sheet_name", WORKBOOK_LAYOUT_SHEET);
  if (delErr) throw delErr;
  if (rows.length === 0) return true;
  for (const row of rows) {
    const { error } = await supabaseClient.from("rack_inventory").insert(row);
    if (error) throw error;
  }
  return true;
}

function getLayoutMapFromDb(layoutRows) {
  const map = {};
  (layoutRows || []).forEach((row) => {
    if (!row.product_name) return;
    map[row.product_name] = {
      leftCols: row.rack_col || 0,
      rightCols: row.pallet_qty || 0,
      maxRows: row.box_qty || 18,
    };
  });
  return map;
}

/** DB 재고·레이아웃만으로 창고 도면 워크북 생성 (엑셀 아카이브 없을 때 모바일 복원용) */
function createWarehouseLayoutSheet(sheetName, totalCols, maxRows) {
  const sheet = {};
  const safeCols = Math.max(4, totalCols || 11);
  const safeRows = Math.max(4, maxRows || 18);
  const lastColIdx = safeCols <= 9 ? 2 * safeCols : 2 * safeCols + 1;
  sheet["!ref"] = `A1:${XLSX.utils.encode_col(lastColIdx + 2)}${7 + 2 * safeRows + 2}`;
  sheet["C4"] = { t: "s", v: `<재고조사표 - ${sheetName}>` };

  for (let c = 1; c <= safeCols; c++) {
    const pColIdx = c <= 9 ? 2 * c - 1 : 2 * c + 1;
    const bColIdx = pColIdx + 1;
    sheet[XLSX.utils.encode_cell({ r: 4, c: pColIdx })] = { t: "n", v: c };
    sheet[XLSX.utils.encode_cell({ r: 5, c: pColIdx })] = { t: "s", v: "제품" };
    sheet[XLSX.utils.encode_cell({ r: 6, c: pColIdx })] = { t: "s", v: "P" };
    sheet[XLSX.utils.encode_cell({ r: 6, c: bColIdx })] = { t: "s", v: "B" };
  }
  for (let r = 1; r <= safeRows; r++) {
    sheet[XLSX.utils.encode_cell({ r: 7 + 2 * (r - 1), c: 0 })] = { t: "n", v: r };
  }
  return sheet;
}

function buildWorkbookFromDbInventory(dbRacks, layoutRows = []) {
  const inventory = getInventoryRacks(dbRacks);
  const layoutMap = getLayoutMapFromDb(layoutRows);
  const sheetNames = [...new Set(inventory.map((r) => r.sheet_name))].filter(Boolean);
  const wb = XLSX.utils.book_new();

  sheetNames.forEach((sheetName) => {
    const racks = inventory.filter((r) => r.sheet_name === sheetName);
    const layout = layoutMap[sheetName];
    const maxRow = layout?.maxRows || Math.max(18, ...racks.map((r) => r.rack_row));
    let totalCols;
    if (layout?.leftCols && layout?.rightCols) {
      totalCols = layout.leftCols + layout.rightCols;
    } else {
      totalCols = Math.max(11, ...racks.map((r) => r.rack_col));
    }
    XLSX.utils.book_append_sheet(
      wb,
      createWarehouseLayoutSheet(sheetName, totalCols, maxRow),
      sheetName
    );
  });

  return wb;
}

function getDbSheetNames(dbRacks) {
  return [...new Set(getInventoryRacks(dbRacks).map((r) => r.sheet_name))].filter(Boolean);
}

function getSheetMismatch(dbRacks, workbook) {
  if (!workbook) return [];
  const wbSheets = new Set(workbook.SheetNames);
  return getDbSheetNames(dbRacks).filter((name) => !wbSheets.has(name));
}

async function tryLoadArchivedWorkbookBytes() {
  const snapshotBytes = await downloadWorkbookBytesFromSnapshotTable();
  if (snapshotBytes) return { bytes: snapshotBytes, source: "snapshot" };

  const storageBytes = await downloadWorkbookBytesFromStorage();
  if (storageBytes) return { bytes: storageBytes, source: "storage" };

  const dbSnapshotBytes = await downloadWorkbookBytesFromDbSnapshot();
  if (dbSnapshotBytes) return { bytes: dbSnapshotBytes, source: "db" };

  const cacheBytes = await loadWorkbookBytesFromLocalCache();
  if (cacheBytes) return { bytes: cacheBytes, source: "cache" };

  return null;
}

async function loadBundledWorkbookBytes() {
  const response = await fetch(`재고조사표.xlsx?v=${Date.now()}`, { cache: "no-store" });
  if (!response.ok) throw new Error("서버에서 재고조사표 엑셀 템플릿을 읽어오는 데 실패했습니다.");
  const arrayBuffer = await response.arrayBuffer();
  return new Uint8Array(arrayBuffer);
}

/** Supabase에서 최신 데이터를 다시 불러옴 (모바일 새로고침용) */
async function reloadFromCloud() {
  if (!isOnline || !supabaseClient) {
    alert("Supabase 연동 후 사용할 수 있습니다. ⚙️ DB 설정을 확인해 주세요.");
    return;
  }
  isUserLocalWorkbook = false;
  isDbImportInProgress = false;
  await loadDataFromSupabase();
}

function rememberLastSheet(sheetName) {
  if (sheetName) localStorage.setItem(LAST_SHEET_KEY, sheetName);
}

function getRememberedSheet() {
  return localStorage.getItem(LAST_SHEET_KEY) || "";
}

async function fetchInventoryRacksFromDb() {
  const { data, error } = await supabaseClient
    .from("rack_inventory")
    .select("sheet_name, rack_row, rack_col, product_name, pallet_qty, box_qty, updated_at")
    .not("sheet_name", "in", `(${SYSTEM_SHEET_NAMES.map((n) => `"${n}"`).join(",")})`);
  if (error) throw error;
  return data || [];
}

async function testSupabaseConnection(url, key) {
  if (!window.supabase) {
    return { ok: false, error: "Supabase SDK 로드 실패" };
  }
  try {
    const client = window.supabase.createClient(url, key);
    const { error } = await client
      .from("rack_inventory")
      .select("id", { count: "exact", head: true });
    if (error) return { ok: false, error: error.message };
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

function loadOptionalConfigScript() {
  return new Promise((resolve) => {
    if (window.SUPABASE_URL && window.SUPABASE_ANON_KEY) {
      resolve();
      return;
    }
    const s = document.createElement("script");
    s.src = `config.js?v=${Date.now()}`;
    s.onload = () => resolve();
    s.onerror = () => resolve();
    document.head.appendChild(s);
  });
}

function updateSyncDiagnostics(inventoryCount, hasWorkbookArchive, templateSource, dbSheetNames = [], wbSheetNames = []) {
  const el = document.getElementById("syncDiagnostics");
  if (!el) return;
  if (!isOnline) {
    el.textContent = "DB 미연결 — ⚙️ DB 설정에서 Supabase URL·Key를 입력하세요.";
    return;
  }
  const dbLabel = dbSheetNames.length ? dbSheetNames.slice(0, 4).join(", ") : "-";
  const wbLabel = wbSheetNames.length ? wbSheetNames.slice(0, 4).join(", ") : "-";
  const parts = [
    `DB 재고 ${inventoryCount}건`,
    hasWorkbookArchive ? "업로드 엑셀 ✓" : "업로드 엑셀 ✗",
    `DB시트: ${dbLabel}`,
    `도면: ${wbLabel}`,
  ];
  el.textContent = parts.join(" · ");
  if (!hasWorkbookArchive && dbSheetNames.length > 0 && templateSource === "bundled") {
    el.textContent += " — ⚠️ PC에서 엑셀 재업로드 필요 (시트 불일치)";
    el.style.color = "#fbbf24";
  } else if (templateSource === "db-layout" || templateSource === "db-generated") {
    el.textContent += " — DB 도면 복원";
    el.style.color = "#10b981";
  } else {
    el.style.color = "var(--text-muted)";
  }
}

/**
 * 1. 이벤트 리스너 등록
 */
fileInput.addEventListener("change", handleFileSelect);
demoBtn.addEventListener("click", loadDemoWorkbook);
sheetSelect.addEventListener("change", (e) => {
  rememberLastSheet(e.target.value);
  renderActiveSheet(e.target.value);
});
searchInput.addEventListener("input", handleSearch);
// 검색창 포커스 시 매칭 목록 재노출 처리
searchInput.addEventListener("focus", handleSearch);
exportBtn.addEventListener("click", exportUpdatedExcel);

// 모달 닫기 이벤트들
closeModal.addEventListener("click", hideModal);
cancelBtn.addEventListener("click", hideModal);
window.addEventListener("click", (e) => {
  if (e.target === editModal) hideModal();
});
saveBtn.addEventListener("click", saveCellChanges);

// DB 설정 모달 리스너
dbConfigBtn.addEventListener("click", showDbModal);
closeDbModal.addEventListener("click", hideDbModal);
saveDbConfigBtn.addEventListener("click", saveDbConfig);
disconnectDbBtn.addEventListener("click", disconnectDb);
window.addEventListener("click", (e) => {
  if (e.target === dbConfigModal) hideDbModal();
});
if (autoSyncOnUploadInput) {
  autoSyncOnUploadInput.addEventListener("change", () => {
    localStorage.setItem(
      "auto_sync_on_upload",
      autoSyncOnUploadInput.checked ? "true" : "false"
    );
  });
}
if (cloudSyncBtn) {
  cloudSyncBtn.addEventListener("click", reloadFromCloud);
}

// 검색창 외부 영역 클릭 시 검색 결과 래퍼 닫기
document.addEventListener("click", (e) => {
  if (!searchInput.contains(e.target) && !searchResultsWrap.contains(e.target)) {
    searchResultsWrap.style.display = "none";
  }
});

// 엑셀 드래그 앤 드롭 업로드
document.body.addEventListener("dragover", (e) => {
  e.preventDefault();
});
document.body.addEventListener("drop", (e) => {
  e.preventDefault();
  const file = e.dataTransfer?.files?.[0];
  if (!file) return;
  if (!/\.xlsx?$/.test(file.name)) {
    alert("엑셀 파일(.xlsx, .xls)만 업로드할 수 있습니다.");
    return;
  }
  loadWorkbookFromFile(file);
});

/**
 * 합계/집계 시트 여부 판별
 */
function isSummarySheetName(sheetName) {
  const upper = String(sheetName).toUpperCase();
  return upper === "TOTAL" || sheetName.includes("합계");
}

/**
 * 시트에 창고 도면(랙 격자) 구조가 있는지 판별
 */
function hasWarehouseLayout(sheet) {
  if (!sheet || !sheet["!ref"]) return false;

  const { leftRacks, rightRacks } = detectRacks(sheet, { allowDefault: false });
  const totalRacks = leftRacks.length + rightRacks.length;
  const maxRows = detectMaxRows(sheet);

  if (totalRacks < 4 || maxRows < 4) return false;

  // '제품' / 'P' 헤더 행이 있어야 재고조사표 도면으로 인정
  let hasProductHeader = false;
  let hasPbHeader = false;

  try {
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    for (let r = 4; r <= 6; r++) {
      for (let c = range.s.c; c <= Math.min(range.e.c, 40); c++) {
        const val = sheet[XLSX.utils.encode_cell({ r, c })]?.v;
        if (val === undefined) continue;
        const text = String(val).trim();
        if (text.includes("제품")) hasProductHeader = true;
        if (text === "P") hasPbHeader = true;
      }
    }
  } catch (e) {
    return false;
  }

  return hasProductHeader && hasPbHeader && totalRacks >= 4 && maxRows >= 4;
}

/**
 * 업로드된 워크북에서 창고 도면 시트 목록만 추출
 */
function getWarehouseSheetNames(workbook) {
  if (!workbook) return [];
  return workbook.SheetNames.filter((name) => {
    if (isSummarySheetName(name)) return false;
    return hasWarehouseLayout(workbook.Sheets[name]);
  });
}

/**
 * 창고 시트의 표시용 제목 — 시트 탭 이름(셀, 루 등)을 우선 사용
 */
function parseSheetDisplayTitle(sheet, sheetName) {
  return sheetName;
}

/**
 * 시트 내부 헤더에서 구역 라벨 추출 (입구 표시용, 탭 이름과 다를 수 있음)
 */
function parseSheetHeaderLabel(sheet, sheetName) {
  if (!sheet || !sheet["!ref"]) return sheetName;

  try {
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    for (let r = range.s.r; r <= Math.min(range.e.r, 6); r++) {
      for (let c = range.s.c; c <= Math.min(range.e.c, 10); c++) {
        const cell = sheet[XLSX.utils.encode_cell({ r, c })];
        if (!cell || cell.v === undefined) continue;
        const text = String(cell.v).trim();
        const match = text.match(/<재고조사표\s*-\s*([^>]+)>/);
        if (match) return match[1].trim();
      }
    }
  } catch (e) {
    console.warn("시트 헤더 라벨 파싱 실패:", e);
  }
  return sheetName;
}

/**
 * 업로드 후 기본으로 열 창고 시트 결정
 */
function resolveDefaultSheetName(workbook, preferredName, dbSheetNames = []) {
  const warehouseSheets = getWarehouseSheetNames(workbook);
  if (warehouseSheets.length === 0) return workbook.SheetNames[0] || "";

  if (preferredName && warehouseSheets.includes(preferredName)) {
    return preferredName;
  }
  for (const name of dbSheetNames) {
    if (warehouseSheets.includes(name)) return name;
  }
  return warehouseSheets[0];
}

/**
 * 엑셀 셀 배경색(RGB) 추출 — xlsx-js-style 서식 기반
 */
function getCellFillRgb(sheet, rowIdx, colIdx) {
  const cell = sheet?.[XLSX.utils.encode_cell({ r: rowIdx, c: colIdx })];
  const rgb = cell?.s?.fill?.fgColor?.rgb || cell?.s?.fgColor?.rgb;
  return rgb ? String(rgb).replace(/^FF/i, "").toUpperCase() : "";
}

/**
 * 엑셀 배경색 → 대시보드 구역(zone) CSS 클래스
 */
function resolveZoneClass(sheet, prodRowIdx, pColIdx) {
  const rgb = getCellFillRgb(sheet, prodRowIdx, pColIdx);
  if (!rgb) return "";

  // 연파랑 계열 (B1 상단 4행 등)
  if (["9DC3E6", "BDD7EE", "DEEAF6", "B4C6E7"].includes(rgb)) {
    return "zone-blue";
  }
  // 연초록 계열
  if (["C5E0B4", "E2EFDA", "A9D08E", "C6E0B4"].includes(rgb)) {
    return "zone-green";
  }

  // 테마 기반 색상 대략적 분류
  const theme = cellThemeHint(sheet, prodRowIdx, pColIdx);
  if (theme === "blue") return "zone-blue";
  if (theme === "green") return "zone-green";
  return "";
}

function cellThemeHint(sheet, rowIdx, colIdx) {
  const cell = sheet?.[XLSX.utils.encode_cell({ r: rowIdx, c: colIdx })];
  const theme = cell?.s?.fill?.fgColor?.theme;
  const tint = cell?.s?.fill?.fgColor?.tint ?? 0;
  if (theme === 4 || (theme === 5 && tint > 0)) return "blue";
  if (theme === 9 || theme === 10) return "green";
  return "";
}

/**
 * File/Blob 객체에서 워크북을 읽어 UI에 반영
 */
function loadWorkbookFromFile(file) {
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    try {
      const workbook = XLSX.read(data, { type: "array", cellStyles: true });
      applyLoadedWorkbook(workbook, undefined, true);
    } catch (err) {
      alert("엑셀 파일을 파싱하는 데 실패했습니다. 파일이 손상되었거나 유효한 엑셀 형식이 아닙니다.");
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
}

/**
 * 파싱된 워크북을 앱 상태에 반영하고 첫 창고 시트 렌더링
 * @param {boolean} fromUserUpload - 사용자가 직접 업로드한 파일이면 Supabase 템플릿 덮어쓰기 방지
 */
function applyLoadedWorkbook(workbook, preferredSheetName, fromUserUpload = false) {
  const previousSheet = currentSheetName;

  if (fromUserUpload) {
    isUserLocalWorkbook = true;
    supabaseLoadToken += 1;
  }

  currentWorkbook = workbook;

  const warehouseSheets = getWarehouseSheetNames(workbook);
  if (warehouseSheets.length === 0) {
    alert("창고 도면 형식의 시트를 찾을 수 없습니다. 재고조사표 양식(B1, B2 등)을 확인해 주세요.");
    return;
  }

  updateSheetDropdown(warehouseSheets);

  emptyGuide.style.display = "none";
  gridBoard.style.display = "grid";
  exportBtn.style.display = "inline-block";

  const defaultSheet = resolveDefaultSheetName(
    workbook,
    preferredSheetName || previousSheet
  );
  sheetSelect.value = defaultSheet;
  renderActiveSheet(defaultSheet);
  rememberLastSheet(defaultSheet);

  if (fromUserUpload) {
    const willAutoSync = isOnline && isAutoSyncOnUploadEnabled();
    if (!willAutoSync) {
      persistUploadedWorkbook(workbook);
    }
  }

  if (isOnline && fromUserUpload && isAutoSyncOnUploadEnabled()) {
    // 렌더링 완료 후 DB에 자동 동기화 (확인창 없이 백그라운드 진행)
    setTimeout(() => {
      importWorkbookToSupabase();
    }, 0);
  }
}

/**
 * 2. 엑셀 파일 로드 및 파싱 처리
 */
function handleFileSelect(e) {
  const file = e.target.files[0];
  if (!file) return;
  loadWorkbookFromFile(file);
  e.target.value = "";
}

/**
 * 데모용 가상 엑셀 워크북을 동적으로 생성하여 로드합니다.
 */
function loadDemoWorkbook() {
  const wb = XLSX.utils.book_new();

  // B2 시트 생성 함수 호출
  const sheetB2 = createMockSheet("B2");
  XLSX.utils.book_append_sheet(wb, sheetB2, "B2");

  // B1 시트도 유사하게 임시 생성하여 추가 (다중 시트 전환 테스트 지원)
  const sheetB1 = createMockSheet("B1");
  XLSX.utils.book_append_sheet(wb, sheetB1, "B1");

  currentWorkbook = wb;

  updateSheetDropdown(getWarehouseSheetNames(wb));
  emptyGuide.style.display = "none";
  gridBoard.style.display = "grid";
  exportBtn.style.display = "inline-block";

  const defaultSheet = resolveDefaultSheetName(wb);
  sheetSelect.value = defaultSheet;
  renderActiveSheet(defaultSheet);
}

/**
 * 특정 시트 명칭에 맞는 모형(Mock) 시트 객체를 생성합니다.
 */
function createMockSheet(sheetName) {
  const sheet = {};
  
  // 전체 셀 영역 바운더리 지정
  sheet["!ref"] = "A1:AM44";

  // 3행 2열: 시트 제목
  sheet["C4"] = { t: "s", v: `<재고조사표 - ${sheetName}>` };

  // 4행: 가로 열 번호 (1~18) 기입
  // 5행: 제품 레이블
  // 6행: P / B 표시
  for (let c = 1; c <= 18; c++) {
    const pColIdx = c <= 9 ? 2 * c - 1 : 2 * c + 1;
    const bColIdx = pColIdx + 1;

    // 4행 열 번호
    const colNumAddr = XLSX.utils.encode_cell({ r: 4, c: pColIdx });
    sheet[colNumAddr] = { t: "n", v: c };

    // 5행 제품 텍스트
    const prodLabelAddr = XLSX.utils.encode_cell({ r: 5, c: pColIdx });
    sheet[prodLabelAddr] = { t: "s", v: "제품" };

    // 6행 P/B
    const pAddr = XLSX.utils.encode_cell({ r: 6, c: pColIdx });
    sheet[pAddr] = { t: "s", v: "P" };
    const bAddr = XLSX.utils.encode_cell({ r: 6, c: bColIdx });
    sheet[bAddr] = { t: "s", v: "B" };
  }

  // 행 라벨(Col 0) 1.0 ~ 18.0 기입
  for (let r = 1; r <= 18; r++) {
    const rowLabelAddr = XLSX.utils.encode_cell({ r: 7 + 2 * (r - 1), c: 0 });
    sheet[rowLabelAddr] = { t: "n", v: r };
  }

  // B2 시트 전용 테스트 데이터 채워넣기 (실제 이미지 데이터 반영)
  if (sheetName === "B2") {
    // 1행 1열: PP300, P:6
    setMockStock(sheet, 1, 1, "PP300", 6, 0);
    // 1행 2열: 이200M, P:10, B:18
    setMockStock(sheet, 1, 2, "이200M", 10, 18);
    // 2행 1열: 팔도300, P:4, B:3
    setMockStock(sheet, 2, 1, "팔도300", 4, 3);
    // 3행 1열: 새콤초500, P:10, B:3
    setMockStock(sheet, 3, 1, "새콤초500", 10, 3);
    // 3행 2열: PP300, P:3, B:9
    setMockStock(sheet, 3, 2, "PP300", 3, 9);
    // 4행 1열: YK460, P:2, B:3
    setMockStock(sheet, 4, 1, "YK460", 2, 3);
    // 4행 2열: S400, P:3, B:3
    setMockStock(sheet, 4, 2, "S400", 3, 3);
    // 5행 1열: D300, P:8
    setMockStock(sheet, 5, 1, "D300", 8, 0);
    // 6행 1열: 남양더건강한, P:6
    setMockStock(sheet, 6, 1, "남양더건강한", 6, 0);

    // 13행 1열: 초300, P:5, B:2
    setMockStock(sheet, 13, 1, "초300", 5, 2);
    // 14행 1열: D300, P:15, B:5
    setMockStock(sheet, 14, 1, "D300", 15, 5);
    // 15행 1열: 새콤초300, P:6, B:7
    setMockStock(sheet, 15, 1, "새콤초300", 6, 7);
    // 15행 2열: D300, P:3
    setMockStock(sheet, 15, 2, "D300", 3, 0);
    // 16행 1열: K500, P:20
    setMockStock(sheet, 16, 1, "K500", 20, 0);
    // 17행 1열: S400, P:17
    setMockStock(sheet, 17, 1, "S400", 17, 0);
    // 18행 1열: 매일연유, P:3
    setMockStock(sheet, 18, 1, "매일연유", 3, 0);

    // 우측 영역 테스트 데이터 (18열 등)
    // 1행 18열: 팔도300, P:18
    setMockStock(sheet, 1, 18, "팔도300", 18, 0);
    // 2행 18열: 초500, P:17
    setMockStock(sheet, 2, 18, "초500", 17, 0);
    // 3행 18열: S250, P:6
    setMockStock(sheet, 3, 18, "S250", 6, 0);
    // 4행 18열: N215, P:4, B:8
    setMockStock(sheet, 4, 18, "N215", 4, 8);
    // 5행 18열: 테코믹300, P:11, B:5
    setMockStock(sheet, 5, 18, "테코믹300", 11, 5);
    // 6행 18열: 사과듬뿍290, P:11, B:4
    setMockStock(sheet, 6, 18, "사과듬뿍290", 11, 4);
    // 7행 18열: 테코믹1000, P:12
    setMockStock(sheet, 7, 18, "테코믹1000", 12, 0);
    // 13행 18열: 원일500, P:5
    setMockStock(sheet, 13, 18, "원일500", 5, 0);
    // 14행 18열: K500, P:7, B:17
    setMockStock(sheet, 14, 18, "K500", 7, 17);
  } else if (sheetName === "B1") {
    // B1 시트용 샘플 데이터 몇 개 추가
    setMockStock(sheet, 1, 1, "샘플제품A", 15, 3);
    setMockStock(sheet, 2, 5, "샘플제품B", 8, 25);
    setMockStock(sheet, 18, 18, "샘플제품C", 50, 0);
  }

  return sheet;
}

/**
 * 랙 위치별로 가상의 재고 데이터를 셀에 기록하는 헬퍼 함수
 */
function setMockStock(sheet, row, col, product, pallet, box) {
  const pColIdx = col <= 9 ? 2 * col - 1 : 2 * col + 1;
  const bColIdx = pColIdx + 1;
  const prodRowIdx = 7 + 2 * (row - 1);
  const qtyRowIdx = 8 + 2 * (row - 1);

  const prodAddr = XLSX.utils.encode_cell({ r: prodRowIdx, c: pColIdx });
  const palletAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: pColIdx });
  const boxAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: bColIdx });

  sheet[prodAddr] = { t: "s", v: product };
  if (pallet > 0) sheet[palletAddr] = { t: "n", v: pallet };
  if (box > 0) sheet[boxAddr] = { t: "n", v: box };
}


/**
 * 시트 변경에 따른 드롭다운 UI 갱신
 */
function updateSheetDropdown(sheetNames) {
  sheetSelect.innerHTML = "";
  sheetNames.forEach((name) => {
    const option = document.createElement("option");
    option.value = name;
    option.textContent = name;
    sheetSelect.appendChild(option);
  });
}

/**
 * 랙 목록을 좌/우 영역으로 분할
 */
function splitRacksByAisle(allRacks) {
  const leftRacks = [];
  const rightRacks = [];

  if (allRacks.length === 0) {
    return { leftRacks, rightRacks };
  }

  let splitIndex = 0;
  let maxGap = 0;

  for (let i = 0; i < allRacks.length - 1; i++) {
    const gap = allRacks[i + 1].colIdx - allRacks[i].colIdx;
    if (gap > maxGap) {
      maxGap = gap;
      splitIndex = i;
    }
  }

  if (maxGap <= 2) {
    splitIndex = Math.floor(allRacks.length / 2) - 1;
  }

  for (let i = 0; i <= splitIndex; i++) {
    leftRacks.push(allRacks[i]);
  }
  for (let i = splitIndex + 1; i < allRacks.length; i++) {
    rightRacks.push(allRacks[i]);
  }

  return { leftRacks, rightRacks };
}

/**
 * 5행(열 번호)이 없을 때 6행 '제품' 헤더로 랙 열 자동 감지
 */
function detectRacksFromProductHeaders(sheet, range) {
  const productRowIdx = 5; // 엑셀 6행
  const allRacks = [];
  let lastColIdx = -10;

  for (let colIdx = range.s.c; colIdx <= range.e.c; colIdx++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: productRowIdx, c: colIdx })];
    const text = cell?.v !== undefined ? String(cell.v).trim() : "";
    if (!text.includes("제품")) continue;

    const pCell = sheet[XLSX.utils.encode_cell({ r: productRowIdx + 1, c: colIdx })];
    const pText = pCell?.v !== undefined ? String(pCell.v).trim() : "";
    if (pText !== "P") continue;

    if (colIdx - lastColIdx >= 2) {
      allRacks.push({ rackNo: 0, colIdx });
      lastColIdx = colIdx;
    }
  }

  if (allRacks.length === 0) {
    return { leftRacks: [], rightRacks: [] };
  }

  const { leftRacks, rightRacks } = splitRacksByAisle(allRacks);
  const rightCount = rightRacks.length;

  leftRacks.forEach((rack, idx) => {
    rack.rackNo = idx + 1;
  });
  rightRacks.forEach((rack, idx) => {
    rack.rackNo = rightCount > 0 ? rightCount - idx : leftRacks.length + idx + 1;
  });

  return { leftRacks, rightRacks };
}

/**
 * 데모/기본 레이아웃 (18열 대칭)
 */
function createDefaultRackLayout() {
  const leftRacks = [];
  const rightRacks = [];
  for (let c = 1; c <= 9; c++) {
    leftRacks.push({ rackNo: c, colIdx: 2 * c - 1 });
  }
  for (let i = 0; i < 9; i++) {
    const c = 10 + i;
    rightRacks.push({ rackNo: 9 - i, colIdx: 2 * c + 1 });
  }
  return { leftRacks, rightRacks };
}

/**
 * 엑셀 시트에서 좌측 영역과 우측 영역의 랙 정보(랙 번호 및 엑셀 열 인덱스)를 동적으로 스캔하여 반환합니다.
 * @param {Object} sheet - SheetJS 시트 객체
 * @returns {Object} { leftRacks, rightRacks }
 */
function detectRacks(sheet, options = {}) {
  const allowDefault = options.allowDefault !== false;

  if (!sheet || !sheet["!ref"]) {
    return allowDefault ? createDefaultRackLayout() : { leftRacks: [], rightRacks: [] };
  }

  try {
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const row4Idx = 4; // 엑셀 5행: 열(랙) 번호
    const allRacks = [];
    let lastColIdx = -10;

    for (let colIdx = range.s.c; colIdx <= range.e.c; colIdx++) {
      const cellAddress = XLSX.utils.encode_cell({ r: row4Idx, c: colIdx });
      const cell = sheet[cellAddress];
      if (cell && cell.v !== undefined) {
        const numVal = parseInt(cell.v, 10);
        if (!isNaN(numVal) && numVal > 0 && numVal <= 99) {
          if (colIdx - lastColIdx >= 2) {
            allRacks.push({ rackNo: numVal, colIdx });
            lastColIdx = colIdx;
          }
        }
      }
    }

    if (allRacks.length === 0) {
      const fromHeaders = detectRacksFromProductHeaders(sheet, range);
      if (fromHeaders.leftRacks.length + fromHeaders.rightRacks.length > 0) {
        return fromHeaders;
      }
      return allowDefault ? createDefaultRackLayout() : { leftRacks: [], rightRacks: [] };
    }

    return splitRacksByAisle(allRacks);
  } catch (e) {
    console.error("랙 배치 구조 분석 오류:", e);
    return allowDefault ? createDefaultRackLayout() : { leftRacks: [], rightRacks: [] };
  }
}

/**
 * 엑셀 시트에서 유효한 최대 열 번호(가장 큰 숫자)를 동적으로 감지합니다.
 * @param {Object} sheet - SheetJS 시트 객체
 * @returns {number} 최대 열 번호 (감지 실패 시 기본값 18)
 */
function detectMaxCols(sheet) {
  const { leftRacks, rightRacks } = detectRacks(sheet);
  return leftRacks.length + rightRacks.length;
}

/**
 * 엑셀 시트에서 유효한 최대 행 번호(가장 큰 숫자)를 A열에서 동적으로 감지합니다.
 * @param {Object} sheet - SheetJS 시트 객체
 * @returns {number} 최대 행 번호 (감지 실패 시 기본값 18)
 */
function detectMaxRows(sheet) {
  if (!sheet || !sheet["!ref"]) return 18;

  try {
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    let maxRowVal = 0;
    let consecutiveMisses = 0;

    for (let r = 1; ; r++) {
      const prodRowIdx = 7 + 2 * (r - 1);
      if (prodRowIdx > range.e.r) break;

      const cellAddress = XLSX.utils.encode_cell({ r: prodRowIdx, c: 0 });
      const cell = sheet[cellAddress];
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

    return maxRowVal || 18;
  } catch (e) {
    console.error("최대 행 수 감지 오류:", e);
    return 18;
  }
}

/**
 * 3. 선택한 시트의 레이아웃 렌더링 및 매핑 알고리즘
 */
function renderActiveSheet(sheetName) {
  currentSheetName = sheetName;
  const sheet = currentWorkbook.Sheets[sheetName];
  currentSheetData = sheet;

  // 그리드 보드 및 검색 필터 초기화
  gridBoard.innerHTML = "";
  searchInput.value = "";

  // 동적으로 랙의 좌측/우측 배치 구조와 순서 배열 파악
  const { leftRacks, rightRacks } = detectRacks(sheet);

  const leftCols = leftRacks.length;
  const rightCols = rightRacks.length;
  const maxRows = detectMaxRows(sheet);
  const displayTitle = parseSheetDisplayTitle(sheet, sheetName);
  const totalCols = leftCols + rightCols;

  // CSS Grid 컬럼·행 — 업로드된 시트의 실제 랙/행 수에 맞춰 동적 구성
  gridBoard.style.gridTemplateColumns = `40px repeat(${leftCols}, minmax(52px, 1fr)) 50px repeat(${rightCols}, minmax(52px, 1fr)) 40px`;
  if (!isMobileViewport()) {
    gridBoard.style.maxWidth = totalCols <= 14 ? "1100px" : totalCols <= 16 ? "1250px" : "1400px";
  }
  applyResponsiveGridLayout(totalCols);
  gridBoard.dataset.sheetName = sheetName;
  gridBoard.dataset.layout = `${leftCols}+${rightCols}x${maxRows}`;

  layoutInfoEl.textContent = `${displayTitle} · 좌${leftCols} / 우${rightCols}열 × ${maxRows}행`;

  // 상단 축 랙 번호 인덱스 라벨 렌더링
  renderGridHeaders(leftRacks, rightRacks);

  // 1부터 시작하는 물리적 순번 기반의 열 매핑 맵 생성
  const colMapping = detectColumnMapping(sheet);

  let totalProducts = 0;
  let totalPallets = 0;
  let totalBoxes = 0;

  // 1행부터 maxRows행까지 루프 돌며 물리적 랙 셀들을 순서대로 배치
  for (let r = 1; r <= maxRows; r++) {
    // 0-based 엑셀 행 인덱스 계산
    const prodRowIdx = 7 + 2 * (r - 1);
    const qtyRowIdx = 8 + 2 * (r - 1);

    // 왼쪽 행 번호 라벨 생성
    const leftRowLabel = document.createElement("div");
    leftRowLabel.className = "row-label";
    leftRowLabel.textContent = r;
    gridBoard.appendChild(leftRowLabel);

    // 좌측 영역 랙들 렌더링 (물리적 순번 1 ~ leftCols)
    leftRacks.forEach((rackInfo, idx) => {
      const colSeq = idx + 1; // 1-based 물리적 순번
      const cellData = getRackCellData(sheet, colMapping, r, colSeq, prodRowIdx, qtyRowIdx);
      const zoneClass = resolveZoneClass(sheet, prodRowIdx, rackInfo.colIdx);
      const cellEl = createRackCellElement(r, colSeq, cellData, rackInfo.rackNo, zoneClass);
      gridBoard.appendChild(cellEl);

      if (cellData.product) {
        totalProducts++;
        totalPallets += cellData.pallet;
        totalBoxes += cellData.box;
      }
    });

    // 중앙 통로 셀 생성 (위치: leftCols + 2번째 컬럼)
    const aisle = document.createElement("div");
    const aisleColIdx = leftCols + 2;
    aisle.style.gridColumn = `${aisleColIdx} / ${aisleColIdx + 1}`;

    if (r === maxRows) {
      aisle.className = "aisle-cell entrance-gate";
      const arrow = document.createElement("div");
      arrow.className = "entrance-arrow";
      arrow.innerHTML = `
        <svg viewBox="0 0 24 24">
          <path d="M4 12l1.41 1.41L11 7.83V20h2V7.83l5.58 5.59L20 12l-8-8-8 8z"/>
        </svg>
        <span>${displayTitle}</span>
      `;
      aisle.appendChild(arrow);
    } else {
      aisle.className = "aisle-cell entrance-path";
    }
    gridBoard.appendChild(aisle);

    // 우측 영역 랙들 렌더링 (물리적 순번 leftCols + 1 ~ leftCols + rightCols)
    rightRacks.forEach((rackInfo, idx) => {
      const colSeq = leftCols + idx + 1; // 1-based 물리적 순번
      const cellData = getRackCellData(sheet, colMapping, r, colSeq, prodRowIdx, qtyRowIdx);
      const zoneClass = resolveZoneClass(sheet, prodRowIdx, rackInfo.colIdx);
      const cellEl = createRackCellElement(r, colSeq, cellData, rackInfo.rackNo, zoneClass);
      gridBoard.appendChild(cellEl);

      if (cellData.product) {
        totalProducts++;
        totalPallets += cellData.pallet;
        totalBoxes += cellData.box;
      }
    });

    // 오른쪽 행 번호 라벨 생성
    const rightRowLabel = document.createElement("div");
    rightRowLabel.className = "row-label";
    rightRowLabel.textContent = r;
    gridBoard.appendChild(rightRowLabel);
  }

  // 하단 축 랙 번호 인덱스 라벨 렌더링
  renderGridFooters(leftRacks, rightRacks);

  // 통계 요약판 수치 반영
  totalProductsEl.textContent = totalProducts;
  totalPalletsEl.textContent = totalPallets;
  totalBoxesEl.textContent = totalBoxes;

  if (mobileScrollHint) {
    mobileScrollHint.style.display = isMobileViewport() ? "block" : "none";
  }
}

/**
 * 엑셀 시트에서 물리적 랙의 열 매핑 객체를 반환합니다.
 * 반환 형식: { 1: colIdx, 2: colIdx, ..., maxCols: colIdx }
 */
function detectColumnMapping(sheet) {
  const { leftRacks, rightRacks } = detectRacks(sheet);
  const allRacks = [...leftRacks, ...rightRacks];
  const mapping = {};

  allRacks.forEach((rackInfo, index) => {
    mapping[index + 1] = rackInfo.colIdx;
  });

  return mapping;
}

function createDefaultMapping(maxCols) {
  const mapping = {};
  const currentMaxCols = maxCols || 18;
  const leftCols = Math.ceil(currentMaxCols / 2);
  for (let c = 1; c <= currentMaxCols; c++) {
    mapping[c] = c <= leftCols ? 2 * c - 1 : 2 * c + 1;
  }
  return mapping;
}

/**
 * 특정 행/열 좌표의 재고 정보(제품명, P수량, B수량)를 시트 객체에서 파싱합니다.
 */
function getRackCellData(sheet, colMapping, row, col, prodRowIdx, qtyRowIdx) {
  const data = {
    product: "",
    pallet: 0,
    box: 0,
    prodCellAddress: "",
    palletCellAddress: "",
    boxCellAddress: ""
  };

  if (!sheet) return data;

  const pColIdx = colMapping[col];
  if (pColIdx === undefined) return data;

  const bColIdx = pColIdx + 1;

  data.prodCellAddress = XLSX.utils.encode_cell({ r: prodRowIdx, c: pColIdx });
  data.palletCellAddress = XLSX.utils.encode_cell({ r: qtyRowIdx, c: pColIdx });
  data.boxCellAddress = XLSX.utils.encode_cell({ r: qtyRowIdx, c: bColIdx });

  const prodCell = sheet[data.prodCellAddress];
  const palletCell = sheet[data.palletCellAddress];
  const boxCell = sheet[data.boxCellAddress];

  if (prodCell && prodCell.v !== undefined) {
    data.product = String(prodCell.v).trim();
  }
  if (palletCell && palletCell.v !== undefined) {
    data.pallet = Number(palletCell.v) || 0;
  }
  if (boxCell && boxCell.v !== undefined) {
    data.box = Number(boxCell.v) || 0;
  }

  return data;
}

/**
 * 랙 셀 DOM 엘리먼트 생성
 */
function createRackCellElement(row, col, cellData, rackNo, zoneClass = "") {
  const cell = document.createElement("div");
  const classes = ["rack-cell"];
  if (cellData.product || cellData.pallet > 0 || cellData.box > 0) {
    classes.push("active-stock");
  }
  if (zoneClass) classes.push(zoneClass);
  cell.className = classes.join(" ");

  // 검색 시 필터링을 쉽게 하기 위해 제품명을 data 속성으로 주입
  cell.dataset.product = cellData.product.toLowerCase();
  cell.dataset.row = row;
  cell.dataset.col = col;
  cell.dataset.rackNo = rackNo; // 표시용 실제 랙 번호

  // 셀 내 제품명 라벨 구성
  const nameEl = document.createElement("div");
  nameEl.className = "cell-product-name";
  nameEl.textContent = cellData.product || "";
  cell.appendChild(nameEl);

  // 셀 내 수량 정보 구성 (P, B)
  const qtyContainer = document.createElement("div");
  qtyContainer.className = "cell-qty-container";

  // 파레트(P) 영역
  const pEl = document.createElement("div");
  pEl.className = "qty-p";
  pEl.innerHTML = `<span>P</span><span class="qty-val">${cellData.pallet || ""}</span>`;
  qtyContainer.appendChild(pEl);

  // 박스(B) 영역
  const bEl = document.createElement("div");
  bEl.className = "qty-b";
  bEl.innerHTML = `<span>B</span><span class="qty-val">${cellData.box || ""}</span>`;
  qtyContainer.appendChild(bEl);

  cell.appendChild(qtyContainer);

  bindRackCellTap(cell, () => {
    showEditModal(row, col, cellData, rackNo);
  });

  return cell;
}

/**
 * 격자 상단 인덱스 라벨들 생성 함수
 */
function renderGridHeaders(leftRacks, rightRacks) {
  // 첫 칸 빈자리 (행 라벨 공간)
  const emptyLeft = document.createElement("div");
  emptyLeft.className = "col-label empty-header";
  gridBoard.appendChild(emptyLeft);

  // 좌측 랙 영역 상단 라벨 (실제 랙 번호 시퀀스)
  leftRacks.forEach((rackInfo) => {
    const label = document.createElement("div");
    label.className = "col-label";
    label.textContent = rackInfo.rackNo;
    gridBoard.appendChild(label);
  });

  const emptyCenter = document.createElement("div");
  emptyCenter.className = "col-label empty-header";
  gridBoard.appendChild(emptyCenter);

  // 우측 랙 영역 상단 라벨 (실제 랙 번호 시퀀스)
  rightRacks.forEach((rackInfo) => {
    const label = document.createElement("div");
    label.className = "col-label";
    label.textContent = rackInfo.rackNo;
    gridBoard.appendChild(label);
  });

  const emptyRight = document.createElement("div");
  emptyRight.className = "col-label empty-header";
  gridBoard.appendChild(emptyRight);
}

/**
 * 격자 하단 대칭형 인덱스 라벨들 생성 함수
 */
function renderGridFooters(leftRacks, rightRacks) {
  const emptyLeft = document.createElement("div");
  emptyLeft.className = "col-label empty-footer";
  gridBoard.appendChild(emptyLeft);

  // 좌측 하단 라벨 (상단과 동일한 실제 랙 번호 시퀀스)
  leftRacks.forEach((rackInfo) => {
    const label = document.createElement("div");
    label.className = "col-label";
    label.textContent = rackInfo.rackNo;
    gridBoard.appendChild(label);
  });

  const emptyCenter = document.createElement("div");
  emptyCenter.className = "col-label empty-footer";
  gridBoard.appendChild(emptyCenter);

  // 우측 하단 라벨 (상단과 동일한 실제 랙 번호 시퀀스)
  rightRacks.forEach((rackInfo) => {
    const label = document.createElement("div");
    label.className = "col-label";
    label.textContent = rackInfo.rackNo;
    gridBoard.appendChild(label);
  });

  const emptyRight = document.createElement("div");
  emptyRight.className = "col-label empty-footer";
  gridBoard.appendChild(emptyRight);
}

/**
 * 4. 재고 검색 및 위치 네비게이션 기능 (전 구역 통합 검색 지원)
 */
function handleSearch(e) {
  const query = searchInput.value.trim().toLowerCase();
  const cells = gridBoard.querySelectorAll(".rack-cell");
  const matches = [];

  // 1. 현재 화면에 표시된 시트의 실시간 네온 깜빡임 피드백 처리
  cells.forEach((cell) => {
    cell.classList.remove("search-highlight");
    if (query !== "") {
      const prodName = cell.dataset.product;
      if (prodName && prodName.includes(query)) {
        cell.classList.add("search-highlight");
      }
    }
  });

  // 2. 워크북이 로드되지 않았거나 검색어가 없으면 결과 영역 숨김 처리
  if (!currentWorkbook || query === "") {
    searchResultsWrap.style.display = "none";
    searchResults.innerHTML = "";
    return;
  }

  // 3. 전체 창고 시트(구역)를 순회하며 일치하는 재고 데이터 및 수량(P, B) 수집
  const warehouseSheets = getWarehouseSheetNames(currentWorkbook);
  warehouseSheets.forEach((sheetName) => {
    const sheet = currentWorkbook.Sheets[sheetName];
    if (!sheet) return;

    // 해당 시트의 실제 최대 행과 열 동적 감지
    const maxCols = detectMaxCols(sheet);
    const maxRows = detectMaxRows(sheet);

    // 해당 시트의 열 좌표 매핑 정보 검출 (동적 maxCols 전달)
    const colMapping = detectColumnMapping(sheet, maxCols);

    // 1행부터 maxRows행까지 순회
    for (let r = 1; r <= maxRows; r++) {
      const prodRowIdx = 7 + 2 * (r - 1);
      const qtyRowIdx = 8 + 2 * (r - 1);

      // 1열부터 maxCols열까지 순회
      for (let c = 1; c <= maxCols; c++) {
        const pColIdx = colMapping[c];
        if (pColIdx === undefined) continue;
        const bColIdx = pColIdx + 1;

        const prodAddr = XLSX.utils.encode_cell({ r: prodRowIdx, c: pColIdx });
        const palletAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: pColIdx });
        const boxAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: bColIdx });

        const cellObj = sheet[prodAddr];
        const palletObj = sheet[palletAddr];
        const boxObj = sheet[boxAddr];

        if (cellObj && cellObj.v !== undefined) {
          const prodNameVal = String(cellObj.v).trim();
          if (prodNameVal.toLowerCase().includes(query)) {
            // 수량 추출 (없거나 형식이 다르면 0으로 방어)
            const pQty = (palletObj && palletObj.v !== undefined) ? (Number(palletObj.v) || 0) : 0;
            const bQty = (boxObj && boxObj.v !== undefined) ? (Number(boxObj.v) || 0) : 0;

            matches.push({
              sheetName: sheetName,
              product: prodNameVal,
              row: r,
              col: c,
              pallet: pQty,
              box: bQty
            });
          }
        }
      }
    }
  });

  // 4. 실시간 검색 결과 패널 동적 렌더링
  searchResults.innerHTML = "";

  if (matches.length > 0) {
    searchResultsWrap.style.display = "block"; // 래퍼를 보여준다 = 창고 도면이 자연스럽게 아래로 밀림

    // 4.1. 전체 매칭 항목에 대해 총 파렛트(P), 총 박스(B) 실시간 요약 연산
    let grandTotalPallets = 0;
    let grandTotalBoxes = 0;
    matches.forEach((m) => {
      grandTotalPallets += m.pallet;
      grandTotalBoxes += m.box;
    });

    // 4.2. 검색 결과 리스트 상단에 고정 요약 헤더바 삽입
    const summaryHeader = document.createElement("div");
    summaryHeader.className = "search-summary-header";
    summaryHeader.innerHTML = `
      <span class="summary-title">🔍 '${searchInput.value}' 검색 요약</span>
      <div class="summary-values">
        <span>매칭: <strong>${matches.length}</strong>곳</span>
        <span class="summary-val-p">파렛트(P): <strong>${grandTotalPallets}</strong></span>
        <span class="summary-val-b">박스(B): <strong>${grandTotalBoxes}</strong></span>
      </div>
    `;
    searchResults.appendChild(summaryHeader);
    
    // 4.3. 개별 매칭 결과 카드를 목록 창에 삽입
    matches.forEach((match) => {
      const item = document.createElement("div");
      item.className = "search-result-item";
      // 리스트에 구역 정보와 함께 우측에 수량 배지(P, B)를 표시
      item.innerHTML = `
        <div style="display: flex; flex-direction: column; gap: 0.2rem;">
          <span class="match-name">${match.product}</span>
          <span class="match-coord">[${match.sheetName} 구역] ${match.row}행 - ${match.col}열</span>
        </div>
        <div class="match-qty-badge">
          <span class="match-qty-p">${match.pallet}P</span>
          <span class="match-qty-b">${match.box}B</span>
        </div>
      `;
      
      // 검색 결과 클릭 시 ➔ 구역(시트) 전환 + 렌더링 + 화면 스크롤 포커스 이동
      item.addEventListener("click", () => {
        // 클릭한 항목의 구역이 현재 활성화된 구역과 다르면 시트 강제 교체 및 화면 재구축
        if (currentSheetName !== match.sheetName) {
          sheetSelect.value = match.sheetName;
          renderActiveSheet(match.sheetName);
        }

        // 해당 좌표의 셀 요소를 돔에서 추적
        const targetCell = gridBoard.querySelector(`.rack-cell[data-row="${match.row}"][data-col="${match.col}"]`);
        
        if (targetCell) {
          // 화면 중앙으로 스무스하게 스크롤 포커스
          targetCell.scrollIntoView({
            behavior: "smooth",
            block: "center",
            inline: "center"
          });

          // 시트 교체 직후 해당 셀을 2.5초간 강력하게 깜빡이게 한 뒤 원복
          targetCell.classList.add("search-highlight");
          setTimeout(() => {
            targetCell.classList.remove("search-highlight");
          }, 2500);
        }

        // 결과 목록 래퍼 닫기
        searchResultsWrap.style.display = "none";
      });
      
      searchResults.appendChild(item);
    });
  } else {
    searchResultsWrap.style.display = "none";
  }
}

/**
 * 5. 셀 편집 모달 동작
 */
function showEditModal(row, col, cellData, rackNo) {
  selectedCellInfo = {
    row: row,
    col: col,
    data: cellData
  };

  // 타이틀 표시 시 물리적 순번(col) 대신 실제 랙 번호(rackNo)를 우선 사용합니다.
  modalCellTitle.textContent = `${row}행 - ${rackNo || col}열 재고 수정`;
  inputProduct.value = cellData.product || "";
  inputPallet.value = cellData.pallet || 0;
  inputBox.value = cellData.box || 0;

  editModal.classList.add("show");
  document.body.classList.add("modal-open");
  if (isMobileViewport()) {
    requestAnimationFrame(() => {
      editModal.querySelector(".modal-container")?.scrollIntoView({ block: "center", behavior: "smooth" });
    });
  } else {
    inputProduct.focus();
  }
}

function hideModal() {
  editModal.classList.remove("show");
  document.body.classList.remove("modal-open");
  selectedCellInfo = null;
}

/**
 * 사용자가 수정한 재고를 엑셀 데이터 구조에 반영하고 화면을 새로고침합니다.
 */
async function saveCellChanges() {
  if (!selectedCellInfo || !currentSheetData) return;

  const newProduct = inputProduct.value.trim();
  const newPallet = parseInt(inputPallet.value, 10) || 0;
  const newBox = parseInt(inputBox.value, 10) || 0;

  const sheet = currentSheetData;
  const cellInfo = selectedCellInfo.data;

  // 1. 엑셀 워크시트 객체에 새 데이터 써넣기
  updateExcelCellValue(sheet, cellInfo.prodCellAddress, newProduct, 's');
  updateExcelCellValue(sheet, cellInfo.palletCellAddress, newPallet, 'n');
  updateExcelCellValue(sheet, cellInfo.boxCellAddress, newBox, 'n');

  // 2. 온라인 모드면 Supabase에 비동기 저장
  if (isOnline) {
    try {
      if (newProduct === "" && newPallet === 0 && newBox === 0) {
        // 비어있는 데이터는 DB에서 삭제 처리
        await supabaseClient
          .from('rack_inventory')
          .delete()
          .match({
            sheet_name: currentSheetName,
            rack_row: selectedCellInfo.row,
            rack_col: selectedCellInfo.col
          });
      } else {
        // 데이터가 존재하면 Upsert 처리
        const { error } = await supabaseClient
          .from('rack_inventory')
          .upsert({
            sheet_name: currentSheetName,
            rack_row: selectedCellInfo.row,
            rack_col: selectedCellInfo.col,
            product_name: newProduct,
            pallet_qty: newPallet,
            box_qty: newBox,
            updated_at: new Date().toISOString()
          }, { onConflict: 'sheet_name,rack_row,rack_col' });
          
        if (error) throw error;
      }
    } catch (e) {
      alert("서버 데이터 동기화에 실패했습니다. (오프라인 변경만 적용됨)");
      console.error(e);
    }
  }

  // 3. 화면 갱신
  aggregateTotalSheet(currentWorkbook);
  renderActiveSheet(currentSheetName);

  // 4. 모달 닫기
  hideModal();
}

/**
 * SheetJS 시트 내 특정 주소의 셀 값을 안전하게 업데이트합니다.
 */
function updateExcelCellValue(sheet, address, val, type) {
  // 수량이 0이거나 제품명이 비어 있는 경우 엑셀 상에서 셀을 지우거나 공백으로 만듭니다.
  if (val === "" || val === 0) {
    if (sheet[address]) {
      // 기존 셀이 존재하면 빈값으로 밀어버림
      sheet[address].v = "";
      sheet[address].t = "s";
      delete sheet[address].w; // 포맷 문자열 제거
    }
    return;
  }

  if (!sheet[address]) {
    // 셀이 생성되어 있지 않았다면 새롭게 추가
    sheet[address] = { t: type, v: val };
  } else {
    // 이미 존재하는 셀이라면 값 및 타입 업데이트
    sheet[address].t = type;
    sheet[address].v = val;
    delete sheet[address].w; // 포맷 문자열 제거
  }
}

/**
 * 6. 수정된 파일 export (다운로드)
 */
function exportUpdatedExcel() {
  if (!currentWorkbook) return;

  try {
    // 다운로드 직전 TOTAL 시트에 전체 구역 파렛트(P)/박스(B) 합산 집계
    aggregateTotalSheet(currentWorkbook);

    // 엑셀 저장 직전에 가독성 디자인(색상, 테두리, 너비 등) 적용
    applyExcelStyles(currentWorkbook);

    // 워크북을 바이너리 배열 형태로 저장 (xlsx-js-style 지원을 위해 cellStyles 활성화)
    const wbout = XLSX.write(currentWorkbook, { bookType: "xlsx", type: "array", cellStyles: true });
    const blob = new Blob([wbout], { type: "application/octet-stream" });
    
    // 다운로드 링크 동적 생성
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    // 기존 파일명에 '_수정본'을 붙이거나 기본 명칭 사용
    a.download = `재고조사표_수정본_${new Date().toISOString().slice(0, 10)}.xlsx`;
    document.body.appendChild(a);
    a.click();
    
    // 리소스 해제
    setTimeout(() => {
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    }, 0);
  } catch (err) {
    alert("엑셀 파일을 생성하고 다운로드하는 도중 오류가 발생했습니다.");
    console.error(err);
  }
}

/**
 * TOTAL 시트에 전체 구역(B1, B2 등) 재고 합산 집계를 기입하는 함수
 *  - 창고 시트에 있는 제품 목록을 기준으로 TOTAL 행을 재구성 (이름 변경·삭제 반영)
 *  - 제품별 P/B 합계 및 창고위치(L열) 기입
 */
function findTotalSheet(workbook) {
  const name = workbook.SheetNames.find(
    (n) => n.toUpperCase() === "TOTAL" || n.includes("합계")
  );
  return name ? workbook.Sheets[name] : null;
}

function totalCellAddr(excelRow, col) {
  return XLSX.utils.encode_cell({ r: excelRow - 1, c: col });
}

function writeTotalCell(totalSheet, excelRow, col, val, type = "n") {
  const addr = totalCellAddr(excelRow, col);
  if (val === undefined || val === null || val === "") {
    delete totalSheet[addr];
    return;
  }
  totalSheet[addr] = { t: type, v: val };
  delete totalSheet[addr].w;
}

function collectWarehouseProductData(workbook) {
  const aggregated = new Map();

  workbook.SheetNames.forEach((sheetName) => {
    if (isSummarySheetName(sheetName)) return;
    const sheet = workbook.Sheets[sheetName];
    if (!sheet || !hasWarehouseLayout(sheet)) return;

    const maxCols = detectMaxCols(sheet);
    const maxRows = detectMaxRows(sheet);
    const colMapping = detectColumnMapping(sheet);

    for (let r = 1; r <= maxRows; r++) {
      const prodRowIdx = 7 + 2 * (r - 1);
      const qtyRowIdx = 8 + 2 * (r - 1);

      for (let c = 1; c <= maxCols; c++) {
        const pColIdx = colMapping[c];
        if (pColIdx === undefined) continue;
        const bColIdx = pColIdx + 1;

        const prodCell = sheet[XLSX.utils.encode_cell({ r: prodRowIdx, c: pColIdx })];
        if (!prodCell || prodCell.v === undefined) continue;
        const prodName = String(prodCell.v).trim();
        if (!prodName) continue;

        const palletCell = sheet[XLSX.utils.encode_cell({ r: qtyRowIdx, c: pColIdx })];
        const boxCell = sheet[XLSX.utils.encode_cell({ r: qtyRowIdx, c: bColIdx })];
        const pQty = palletCell?.v !== undefined ? Number(palletCell.v) || 0 : 0;
        const bQty = boxCell?.v !== undefined ? Number(boxCell.v) || 0 : 0;

        if (!aggregated.has(prodName)) {
          aggregated.set(prodName, { pallet: 0, box: 0, locations: new Set() });
        }
        const entry = aggregated.get(prodName);
        entry.pallet += pQty;
        entry.box += bQty;
        entry.locations.add(sheetName);
      }
    }
  });

  return aggregated;
}

function readTotalSheetRowMetadata(totalSheet) {
  const meta = new Map();
  for (let excelRow = TOTAL_DATA_START_ROW; excelRow <= 300; excelRow++) {
    const aCell = totalSheet[totalCellAddr(excelRow, TOTAL_COL.PRODUCT)];
    if (!aCell || aCell.v === undefined || String(aCell.v).trim() === "") break;
    const prodName = String(aCell.v).trim();
    const readVal = (col) => {
      const cell = totalSheet[totalCellAddr(excelRow, col)];
      return cell?.v !== undefined ? cell.v : undefined;
    };
    meta.set(prodName, {
      ipsuryang: readVal(TOTAL_COL.IPSUR),
      box_danyi: readVal(TOTAL_COL.BOX_UNIT),
      janryang: readVal(TOTAL_COL.JAN),
      chulgo: readVal(TOTAL_COL.CHULGO),
      panmae_ilbo: readVal(TOTAL_COL.PANMAE),
    });
  }
  return meta;
}

function getTotalSheetLastDataRow(totalSheet) {
  let lastRow = TOTAL_DATA_START_ROW - 1;
  for (let excelRow = TOTAL_DATA_START_ROW; excelRow <= 300; excelRow++) {
    const aCell = totalSheet[totalCellAddr(excelRow, TOTAL_COL.PRODUCT)];
    if (!aCell || aCell.v === undefined || String(aCell.v).trim() === "") break;
    lastRow = excelRow;
  }
  return lastRow;
}

function clearTotalSheetDataRows(totalSheet, fromRow, toRow) {
  for (let excelRow = fromRow; excelRow <= toRow; excelRow++) {
    for (let col = TOTAL_COL.PRODUCT; col <= TOTAL_COL.LOCATION; col++) {
      delete totalSheet[totalCellAddr(excelRow, col)];
    }
  }
}

function ensureTotalLocationHeader(totalSheet) {
  writeTotalCell(totalSheet, TOTAL_HEADER_ROW, TOTAL_COL.LOCATION, "창고위치", "s");
}

function aggregateTotalSheet(workbook) {
  const totalSheet = findTotalSheet(workbook);
  if (!totalSheet) return;

  ensureTotalLocationHeader(totalSheet);

  const warehouseProducts = collectWarehouseProductData(workbook);
  const preservedMeta = readTotalSheetRowMetadata(totalSheet);
  const oldLastRow = getTotalSheetLastDataRow(totalSheet);

  const sortedNames = [...warehouseProducts.keys()].sort((a, b) =>
    a.localeCompare(b, "ko")
  );

  const clearUntil = Math.max(
    oldLastRow,
    sortedNames.length > 0
      ? TOTAL_DATA_START_ROW + sortedNames.length - 1
      : TOTAL_DATA_START_ROW - 1
  );
  if (clearUntil >= TOTAL_DATA_START_ROW) {
    clearTotalSheetDataRows(totalSheet, TOTAL_DATA_START_ROW, clearUntil);
  }

  sortedNames.forEach((prodName, idx) => {
    const excelRow = TOTAL_DATA_START_ROW + idx;
    const totals = warehouseProducts.get(prodName);
    const meta = preservedMeta.get(prodName) || {};
    const locationText = [...totals.locations]
      .sort((a, b) => a.localeCompare(b, "ko"))
      .join(", ");

    writeTotalCell(totalSheet, excelRow, TOTAL_COL.PRODUCT, prodName, "s");
    writeTotalCell(totalSheet, excelRow, TOTAL_COL.IPSUR, meta.ipsuryang);
    writeTotalCell(totalSheet, excelRow, TOTAL_COL.BOX_UNIT, meta.box_danyi);
    writeTotalCell(totalSheet, excelRow, TOTAL_COL.PALLET, totals.pallet);
    writeTotalCell(totalSheet, excelRow, TOTAL_COL.BOX, totals.box);
    writeTotalCell(totalSheet, excelRow, TOTAL_COL.JAN, meta.janryang);
    writeTotalCell(totalSheet, excelRow, TOTAL_COL.CHULGO, meta.chulgo);
    writeTotalCell(totalSheet, excelRow, TOTAL_COL.PANMAE, meta.panmae_ilbo);
    writeTotalCell(totalSheet, excelRow, TOTAL_COL.LOCATION, locationText, "s");

    totalSheet[totalCellAddr(excelRow, TOTAL_COL.STOCK)] = {
      t: "n",
      f: `((D${excelRow}*C${excelRow})+E${excelRow})*B${excelRow}+F${excelRow}`,
    };
    totalSheet[totalCellAddr(excelRow, TOTAL_COL.SUM)] = {
      t: "n",
      f: `G${excelRow}+H${excelRow}`,
    };
    totalSheet[totalCellAddr(excelRow, TOTAL_COL.DIFF)] = {
      t: "n",
      f: `I${excelRow}-J${excelRow}`,
    };
  });

  if (totalSheet["!ref"]) {
    try {
      const range = XLSX.utils.decode_range(totalSheet["!ref"]);
      const newLastRow =
        sortedNames.length > 0
          ? TOTAL_DATA_START_ROW + sortedNames.length - 1
          : TOTAL_DATA_START_ROW - 1;
      range.e.r = Math.max(range.e.r, newLastRow - 1);
      range.e.c = Math.max(range.e.c, TOTAL_COL.LOCATION);
      totalSheet["!ref"] = XLSX.utils.encode_range(range);
    } catch (e) {
      console.error("TOTAL 시트의 범위(!ref)를 갱신하는 데 실패했습니다.", e);
    }
  }

  console.log("✅ TOTAL 시트 집계 완료:", sortedNames.length, "품목");
}


/**
 * xlsx-js-style 라이브러리를 활용하여 다운로드 엑셀 파일에 격자 및 색상 서식을 입힙니다.
 */
/**
 * xlsx-js-style 라이브러리를 활용하여 다운로드 엑셀 파일에 격자 및 색상 서식을 입힙니다.
 */
function applyExcelStyles(workbook) {
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;

    // TOTAL 시트는 전용 함수로 처리
    if (sheetName.toUpperCase() === "TOTAL" || sheetName.includes("합계")) {
      applyTotalSheetStyles(sheet);
      continue;
    }

    const maxCols = detectMaxCols(sheet);
    const maxRows = detectMaxRows(sheet);
    const colMapping = detectColumnMapping(sheet, maxCols);

    const lastColIdx = colMapping[maxCols] || (2 * maxCols + 1);
    const rightLabelCol = lastColIdx + 1; // 랙들의 우측 끝 바로 다음 열을 행 번호 라벨용으로 지정

    // 1. 열 너비 지정 (각 시트의 폭을 최적화)
    const colsWidths = [];
    colsWidths[0] = { wch: 6 }; // 좌측 행 번호 라벨 열 너비
    
    // colMapping에 따라 각 랙 열 너비 할당 (P: 14, B: 6)
    for (let c = 1; c <= maxCols; c++) {
      const pCol = colMapping[c];
      if (pCol !== undefined) {
        colsWidths[pCol] = { wch: 14 }; // 제품명 및 파렛트(P) 열
        colsWidths[pCol + 1] = { wch: 6 };  // 박스(B) 열
      }
    }
    
    // 통로 및 기타 미지정 열 너비 기본값 채우기
    for (let i = 0; i <= rightLabelCol; i++) {
      if (!colsWidths[i]) {
        colsWidths[i] = { wch: 4 }; // 콤팩트 통로
      }
    }
    colsWidths[rightLabelCol] = { wch: 6 }; // 우측 행 번호 라벨 열 너비
    sheet["!cols"] = colsWidths;

    // 2. 셀 스타일 정의
    // 얇은 실선 테두리 스타일
    const borderGray = {
      top: { style: "thin", color: { rgb: "D1D5DB" } },
      bottom: { style: "thin", color: { rgb: "D1D5DB" } },
      left: { style: "thin", color: { rgb: "D1D5DB" } },
      right: { style: "thin", color: { rgb: "D1D5DB" } }
    };
    
    const borderMint = {
      top: { style: "thin", color: { rgb: "A5D6A7" } },
      bottom: { style: "thin", color: { rgb: "A5D6A7" } },
      left: { style: "thin", color: { rgb: "A5D6A7" } },
      right: { style: "thin", color: { rgb: "A5D6A7" } }
    };

    const borderBlue = {
      top: { style: "thin", color: { rgb: "90CAF9" } },
      bottom: { style: "thin", color: { rgb: "90CAF9" } },
      left: { style: "thin", color: { rgb: "90CAF9" } },
      right: { style: "thin", color: { rgb: "90CAF9" } }
    };

    const borderPink = {
      top: { style: "thin", color: { rgb: "FFCDD2" } },
      bottom: { style: "thin", color: { rgb: "FFCDD2" } },
      left: { style: "thin", color: { rgb: "FFCDD2" } },
      right: { style: "thin", color: { rgb: "FFCDD2" } }
    };

    const fontName = "맑은 고딕";

    // 스타일 프리셋 (xlsx-js-style 배경색 채우기를 위해 patternType: "solid" 적용)
    const sHeader = {
      font: { name: fontName, size: 9, bold: true, color: { rgb: "1F2937" } },
      fill: { patternType: "solid", fgColor: { rgb: "E2E8F0" } }, // 옅은 회색 배경
      alignment: { vertical: "center", horizontal: "center" },
      border: borderGray
    };

    const sProduct = {
      font: { name: fontName, size: 9, bold: true, color: { rgb: "1E293B" } },
      fill: { patternType: "solid", fgColor: { rgb: "E8F5E9" } }, // 민트 파스텔 연초록
      alignment: { vertical: "center", horizontal: "center" },
      border: borderMint
    };

    const sPallet = {
      font: { name: fontName, size: 9, bold: true, color: { rgb: "1D4ED8" } }, // 파란 수량 글씨
      fill: { patternType: "solid", fgColor: { rgb: "E3F2FD" } }, // 연하늘색 배경
      alignment: { vertical: "center", horizontal: "center" },
      border: borderBlue
    };

    const sBox = {
      font: { name: fontName, size: 9, bold: true, color: { rgb: "BE123C" } }, // 붉은 수량 글씨
      fill: { patternType: "solid", fgColor: { rgb: "FFEBEE" } }, // 연분홍색 배경
      alignment: { vertical: "center", horizontal: "center" },
      border: borderPink
    };

    const sEmptyGrid = {
      fill: { patternType: "solid", fgColor: { rgb: "FFFFFF" } }, // 흰색 배경으로 격자 테두리 유지
      border: borderGray
    };

    // 1행부터 maxRows행까지 순회하며 격자 셀 스타일링
    for (let r = 1; r <= maxRows; r++) {
      const prodRowIdx = 7 + 2 * (r - 1);
      const qtyRowIdx = 8 + 2 * (r - 1);

      // 행 라벨 (좌/우) 스타일링
      const leftLabelAddr = XLSX.utils.encode_cell({ r: prodRowIdx, c: 0 });
      const leftQtyLabelAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: 0 });
      const rightLabelAddr = XLSX.utils.encode_cell({ r: prodRowIdx, c: rightLabelCol });
      const rightQtyLabelAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: rightLabelCol });

      [leftLabelAddr, leftQtyLabelAddr, rightLabelAddr, rightQtyLabelAddr].forEach((addr) => {
        if (!sheet[addr]) sheet[addr] = { t: "n", v: r };
        sheet[addr].s = sHeader;
      });

      // 1~maxCols열 격자 루프
      for (let c = 1; c <= maxCols; c++) {
        const pCol = colMapping[c];
        if (pCol === undefined) continue;
        const bCol = pCol + 1;

        const prodAddr = XLSX.utils.encode_cell({ r: prodRowIdx, c: pCol });
        const palletAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: pCol });
        const boxAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: bCol });

        // 데이터가 없는 셀도 얇은 선명 격자를 만들기 위해 셀 구조 선언
        if (!sheet[prodAddr]) sheet[prodAddr] = { t: "s", v: "" };
        if (!sheet[palletAddr]) sheet[palletAddr] = { t: "s", v: "" };
        if (!sheet[boxAddr]) sheet[boxAddr] = { t: "s", v: "" };

        const productVal = String(sheet[prodAddr].v || "").trim();

        if (productVal !== "") {
          sheet[prodAddr].s = sProduct;
          sheet[palletAddr].s = sPallet;
          sheet[boxAddr].s = sBox;
        } else {
          sheet[prodAddr].s = sEmptyGrid;
          sheet[palletAddr].s = sEmptyGrid;
          sheet[boxAddr].s = sEmptyGrid;
        }
      }
    }

    // 상단 축 헤더 라벨링 스타일 입히기
    for (let c = 1; c <= maxCols; c++) {
      const pCol = colMapping[c];
      if (pCol === undefined) continue;
      const bCol = pCol + 1;

      // 4행 열 번호
      const colNumAddr = XLSX.utils.encode_cell({ r: 4, c: pCol });
      if (sheet[colNumAddr]) sheet[colNumAddr].s = sHeader;

      // 5행 "제품"
      const prodLabelAddr = XLSX.utils.encode_cell({ r: 5, c: pCol });
      if (sheet[prodLabelAddr]) sheet[prodLabelAddr].s = sHeader;

      // 6행 P / B
      const pAddr = XLSX.utils.encode_cell({ r: 6, c: pCol });
      if (sheet[pAddr]) sheet[pAddr].s = sHeader;
      const bAddr = XLSX.utils.encode_cell({ r: 6, c: bCol });
      if (sheet[bAddr]) sheet[bAddr].s = sHeader;
    }
  }
}

/**
 * TOTAL 시트 전용 디자인 서식 적용 함수
 *  - 열 너비를 콘텐츠에 맞게 최적화
 *  - 4행 헤더: 짙은 네이비 배경 + 흰 글씨 + 굵게
 *  - 데이터 행: 짝수/홀수 줄무늬, A열(제품명), D열(P-파란), E열(B-분홍) 강조
 *  - 전체 테두리 적용 및 정렬 통일
 */
function applyTotalSheetStyles(sheet) {
  const fontName = "맑은 고딕";

  // ── 열 너비 최적화 ──────────────────────────────────────────
  // A:제품명, B:입수량, C:박스단위, D:P, E:B, F:잔량, G:재고, H:출고, I:합계, J:판매일보, K:차이, L:창고위치
  sheet["!cols"] = [
    { wch: 14 }, // A: 제품명
    { wch: 8  }, // B: 입수량
    { wch: 8  }, // C: 박스단위
    { wch: 8  }, // D: P(파렛트)
    { wch: 8  }, // E: B(박스)
    { wch: 10 }, // F: 잔량(ea)
    { wch: 12 }, // G: 재고
    { wch: 12 }, // H: 출고
    { wch: 12 }, // I: 합계
    { wch: 12 }, // J: 판매일보
    { wch: 10 }, // K: 차이
    { wch: 16 }, // L: 창고위치
  ];

  // ── 공통 테두리 (진한 회색) ────────────────────────────────
  const border = {
    top:    { style: "thin", color: { rgb: "BDBDBD" } },
    bottom: { style: "thin", color: { rgb: "BDBDBD" } },
    left:   { style: "thin", color: { rgb: "BDBDBD" } },
    right:  { style: "thin", color: { rgb: "BDBDBD" } }
  };

  // ── 스타일 프리셋 ─────────────────────────────────────────
  // 헤더(4행): 진한 청남색 배경 + 흰 글씨
  const sHead = {
    font: { name: fontName, sz: 10, bold: true, color: { rgb: "FFFFFF" } },
    fill: { patternType: "solid", fgColor: { rgb: "1E3A5F" } },
    alignment: { vertical: "center", horizontal: "center", wrapText: true },
    border
  };

  // 제품명(A열): 연민트 + 중앙 정렬
  const sProd = (isOdd) => ({
    font: { name: fontName, sz: 9, bold: true, color: { rgb: "1B4332" } },
    fill: { patternType: "solid", fgColor: { rgb: isOdd ? "F0FDF4" : "DCFCE7" } },
    alignment: { vertical: "center", horizontal: "center" },
    border
  });

  // 파렛트(D열): 연하늘 + 파란 글씨
  const sPallet = (isOdd) => ({
    font: { name: fontName, sz: 10, bold: true, color: { rgb: "1D4ED8" } },
    fill: { patternType: "solid", fgColor: { rgb: isOdd ? "EFF6FF" : "DBEAFE" } },
    alignment: { vertical: "center", horizontal: "center" },
    border
  });

  // 박스(E열): 연분홍 + 붉은 글씨
  const sBox = (isOdd) => ({
    font: { name: fontName, sz: 10, bold: true, color: { rgb: "BE123C" } },
    fill: { patternType: "solid", fgColor: { rgb: isOdd ? "FFF1F2" : "FFE4E6" } },
    alignment: { vertical: "center", horizontal: "center" },
    border
  });

  // 일반 숫자 셀: 줄무늬 흰/연회색
  const sNum = (isOdd) => ({
    font: { name: fontName, sz: 9, color: { rgb: "374151" } },
    fill: { patternType: "solid", fgColor: { rgb: isOdd ? "FFFFFF" : "F9FAFB" } },
    alignment: { vertical: "center", horizontal: "center" },
    border
  });

  // 창고위치(L열): 연보라 + 작은 글씨
  const sLocation = (isOdd) => ({
    font: { name: fontName, sz: 9, color: { rgb: "5B21B6" } },
    fill: { patternType: "solid", fgColor: { rgb: isOdd ? "FAF5FF" : "F3E8FF" } },
    alignment: { vertical: "center", horizontal: "center", wrapText: true },
    border
  });

  // ── 4행 헤더 스타일 적용 (A~L 열, 인덱스 0~11) ──────────
  for (let col = 0; col <= TOTAL_COL.LOCATION; col++) {
    const addr = XLSX.utils.encode_cell({ r: 3, c: col }); // 4행 = r:3
    if (!sheet[addr]) sheet[addr] = { t: "s", v: "" };
    sheet[addr].s = sHead;
  }

  // ── 5행~데이터 마지막 행까지 스타일 적용 ─────────────────
  for (let excelRow = 5; excelRow <= 300; excelRow++) {
    const aAddr = XLSX.utils.encode_cell({ r: excelRow - 1, c: 0 }); // A열
    const aCell = sheet[aAddr];
    // A열이 비어있으면 데이터 끝으로 판단하고 중단
    if (!aCell || aCell.v === undefined || String(aCell.v).trim() === "") break;

    const isOdd = excelRow % 2 !== 0;

    // 열별 스타일 매핑
    const colStyles = [
      sProd(isOdd),    // A: 제품명 (col 0)
      sNum(isOdd),     // B: 입수량 (col 1)
      sNum(isOdd),     // C: 박스단위 (col 2)
      sPallet(isOdd),  // D: P 파렛트 (col 3)
      sBox(isOdd),     // E: B 박스 (col 4)
      sNum(isOdd),     // F: 잔량 (col 5)
      sNum(isOdd),     // G: 재고 (col 6)
      sNum(isOdd),     // H: 출고 (col 7)
      sNum(isOdd),     // I: 합계 (col 8)
      sNum(isOdd),     // J: 판매일보 (col 9)
      sNum(isOdd),     // K: 차이 (col 10)
      sLocation(isOdd), // L: 창고위치 (col 11)
    ];

    colStyles.forEach((style, colIdx) => {
      const addr = XLSX.utils.encode_cell({ r: excelRow - 1, c: colIdx });

      if (sheet[addr]) {
        // 기존 셀이 있으면 수식/값 보존, 스타일만 입힌
        sheet[addr].s = style;
      } else if (colIdx <= 5) {
        // A~F열(입력 데이터 영역)만: 비어있는 셀은 숫자 0으로 초기화
        // (빈 문자열 ""로 초기화하면 수식에서 #VALUE! 발생)
        sheet[addr] = { t: "n", v: 0, s: style };
      }
      // G~K열(수식 영역)은 셀이 없으면 절대 새로 만들지 않음 → 수식 안전 보존
    });
  }

  // ── 행 높이 설정 ─────────────────────────────────────────
  if (!sheet["!rows"]) sheet["!rows"] = [];
  sheet["!rows"][3] = { hpt: 22 }; // 헤더 행(4행) 높이
}

/**
 * 7. Supabase 실시간 동기화 유틸리티 함수군
 */

// Supabase 연결 초기화 및 연동 상태 체크
function initSupabase() {
  const url = window.SUPABASE_URL || localStorage.getItem("supabase_url") || "";
  const key = window.SUPABASE_ANON_KEY || localStorage.getItem("supabase_anon_key") || "";

  // URL과 Key는 입력되었는데 Supabase 라이브러리가 로드되지 않은 경우 경고
  if (url && key && !window.supabase) {
    alert("Supabase 라이브러리(SDK)가 브라우저에 로드되지 않았습니다.\n네트워크 상태나 광고 차단(AdBlock) 프로그램이 CDN 주소를 차단했는지 확인해 주세요.");
    setOfflineMode();
    return;
  }

  if (url && key && window.supabase) {
    try {
      supabaseClient = window.supabase.createClient(url, key);
      isOnline = true;
      syncStatusEl.textContent = "🟢 실시간 동기화 중";
      syncStatusEl.style.background = "rgba(16, 185, 129, 0.1)";
      syncStatusEl.style.border = "1px solid rgba(16, 185, 129, 0.4)";
      syncStatusEl.style.color = "#10b981";
      
      // 실시간 데이터 변경 채널 구독
      subscribeToRealtime();
      // DB 서버로부터 데이터 내려받기
      loadDataFromSupabase();
    } catch (e) {
      alert("Supabase 클라이언트 초기화 중 오류가 발생했습니다:\n" + e.message);
      console.error("Supabase 초기화 실패:", e);
      setOfflineMode();
    }
  } else {
    setOfflineMode();
  }
}

// 오프라인 상태 설정
function setOfflineMode() {
  isOnline = false;
  supabaseClient = null;
  syncStatusEl.textContent = "🔴 오프라인 모드";
  syncStatusEl.style.background = "rgba(239, 68, 68, 0.1)";
  syncStatusEl.style.border = "1px solid rgba(239, 68, 68, 0.4)";
  syncStatusEl.style.color = "#ef4444";
  updateSyncDiagnostics(0, false, "bundled");
}

// Supabase 서버 데이터 로드
async function loadDataFromSupabase() {
  const loadToken = ++supabaseLoadToken;

  try {
    if (isUserLocalWorkbook || isDbImportInProgress) {
      syncStatusEl.textContent = "🟢 실시간 동기화 중 (로컬 엑셀 우선)";
      return;
    }

    syncStatusEl.textContent = "🟡 데이터 로드 중...";

    // 1. DB 재고 먼저 조회 (가벼움) — PC/모바일 시트 이름 확인용
    let dbRacks;
    try {
      dbRacks = await fetchInventoryRacksFromDb();
    } catch (e) {
      throw new Error(
        "Supabase 서버(rack_inventory) 연결에 실패했습니다.\n" +
          (e.message || e) +
          "\n모바일: Wi-Fi/데이터 연결, ⚙️ DB 설정(URL·Key)을 확인해 주세요."
      );
    }
    if (loadToken !== supabaseLoadToken || isUserLocalWorkbook || isDbImportInProgress) return;

    const inventoryRacks = getInventoryRacks(dbRacks);
    const dbSheetNames = getDbSheetNames(dbRacks);
    const layoutRows = await fetchLayoutRowsFromDb();
    if (loadToken !== supabaseLoadToken || isUserLocalWorkbook || isDbImportInProgress) return;

    // 2. 업로드 엑셀 복원 → 없으면 DB 재고로 도면 자동 생성 (B1 vs 셀 불일치 방지)
    let templateSource = "bundled";
    let workbookBytes = null;

    const archived = await tryLoadArchivedWorkbookBytes();
    if (loadToken !== supabaseLoadToken || isUserLocalWorkbook || isDbImportInProgress) return;

    if (archived) {
      workbookBytes = archived.bytes;
      templateSource = archived.source;
    } else if (inventoryRacks.length > 0) {
      currentWorkbook = buildWorkbookFromDbInventory(dbRacks, layoutRows);
      templateSource = layoutRows.length > 0 ? "db-layout" : "db-generated";
    } else {
      try {
        workbookBytes = await loadBundledWorkbookBytes();
      } catch (e) {
        throw new Error("엑셀 템플릿 파일(재고조사표.xlsx)을 불러오는 데 실패했습니다 (네트워크 오류).");
      }
    }

    if (loadToken !== supabaseLoadToken || isUserLocalWorkbook || isDbImportInProgress) return;

    if (workbookBytes) {
      currentWorkbook = XLSX.read(workbookBytes, { type: "array", cellStyles: true });
    }

    let sheetMismatch = getSheetMismatch(dbRacks, currentWorkbook);
    if (sheetMismatch.length > 0 && inventoryRacks.length > 0) {
      currentWorkbook = buildWorkbookFromDbInventory(dbRacks, layoutRows);
      templateSource = layoutRows.length > 0 ? "db-layout" : "db-generated";
      sheetMismatch = getSheetMismatch(dbRacks, currentWorkbook);
    }

    // 3. Supabase에서 최신 제품 메타데이터 조회
    let dbMeta, metaError;
    try {
      const res = await supabaseClient
        .from('product_metadata')
        .select('*');
      dbMeta = res.data;
      metaError = res.error;
    } catch (e) {
      throw new Error("Supabase 서버(product_metadata) 연결에 실패했습니다.");
    }
    if (metaError) throw metaError;

    if (loadToken !== supabaseLoadToken || isUserLocalWorkbook || isDbImportInProgress) return;

    // 4. DB 데이터 반영
    if (templateSource === "bundled" && sheetMismatch.length > 0) {
      productsMetadataCache = {};
      (dbMeta || []).forEach((meta) => {
        productsMetadataCache[meta.product_name] = meta;
      });
    } else if (templateSource === "bundled" || inventoryRacks.length > 0) {
      applyDbDataToWorkbook(dbRacks, dbMeta);
    } else {
      productsMetadataCache = {};
      (dbMeta || []).forEach((meta) => {
        productsMetadataCache[meta.product_name] = meta;
      });
    }

    if (templateSource !== "bundled") {
      isUserLocalWorkbook = true;
    }

    // 5. 가이드 숨기고 그리드 표시
    const wbSheetNames = getWarehouseSheetNames(currentWorkbook);
    updateSheetDropdown(wbSheetNames);
    emptyGuide.style.display = "none";
    gridBoard.style.display = "grid";
    exportBtn.style.display = "inline-block";

    const defaultSheet = resolveDefaultSheetName(
      currentWorkbook,
      getRememberedSheet(),
      dbSheetNames
    );
    sheetSelect.value = defaultSheet;
    renderActiveSheet(defaultSheet);
    rememberLastSheet(defaultSheet);

    const hasWorkbookArchive =
      templateSource !== "bundled" &&
      templateSource !== "db-generated" &&
      templateSource !== "db-layout";

    if (sheetMismatch.length > 0 && templateSource === "bundled") {
      syncStatusEl.textContent = `🟡 시트 불일치 (DB: ${sheetMismatch.slice(0, 2).join(", ")}) — PC에서 엑셀 재업로드`;
      syncStatusEl.style.background = "rgba(245, 158, 11, 0.12)";
      syncStatusEl.style.border = "1px solid rgba(245, 158, 11, 0.4)";
      syncStatusEl.style.color = "#fbbf24";
      if (isMobileViewport()) {
        alert(
          `PC와 모바일 도면이 다릅니다.\n\nDB 재고 시트: ${dbSheetNames.join(", ")}\n모바일 도면: ${wbSheetNames.join(", ")}\n\nPC에서 엑셀을 다시 업로드한 뒤 모바일에서 🔄 클라우드 동기화를 눌러 주세요.`
        );
      }
    } else if (templateSource === "db-layout" || templateSource === "db-generated") {
      syncStatusEl.textContent = "🟢 실시간 동기화 중 (DB 재고 도면 복원)";
      syncStatusEl.style.background = "rgba(16, 185, 129, 0.1)";
      syncStatusEl.style.border = "1px solid rgba(16, 185, 129, 0.4)";
      syncStatusEl.style.color = "#10b981";
    } else if (templateSource === "snapshot" || templateSource === "storage" || templateSource === "db") {
      syncStatusEl.textContent = "🟢 실시간 동기화 중 (업로드 엑셀 복원)";
      syncStatusEl.style.background = "rgba(16, 185, 129, 0.1)";
      syncStatusEl.style.border = "1px solid rgba(16, 185, 129, 0.4)";
      syncStatusEl.style.color = "#10b981";
    } else if (templateSource === "cache") {
      syncStatusEl.textContent = "🟢 실시간 동기화 중 (저장된 엑셀 복원)";
      syncStatusEl.style.background = "rgba(16, 185, 129, 0.1)";
      syncStatusEl.style.border = "1px solid rgba(16, 185, 129, 0.4)";
      syncStatusEl.style.color = "#10b981";
    } else if (inventoryRacks.length > 0) {
      syncStatusEl.textContent = "🟢 실시간 동기화 중 (재고 데이터 연결됨)";
      syncStatusEl.style.background = "rgba(16, 185, 129, 0.1)";
      syncStatusEl.style.border = "1px solid rgba(16, 185, 129, 0.4)";
      syncStatusEl.style.color = "#10b981";
    } else if (isMobileViewport()) {
      syncStatusEl.textContent = "🟡 PC에서 엑셀 업로드 후 🔄 클라우드 동기화";
      syncStatusEl.style.background = "rgba(245, 158, 11, 0.12)";
      syncStatusEl.style.border = "1px solid rgba(245, 158, 11, 0.4)";
      syncStatusEl.style.color = "#fbbf24";
    } else {
      syncStatusEl.textContent = "🟢 실시간 동기화 중 (기본 템플릿)";
      syncStatusEl.style.background = "rgba(16, 185, 129, 0.1)";
      syncStatusEl.style.border = "1px solid rgba(16, 185, 129, 0.4)";
      syncStatusEl.style.color = "#10b981";
    }

    updateSyncDiagnostics(
      inventoryRacks.length,
      hasWorkbookArchive,
      templateSource,
      dbSheetNames,
      wbSheetNames
    );
  } catch (err) {
    if (loadToken !== supabaseLoadToken) return;
    console.error("데이터베이스 로드 실패:", err);
    syncStatusEl.textContent = `🔴 동기화 실패 — ${err.message || err}`;
    syncStatusEl.style.background = "rgba(239, 68, 68, 0.1)";
    syncStatusEl.style.border = "1px solid rgba(239, 68, 68, 0.4)";
    syncStatusEl.style.color = "#ef4444";
    updateSyncDiagnostics(0, false, "bundled");
    if (isMobileViewport()) {
      alert(
        "모바일 동기화 실패\n\n" +
          (err.message || err) +
          "\n\n① ⚙️ DB 설정 → PC와 동일한 URL·Key 입력\n② PC에서 엑셀 업로드\n③ 🔄 클라우드 동기화 버튼"
      );
    }
  }
}

// 데이터베이스 조회 정보를 엑셀 워크북에 매핑해 넣는 로직
function applyDbDataToWorkbook(dbRacks, dbMeta) {
  if (isDbImportInProgress) return;

  const totalSheetName = currentWorkbook.SheetNames.find(
    (n) => n.toUpperCase() === "TOTAL" || n.includes("합계")
  );
  
  // 제품 메타데이터 로컬 캐시 주입
  productsMetadataCache = {};
  dbMeta.forEach((meta) => {
    productsMetadataCache[meta.product_name] = meta;
  });

  // 모든 구역 랙 셀 비우고 초기화
  getWarehouseSheetNames(currentWorkbook).forEach((sheetName) => {
    const sheet = currentWorkbook.Sheets[sheetName];
    if (!sheet) return;
    const maxCols = detectMaxCols(sheet);
    const maxRows = detectMaxRows(sheet);
    const colMapping = detectColumnMapping(sheet, maxCols);

    for (let r = 1; r <= maxRows; r++) {
      const prodRowIdx = 7 + 2 * (r - 1);
      const qtyRowIdx  = 8 + 2 * (r - 1);
      for (let c = 1; c <= maxCols; c++) {
        const pCol = colMapping[c];
        if (pCol === undefined) continue;
        const bCol = pCol + 1;
        const prodAddr   = XLSX.utils.encode_cell({ r: prodRowIdx, c: pCol });
        const palletAddr = XLSX.utils.encode_cell({ r: qtyRowIdx,  c: pCol });
        const boxAddr    = XLSX.utils.encode_cell({ r: qtyRowIdx,  c: bCol });

        updateExcelCellValue(sheet, prodAddr, "", "s");
        updateExcelCellValue(sheet, palletAddr, 0, "n");
        updateExcelCellValue(sheet, boxAddr, 0, "n");
      }
    }
  });

  // DB 랙 데이터를 엑셀 메모리에 덮어쓰기 (워크북 아카이브 행 제외)
  dbRacks
    .filter((rack) => !SYSTEM_SHEET_NAMES.includes(rack.sheet_name))
    .forEach((rack) => {
    const sheet = currentWorkbook.Sheets[rack.sheet_name];
    if (!sheet) return;
    const colMapping = detectColumnMapping(sheet);
    const pColIdx = colMapping[rack.rack_col];
    if (pColIdx === undefined) return;
    const bColIdx = pColIdx + 1;

    const prodRowIdx = 7 + 2 * (rack.rack_row - 1);
    const qtyRowIdx = 8 + 2 * (rack.rack_row - 1);

    const prodAddr = XLSX.utils.encode_cell({ r: prodRowIdx, c: pColIdx });
    const palletAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: pColIdx });
    const boxAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: bColIdx });

    updateExcelCellValue(sheet, prodAddr, rack.product_name, 's');
    updateExcelCellValue(sheet, palletAddr, rack.pallet_qty, 'n');
    updateExcelCellValue(sheet, boxAddr, rack.box_qty, 'n');
  });
}

// 엑셀 업로드 시 엑셀 데이터를 Supabase 클라우드로 일괄 가져오는 적재기
async function importWorkbookToSupabase() {
  if (!currentWorkbook || !supabaseClient) return;

  const savedSheetName = currentSheetName;
  isDbImportInProgress = true;
  isUserLocalWorkbook = true;
  supabaseLoadToken += 1;
  pauseRealtimeSync();

  try {
    syncStatusEl.textContent = "🟡 DB 덮어쓰기 중...";

    // TOTAL 시트를 창고 재고 기준으로 먼저 동기화 (제품명 변경·삭제 반영)
    aggregateTotalSheet(currentWorkbook);
    
    // 1. 기존 데이터 초기화 (완전 덮어쓰기)
    const { error: delRackErr } = await supabaseClient.from('rack_inventory').delete().neq('sheet_name', '');
    if (delRackErr) throw delRackErr;

    const { error: delMetaErr } = await supabaseClient.from('product_metadata').delete().neq('product_name', '');
    if (delMetaErr) throw delMetaErr;

    const totalSheetName = currentWorkbook.SheetNames.find(
      (n) => n.toUpperCase() === "TOTAL" || n.includes("합계")
    );
    const totalSheet = totalSheetName ? currentWorkbook.Sheets[totalSheetName] : null;

    const rackInserts = [];
    const metaInserts = [];

    // TOTAL 시트에서 메타데이터 읽기
    if (totalSheet) {
      for (let excelRow = 5; excelRow <= 300; excelRow++) {
        const aAddr = XLSX.utils.encode_cell({ r: excelRow - 1, c: 0 }); // A열
        const aCell = totalSheet[aAddr];
        if (!aCell || aCell.v === undefined || String(aCell.v).trim() === "") break;
        const prodName = String(aCell.v).trim();

        const bCell = totalSheet[XLSX.utils.encode_cell({ r: excelRow - 1, c: 1 })]; // B열
        const cCell = totalSheet[XLSX.utils.encode_cell({ r: excelRow - 1, c: 2 })]; // C열
        const fCell = totalSheet[XLSX.utils.encode_cell({ r: excelRow - 1, c: 5 })]; // F열
        const hCell = totalSheet[XLSX.utils.encode_cell({ r: excelRow - 1, c: 7 })]; // H열
        const jCell = totalSheet[XLSX.utils.encode_cell({ r: excelRow - 1, c: 9 })]; // J열

        metaInserts.push({
          product_name: prodName,
          ipsuryang: bCell && bCell.v !== undefined ? Number(bCell.v) || 0 : 0,
          box_danyi: cCell && cCell.v !== undefined ? Number(cCell.v) || 0 : 0,
          janryang: fCell && fCell.v !== undefined ? Number(fCell.v) || 0 : 0,
          chulgo: hCell && hCell.v !== undefined ? Number(hCell.v) || 0 : 0,
          panmae_ilbo: jCell && jCell.v !== undefined ? Number(jCell.v) || 0 : 0
        });
      }
    }

    // 각 구역 시트에서 랙 리스트 긁기
    getWarehouseSheetNames(currentWorkbook).forEach((sheetName) => {
      const sheet = currentWorkbook.Sheets[sheetName];
      if (!sheet) return;

      const maxCols = detectMaxCols(sheet);
      const maxRows = detectMaxRows(sheet);
      const colMapping = detectColumnMapping(sheet, maxCols);
      for (let r = 1; r <= maxRows; r++) {
        const prodRowIdx = 7 + 2 * (r - 1);
        const qtyRowIdx  = 8 + 2 * (r - 1);

        for (let c = 1; c <= maxCols; c++) {
          const cellData = getRackCellData(sheet, colMapping, r, c, prodRowIdx, qtyRowIdx);
          if (cellData.product) {
            rackInserts.push({
              sheet_name: sheetName,
              rack_row: r,
              rack_col: c,
              product_name: cellData.product,
              pallet_qty: cellData.pallet,
              box_qty: cellData.box
            });
          }
        }
      }
    });

    // 2. Supabase 데이터 주입
    if (metaInserts.length > 0) {
      const { error: metaInsertErr } = await supabaseClient.from('product_metadata').insert(metaInserts);
      if (metaInsertErr) throw metaInsertErr;
    }

    if (rackInserts.length > 0) {
      // 대량 데이터 분할 주입 (Supabase 1회 제한 방지)
      const chunkSize = 100;
      for (let i = 0; i < rackInserts.length; i += chunkSize) {
        const chunk = rackInserts.slice(i, i + chunkSize);
        const { error: rackInsertErr } = await supabaseClient.from('rack_inventory').insert(chunk);
        if (rackInsertErr) throw rackInsertErr;
      }
    }

    await saveSheetLayoutsToDb(currentWorkbook);

    // 3. 업로드한 워크북/레이아웃 유지 — 서버 템플릿 재로드 없음
    isUserLocalWorkbook = true;
    if (savedSheetName && currentWorkbook.Sheets[savedSheetName]) {
      sheetSelect.value = savedSheetName;
      renderActiveSheet(savedSheetName);
    } else {
      renderActiveSheet(currentSheetName);
    }

    syncStatusEl.textContent = "🟢 업로드 → DB 자동 동기화 완료";
    syncStatusEl.style.background = "rgba(16, 185, 129, 0.1)";
    syncStatusEl.style.border = "1px solid rgba(16, 185, 129, 0.4)";
    syncStatusEl.style.color = "#10b981";

    const workbookSaved = await persistUploadedWorkbook(currentWorkbook);
    if (!workbookSaved) {
      syncStatusEl.textContent = "🟡 재고 동기화됨 — 엑셀 파일 저장 실패 (모바일은 DB 도면 복원)";
      syncStatusEl.style.background = "rgba(245, 158, 11, 0.12)";
      syncStatusEl.style.border = "1px solid rgba(245, 158, 11, 0.4)";
      syncStatusEl.style.color = "#fbbf24";
      alert(
        "재고·도면 레이아웃은 DB에 저장됐습니다.\n엑셀 원본 파일 저장만 실패했 — 모바일은 DB 재고 도면으로 표시됩니다.\n\n더 정확한 도면을 위해 엑셀을 한 번 더 업로드해 주세요."
      );
    }
  } catch (err) {
    alert("서버에 데이터를 업로드하는 중 오류가 발생했습니다.");
    console.error(err);
    setOfflineMode();
  } finally {
    isDbImportInProgress = false;
    isUserLocalWorkbook = true;
    // 서버 템플릿 재로드는 막되, 셀 단위 실시간 변경은 PC·모바일 간 수신
    subscribeToRealtime();
  }
}

// 실시간 DB 변경사항 구독 등록
let realtimeChannel = null;

function pauseRealtimeSync() {
  if (realtimeChannel && supabaseClient) {
    supabaseClient.removeChannel(realtimeChannel);
    realtimeChannel = null;
  }
}

function subscribeToRealtime() {
  if (!isOnline || !supabaseClient || isDbImportInProgress) return;

  pauseRealtimeSync();

  realtimeChannel = supabaseClient
    .channel('public:rack_inventory')
    .on('postgres_changes', { event: '*', schema: 'public', table: 'rack_inventory' }, (payload) => {
      handleRealtimeRackChange(payload);
    })
    .subscribe();
}

// 실시간 기기 변경에 대한 수신 처리기
function handleRealtimeRackChange(payload) {
  if (isDbImportInProgress) return;
  if (!currentWorkbook) return;

  const { eventType, new: newRow, old: oldRow } = payload;
  
  // 1. 메모리 상의 워크북 객체 실시간 수정
  if (eventType === 'DELETE') {
    if (oldRow.sheet_name === WORKBOOK_ARCHIVE_SHEET) return;
    const sheet = currentWorkbook.Sheets[oldRow.sheet_name];
    if (sheet) {
      const colMapping = detectColumnMapping(sheet);
      const pColIdx = colMapping[oldRow.rack_col];
      if (pColIdx !== undefined) {
        const prodRowIdx = 7 + 2 * (oldRow.rack_row - 1);
        const qtyRowIdx = 8 + 2 * (oldRow.rack_row - 1);
        
        const prodAddr = XLSX.utils.encode_cell({ r: prodRowIdx, c: pColIdx });
        const palletAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: pColIdx });
        const boxAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: pColIdx + 1 });

        updateExcelCellValue(sheet, prodAddr, "", "s");
        updateExcelCellValue(sheet, palletAddr, 0, "n");
        updateExcelCellValue(sheet, boxAddr, 0, "n");
      }
    }
  } else {
    // INSERT, UPDATE
    if (newRow.sheet_name === WORKBOOK_ARCHIVE_SHEET) return;
    const sheet = currentWorkbook.Sheets[newRow.sheet_name];
    if (sheet) {
      const colMapping = detectColumnMapping(sheet);
      const pColIdx = colMapping[newRow.rack_col];
      if (pColIdx !== undefined) {
        const prodRowIdx = 7 + 2 * (newRow.rack_row - 1);
        const qtyRowIdx = 8 + 2 * (newRow.rack_row - 1);
        
        const prodAddr = XLSX.utils.encode_cell({ r: prodRowIdx, c: pColIdx });
        const palletAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: pColIdx });
        const boxAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: pColIdx + 1 });

        updateExcelCellValue(sheet, prodAddr, newRow.product_name, "s");
        updateExcelCellValue(sheet, palletAddr, newRow.pallet_qty, "n");
        updateExcelCellValue(sheet, boxAddr, newRow.box_qty, "n");
      }
    }
  }

  // 2. 화면 동적 새로고침 (현재 사용자가 해당 시트를 보는 중이고 수정 모달이 닫혀있는 상태일 때)
  const activeSheetName = currentSheetName;
  const isEditing = editModal.classList.contains("show");
  const affectedSheet = eventType === 'DELETE' ? oldRow.sheet_name : newRow.sheet_name;
  
  if (activeSheetName === affectedSheet && !isEditing) {
    renderActiveSheet(currentSheetName);
  }
}

// DB 설정 모달 열기/닫기/제어
function showDbModal() {
  inputDbUrl.value = window.SUPABASE_URL || localStorage.getItem("supabase_url") || "";
  inputDbKey.value = window.SUPABASE_ANON_KEY || localStorage.getItem("supabase_anon_key") || "";
  if (autoSyncOnUploadInput) {
    autoSyncOnUploadInput.checked = isAutoSyncOnUploadEnabled();
  }
  dbConfigModal.classList.add("show");
}

function hideDbModal() {
  dbConfigModal.classList.remove("show");
}

function saveDbConfig() {
  const url = inputDbUrl.value.trim();
  const key = inputDbKey.value.trim();
  
  if (!url || !key) {
    alert("URL과 Anon Key를 모두 입력해야 저장됩니다.");
    return;
  }

  saveDbConfigBtn.disabled = true;
  saveDbConfigBtn.textContent = "연결 확인 중...";

  testSupabaseConnection(url, key).then(async (test) => {
    saveDbConfigBtn.disabled = false;
    saveDbConfigBtn.textContent = "연동 완료";

    if (!test.ok) {
      alert(
        "Supabase 연결에 실패했습니다.\n\n" +
          (test.error || "알 수 없는 오류") +
          "\n\nURL·Anon Key를 PC Supabase 설정과 동일하게 입력했는지 확인해 주세요."
      );
      return;
    }

    localStorage.setItem("supabase_url", url);
    localStorage.setItem("supabase_anon_key", key);
    if (autoSyncOnUploadInput) {
      localStorage.setItem("auto_sync_on_upload", autoSyncOnUploadInput.checked ? "true" : "false");
    }
    hideDbModal();
    isUserLocalWorkbook = false;
    isDbImportInProgress = false;
    supabaseLoadToken += 1;
    initSupabase();
    alert("연결 성공! 데이터를 불러옵니다.");
  });
}

function disconnectDb() {
  if (confirm("Supabase 실시간 연동을 해제하시겠습니까? (로컬 오프라인 엑셀 방식으로 돌아갑니다.)")) {
    localStorage.removeItem("supabase_url");
    localStorage.removeItem("supabase_anon_key");
    hideDbModal();
    alert("연동 해제 완료.");
    window.location.reload();
  }
}

// 앱 구동 — config.js 로드 후 Supabase 초기화
async function bootApplication() {
  await loadOptionalConfigScript();
  initSupabase();
}

window.addEventListener("resize", () => {
  if (currentSheetName && gridBoard.style.display !== "none") {
    const layout = gridBoard.dataset.layout || "";
    const colMatch = layout.match(/^(\d+)\+(\d+)/);
    if (colMatch) {
      applyResponsiveGridLayout(parseInt(colMatch[1], 10) + parseInt(colMatch[2], 10));
    }
  }
});

bootApplication();

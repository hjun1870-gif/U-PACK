/**
 * app.js
 * 창고 재고조사표 시각화 및 편집 애플리케이션의 핵심 비즈니스 로직
 * - SheetJS (xlsx.js) 라이브러리를 활용하여 로컬 엑셀 파일을 파싱하고 갱신
 * - 시트별 동적 컬럼 매핑으로 구조 변경 유연성 제공
 * - 실시간 재고 검색, 수정 및 다운로드 기능 제공
 */

// 애플리케이션 상태 관리 전역 변수
let currentWorkbook = null; // 로드된 전체 엑셀 워크북 객체
/** 업로드/복원한 원본 엑셀 바이트 — 다운로드 시 레이아웃·서식 유지용 */
let pristineLayoutBytes = null;
let productCatalogCache = [];
let autocompleteActiveIndex = -1;
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
const searchAutocomplete = document.getElementById("searchAutocomplete");
const searchResults = document.getElementById("searchResults");     // 결과 목록 패널
const searchResultsWrap = document.getElementById("searchResultsWrap"); // 래퍼: 이것을 show/hide하여 창고 도면을 밀어냄
const exportBtn = document.getElementById("exportBtn");
const exportHtmlBtn = document.getElementById("exportHtmlBtn");
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
const deleteBtn = document.getElementById("deleteBtn");

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
const rackFilterGroup = document.getElementById("rackFilterGroup");
const undoBtn = document.getElementById("undoBtn");
const occupancyRateEl = document.getElementById("occupancyRate");
const totalMatchEl = document.getElementById("totalMatchStatus");
const searchNavBar = document.getElementById("searchNavBar");
const searchPrevBtn = document.getElementById("searchPrevBtn");
const searchNextBtn = document.getElementById("searchNextBtn");
const searchNavLabel = document.getElementById("searchNavLabel");
const moveBtn = document.getElementById("moveBtn");
const moveModal = document.getElementById("moveModal");
const closeMoveModal = document.getElementById("closeMoveModal");
const cancelMoveBtn = document.getElementById("cancelMoveBtn");
const confirmMoveBtn = document.getElementById("confirmMoveBtn");
const moveSheetSelect = document.getElementById("moveSheetSelect");
const moveRowInput = document.getElementById("moveRow");
const moveColInput = document.getElementById("moveCol");
const toastContainer = document.getElementById("toastContainer");

let rackDisplayFilter = "all";
let lastSearchMatches = [];
let lastSearchMatchIndex = -1;
let lastUndoSnapshot = null;
let syncFailedState = false;
let globalLoadingCount = 0;

function showToast(message, type = "info", duration = 3000) {
  if (!toastContainer) return;
  const toast = document.createElement("div");
  toast.className = `toast toast-${type}`;
  toast.textContent = message;
  toastContainer.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = "0";
    toast.style.transition = "opacity 0.3s";
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

function setGlobalLoading(loading, statusText) {
  if (loading) {
    globalLoadingCount += 1;
  } else {
    globalLoadingCount = Math.max(0, globalLoadingCount - 1);
  }
  const busy = globalLoadingCount > 0;
  document.body.classList.toggle("is-loading", busy);
  [exportBtn, exportHtmlBtn, cloudSyncBtn, demoBtn].forEach((btn) => {
    if (btn) btn.disabled = busy;
  });
  const uploadBtn = document.querySelector(".file-upload-btn");
  if (uploadBtn) uploadBtn.disabled = busy;
  if (statusText && syncStatusEl && busy) {
    syncStatusEl.textContent = statusText;
  }
}

function markSyncFailed(message) {
  syncFailedState = true;
  if (!syncStatusEl) return;
  syncStatusEl.textContent = message || "🔴 동기화 실패 — 클릭하여 재시도";
  syncStatusEl.classList.add("sync-failed");
  syncStatusEl.title = "클릭하여 동기화 재시도";
}

function clearSyncFailed() {
  syncFailedState = false;
  if (syncStatusEl) {
    syncStatusEl.classList.remove("sync-failed");
    syncStatusEl.title = "";
  }
}

function buildDuplicateProductMap() {
  const counts = new Map();
  if (!currentWorkbook) return counts;
  getWarehouseSheetNames(currentWorkbook).forEach((sheetName) => {
    const sheet = currentWorkbook.Sheets[sheetName];
    if (!sheet) return;
    const maxCols = detectMaxCols(sheet);
    const maxRows = detectMaxRows(sheet);
    const colMapping = detectColumnMapping(sheet);
    for (let r = 1; r <= maxRows; r++) {
      const prodRowIdx = 7 + 2 * (r - 1);
      const qtyRowIdx = 8 + 2 * (r - 1);
      for (let c = 1; c <= maxCols; c++) {
        const data = getRackCellData(sheet, colMapping, r, c, prodRowIdx, qtyRowIdx);
        if (!data.product) continue;
        const key = data.product.toLowerCase();
        counts.set(key, (counts.get(key) || 0) + 1);
      }
    }
  });
  return counts;
}

function applyRackFilterToCell(cell, cellData) {
  const hasStock = Boolean(cellData.product || cellData.pallet > 0 || cellData.box > 0);
  cell.classList.toggle("rack-filter-hidden", (
    (rackDisplayFilter === "stocked" && !hasStock) ||
    (rackDisplayFilter === "empty" && hasStock)
  ));
}

function updateUndoButtonVisibility() {
  if (undoBtn) undoBtn.style.display = lastUndoSnapshot ? "inline-block" : "none";
}

function saveUndoSnapshot() {
  if (!selectedCellInfo || !currentSheetName) return;
  lastUndoSnapshot = {
    sheetName: currentSheetName,
    row: selectedCellInfo.row,
    col: selectedCellInfo.col,
    product: selectedCellInfo.data.product || "",
    pallet: selectedCellInfo.data.pallet || 0,
    box: selectedCellInfo.data.box || 0,
  };
  updateUndoButtonVisibility();
}

async function performUndo() {
  if (!lastUndoSnapshot || !currentWorkbook) return;
  const snap = lastUndoSnapshot;
  const sheet = currentWorkbook.Sheets[snap.sheetName];
  if (!sheet) {
    showToast("되돌릴 시트를 찾을 수 없습니다.", "error");
    return;
  }
  if (snap.sheetName !== currentSheetName) {
    sheetSelect.value = snap.sheetName;
    renderActiveSheet(snap.sheetName);
  }
  const colMapping = detectColumnMapping(sheet);
  const prodRowIdx = 7 + 2 * (snap.row - 1);
  const qtyRowIdx = 8 + 2 * (snap.row - 1);
  const cellData = getRackCellData(sheet, colMapping, snap.row, snap.col, prodRowIdx, qtyRowIdx);
  selectedCellInfo = { row: snap.row, col: snap.col, data: cellData };
  lastUndoSnapshot = null;
  updateUndoButtonVisibility();
  await applyCellChanges(snap.product, snap.pallet, snap.box, { skipUndo: true });
  showToast("마지막 편집을 되돌렸습니다.", "success");
}

function focusSearchMatchAt(index) {
  if (!lastSearchMatches.length) return;
  const safeIndex = ((index % lastSearchMatches.length) + lastSearchMatches.length) % lastSearchMatches.length;
  lastSearchMatchIndex = safeIndex;
  const match = lastSearchMatches[safeIndex];
  if (searchNavLabel) {
    searchNavLabel.textContent = `${safeIndex + 1} / ${lastSearchMatches.length}`;
  }
  if (currentSheetName !== match.sheetName) {
    sheetSelect.value = match.sheetName;
    renderActiveSheet(match.sheetName);
  }
  searchInput.value = match.product;
  handleSearchInput();
  const targetCell = gridBoard.querySelector(`.rack-cell[data-row="${match.row}"][data-col="${match.col}"]`);
  if (targetCell) {
    targetCell.scrollIntoView({ behavior: "smooth", block: "center", inline: "center" });
    targetCell.classList.add("search-highlight");
    setTimeout(() => targetCell.classList.remove("search-highlight"), 2500);
  }
}

function updateSearchNavBar() {
  if (!searchNavBar) return;
  const show = lastSearchMatches.length > 0 && searchInput.value.trim() !== "";
  searchNavBar.style.display = show ? "flex" : "none";
  if (searchPrevBtn) searchPrevBtn.disabled = !show;
  if (searchNextBtn) searchNextBtn.disabled = !show;
  if (searchNavLabel && show) {
    searchNavLabel.textContent = `${Math.max(0, lastSearchMatchIndex) + 1} / ${lastSearchMatches.length}`;
  }
}

function showMoveModal() {
  if (!selectedCellInfo || !currentWorkbook) return;
  const hasData = Boolean(selectedCellInfo.data.product || selectedCellInfo.data.pallet || selectedCellInfo.data.box);
  if (!hasData) {
    showToast("이동할 재고가 없습니다.", "warning");
    return;
  }
  const sheets = getWarehouseSheetNames(currentWorkbook);
  moveSheetSelect.innerHTML = sheets.map((n) => `<option value="${n}">${n}</option>`).join("");
  moveSheetSelect.value = currentSheetName;
  moveRowInput.value = selectedCellInfo.row;
  moveColInput.value = selectedCellInfo.col;
  moveModal.classList.add("show");
  document.body.classList.add("modal-open");
}

function hideMoveModal() {
  moveModal.classList.remove("show");
  if (!editModal.classList.contains("show") && !dbConfigModal.classList.contains("show")) {
    document.body.classList.remove("modal-open");
  }
}

async function confirmMoveStock() {
  if (!selectedCellInfo || !currentWorkbook) return;
  const srcSheet = currentSheetName;
  const srcRow = selectedCellInfo.row;
  const srcCol = selectedCellInfo.col;
  const product = inputProduct.value.trim() || selectedCellInfo.data.product;
  const pallet = parseInt(inputPallet.value, 10) || 0;
  const box = parseInt(inputBox.value, 10) || 0;
  const destSheet = moveSheetSelect.value;
  const destRow = parseInt(moveRowInput.value, 10);
  const destCol = parseInt(moveColInput.value, 10);

  if (!destSheet || !destRow || !destCol) {
    showToast("이동 대상 위치를 입력해 주세요.", "warning");
    return;
  }
  if (srcSheet === destSheet && srcRow === destRow && srcCol === destCol) {
    showToast("같은 위치로는 이동할 수 없습니다.", "warning");
    return;
  }

  saveUndoSnapshot();
  const destSheetObj = currentWorkbook.Sheets[destSheet];
  if (!destSheetObj) {
    showToast("대상 구역을 찾을 수 없습니다.", "error");
    return;
  }
  const destMapping = detectColumnMapping(destSheetObj);
  const destProdRowIdx = 7 + 2 * (destRow - 1);
  const destQtyRowIdx = 8 + 2 * (destRow - 1);
  const destData = getRackCellData(destSheetObj, destMapping, destRow, destCol, destProdRowIdx, destQtyRowIdx);

  updateExcelCellValue(destSheetObj, destData.prodCellAddress, product, "s");
  updateExcelCellValue(destSheetObj, destData.palletCellAddress, pallet, "n");
  updateExcelCellValue(destSheetObj, destData.boxCellAddress, box, "n");

  if (isOnline && supabaseClient) {
    try {
      await supabaseClient.from("rack_inventory").upsert({
        sheet_name: destSheet,
        rack_row: destRow,
        rack_col: destCol,
        product_name: product,
        pallet_qty: pallet,
        box_qty: box,
        updated_at: new Date().toISOString(),
      }, { onConflict: "sheet_name,rack_row,rack_col" });
    } catch (e) {
      showToast("대상 위치 서버 동기화 실패 (로컬만 반영)", "warning");
    }
  }

  selectedCellInfo = { row: srcRow, col: srcCol, data: selectedCellInfo.data };
  currentSheetData = currentWorkbook.Sheets[srcSheet];
  await applyCellChanges("", 0, 0, { skipUndo: true });

  sheetSelect.value = destSheet;
  renderActiveSheet(destSheet);
  hideMoveModal();
  hideModal();
  showToast(`[${destSheet}] ${destRow}행-${destCol}열로 이동했습니다.`, "success");
}

function showWorkbookControls() {
  if (rackFilterGroup) rackFilterGroup.style.display = "flex";
  if (undoBtn && lastUndoSnapshot) undoBtn.style.display = "inline-block";
}

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
/** 입수량·박스단위 참조 시트 (헤더 2행, 데이터 3행~) */
const PACK_UNIT_SHEET_HINT = "입수량";
const PACK_UNIT_BOX_HINT = "박스단위";
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
  CHECK: 12,
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

async function saveBytesToLocalCache(bytes) {
  if (!bytes || bytes.length < 100) return;
  try {
    const payload = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const db = await openWorkbookIdb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(WORKBOOK_IDB_STORE, "readwrite");
      tx.objectStore(WORKBOOK_IDB_STORE).put(payload, WORKBOOK_IDB_KEY);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
    db.close();
    localStorage.setItem("workbook_cache_active", "true");
  } catch (e) {
    console.warn("로컬 워크북 캐시 저장 실패:", e);
  }
}

async function saveWorkbookToLocalCache(workbook) {
  if (pristineLayoutBytes && pristineLayoutBytes.length > 100) {
    await saveBytesToLocalCache(pristineLayoutBytes);
    return;
  }
  if (!workbook) return;
  try {
    await saveBytesToLocalCache(workbookToBytes(workbook));
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

async function uploadRawWorkbookBytesToStorage(bytes) {
  if (!supabaseClient || !bytes || bytes.length < 100) return false;
  try {
    const payload = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const blob = new Blob([payload], {
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

async function uploadWorkbookToStorage(workbook) {
  if (pristineLayoutBytes && pristineLayoutBytes.length > 100) {
    return uploadRawWorkbookBytesToStorage(pristineLayoutBytes);
  }
  if (!supabaseClient || !workbook) return false;
  try {
    return await uploadRawWorkbookBytesToStorage(workbookToBytes(workbook));
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

async function uploadRawWorkbookBytesToSnapshotTable(bytes) {
  if (!supabaseClient || !bytes || bytes.length < 100) return false;
  try {
    const payload = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const base64 = uint8ArrayToBase64(payload);
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

async function uploadWorkbookToSnapshotTable(workbook) {
  if (pristineLayoutBytes && pristineLayoutBytes.length > 100) {
    return uploadRawWorkbookBytesToSnapshotTable(pristineLayoutBytes);
  }
  if (!supabaseClient || !workbook) return false;
  try {
    // 원본 바이트가 없으면 재인코딩 (레이아웃 손실 가능)
    const bytes = workbookToBytes(workbook);
    if (!bytes || bytes.length < 100) throw new Error("엑셀 변환 결과가 비어 있습니다.");
    return uploadRawWorkbookBytesToSnapshotTable(bytes);
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
  if (!supabaseClient) return false;
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      // 원본 바이트 우선 — 재인코딩하면 도면 레이아웃이 깨짐
      const bytes =
        pristineLayoutBytes && pristineLayoutBytes.length > 100
          ? pristineLayoutBytes
          : workbook
            ? workbookToArchiveBytes(workbook)
            : null;
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
  // 레이아웃에만 있고 재고가 비어 있는 구역(B1/B2/B3 등)도 시트 목록에 포함
  const layoutSheetNames = (layoutRows || [])
    .slice()
    .sort((a, b) => (a.rack_row || 0) - (b.rack_row || 0))
    .map((row) => row.product_name)
    .filter(Boolean);
  const inventorySheetNames = [...new Set(inventory.map((r) => r.sheet_name))].filter(Boolean);
  const sheetNames = [...new Set([...layoutSheetNames, ...inventorySheetNames])];
  const wb = XLSX.utils.book_new();

  sheetNames.forEach((sheetName) => {
    const racks = inventory.filter((r) => r.sheet_name === sheetName);
    const layout = layoutMap[sheetName];
    const maxRow =
      layout?.maxRows ||
      (racks.length ? Math.max(18, ...racks.map((r) => r.rack_row)) : 18);
    let totalCols;
    if (layout?.leftCols || layout?.rightCols) {
      totalCols = (layout.leftCols || 0) + (layout.rightCols || 0) || 11;
    } else if (racks.length) {
      totalCols = Math.max(11, ...racks.map((r) => r.rack_col));
    } else {
      totalCols = 11;
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
  // Storage·로컬 캐시는 서식 포함 가능성이 높아 우선 사용
  const storageBytes = await downloadWorkbookBytesFromStorage();
  if (storageBytes) return { bytes: storageBytes, source: "storage" };

  const snapshotBytes = await downloadWorkbookBytesFromSnapshotTable();
  if (snapshotBytes) return { bytes: snapshotBytes, source: "snapshot" };

  const cacheBytes = await loadWorkbookBytesFromLocalCache();
  if (cacheBytes) return { bytes: cacheBytes, source: "cache" };

  const dbSnapshotBytes = await downloadWorkbookBytesFromDbSnapshot();
  if (dbSnapshotBytes) return { bytes: dbSnapshotBytes, source: "db" };

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
    showToast("Supabase 연동 후 사용할 수 있습니다. ⚙️ DB 설정을 확인해 주세요.", "warning");
    return;
  }
  isUserLocalWorkbook = false;
  isDbImportInProgress = false;
  clearSyncFailed();
  setGlobalLoading(true, "🟡 클라우드 동기화 중...");
  try {
    await loadDataFromSupabase();
    showToast("클라우드 데이터를 불러왔습니다.", "success");
  } catch (e) {
    markSyncFailed("🔴 동기화 실패 — 클릭하여 재시도");
  } finally {
    setGlobalLoading(false);
  }
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
  el.classList.remove("is-warning", "is-ok");

  const renderChips = (items) => {
    el.innerHTML = items.map((text) => `<span class="sync-diag-chip">${text}</span>`).join("");
  };

  if (!isOnline) {
    renderChips(["DB 미연결", "⚙️ DB 설정에서 Supabase URL·Key를 입력하세요"]);
    el.classList.add("is-warning");
    return;
  }

  const dbLabel = dbSheetNames.length ? dbSheetNames.slice(0, 4).join(", ") : "-";
  const wbLabel = wbSheetNames.length ? wbSheetNames.slice(0, 4).join(", ") : "-";
  const parts = [
    `DB 재고 ${inventoryCount}건`,
    hasWorkbookArchive ? "업로드 엑셀 ✓" : "업로드 엑셀 ✗",
    `DB시트 ${dbLabel}`,
    `도면 ${wbLabel}`,
  ];

  if (!hasWorkbookArchive && dbSheetNames.length > 0 && templateSource === "bundled") {
    parts.push("⚠️ PC에서 엑셀 재업로드 필요");
    el.classList.add("is-warning");
  } else if (templateSource === "db-layout" || templateSource === "db-generated") {
    parts.push("DB 도면 복원");
    el.classList.add("is-ok");
  }

  renderChips(parts);
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
searchInput.addEventListener("input", handleSearchInput);
searchInput.addEventListener("focus", handleSearchInput);
searchInput.addEventListener("keydown", handleSearchKeydown);
exportBtn.addEventListener("click", exportUpdatedExcel);
exportHtmlBtn.addEventListener("click", exportWarehouseHtmlReport);

// 모달 닫기 이벤트들
closeModal.addEventListener("click", hideModal);
cancelBtn.addEventListener("click", hideModal);
window.addEventListener("click", (e) => {
  if (e.target === editModal) hideModal();
});
saveBtn.addEventListener("click", saveCellChanges);
if (deleteBtn) {
  deleteBtn.addEventListener("click", deleteCellChanges);
}

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

if (syncStatusEl) {
  syncStatusEl.addEventListener("click", () => {
    if (syncFailedState) reloadFromCloud();
  });
}

if (rackFilterGroup) {
  rackFilterGroup.querySelectorAll(".rack-filter-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      rackFilterGroup.querySelectorAll(".rack-filter-btn").forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      rackDisplayFilter = btn.dataset.filter || "all";
      if (currentSheetName) renderActiveSheet(currentSheetName);
    });
  });
}

if (undoBtn) undoBtn.addEventListener("click", performUndo);

if (searchPrevBtn) {
  searchPrevBtn.addEventListener("click", () => focusSearchMatchAt(lastSearchMatchIndex - 1));
}
if (searchNextBtn) {
  searchNextBtn.addEventListener("click", () => focusSearchMatchAt(lastSearchMatchIndex + 1));
}

if (moveBtn) moveBtn.addEventListener("click", showMoveModal);
if (closeMoveModal) closeMoveModal.addEventListener("click", hideMoveModal);
if (cancelMoveBtn) cancelMoveBtn.addEventListener("click", hideMoveModal);
if (confirmMoveBtn) confirmMoveBtn.addEventListener("click", confirmMoveStock);
window.addEventListener("click", (e) => {
  if (e.target === moveModal) hideMoveModal();
});

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    if (moveModal?.classList.contains("show")) hideMoveModal();
    else if (dbConfigModal?.classList.contains("show")) hideDbModal();
    else if (editModal?.classList.contains("show")) hideModal();
    else hideSearchAutocomplete();
  }
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z" && !e.shiftKey) {
    if (["INPUT", "SELECT", "TEXTAREA"].includes(e.target.tagName)) return;
    e.preventDefault();
    performUndo();
  }
  if (e.key === "F3") {
    e.preventDefault();
    if (e.shiftKey) focusSearchMatchAt(lastSearchMatchIndex - 1);
    else focusSearchMatchAt(lastSearchMatchIndex + 1);
  }
});

// 검색창 외부 영역 클릭 시 검색 결과·자동완성 닫기
document.addEventListener("click", (e) => {
  const searchArea = searchInput?.closest(".search-wrapper");
  if (searchArea && !searchArea.contains(e.target) && !searchResultsWrap.contains(e.target)) {
    searchResultsWrap.style.display = "none";
  }
  if (searchArea && !searchArea.contains(e.target)) {
    hideSearchAutocomplete();
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
    showToast("엑셀 파일(.xlsx, .xls)만 업로드할 수 있습니다.", "warning");
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

/** 입수량·박스단위 참조 시트 여부 */
function isPackUnitSheetName(sheetName) {
  const name = String(sheetName);
  return name.includes(PACK_UNIT_SHEET_HINT) && name.includes(PACK_UNIT_BOX_HINT);
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
    if (isPackUnitSheetName(name)) return false;
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
  setGlobalLoading(true, "🟡 엑셀 파일 읽는 중...");
  const reader = new FileReader();
  reader.onerror = () => {
    setGlobalLoading(false);
    showToast("파일을 읽는 중 오류가 발생했습니다.", "error");
  };
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    try {
      rememberPristineLayoutBytes(data);
      const workbook = XLSX.read(data, { type: "array", cellStyles: true });
      applyLoadedWorkbook(workbook, undefined, true);
      showToast("엑셀 파일을 불러왔습니다.", "success");
    } catch (err) {
      alert("엑셀 파일을 파싱하는 데 실패했습니다. 파일이 손상되었거나 유효한 엑셀 형식이 아닙니다.");
      console.error(err);
    } finally {
      setGlobalLoading(false);
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
  rebuildProductCatalog();

  const warehouseSheets = getWarehouseSheetNames(workbook);
  if (warehouseSheets.length === 0) {
    alert("창고 도면 형식의 시트를 찾을 수 없습니다. 재고조사표 양식(B1, B2 등)을 확인해 주세요.");
    return;
  }

  const totalCreated = ensureTotalSheet(workbook);
  aggregateTotalSheet(workbook);
  if (totalCreated) {
    showToast("TOTAL 시트가 없어 새로 생성했습니다.", "info");
  }

  updateSheetDropdown(warehouseSheets);

  emptyGuide.style.display = "none";
  gridBoard.style.display = "grid";
  exportBtn.style.display = "inline-block";
  if (exportHtmlBtn) exportHtmlBtn.style.display = "inline-block";
  showWorkbookControls();

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
  pristineLayoutBytes = null;
  const wb = XLSX.utils.book_new();
  const sheetB2 = createMockSheet("B2");
  XLSX.utils.book_append_sheet(wb, sheetB2, "B2");
  const sheetB1 = createMockSheet("B1");
  XLSX.utils.book_append_sheet(wb, sheetB1, "B1");
  applyLoadedWorkbook(wb, resolveDefaultSheetName(wb), false);
  showToast("데모 데이터를 불러왔습니다.", "info");
}

function rememberPristineLayoutBytes(bytes) {
  if (!bytes || bytes.length < 100) {
    pristineLayoutBytes = null;
    return;
  }
  pristineLayoutBytes = bytes instanceof Uint8Array ? bytes.slice() : new Uint8Array(bytes);
}

/** 다운로드용 원본으로 쓸 만큼 랙 열이 충분한지 검사 (DB 1열 도면 거부) */
function isWorkbookLayoutHealthy(bytesOrWorkbook) {
  try {
    const wb =
      bytesOrWorkbook && bytesOrWorkbook.SheetNames
        ? bytesOrWorkbook
        : XLSX.read(bytesOrWorkbook, { type: "array", cellStyles: false });
    const names = getWarehouseSheetNames(wb);
    if (names.length === 0) return false;
    // 창고 시트 중 하나라도 랙 5열 이상이면 정상 도면으로 간주
    return names.some((name) => {
      const sheet = wb.Sheets[name];
      return sheet && hasWarehouseLayout(sheet) && detectMaxCols(sheet) >= 5;
    });
  } catch (e) {
    return false;
  }
}

/** 다운로드 직전에 건강한 원본 바이트 확보 (업로드본 → 번들 템플릿) */
async function ensurePristineLayoutForExport() {
  if (pristineLayoutBytes && isWorkbookLayoutHealthy(pristineLayoutBytes)) {
    return "cached";
  }
  try {
    const bundled = await loadBundledWorkbookBytes();
    if (!isWorkbookLayoutHealthy(bundled)) {
      throw new Error("기본 템플릿 레이아웃이 손상되었습니다.");
    }
    rememberPristineLayoutBytes(bundled);
    return "bundled";
  } catch (e) {
    throw new Error(
      "원본 도면 엑셀이 없습니다. 레이아웃이 있는 엑셀 파일을 업로드한 뒤 다시 다운로드해 주세요."
    );
  }
}

/**
 * 원본 레이아웃 워크북에 현재 재고 값만 덮어씀 (병합·열너비·색상 등 서식 유지)
 */
function overlayWarehouseInventoryValues(sourceWb, targetWb) {
  if (!sourceWb || !targetWb) return;

  getWarehouseSheetNames(sourceWb).forEach((sheetName) => {
    const src = sourceWb.Sheets[sheetName];
    const dst = targetWb.Sheets[sheetName];
    if (!src || !dst || !hasWarehouseLayout(src) || !hasWarehouseLayout(dst)) return;

    const maxCols = detectMaxCols(src);
    const maxRows = detectMaxRows(src);
    const colMapping = detectColumnMapping(src, maxCols);

    for (let r = 1; r <= maxRows; r++) {
      const prodRowIdx = 7 + 2 * (r - 1);
      const qtyRowIdx = 8 + 2 * (r - 1);

      for (let c = 1; c <= maxCols; c++) {
        const pCol = colMapping[c];
        if (pCol === undefined) continue;
        const bCol = pCol + 1;
        copyCellValueKeepStyle(src, dst, XLSX.utils.encode_cell({ r: prodRowIdx, c: pCol }));
        copyCellValueKeepStyle(src, dst, XLSX.utils.encode_cell({ r: qtyRowIdx, c: pCol }));
        copyCellValueKeepStyle(src, dst, XLSX.utils.encode_cell({ r: qtyRowIdx, c: bCol }));
      }
    }
  });
}

/** 셀 값만 복사하고 대상 셀의 기존 서식(.s)은 유지 */
function copyCellValueKeepStyle(srcSheet, dstSheet, addr) {
  const srcCell = srcSheet[addr];
  const srcVal = srcCell?.v;
  const isEmpty =
    srcCell == null ||
    srcVal === undefined ||
    srcVal === null ||
    (typeof srcVal === "string" && srcVal.trim() === "");

  if (isEmpty) {
    if (!dstSheet[addr]) return;
    const keepStyle = dstSheet[addr].s;
    dstSheet[addr].v = "";
    dstSheet[addr].t = "s";
    delete dstSheet[addr].w;
    delete dstSheet[addr].f;
    if (keepStyle) dstSheet[addr].s = keepStyle;
    return;
  }

  const type = srcCell.t || (typeof srcVal === "number" ? "n" : "s");
  if (!dstSheet[addr]) {
    dstSheet[addr] = { t: type, v: srcVal };
  } else {
    const keepStyle = dstSheet[addr].s;
    dstSheet[addr].t = type;
    dstSheet[addr].v = srcVal;
    delete dstSheet[addr].w;
    delete dstSheet[addr].f;
    if (keepStyle) dstSheet[addr].s = keepStyle;
  }
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
  hideSearchAutocomplete();

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
  let stockedRacks = 0;
  const totalRacks = totalCols * maxRows;
  const duplicateMap = buildDuplicateProductMap();

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
      const cellEl = createRackCellElement(r, colSeq, cellData, rackInfo.rackNo, zoneClass, duplicateMap);
      gridBoard.appendChild(cellEl);
      applyRackFilterToCell(cellEl, cellData);

      if (cellData.product) {
        totalProducts++;
        totalPallets += cellData.pallet;
        totalBoxes += cellData.box;
      }
      if (cellData.product || cellData.pallet > 0 || cellData.box > 0) stockedRacks++;
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
      const cellEl = createRackCellElement(r, colSeq, cellData, rackInfo.rackNo, zoneClass, duplicateMap);
      gridBoard.appendChild(cellEl);
      applyRackFilterToCell(cellEl, cellData);

      if (cellData.product) {
        totalProducts++;
        totalPallets += cellData.pallet;
        totalBoxes += cellData.box;
      }
      if (cellData.product || cellData.pallet > 0 || cellData.box > 0) stockedRacks++;
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
  if (occupancyRateEl) {
    const pct = totalRacks > 0 ? Math.round((stockedRacks / totalRacks) * 100) : 0;
    occupancyRateEl.textContent = `${stockedRacks}/${totalRacks} (${pct}%)`;
  }

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
function createRackCellElement(row, col, cellData, rackNo, zoneClass = "", duplicateMap = null) {
  const cell = document.createElement("div");
  const classes = ["rack-cell"];
  const hasStock = cellData.product || cellData.pallet > 0 || cellData.box > 0;
  if (hasStock) {
    classes.push("active-stock");
  }
  if (zoneClass) classes.push(zoneClass);
  if (duplicateMap && cellData.product && (duplicateMap.get(cellData.product.toLowerCase()) || 0) > 1) {
    classes.push("has-duplicate");
  }
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

function includesProductSearch(productName, query) {
  if (!query) return false;
  return String(productName ?? "").trim().toLowerCase().includes(query.trim().toLowerCase());
}

function matchesProductSearch(productName, query) {
  if (!query) return false;
  return String(productName ?? "").trim().toLowerCase() === query.trim().toLowerCase();
}

function rebuildProductCatalog() {
  const names = new Set();
  if (!currentWorkbook) {
    productCatalogCache = [];
    return;
  }

  const warehouseSheets = getWarehouseSheetNames(currentWorkbook);
  warehouseSheets.forEach((sheetName) => {
    const sheet = currentWorkbook.Sheets[sheetName];
    if (!sheet) return;

    const maxCols = detectMaxCols(sheet);
    const maxRows = detectMaxRows(sheet);
    const colMapping = detectColumnMapping(sheet, maxCols);

    for (let r = 1; r <= maxRows; r++) {
      const prodRowIdx = 7 + 2 * (r - 1);
      for (let c = 1; c <= maxCols; c++) {
        const pColIdx = colMapping[c];
        if (pColIdx === undefined) continue;

        const prodAddr = XLSX.utils.encode_cell({ r: prodRowIdx, c: pColIdx });
        const cellObj = sheet[prodAddr];
        if (cellObj && cellObj.v !== undefined) {
          const name = String(cellObj.v).trim();
          if (name) names.add(name);
        }
      }
    }
  });

  productCatalogCache = [...names].sort((a, b) => a.localeCompare(b, "ko"));
}

function getProductSuggestions(query, limit = 10) {
  const q = query.trim().toLowerCase();
  if (!q) return [];

  const matched = productCatalogCache.filter((name) => includesProductSearch(name, q));
  matched.sort((a, b) => {
    const aLower = a.toLowerCase();
    const bLower = b.toLowerCase();
    const rank = (value) => {
      if (value === q) return 0;
      if (value.startsWith(q)) return 1;
      return 2;
    };
    const diff = rank(aLower) - rank(bLower);
    return diff !== 0 ? diff : a.localeCompare(b, "ko");
  });

  return matched.slice(0, limit);
}

function appendAutocompleteHighlight(parent, name, query) {
  const lowerName = name.toLowerCase();
  const lowerQuery = query.trim().toLowerCase();
  const matchIndex = lowerName.indexOf(lowerQuery);

  if (matchIndex < 0) {
    parent.textContent = name;
    return;
  }

  if (matchIndex > 0) {
    parent.appendChild(document.createTextNode(name.slice(0, matchIndex)));
  }

  const mark = document.createElement("mark");
  mark.textContent = name.slice(matchIndex, matchIndex + lowerQuery.length);
  parent.appendChild(mark);

  const rest = name.slice(matchIndex + lowerQuery.length);
  if (rest) {
    parent.appendChild(document.createTextNode(rest));
  }
}

function hideSearchAutocomplete() {
  if (!searchAutocomplete) return;
  searchAutocomplete.style.display = "none";
  searchAutocomplete.innerHTML = "";
  autocompleteActiveIndex = -1;
  searchInput?.setAttribute("aria-expanded", "false");
}

function selectAutocompleteSuggestion(value) {
  searchInput.value = value;
  hideSearchAutocomplete();
  handleSearch();
  searchInput.focus();
}

function updateSearchAutocomplete() {
  if (!searchAutocomplete) return;

  const query = searchInput.value.trim();
  if (!query || !currentWorkbook) {
    hideSearchAutocomplete();
    return;
  }

  const suggestions = getProductSuggestions(query);
  searchAutocomplete.innerHTML = "";
  if (suggestions.length === 0) {
    hideSearchAutocomplete();
    return;
  }

  suggestions.forEach((name, index) => {
    const item = document.createElement("li");
    item.className = "search-autocomplete-item";
    item.role = "option";
    item.dataset.value = name;
    if (index === autocompleteActiveIndex) {
      item.classList.add("active");
    }
    appendAutocompleteHighlight(item, name, query);
    item.addEventListener("mousedown", (e) => {
      e.preventDefault();
      selectAutocompleteSuggestion(name);
    });
    searchAutocomplete.appendChild(item);
  });

  searchAutocomplete.style.display = "block";
  searchInput.setAttribute("aria-expanded", "true");
}

function handleSearchInput() {
  autocompleteActiveIndex = -1;
  updateSearchAutocomplete();
  handleSearch();
}

function handleSearchKeydown(e) {
  const items = searchAutocomplete?.querySelectorAll(".search-autocomplete-item") || [];
  const autocompleteVisible = searchAutocomplete && searchAutocomplete.style.display !== "none" && items.length > 0;

  if (autocompleteVisible) {
    if (e.key === "ArrowDown") {
      e.preventDefault();
      autocompleteActiveIndex = (autocompleteActiveIndex + 1) % items.length;
      items.forEach((el, i) => el.classList.toggle("active", i === autocompleteActiveIndex));
      items[autocompleteActiveIndex]?.scrollIntoView({ block: "nearest" });
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      autocompleteActiveIndex = autocompleteActiveIndex <= 0 ? items.length - 1 : autocompleteActiveIndex - 1;
      items.forEach((el, i) => el.classList.toggle("active", i === autocompleteActiveIndex));
      items[autocompleteActiveIndex]?.scrollIntoView({ block: "nearest" });
    } else if (e.key === "Enter" && autocompleteActiveIndex >= 0) {
      e.preventDefault();
      selectAutocompleteSuggestion(items[autocompleteActiveIndex].dataset.value);
    } else if (e.key === "Escape") {
      hideSearchAutocomplete();
    }
    return;
  }

  if (e.key === "Enter" && searchInput.value.trim()) {
    e.preventDefault();
    handleSearch();
    if (lastSearchMatches.length > 0) {
      focusSearchMatchAt(0);
    }
  } else if (e.key === "Escape") {
    hideSearchAutocomplete();
  }
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
      if (includesProductSearch(prodName, query)) {
        cell.classList.add("search-highlight");
      }
    }
  });

  // 2. 워크북이 로드되지 않았거나 검색어가 없으면 결과 영역 숨김 처리
  if (!currentWorkbook || query === "") {
    searchResultsWrap.style.display = "none";
    searchResults.innerHTML = "";
    lastSearchMatches = [];
    lastSearchMatchIndex = -1;
    updateSearchNavBar();
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
          if (includesProductSearch(prodNameVal, query)) {
            // 수량 추출 (없거나 형식이 다르면 0으로 방어)
            const pQty = (palletObj && palletObj.v !== undefined) ? (Number(palletObj.v) || 0) : 0;
            const bQty = (boxObj && boxObj.v !== undefined) ? (Number(boxObj.v) || 0) : 0;

            matches.push({
              sheetName: sheetName,
              product: prodNameVal,
              row: r,
              col: c,
              pallet: pQty,
              box: bQty,
              exactMatch: matchesProductSearch(prodNameVal, query),
            });
          }
        }
      }
    }
  });

  // 4. 실시간 검색 결과 패널 동적 렌더링
  searchResults.innerHTML = "";
  lastSearchMatches = matches;
  lastSearchMatchIndex = matches.length > 0 ? 0 : -1;
  updateSearchNavBar();

  if (matches.length > 0) {
    searchResultsWrap.style.display = "block"; // 래퍼를 보여준다 = 창고 도면이 자연스럽게 아래로 밀림

    // 4.1. 완전 일치 항목만 요약 카운터(P, B)에 반영
    const exactMatches = matches.filter((m) => m.exactMatch);
    let grandTotalPallets = 0;
    let grandTotalBoxes = 0;
    exactMatches.forEach((m) => {
      grandTotalPallets += m.pallet;
      grandTotalBoxes += m.box;
    });

    // 4.2. 검색 결과 리스트 상단에 고정 요약 헤더바 삽입
    const summaryHeader = document.createElement("div");
    summaryHeader.className = "search-summary-header";
    summaryHeader.innerHTML = `
      <span class="summary-title">🔍 '${searchInput.value}' 검색 요약</span>
      <div class="summary-values">
        <span>매칭: <strong>${exactMatches.length}</strong>곳</span>
        <span class="summary-val-p">파렛트(P): <strong>${grandTotalPallets}</strong></span>
        <span class="summary-val-b">박스(B): <strong>${grandTotalBoxes}</strong></span>
      </div>
    `;
    searchResults.appendChild(summaryHeader);
    
    // 4.3. 개별 매칭 결과 카드를 목록 창에 삽입
    matches.forEach((match) => {
      const item = document.createElement("div");
      item.className = "search-result-item" + (match.exactMatch ? " exact-match-item" : " partial-match");
      const relatedBadge = match.exactMatch ? "" : `<span class="match-related-badge">연관</span>`;
      item.innerHTML = `
        <div style="display: flex; flex-direction: column; gap: 0.2rem;">
          <span class="match-name">${match.product}${relatedBadge}</span>
          <span class="match-coord">[${match.sheetName} 구역] ${match.row}행 - ${match.col}열</span>
        </div>
        <div class="match-qty-badge">
          <span class="match-qty-p">${match.pallet}P</span>
          <span class="match-qty-b">${match.box}B</span>
        </div>
      `;
      
      item.addEventListener("click", () => {
        lastSearchMatchIndex = matches.indexOf(match);
        updateSearchNavBar();
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
    lastSearchMatches = [];
    lastSearchMatchIndex = -1;
    updateSearchNavBar();
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

  const hasData = Boolean(cellData.product || cellData.pallet || cellData.box);
  if (deleteBtn) {
    deleteBtn.disabled = !hasData;
    deleteBtn.style.display = hasData ? "" : "none";
  }
  if (moveBtn) {
    moveBtn.style.display = hasData ? "" : "none";
  }

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
async function applyCellChanges(newProduct, newPallet, newBox, options = {}) {
  if (!selectedCellInfo || !currentSheetData) return;

  if (!options.skipUndo) {
    saveUndoSnapshot();
  }

  const sheet = currentSheetData;
  const cellInfo = selectedCellInfo.data;

  updateExcelCellValue(sheet, cellInfo.prodCellAddress, newProduct, "s");
  updateExcelCellValue(sheet, cellInfo.palletCellAddress, newPallet, "n");
  updateExcelCellValue(sheet, cellInfo.boxCellAddress, newBox, "n");

  if (isOnline) {
    try {
      if (newProduct === "" && newPallet === 0 && newBox === 0) {
        await supabaseClient
          .from("rack_inventory")
          .delete()
          .match({
            sheet_name: currentSheetName,
            rack_row: selectedCellInfo.row,
            rack_col: selectedCellInfo.col,
          });
      } else {
        const { error } = await supabaseClient.from("rack_inventory").upsert(
          {
            sheet_name: currentSheetName,
            rack_row: selectedCellInfo.row,
            rack_col: selectedCellInfo.col,
            product_name: newProduct,
            pallet_qty: newPallet,
            box_qty: newBox,
            updated_at: new Date().toISOString(),
          },
          { onConflict: "sheet_name,rack_row,rack_col" }
        );
        if (error) throw error;
      }
    } catch (e) {
      showToast("서버 동기화 실패 — 로컬 변경만 저장됨", "warning");
      console.error(e);
    }
  }

  aggregateTotalSheet(currentWorkbook);
  rebuildProductCatalog();
  renderActiveSheet(currentSheetName);
  hideModal();
  if (!options.skipUndo) {
    showToast("재고를 저장했습니다.", "success");
  }
}

async function saveCellChanges() {
  await applyCellChanges(
    inputProduct.value.trim(),
    parseInt(inputPallet.value, 10) || 0,
    parseInt(inputBox.value, 10) || 0
  );
}

async function deleteCellChanges() {
  if (!selectedCellInfo) return;
  const cellData = selectedCellInfo.data || {};
  if (!cellData.product && !cellData.pallet && !cellData.box) {
    hideModal();
    return;
  }
  if (!confirm("이 셀의 재고 정보를 삭제할까요?")) return;
  await applyCellChanges("", 0, 0);
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
 * ExcelJS 셀에 SheetJS 셀 값을 넣되, 기존 스타일·병합은 건드리지 않음
 */
function setExcelJsCellValue(ws, row0, col0, srcCell) {
  const cell = ws.getCell(row0 + 1, col0 + 1);
  const val = srcCell?.v;
  const isEmpty =
    srcCell == null ||
    val === undefined ||
    val === null ||
    (typeof val === "string" && val.trim() === "");

  if (isEmpty) {
    cell.value = null;
    return;
  }
  if (srcCell.f) {
    cell.value = { formula: String(srcCell.f).replace(/^=/, "") };
    return;
  }
  if (srcCell.t === "n" || typeof val === "number") {
    cell.value = Number(val) || 0;
    return;
  }
  cell.value = String(val);
}

/**
 * 원본 xlsx 바이트를 ExcelJS로 열어 재고 값만 패치 (레이아웃 재구성 없음)
 * 셀 좌표는 원본 도면 기준, 값은 현재 워크북에서 가져온다.
 */
async function buildExportBufferPreservingLayout() {
  if (!window.ExcelJS) {
    throw new Error("ExcelJS 라이브러리를 불러오지 못했습니다. 페이지를 새로고침해 주세요.");
  }
  if (!pristineLayoutBytes || pristineLayoutBytes.length < 100) {
    throw new Error("원본 엑셀이 없습니다. 레이아웃이 있는 엑셀 파일을 다시 업로드해 주세요.");
  }

  aggregateTotalSheet(currentWorkbook);

  const layoutWb = XLSX.read(pristineLayoutBytes, { type: "array", cellStyles: true });
  const excelWb = new window.ExcelJS.Workbook();
  await excelWb.xlsx.load(pristineLayoutBytes);

  getWarehouseSheetNames(layoutWb).forEach((sheetName) => {
    const layoutSheet = layoutWb.Sheets[sheetName];
    const dataSheet = currentWorkbook.Sheets[sheetName];
    const ws = excelWb.getWorksheet(sheetName);
    if (!layoutSheet || !ws || !hasWarehouseLayout(layoutSheet)) return;
    // 현재 워크북에 해당 시트가 없으면 원본 시트 그대로 유지
    if (!dataSheet || !hasWarehouseLayout(dataSheet)) return;

    const maxCols = detectMaxCols(layoutSheet);
    const maxRows = detectMaxRows(layoutSheet);
    const layoutMapping = detectColumnMapping(layoutSheet, maxCols);
    const dataMapping = detectColumnMapping(dataSheet);

    for (let r = 1; r <= maxRows; r++) {
      const prodRowIdx = 7 + 2 * (r - 1);
      const qtyRowIdx = 8 + 2 * (r - 1);

      for (let c = 1; c <= maxCols; c++) {
        const layoutPCol = layoutMapping[c];
        if (layoutPCol === undefined) continue;
        const layoutBCol = layoutPCol + 1;

        // 논리 랙(행/열) 기준으로 현재 재고를 읽어 원본 좌표에 기록
        const cellData = getRackCellData(
          dataSheet,
          dataMapping,
          r,
          c,
          prodRowIdx,
          qtyRowIdx
        );
        setExcelJsCellValue(
          ws,
          prodRowIdx,
          layoutPCol,
          cellData.product ? { t: "s", v: cellData.product } : null
        );
        setExcelJsCellValue(
          ws,
          qtyRowIdx,
          layoutPCol,
          cellData.product || cellData.pallet
            ? { t: "n", v: cellData.pallet || 0 }
            : null
        );
        setExcelJsCellValue(
          ws,
          qtyRowIdx,
          layoutBCol,
          cellData.product || cellData.box
            ? { t: "n", v: cellData.box || 0 }
            : null
        );
      }
    }
  });

  const totalName = currentWorkbook.SheetNames.find(
    (n) => n.toUpperCase() === "TOTAL" || n.includes("합계")
  );
  const srcTotal = totalName ? currentWorkbook.Sheets[totalName] : null;
  const wsTotal = totalName ? excelWb.getWorksheet(totalName) : null;
  if (srcTotal && wsTotal) {
    const lastRow = getTotalSheetLastDataRow(srcTotal);
    for (let excelRow = TOTAL_DATA_START_ROW; excelRow <= Math.max(lastRow, TOTAL_DATA_START_ROW); excelRow++) {
      for (let col = TOTAL_COL.PRODUCT; col <= TOTAL_COL.CHECK; col++) {
        const addr = totalCellAddr(excelRow, col);
        setExcelJsCellValue(wsTotal, excelRow - 1, col, srcTotal[addr] || null);
      }
    }
  }

  return excelWb.xlsx.writeBuffer();
}

/** 열 치수가 깨진 시트인지 (숨김·초협폭 다수) */
function sheetHasBrokenColumnDims(sheet) {
  const cols = sheet?.["!cols"] || [];
  if (!cols.length) return false;
  let bad = 0;
  cols.forEach((c) => {
    if (!c) return;
    const w = c.wch != null ? c.wch : c.width;
    if (c.hidden || (w != null && w < 3)) bad += 1;
  });
  return bad >= 4;
}

/** SheetJS 원본 열 너비(wch) — 숨김/초협폭은 fallback */
function getSourceColWidth(sheet, col0, fallback = 6) {
  const col = sheet?.["!cols"]?.[col0];
  if (!col || col.hidden) return fallback;
  let w = null;
  if (col.wch != null && Number.isFinite(col.wch)) w = col.wch;
  else if (col.width != null && Number.isFinite(col.width)) w = col.width;
  if (w == null || w < 3) return fallback;
  return w;
}

/** SheetJS 원본 행 높이(pt) */
function getSourceRowHeight(sheet, row0, fallback = 15) {
  const row = sheet?.["!rows"]?.[row0];
  if (!row) return fallback;
  if (row.hpt != null && Number.isFinite(row.hpt)) return row.hpt;
  if (row.hpx != null && Number.isFinite(row.hpx)) return row.hpx * 0.75;
  return fallback;
}

/**
 * 원본 시트의 열너비·행높이를 HTML형 도면 좌표에 매핑
 */
function applyOriginalSheetDimensions(ws, sourceSheet, reportData, layout) {
  const { leftRacks, rightRacks, maxRows } = reportData;
  // 치수 원본 시트의 랙 좌표 기준으로 너비 조회
  const dimMapping = hasWarehouseLayout(sourceSheet)
    ? detectColumnMapping(sourceSheet)
    : reportData.colMapping;
  const {
    aisleCol,
    rightLabelCol,
    rackPCol,
    headRow,
    pbRow,
    prodHeadRow,
    dataStartRow,
  } = layout;

  // 정상 원본 기준 최소 너비 (깨진 캐시·인쇄용 축소값 방지)
  const MIN = { label: 4.5, p: 10, b: 6, aisle: 3 };
  const setCol = (excelCol1, width, minW) => {
    const col = ws.getColumn(excelCol1);
    col.width = Math.max(minW, Number(width) || minW);
    col.hidden = false;
  };

  const useDims = sourceSheet && !sheetHasBrokenColumnDims(sourceSheet);

  // A열(행번호) · 우측 행번호
  setCol(1, useDims ? getSourceColWidth(sourceSheet, 0, MIN.label) : MIN.label, MIN.label);
  const srcLastP =
    dimMapping[leftRacks.length + rightRacks.length] ??
    (rightRacks[rightRacks.length - 1]?.colIdx ?? leftRacks[leftRacks.length - 1]?.colIdx ?? 0);
  setCol(
    rightLabelCol,
    useDims ? getSourceColWidth(sourceSheet, srcLastP + 2, MIN.label) : MIN.label,
    MIN.label
  );

  // 랙 P/B 열
  leftRacks.forEach((_, idx) => {
    const srcP = dimMapping[idx + 1];
    const pW = useDims && srcP !== undefined ? getSourceColWidth(sourceSheet, srcP, MIN.p) : MIN.p;
    const bW =
      useDims && srcP !== undefined ? getSourceColWidth(sourceSheet, srcP + 1, MIN.b) : MIN.b;
    setCol(rackPCol("left", idx), pW, MIN.p);
    setCol(rackPCol("left", idx) + 1, bW, MIN.b);
  });
  rightRacks.forEach((_, idx) => {
    const srcP = dimMapping[leftRacks.length + idx + 1];
    const pW = useDims && srcP !== undefined ? getSourceColWidth(sourceSheet, srcP, MIN.p) : MIN.p;
    const bW =
      useDims && srcP !== undefined ? getSourceColWidth(sourceSheet, srcP + 1, MIN.b) : MIN.b;
    setCol(rackPCol("right", idx), pW, MIN.p);
    setCol(rackPCol("right", idx) + 1, bW, MIN.b);
  });

  // 통로
  const aisleCols = useDims ? [...getAisleColumnSet(sourceSheet)].sort((a, b) => a - b) : [];
  let aisleWidth = MIN.aisle;
  if (aisleCols.length) {
    aisleWidth = Math.max(
      MIN.aisle,
      aisleCols.reduce((sum, c) => sum + getSourceColWidth(sourceSheet, c, MIN.aisle), 0)
    );
  }
  setCol(aisleCol, aisleWidth, MIN.aisle);

  // 행 높이 — 원본 타이틀/헤더/데이터 행 높이 매핑
  // 원본: 3행 부제(~), 4행 제목, 5행 열번호, 6행 제품, 7행 P/B, 8행~ 데이터
  ws.getRow(1).height = getSourceRowHeight(sourceSheet, 3, 20); // 제목
  ws.getRow(2).height = getSourceRowHeight(sourceSheet, 2, 14); // 부제
  ws.getRow(3).height = getSourceRowHeight(sourceSheet, 2, 10);
  ws.getRow(headRow).height = getSourceRowHeight(sourceSheet, 4, 13); // 열번호
  ws.getRow(pbRow).height = getSourceRowHeight(sourceSheet, 6, 13); // P/B
  ws.getRow(prodHeadRow).height = getSourceRowHeight(sourceSheet, 5, 13); // 제품

  for (let r = 1; r <= maxRows; r++) {
    const srcProdRowIdx = 7 + 2 * (r - 1);
    const srcQtyRowIdx = 8 + 2 * (r - 1);
    const prodExcelRow = dataStartRow + (r - 1) * 2;
    const qtyExcelRow = prodExcelRow + 1;
    ws.getRow(prodExcelRow).height = getSourceRowHeight(sourceSheet, srcProdRowIdx, 15);
    ws.getRow(qtyExcelRow).height = getSourceRowHeight(sourceSheet, srcQtyRowIdx, 13);
  }
}

/**
 * HTML 보고서 도면 배치도와 동일한 격자를 Excel 시트로 그린다.
 * 열너비·행높이는 원본(업로드) 시트 값을 매핑한다.
 * @param dimSourceSheet 치수 참조용 원본 시트 (없으면 sourceSheet 사용)
 */
function paintWarehouseSheetLikeHtml(ws, reportData, sourceSheet, dimSourceSheet = null) {
  const dimSheet = dimSourceSheet || sourceSheet;
  const { leftRacks, rightRacks, leftCols, maxRows, colMapping, sheetName, displayTitle, headerLabel } =
    reportData;
  const navy = "1E3A5F";
  const light = "DDEBF7";
  const aisleBg = "E2E8F0";
  const white = "FFFFFF";
  const border = {
    top: { style: "thin", color: { argb: "FF000000" } },
    bottom: { style: "thin", color: { argb: "FF000000" } },
    left: { style: "thin", color: { argb: "FF000000" } },
    right: { style: "thin", color: { argb: "FF000000" } },
  };
  const fill = (hex) => ({ type: "pattern", pattern: "solid", fgColor: { argb: "FF" + hex } });
  const font = (opts) => ({ name: "맑은 고딕", size: 9, ...opts });
  const center = { vertical: "middle", horizontal: "center", wrapText: true };

  const leftStartCol = 2; // A=행번호
  const aisleCol = leftStartCol + leftRacks.length * 2;
  const rightStartCol = aisleCol + 1;
  const rightLabelCol = rightStartCol + rightRacks.length * 2;
  const lastCol = rightLabelCol;

  const rackPCol = (side, idx) =>
    side === "left" ? leftStartCol + idx * 2 : rightStartCol + idx * 2;

  const headRow = 4; // 랙 번호
  const pbRow = 5;
  const prodHeadRow = 6;
  const dataStartRow = 7;

  // 원본 열너비·행높이 적용 (고정값 대신)
  applyOriginalSheetDimensions(ws, dimSheet, reportData, {
    leftStartCol,
    aisleCol,
    rightStartCol,
    rightLabelCol,
    rackPCol,
    headRow,
    pbRow,
    prodHeadRow,
    dataStartRow,
  });

  // 제목/부제
  ws.mergeCells(1, 1, 1, lastCol);
  const titleCell = ws.getCell(1, 1);
  titleCell.value = displayTitle || `<재고조사표 - ${sheetName}>`;
  titleCell.font = font({ size: 14, bold: true, color: { argb: "FF" + navy } });
  titleCell.alignment = center;

  ws.mergeCells(2, 1, 2, lastCol);
  const subCell = ws.getCell(2, 1);
  subCell.value = `${headerLabel || sheetName} · 좌${leftCols} / 우${rightRacks.length}열 × ${maxRows}행   │   P=파렛트  B=박스`;
  subCell.font = font({ size: 8, color: { argb: "FF475569" } });
  subCell.alignment = center;

  const styleHeadNavy = (cell) => {
    cell.fill = fill(navy);
    cell.font = font({ bold: true, color: { argb: "FFFFFFFF" }, size: 9 });
    cell.alignment = center;
    cell.border = border;
  };
  const styleHeadLight = (cell) => {
    cell.fill = fill(light);
    cell.font = font({ bold: true, color: { argb: "FF000000" }, size: 8 });
    cell.alignment = center;
    cell.border = border;
  };

  // 모서리
  [headRow, pbRow, prodHeadRow].forEach((r) => {
    styleHeadNavy(ws.getCell(r, 1));
    styleHeadNavy(ws.getCell(r, rightLabelCol));
  });

  // 좌측 랙 헤더
  leftRacks.forEach((rack, idx) => {
    const c = rackPCol("left", idx);
    ws.mergeCells(headRow, c, headRow, c + 1);
    const numCell = ws.getCell(headRow, c);
    numCell.value = rack.rackNo;
    styleHeadNavy(numCell);
    styleHeadNavy(ws.getCell(headRow, c + 1));

    const pCell = ws.getCell(pbRow, c);
    const bCell = ws.getCell(pbRow, c + 1);
    pCell.value = "P";
    bCell.value = "B";
    styleHeadNavy(pCell);
    styleHeadNavy(bCell);

    ws.mergeCells(prodHeadRow, c, prodHeadRow, c + 1);
    const prodH = ws.getCell(prodHeadRow, c);
    prodH.value = "제품";
    styleHeadLight(prodH);
    styleHeadLight(ws.getCell(prodHeadRow, c + 1));
  });

  // 통로 (헤더~데이터 끝까지 병합)
  const lastDataExcelRow = dataStartRow + maxRows * 2 - 1;
  ws.mergeCells(headRow, aisleCol, lastDataExcelRow, aisleCol);
  const aisleCell = ws.getCell(headRow, aisleCol);
  aisleCell.value = "통로";
  aisleCell.fill = fill(aisleBg);
  aisleCell.font = font({ bold: true, size: 8, color: { argb: "FF64748B" } });
  aisleCell.alignment = { vertical: "middle", horizontal: "center", textRotation: 90 };
  aisleCell.border = border;

  // 우측 랙 헤더
  rightRacks.forEach((rack, idx) => {
    const c = rackPCol("right", idx);
    ws.mergeCells(headRow, c, headRow, c + 1);
    const numCell = ws.getCell(headRow, c);
    numCell.value = rack.rackNo;
    styleHeadNavy(numCell);
    styleHeadNavy(ws.getCell(headRow, c + 1));

    const pCell = ws.getCell(pbRow, c);
    const bCell = ws.getCell(pbRow, c + 1);
    pCell.value = "P";
    bCell.value = "B";
    styleHeadNavy(pCell);
    styleHeadNavy(bCell);

    ws.mergeCells(prodHeadRow, c, prodHeadRow, c + 1);
    const prodH = ws.getCell(prodHeadRow, c);
    prodH.value = "제품";
    styleHeadLight(prodH);
    styleHeadLight(ws.getCell(prodHeadRow, c + 1));
  });

  // 데이터 행 (HTML과 동일: 제품행 → P/B행)
  for (let r = 1; r <= maxRows; r++) {
    const prodExcelRow = dataStartRow + (r - 1) * 2;
    const qtyExcelRow = prodExcelRow + 1;
    const srcProdRowIdx = 7 + 2 * (r - 1);
    const srcQtyRowIdx = 8 + 2 * (r - 1);

    // 행 높이는 applyOriginalSheetDimensions에서 원본값 적용

    // HTML과 동일: 제품행=행번호, 수량행=P/B (병합하지 않음)
    const leftLab = ws.getCell(prodExcelRow, 1);
    leftLab.value = r;
    styleHeadLight(leftLab);
    const leftPb = ws.getCell(qtyExcelRow, 1);
    leftPb.value = "P/B";
    styleHeadLight(leftPb);

    const rightLab = ws.getCell(prodExcelRow, rightLabelCol);
    rightLab.value = r;
    styleHeadLight(rightLab);
    const rightPb = ws.getCell(qtyExcelRow, rightLabelCol);
    rightPb.value = "P/B";
    styleHeadLight(rightPb);

    const paintRack = (side, idx, colSeq) => {
      const c = rackPCol(side, idx);
      const cell = getRackCellData(
        sourceSheet,
        colMapping,
        r,
        colSeq,
        srcProdRowIdx,
        srcQtyRowIdx
      );
      const stocked = !!(cell.product || cell.pallet > 0 || cell.box > 0);

      ws.mergeCells(prodExcelRow, c, prodExcelRow, c + 1);
      const prodCell = ws.getCell(prodExcelRow, c);
      prodCell.value = cell.product || "";
      prodCell.font = font({ bold: stocked, size: 8 });
      prodCell.fill = fill(white);
      prodCell.alignment = center;
      prodCell.border = border;
      ws.getCell(prodExcelRow, c + 1).border = border;
      ws.getCell(prodExcelRow, c + 1).fill = fill(white);

      const pCell = ws.getCell(qtyExcelRow, c);
      const bCell = ws.getCell(qtyExcelRow, c + 1);
      pCell.value = cell.pallet || "";
      bCell.value = cell.box || "";
      [pCell, bCell].forEach((qc) => {
        qc.font = font({ size: 8 });
        qc.fill = fill(white);
        qc.alignment = center;
        qc.border = border;
      });
    };

    leftRacks.forEach((_, idx) => paintRack("left", idx, idx + 1));
    rightRacks.forEach((_, idx) => paintRack("right", idx, leftCols + idx + 1));
  }

  ws.pageSetup = {
    paperSize: 9,
    orientation: "landscape",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    margins: { left: 0.25, right: 0.25, top: 0.3, bottom: 0.3, header: 0.1, footer: 0.1 },
  };
  ws.views = [{ state: "normal", showGridLines: true }];

  // 셀 기록 후 치수 재적용 (ExcelJS가 덮어쓰는 경우 방지)
  applyOriginalSheetDimensions(ws, dimSheet, reportData, {
    leftStartCol,
    aisleCol,
    rightStartCol,
    rightLabelCol,
    rackPCol,
    headRow,
    pbRow,
    prodHeadRow,
    dataStartRow,
  });
}

/** SheetJS 시트를 ExcelJS로 값 복사 (폴백용) */
function copySheetJsToExcelJsWorksheet(sheetJs, ws) {
  if (!sheetJs || !sheetJs["!ref"]) return;
  const range = XLSX.utils.decode_range(sheetJs["!ref"]);
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const src = sheetJs[XLSX.utils.encode_cell({ r, c })];
      if (!src || (src.v === undefined && !src.f)) continue;
      const cell = ws.getCell(r + 1, c + 1);
      if (src.f) cell.value = { formula: String(src.f).replace(/^=/, "") };
      else if (src.t === "n" || typeof src.v === "number") cell.value = Number(src.v) || 0;
      else cell.value = String(src.v);
    }
  }
  (sheetJs["!merges"] || []).forEach((m) => {
    try {
      ws.mergeCells(m.s.r + 1, m.s.c + 1, m.e.r + 1, m.e.c + 1);
    } catch (_) {
      /* 중복 병합 무시 */
    }
  });
  const cols = sheetJs["!cols"] || [];
  cols.forEach((col, idx) => {
    if (!col) return;
    ws.getColumn(idx + 1).width = col.wch || col.width || 10;
    ws.getColumn(idx + 1).hidden = false;
  });
}

const EXCELJS_THIN_BLACK = {
  top: { style: "thin", color: { argb: "FF000000" } },
  bottom: { style: "thin", color: { argb: "FF000000" } },
  left: { style: "thin", color: { argb: "FF000000" } },
  right: { style: "thin", color: { argb: "FF000000" } },
};
const EXCELJS_THIN_GRAY = {
  top: { style: "thin", color: { argb: "FFBDBDBD" } },
  bottom: { style: "thin", color: { argb: "FFBDBDBD" } },
  left: { style: "thin", color: { argb: "FFBDBDBD" } },
  right: { style: "thin", color: { argb: "FFBDBDBD" } },
};

function excelJsFill(hex) {
  return { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + hex } };
}

/**
 * 입수량 박스단위 시트 — 원본(수정본)과 동일한 레이아웃·서식
 * 헤더 2행(제품명/입수량/박스단위), 데이터 3행~, 네이비 헤더·검정 테두리
 */
function paintPackUnitSheetLikeOriginal(ws, sheetJs) {
  const headerRowIdx = detectPackUnitHeaderRow(sheetJs); // 0-based
  const dataStartIdx = headerRowIdx + 1;
  const lastDataIdx = detectPackUnitLastDataRow(sheetJs, dataStartIdx);
  const headerExcelRow = headerRowIdx + 1;
  const dataStartExcelRow = dataStartIdx + 1;

  // 원본 치수 (재고조사표_수정본 기준) + SheetJS !cols 우선
  const srcCols = sheetJs["!cols"] || [];
  ws.getColumn(1).width = Math.max(14, srcCols[0]?.wch || 19);
  ws.getColumn(2).width = Math.max(8, srcCols[1]?.wch || 10);
  ws.getColumn(3).width = Math.max(10, srcCols[2]?.wch || 12);
  [1, 2, 3].forEach((c) => {
    ws.getColumn(c).hidden = false;
  });

  const headers = ["제품명", "입수량", "박스단위"];
  for (let c = 0; c < 3; c++) {
    const cell = ws.getCell(headerExcelRow, c + 1);
    const src = sheetJs[XLSX.utils.encode_cell({ r: headerRowIdx, c })];
    cell.value = src?.v != null ? String(src.v) : headers[c];
    cell.fill = excelJsFill("1E3A5F");
    cell.font = { name: "맑은 고딕", size: 9, bold: true, color: { argb: "FFFFFFFF" } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = EXCELJS_THIN_BLACK;
  }
  ws.getRow(headerExcelRow).height = getSourceRowHeight(sheetJs, headerRowIdx, 18);

  let rowNum = 0;
  for (let r = dataStartIdx; r <= lastDataIdx; r++) {
    const excelRow = r + 1;
    const odd = rowNum % 2 === 0;
    rowNum += 1;
    ws.getRow(excelRow).height = getSourceRowHeight(sheetJs, r, 16);

    const bg = odd ? "FFFFFF" : "F8FAFC";
    const prodSrc = sheetJs[XLSX.utils.encode_cell({ r, c: 0 })];
    const ipsurSrc = sheetJs[XLSX.utils.encode_cell({ r, c: 1 })];
    const boxSrc = sheetJs[XLSX.utils.encode_cell({ r, c: 2 })];

    const prodCell = ws.getCell(excelRow, 1);
    prodCell.value = prodSrc?.v != null ? String(prodSrc.v) : "";
    prodCell.fill = excelJsFill(bg);
    prodCell.font = { name: "맑은 고딕", size: 9, bold: true, color: { argb: "FF000000" } };
    prodCell.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
    prodCell.border = EXCELJS_THIN_BLACK;

    const ipsurCell = ws.getCell(excelRow, 2);
    ipsurCell.value =
      ipsurSrc?.v !== undefined && ipsurSrc?.v !== "" ? Number(ipsurSrc.v) || 0 : "";
    ipsurCell.fill = excelJsFill(bg);
    ipsurCell.font = { name: "맑은 고딕", size: 9, color: { argb: "FF000000" } };
    ipsurCell.alignment = { vertical: "middle", horizontal: "center" };
    ipsurCell.border = EXCELJS_THIN_BLACK;

    const boxCell = ws.getCell(excelRow, 3);
    boxCell.value = boxSrc?.v !== undefined && boxSrc?.v !== "" ? Number(boxSrc.v) || 0 : "";
    boxCell.fill = excelJsFill(bg);
    boxCell.font = { name: "맑은 고딕", size: 9, color: { argb: "FF000000" } };
    boxCell.alignment = { vertical: "middle", horizontal: "center" };
    boxCell.border = EXCELJS_THIN_BLACK;
  }

  // 원본에 제목 행이 있으면 유지
  if (headerRowIdx > 0) {
    const titleSrc = sheetJs[XLSX.utils.encode_cell({ r: headerRowIdx - 1, c: 0 })];
    if (titleSrc?.v) {
      try {
        ws.mergeCells(headerExcelRow - 1, 1, headerExcelRow - 1, 3);
      } catch (_) {
        /* ignore */
      }
      const titleCell = ws.getCell(headerExcelRow - 1, 1);
      titleCell.value = String(titleSrc.v);
      titleCell.font = { name: "맑은 고딕", size: 12, bold: true };
      titleCell.alignment = { vertical: "middle", horizontal: "center" };
      ws.getRow(headerExcelRow - 1).height = getSourceRowHeight(sheetJs, headerRowIdx - 1, 22);
    }
  }

  ws.views = [{ state: "normal", showGridLines: true }];
}

/**
 * TOTAL 시트 — 원본(수정본)과 동일한 레이아웃·서식·수식
 * G=((D*C)+E)*B+F , I=G+H , K=I-J
 */
function paintTotalSheetLikeOriginal(ws, sheetJs) {
  const headers = [
    "제품명",
    "입수량",
    "박스단위",
    "P",
    "B",
    "잔량(ea)",
    "재고",
    "출고",
    "합계",
    "판매일보",
    "차이",
    "창고위치",
    "창고일치",
  ];
  const colWidths = [14, 8, 8, 8, 8, 10, 12, 12, 12, 12, 10, 16, 10];
  const srcCols = sheetJs["!cols"] || [];
  colWidths.forEach((w, idx) => {
    ws.getColumn(idx + 1).width = Math.max(w, srcCols[idx]?.wch || w);
    ws.getColumn(idx + 1).hidden = false;
  });

  // 헤더 4행
  ws.getRow(TOTAL_HEADER_ROW).height = getSourceRowHeight(sheetJs, TOTAL_HEADER_ROW - 1, 22);
  for (let c = 0; c < headers.length; c++) {
    const cell = ws.getCell(TOTAL_HEADER_ROW, c + 1);
    const src = sheetJs[XLSX.utils.encode_cell({ r: TOTAL_HEADER_ROW - 1, c })];
    cell.value = src?.v != null ? String(src.v) : headers[c];
    cell.fill = excelJsFill("1E3A5F");
    cell.font = { name: "맑은 고딕", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = EXCELJS_THIN_GRAY;
  }

  const lastRow = getTotalSheetLastDataRow(sheetJs);
  for (let excelRow = TOTAL_DATA_START_ROW; excelRow <= lastRow; excelRow++) {
    const isOdd = excelRow % 2 !== 0;
    ws.getRow(excelRow).height = getSourceRowHeight(sheetJs, excelRow - 1, 15);

    const readVal = (col0) => {
      const src = sheetJs[totalCellAddr(excelRow, col0)];
      return src;
    };

    const prodSrc = readVal(TOTAL_COL.PRODUCT);
    if (!prodSrc || prodSrc.v === undefined || String(prodSrc.v).trim() === "") break;

    const applyStyle = (cell, fillHex, fontOpts, align = "center") => {
      cell.fill = excelJsFill(fillHex);
      cell.font = { name: "맑은 고딕", size: 9, ...fontOpts };
      cell.alignment = { vertical: "middle", horizontal: align, wrapText: align === "center" };
      cell.border = EXCELJS_THIN_GRAY;
    };

    // A 제품명
    const a = ws.getCell(excelRow, 1);
    a.value = String(prodSrc.v).trim();
    applyStyle(a, isOdd ? "F0FDF4" : "DCFCE7", { bold: true, size: 9, color: { argb: "FF1B4332" } });

    // B 입수량, C 박스단위
    const b = ws.getCell(excelRow, 2);
    const c = ws.getCell(excelRow, 3);
    const bSrc = readVal(TOTAL_COL.IPSUR);
    const cSrc = readVal(TOTAL_COL.BOX_UNIT);
    b.value = bSrc?.v !== undefined && bSrc?.v !== "" ? Number(bSrc.v) || 0 : "";
    c.value = cSrc?.v !== undefined && cSrc?.v !== "" ? Number(cSrc.v) || 0 : "";
    applyStyle(b, isOdd ? "FFFFFF" : "F9FAFB", { size: 9, color: { argb: "FF374151" } });
    applyStyle(c, isOdd ? "FFFFFF" : "F9FAFB", { size: 9, color: { argb: "FF374151" } });

    // D P, E B
    const d = ws.getCell(excelRow, 4);
    const e = ws.getCell(excelRow, 5);
    const dSrc = readVal(TOTAL_COL.PALLET);
    const eSrc = readVal(TOTAL_COL.BOX);
    d.value = dSrc?.v !== undefined && dSrc?.v !== "" ? Number(dSrc.v) || 0 : "";
    e.value = eSrc?.v !== undefined && eSrc?.v !== "" ? Number(eSrc.v) || 0 : "";
    applyStyle(d, isOdd ? "EFF6FF" : "DBEAFE", {
      bold: true,
      size: 10,
      color: { argb: "FF1D4ED8" },
    });
    applyStyle(e, isOdd ? "FFF1F2" : "FFE4E6", {
      bold: true,
      size: 10,
      color: { argb: "FFBE123C" },
    });

    // F 잔량
    const f = ws.getCell(excelRow, 6);
    const fSrc = readVal(TOTAL_COL.JAN);
    f.value = fSrc?.v !== undefined && fSrc?.v !== "" ? Number(fSrc.v) || 0 : "";
    applyStyle(f, isOdd ? "FFFFFF" : "F9FAFB", { size: 9, color: { argb: "FF374151" } });

    // G 재고 수식, H 출고, I 합계 수식, J 판매일보, K 차이 수식
    const g = ws.getCell(excelRow, 7);
    const h = ws.getCell(excelRow, 8);
    const i = ws.getCell(excelRow, 9);
    const j = ws.getCell(excelRow, 10);
    const k = ws.getCell(excelRow, 11);
    g.value = { formula: `((D${excelRow}*C${excelRow})+E${excelRow})*B${excelRow}+F${excelRow}` };
    const hSrc = readVal(TOTAL_COL.CHULGO);
    h.value = hSrc?.v !== undefined && hSrc?.v !== "" ? Number(hSrc.v) || 0 : "";
    i.value = { formula: `G${excelRow}+H${excelRow}` };
    const jSrc = readVal(TOTAL_COL.PANMAE);
    j.value = jSrc?.v !== undefined && jSrc?.v !== "" ? Number(jSrc.v) || 0 : "";
    k.value = { formula: `I${excelRow}-J${excelRow}` };
    [g, h, i, j, k].forEach((cell) =>
      applyStyle(cell, isOdd ? "FFFFFF" : "F9FAFB", { size: 9, color: { argb: "FF374151" } })
    );

    // L 창고위치
    const l = ws.getCell(excelRow, 12);
    const lSrc = readVal(TOTAL_COL.LOCATION);
    l.value = lSrc?.v != null ? String(lSrc.v) : "";
    applyStyle(l, isOdd ? "FAF5FF" : "F3E8FF", { size: 9, color: { argb: "FF5B21B6" } });

    // M 창고일치
    const m = ws.getCell(excelRow, 13);
    const mSrc = readVal(TOTAL_COL.CHECK);
    const checkLabel = mSrc?.v != null ? String(mSrc.v).trim() : "";
    m.value = checkLabel;
    if (checkLabel === "불일치") {
      applyStyle(m, isOdd ? "FEF2F2" : "FEE2E2", {
        bold: true,
        size: 9,
        color: { argb: "FFB91C1C" },
      });
    } else {
      applyStyle(m, isOdd ? "F0FDF4" : "DCFCE7", {
        bold: true,
        size: 9,
        color: { argb: "FF166534" },
      });
    }
  }

  ws.views = [{ state: "normal", showGridLines: true }];
}

/**
 * HTML 도면과 같은 레이아웃으로 엑셀 버퍼 생성
 */
async function buildExportBufferMatchingHtmlLayout() {
  if (!window.ExcelJS) {
    throw new Error("ExcelJS 라이브러리를 불러오지 못했습니다. 페이지를 새로고침(Ctrl+F5)해 주세요.");
  }
  if (!currentWorkbook) {
    throw new Error("내보낼 워크북이 없습니다.");
  }

  aggregateTotalSheet(currentWorkbook);
  const excelWb = new window.ExcelJS.Workbook();
  excelWb.creator = "U-PACK";
  excelWb.created = new Date();

  // 열너비·행높이 참조: 깨진 캐시보다 번들(정상 템플릿) 우선
  let dimWorkbook = null;
  try {
    const bundledBytes = await loadBundledWorkbookBytes();
    if (isWorkbookLayoutHealthy(bundledBytes)) {
      dimWorkbook = XLSX.read(bundledBytes, { type: "array", cellStyles: true });
    }
  } catch (_) {
    dimWorkbook = null;
  }
  try {
    await ensurePristineLayoutForExport();
    if (pristineLayoutBytes && isWorkbookLayoutHealthy(pristineLayoutBytes)) {
      const uploaded = XLSX.read(pristineLayoutBytes, { type: "array", cellStyles: true });
      // 업로드본 치수가 정상일 때만 사용 (숨김/초협폭 캐시 제외)
      const sample = uploaded.Sheets[uploaded.SheetNames[0]];
      if (sample && !sheetHasBrokenColumnDims(sample)) {
        dimWorkbook = uploaded;
      }
    }
  } catch (_) {
    /* keep bundled dims */
  }

  currentWorkbook.SheetNames.forEach((sheetName) => {
    const sheet = currentWorkbook.Sheets[sheetName];
    if (!sheet) return;

    if (
      hasWarehouseLayout(sheet) &&
      !isSummarySheetName(sheetName) &&
      !isPackUnitSheetName(sheetName)
    ) {
      const reportData = collectWarehouseReportData(currentWorkbook, sheetName);
      if (!reportData || reportData.maxCols < 1) return;
      const ws = excelWb.addWorksheet(sheetName);
      let dimSheet = sheet;
      const candidate = dimWorkbook?.Sheets?.[sheetName];
      if (candidate && hasWarehouseLayout(candidate) && !sheetHasBrokenColumnDims(candidate)) {
        dimSheet = candidate;
      } else if (sheetHasBrokenColumnDims(sheet) && candidate && hasWarehouseLayout(candidate)) {
        dimSheet = candidate;
      }
      paintWarehouseSheetLikeHtml(ws, reportData, sheet, dimSheet);
      return;
    }

    // 입수량 박스단위 — 원본과 동일 레이아웃·서식
    if (isPackUnitSheetName(sheetName)) {
      const ws = excelWb.addWorksheet(sheetName);
      paintPackUnitSheetLikeOriginal(ws, sheet);
      return;
    }

    // TOTAL — 원본과 동일 레이아웃·서식·수식
    if (sheetName.toUpperCase() === "TOTAL" || sheetName.includes("합계")) {
      const ws = excelWb.addWorksheet(sheetName);
      paintTotalSheetLikeOriginal(ws, sheet);
      return;
    }

    // 기타 시트 폴백
    const ws = excelWb.addWorksheet(sheetName);
    copySheetJsToExcelJsWorksheet(sheet, ws);
  });

  if (excelWb.worksheets.length === 0) {
    throw new Error("내보낼 창고 도면 시트가 없습니다.");
  }

  // 숨김 해제 + 최소 너비 강제 (Excel에서 열이 사라지지 않게)
  excelWb.eachSheet((ws) => {
    if (ws.name.toUpperCase() === "TOTAL" || ws.name.includes("합계") || isPackUnitSheetName(ws.name)) {
      return;
    }
    const maxC = Math.max(ws.actualColumnCount || 0, ws.columnCount || 0, 30);
    for (let i = 1; i <= maxC; i++) {
      const col = ws.getColumn(i);
      col.hidden = false;
      if (col.width == null || col.width < 3) {
        col.width = i === 1 ? 4.5 : 6;
      }
    }
  });

  return excelWb.xlsx.writeBuffer();
}

/**
 * 6. 수정된 파일 export (다운로드)
 * HTML 보고서 도면 배치도와 동일한 격자로 엑셀을 생성한다.
 */
async function exportUpdatedExcel() {
  if (!currentWorkbook) return;

  setGlobalLoading(true, "🟡 엑셀 생성 중...");
  try {
    const buffer = await buildExportBufferMatchingHtmlLayout();
    const bytes = new Uint8Array(buffer);

    // 검증: 창고 시트 랙 열·숨김 열
    if (window.ExcelJS) {
      const checkWb = new window.ExcelJS.Workbook();
      await checkWb.xlsx.load(bytes);
      let hiddenCount = 0;
      let maxRackHeaders = 0;
      checkWb.eachSheet((ws) => {
        if (ws.name.toUpperCase() === "TOTAL" || ws.name.includes("합계") || isPackUnitSheetName(ws.name)) {
          return;
        }
        for (let i = 1; i <= Math.max(ws.columnCount, 30); i++) {
          if (ws.getColumn(i).hidden) hiddenCount++;
        }
        // 4행 랙번호 개수
        let headers = 0;
        for (let c = 1; c <= 40; c++) {
          const v = ws.getCell(4, c).value;
          if (v !== null && v !== undefined && v !== "" && !Number.isNaN(Number(v))) headers++;
        }
        maxRackHeaders = Math.max(maxRackHeaders, headers);
      });
      let tooNarrow = 0;
      checkWb.eachSheet((ws) => {
        if (ws.name.toUpperCase() === "TOTAL" || ws.name.includes("합계") || isPackUnitSheetName(ws.name)) {
          return;
        }
        for (let i = 1; i <= Math.min(ws.columnCount || 20, 25); i++) {
          const w = ws.getColumn(i).width;
          if (w != null && w < 3) tooNarrow += 1;
        }
      });
      if (hiddenCount > 0) {
        throw new Error(`다운로드 파일에 숨김 열이 ${hiddenCount}개 있습니다. Ctrl+F5 후 다시 다운로드해 주세요.`);
      }
      if (tooNarrow >= 4) {
        throw new Error("다운로드 파일 열 너비가 비정상적으로 좁습니다. Ctrl+F5 후 다시 다운로드해 주세요.");
      }
      if (maxRackHeaders < 2) {
        throw new Error("다운로드 파일 도면 열이 부족합니다. 화면의 구역 도면을 확인한 뒤 다시 시도해 주세요.");
      }
    }

    const blob = new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const filename = `재고조사표_화면도면_${new Date().toISOString().slice(0, 10)}.xlsx`;
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    }, 0);
    showToast(`다운로드: ${filename} (열너비 정상 · 숨김열 없음)`, "success", 4000);
  } catch (err) {
    alert(err.message || "엑셀 파일을 생성하고 다운로드하는 도중 오류가 발생했습니다.");
    console.error(err);
  } finally {
    setGlobalLoading(false);
  }
}

function escapeHtml(text) {
  return String(text ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function formatReportNumber(value) {
  const num = Number(value) || 0;
  return num.toLocaleString("ko-KR");
}

/**
 * 단일 창고 시트의 사장님 보고용 데이터 수집
 */
function collectWarehouseReportData(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return null;

  const { leftRacks, rightRacks } = detectRacks(sheet);
  const leftCols = leftRacks.length;
  const rightCols = rightRacks.length;
  const maxRows = detectMaxRows(sheet);
  const colMapping = detectColumnMapping(sheet);
  const maxCols = leftCols + rightCols;
  const totalRacks = maxCols * maxRows;

  let totalProducts = 0;
  let totalPallets = 0;
  let totalBoxes = 0;
  let stockedRacks = 0;
  const entries = [];

  for (let r = 1; r <= maxRows; r++) {
    const prodRowIdx = 7 + 2 * (r - 1);
    const qtyRowIdx = 8 + 2 * (r - 1);

    leftRacks.forEach((rackInfo, idx) => {
      const colSeq = idx + 1;
      const cellData = getRackCellData(sheet, colMapping, r, colSeq, prodRowIdx, qtyRowIdx);
      if (!cellData.product && cellData.pallet <= 0 && cellData.box <= 0) return;
      if (cellData.product) {
        totalProducts++;
        totalPallets += cellData.pallet;
        totalBoxes += cellData.box;
      }
      stockedRacks++;
      entries.push({
        row: r,
        colSeq,
        rackNo: rackInfo.rackNo,
        side: "좌",
        product: cellData.product || "(미입력)",
        pallet: cellData.pallet,
        box: cellData.box,
      });
    });

    rightRacks.forEach((rackInfo, idx) => {
      const colSeq = leftCols + idx + 1;
      const cellData = getRackCellData(sheet, colMapping, r, colSeq, prodRowIdx, qtyRowIdx);
      if (!cellData.product && cellData.pallet <= 0 && cellData.box <= 0) return;
      if (cellData.product) {
        totalProducts++;
        totalPallets += cellData.pallet;
        totalBoxes += cellData.box;
      }
      stockedRacks++;
      entries.push({
        row: r,
        colSeq,
        rackNo: rackInfo.rackNo,
        side: "우",
        product: cellData.product || "(미입력)",
        pallet: cellData.pallet,
        box: cellData.box,
      });
    });
  }

  entries.sort((a, b) => a.row - b.row || a.colSeq - b.colSeq);

  const occupancyPct = totalRacks > 0 ? Math.round((stockedRacks / totalRacks) * 100) : 0;

  return {
    sheetName,
    displayTitle: parseSheetDisplayTitle(sheet, sheetName),
    headerLabel: parseSheetHeaderLabel(sheet, sheetName),
    leftCols,
    rightCols,
    maxRows,
    maxCols,
    leftRacks,
    rightRacks,
    colMapping,
    totalProducts,
    totalPallets,
    totalBoxes,
    stockedRacks,
    totalRacks,
    occupancyPct,
    entries,
  };
}

function buildWarehouseReportGridHtml(data) {
  const { sheet, leftRacks, rightRacks, leftCols, maxRows, colMapping } = data;
  const rackHeader = (racks) =>
    racks
      .map((rack) => `<th colspan="2">${escapeHtml(rack.rackNo)}</th>`)
      .join("");

  let html = `<table class="rack-grid"><thead><tr>`;
  html += `<th class="corner"></th>${rackHeader(leftRacks)}`;
  html += `<th class="aisle-head" rowspan="3">통<br>로</th>${rackHeader(rightRacks)}`;
  html += `<th class="corner"></th></tr><tr>`;
  html += `<th class="corner"></th>`;
  leftRacks.forEach(() => {
    html += `<th>P</th><th>B</th>`;
  });
  rightRacks.forEach(() => {
    html += `<th>P</th><th>B</th>`;
  });
  html += `<th class="corner"></th></tr><tr>`;
  html += `<th class="corner"></th>`;
  leftRacks.forEach(() => {
    html += `<th class="subhead">제품</th><th class="subhead"></th>`;
  });
  rightRacks.forEach(() => {
    html += `<th class="subhead">제품</th><th class="subhead"></th>`;
  });
  html += `<th class="corner"></th></tr></thead><tbody>`;

  for (let r = 1; r <= maxRows; r++) {
    const prodRowIdx = 7 + 2 * (r - 1);
    const qtyRowIdx = 8 + 2 * (r - 1);

    html += `<tr><td class="row-label">${r}</td>`;

    leftRacks.forEach((rackInfo, idx) => {
      const colSeq = idx + 1;
      const cell = getRackCellData(sheet, colMapping, r, colSeq, prodRowIdx, qtyRowIdx);
      const stocked = cell.product || cell.pallet > 0 || cell.box > 0;
      html += `<td class="prod ${stocked ? "stocked" : ""}" colspan="2">${escapeHtml(cell.product)}</td>`;
    });

    html += `<td class="aisle"></td>`;

    rightRacks.forEach((rackInfo, idx) => {
      const colSeq = leftCols + idx + 1;
      const cell = getRackCellData(sheet, colMapping, r, colSeq, prodRowIdx, qtyRowIdx);
      const stocked = cell.product || cell.pallet > 0 || cell.box > 0;
      html += `<td class="prod ${stocked ? "stocked" : ""}" colspan="2">${escapeHtml(cell.product)}</td>`;
    });

    html += `<td class="row-label">${r}</td></tr>`;

    html += `<tr><td class="row-label pb-label">P/B</td>`;
    leftRacks.forEach((_, idx) => {
      const colSeq = idx + 1;
      const cell = getRackCellData(sheet, colMapping, r, colSeq, prodRowIdx, qtyRowIdx);
      html += `<td class="qty">${cell.pallet || ""}</td><td class="qty">${cell.box || ""}</td>`;
    });
    html += `<td class="aisle"></td>`;
    rightRacks.forEach((_, idx) => {
      const colSeq = leftCols + idx + 1;
      const cell = getRackCellData(sheet, colMapping, r, colSeq, prodRowIdx, qtyRowIdx);
      html += `<td class="qty">${cell.pallet || ""}</td><td class="qty">${cell.box || ""}</td>`;
    });
    html += `<td class="row-label pb-label">P/B</td></tr>`;
  }

  html += `</tbody></table>`;
  return html;
}

function buildWarehouseReportSectionHtml(reportData, workbook) {
  const sheet = workbook.Sheets[reportData.sheetName];
  const gridHtml = buildWarehouseReportGridHtml({ ...reportData, sheet });

  const rowsHtml = reportData.entries
    .map(
      (entry, idx) => `
      <tr>
        <td class="num">${idx + 1}</td>
        <td>${entry.row}행 · ${entry.side}${entry.colSeq}열</td>
        <td>${escapeHtml(entry.rackNo)}</td>
        <td class="product">${escapeHtml(entry.product)}</td>
        <td class="num">${formatReportNumber(entry.pallet)}</td>
        <td class="num">${formatReportNumber(entry.box)}</td>
      </tr>`
    )
    .join("");

  const emptyRow =
    reportData.entries.length === 0
      ? `<tr><td colspan="6" class="empty">적재된 재고가 없습니다.</td></tr>`
      : rowsHtml;

  return `
    <section class="warehouse-section">
      <header class="section-head">
        <div>
          <p class="section-eyebrow">창고 구역</p>
          <h2>${escapeHtml(reportData.displayTitle)}</h2>
          <p class="section-sub">${escapeHtml(reportData.headerLabel)} · 좌${reportData.leftCols} / 우${reportData.rightCols}열 × ${reportData.maxRows}행</p>
        </div>
        <div class="kpi-row">
          <div class="kpi"><span>적재 품목</span><strong>${formatReportNumber(reportData.totalProducts)}</strong></div>
          <div class="kpi"><span>파렛트(P)</span><strong>${formatReportNumber(reportData.totalPallets)}</strong></div>
          <div class="kpi"><span>박스(B)</span><strong>${formatReportNumber(reportData.totalBoxes)}</strong></div>
          <div class="kpi"><span>점유율</span><strong>${reportData.stockedRacks}/${reportData.totalRacks} (${reportData.occupancyPct}%)</strong></div>
        </div>
      </header>

      <h3 class="block-title">재고 상세 목록</h3>
      <table class="detail-table">
        <thead>
          <tr>
            <th>No</th>
            <th>위치</th>
            <th>랙번호</th>
            <th>제품명</th>
            <th>P</th>
            <th>B</th>
          </tr>
        </thead>
        <tbody>${emptyRow}</tbody>
      </table>

      <h3 class="block-title">도면 배치도</h3>
      <div class="grid-wrap">${gridHtml}</div>
    </section>`;
}

function buildBossReportHtmlDocument(workbook) {
  const printDate = new Date().toLocaleString("ko-KR", {
    year: "numeric",
    month: "long",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
  const sheetNames = getWarehouseSheetNames(workbook);
  const reports = sheetNames
    .map((name) => collectWarehouseReportData(workbook, name))
    .filter(Boolean);

  const grand = reports.reduce(
    (acc, r) => {
      acc.products += r.totalProducts;
      acc.pallets += r.totalPallets;
      acc.boxes += r.totalBoxes;
      acc.stocked += r.stockedRacks;
      acc.total += r.totalRacks;
      return acc;
    },
    { products: 0, pallets: 0, boxes: 0, stocked: 0, total: 0 }
  );
  const grandPct = grand.total > 0 ? Math.round((grand.stocked / grand.total) * 100) : 0;

  const overviewRows = reports
    .map(
      (r) => `
      <tr>
        <td>${escapeHtml(r.displayTitle)}</td>
        <td class="num">${formatReportNumber(r.totalProducts)}</td>
        <td class="num">${formatReportNumber(r.totalPallets)}</td>
        <td class="num">${formatReportNumber(r.totalBoxes)}</td>
        <td class="num">${r.stockedRacks}/${r.totalRacks} (${r.occupancyPct}%)</td>
      </tr>`
    )
    .join("");

  const sectionsHtml = reports
    .map((data) => buildWarehouseReportSectionHtml(data, workbook))
    .join("");

  return `<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>재고조사 보고서 ${new Date().toISOString().slice(0, 10)}</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: "Malgun Gothic", "Apple SD Gothic Neo", sans-serif;
      color: #1e293b;
      background: #f1f5f9;
      line-height: 1.5;
    }
    .report-page {
      max-width: 1200px;
      margin: 0 auto;
      padding: 24px 20px 48px;
    }
    .report-cover {
      background: linear-gradient(135deg, #1e3a5f, #2563eb);
      color: #fff;
      border-radius: 12px;
      padding: 28px 32px;
      margin-bottom: 24px;
      box-shadow: 0 8px 24px rgba(30, 58, 95, 0.25);
    }
    .report-cover h1 { font-size: 1.75rem; margin-bottom: 6px; }
    .report-cover .meta { opacity: 0.9; font-size: 0.95rem; }
    .report-cover .grand-kpi {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 12px;
      margin-top: 20px;
    }
    .report-cover .grand-kpi div {
      background: rgba(255,255,255,0.12);
      border-radius: 8px;
      padding: 12px 14px;
    }
    .report-cover .grand-kpi span { display: block; font-size: 0.78rem; opacity: 0.85; }
    .report-cover .grand-kpi strong { font-size: 1.35rem; }
    .overview-table, .detail-table, .rack-grid {
      width: 100%;
      border-collapse: collapse;
      background: #fff;
    }
    .overview-wrap {
      background: #fff;
      border-radius: 10px;
      padding: 16px;
      margin-bottom: 28px;
      border: 1px solid #e2e8f0;
    }
    .overview-wrap h2 { font-size: 1.05rem; margin-bottom: 12px; color: #1e3a5f; }
    .overview-table th, .overview-table td,
    .detail-table th, .detail-table td {
      border: 1px solid #cbd5e1;
      padding: 8px 10px;
      font-size: 0.88rem;
    }
    .overview-table th, .detail-table th {
      background: #1e3a5f;
      color: #fff;
      font-weight: 600;
    }
    .overview-table tbody tr:nth-child(even) { background: #f8fafc; }
    .warehouse-section {
      background: #fff;
      border: 1px solid #e2e8f0;
      border-radius: 12px;
      padding: 22px 20px 28px;
      margin-bottom: 28px;
      page-break-inside: avoid;
    }
    .section-head {
      display: flex;
      flex-wrap: wrap;
      justify-content: space-between;
      gap: 16px;
      margin-bottom: 18px;
      padding-bottom: 14px;
      border-bottom: 2px solid #1e3a5f;
    }
    .section-eyebrow { font-size: 0.75rem; color: #64748b; letter-spacing: 0.04em; }
    .section-head h2 { font-size: 1.45rem; color: #1e3a5f; margin: 2px 0; }
    .section-sub { font-size: 0.88rem; color: #64748b; }
    .kpi-row {
      display: grid;
      grid-template-columns: repeat(4, minmax(110px, 1fr));
      gap: 10px;
    }
    .kpi {
      background: #ddebf7;
      border-radius: 8px;
      padding: 10px 12px;
      text-align: center;
    }
    .kpi span { display: block; font-size: 0.72rem; color: #475569; }
    .kpi strong { font-size: 1.05rem; color: #1e3a5f; }
    .block-title {
      font-size: 0.95rem;
      color: #1e3a5f;
      margin: 18px 0 10px;
    }
    .detail-table td.product { font-weight: 600; text-align: left; }
    .detail-table td.num, .overview-table td.num { text-align: right; font-variant-numeric: tabular-nums; }
    .detail-table td.empty { text-align: center; color: #94a3b8; padding: 20px; }
    .grid-wrap { overflow-x: auto; border: 1px solid #cbd5e1; border-radius: 8px; }
    .rack-grid { font-size: 0.72rem; min-width: 720px; }
    .rack-grid th, .rack-grid td {
      border: 1px solid #000;
      padding: 4px 3px;
      text-align: center;
      vertical-align: middle;
    }
    .rack-grid thead th {
      background: #1e3a5f;
      color: #fff;
      font-weight: 700;
    }
    .rack-grid .subhead { background: #1e3a5f; font-size: 0.68rem; }
    .rack-grid .row-label, .rack-grid .pb-label {
      background: #ddebf7;
      font-weight: 700;
      width: 28px;
    }
    .rack-grid .aisle-head, .rack-grid .aisle {
      background: #f1f5f9;
      color: #64748b;
      font-weight: 600;
      width: 22px;
    }
    .rack-grid .prod { background: #fff; min-width: 52px; height: 28px; }
    .rack-grid .prod.stocked { font-weight: 700; }
    .rack-grid .qty { background: #fff; font-variant-numeric: tabular-nums; }
    .report-footer {
      text-align: center;
      color: #64748b;
      font-size: 0.8rem;
      margin-top: 12px;
    }
    @media print {
      body { background: #fff; }
      .report-page { padding: 0; max-width: none; }
      .warehouse-section { page-break-before: always; box-shadow: none; }
      .warehouse-section:first-of-type { page-break-before: auto; }
    }
    @media (max-width: 768px) {
      .report-cover .grand-kpi, .kpi-row { grid-template-columns: repeat(2, 1fr); }
      .section-head { flex-direction: column; }
    }
  </style>
</head>
<body>
  <div class="report-page">
    <header class="report-cover">
      <h1>창고 재고조사 보고서</h1>
      <p class="meta">출력일시: ${escapeHtml(printDate)} · 구역 ${reports.length}개</p>
      <div class="grand-kpi">
        <div><span>전체 적재 품목</span><strong>${formatReportNumber(grand.products)}</strong></div>
        <div><span>전체 파렛트(P)</span><strong>${formatReportNumber(grand.pallets)}</strong></div>
        <div><span>전체 박스(B)</span><strong>${formatReportNumber(grand.boxes)}</strong></div>
        <div><span>전체 점유율</span><strong>${grand.stocked}/${grand.total} (${grandPct}%)</strong></div>
      </div>
    </header>

    <div class="overview-wrap">
      <h2>구역별 요약</h2>
      <table class="overview-table">
        <thead>
          <tr>
            <th>구역</th>
            <th>적재 품목</th>
            <th>파렛트(P)</th>
            <th>박스(B)</th>
            <th>점유율</th>
          </tr>
        </thead>
        <tbody>${overviewRows}</tbody>
      </table>
    </div>

    ${sectionsHtml}

    <p class="report-footer">U-PACK 창고 재고조사표 시각화 시스템 · 사장님 보고용</p>
  </div>
</body>
</html>`;
}

/**
 * 창고별 재고 현황을 사장님 보고용 HTML 파일로 내보내기
 */
function exportWarehouseHtmlReport() {
  if (!currentWorkbook) return;

  const sheetNames = getWarehouseSheetNames(currentWorkbook);
  if (sheetNames.length === 0) {
    alert("내보낼 창고 도면 시트가 없습니다.");
    return;
  }

  setGlobalLoading(true, "HTML 보고서 생성 중...");
  try {
    const html = buildBossReportHtmlDocument(currentWorkbook);
    const blob = new Blob([html], { type: "text/html;charset=utf-8" });
    const filename = `재고조사_사장님보고_${new Date().toISOString().slice(0, 10)}.html`;
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    }, 0);
    showToast(`다운로드: ${filename}`, "success", 4000);
  } catch (err) {
    alert("HTML 보고서를 생성하는 도중 오류가 발생했습니다.");
    console.error(err);
  } finally {
    setGlobalLoading(false);
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

function findPackUnitSheet(workbook) {
  if (!workbook) return null;
  const exact = workbook.SheetNames.find((n) => n === "입수량 박스단위");
  if (exact) return workbook.Sheets[exact];
  const name = workbook.SheetNames.find((n) => isPackUnitSheetName(n));
  return name ? workbook.Sheets[name] : null;
}

function parseOptionalNumber(val) {
  if (val === undefined || val === null || String(val).trim() === "") return undefined;
  const num = Number(val);
  return Number.isFinite(num) ? num : undefined;
}

/** 입수량 박스단위 시트 → 제품명별 입수량·박스단위 조회 맵 */
function readPackUnitLookup(workbook) {
  const map = new Map();
  const sheet = findPackUnitSheet(workbook);
  if (!sheet) return map;

  let dataStartRow = 3;
  const headerRowIdx = detectPackUnitHeaderRow(sheet);
  dataStartRow = headerRowIdx + 2;

  for (let excelRow = dataStartRow; excelRow <= 500; excelRow++) {
    const aCell = sheet[XLSX.utils.encode_cell({ r: excelRow - 1, c: 0 })];
    if (!aCell || aCell.v === undefined || String(aCell.v).trim() === "") break;
    const prodName = String(aCell.v).trim();
    const bCell = sheet[XLSX.utils.encode_cell({ r: excelRow - 1, c: 1 })];
    const cCell = sheet[XLSX.utils.encode_cell({ r: excelRow - 1, c: 2 })];
    map.set(prodName, {
      ipsuryang: parseOptionalNumber(bCell?.v),
      box_danyi: parseOptionalNumber(cCell?.v),
    });
  }
  return map;
}

/** TOTAL 시트가 없으면 헤더 포함 새 시트 생성 */
function ensureTotalSheet(workbook) {
  if (findTotalSheet(workbook)) return false;

  const sheet = {};
  const headers = [
    [TOTAL_COL.PRODUCT, "제품명"],
    [TOTAL_COL.IPSUR, "입수량"],
    [TOTAL_COL.BOX_UNIT, "박스단위"],
    [TOTAL_COL.PALLET, "P"],
    [TOTAL_COL.BOX, "B"],
    [TOTAL_COL.JAN, "잔량(ea)"],
    [TOTAL_COL.STOCK, "재고"],
    [TOTAL_COL.CHULGO, "출고"],
    [TOTAL_COL.SUM, "합계"],
    [TOTAL_COL.PANMAE, "판매일보"],
    [TOTAL_COL.DIFF, "차이"],
    [TOTAL_COL.LOCATION, "창고위치"],
    [TOTAL_COL.CHECK, "창고일치"],
  ];
  headers.forEach(([col, label]) => {
    writeTotalCell(sheet, TOTAL_HEADER_ROW, col, label, "s");
  });
  sheet["!ref"] = `A${TOTAL_HEADER_ROW}:${XLSX.utils.encode_col(TOTAL_COL.CHECK)}${TOTAL_DATA_START_ROW - 1}`;
  XLSX.utils.book_append_sheet(workbook, sheet, "TOTAL");
  return true;
}

function ensureTotalCheckHeader(totalSheet) {
  writeTotalCell(totalSheet, TOTAL_HEADER_ROW, TOTAL_COL.CHECK, "창고일치", "s");
}

function readTotalSheetQuantities(totalSheet) {
  const data = new Map();
  for (let excelRow = TOTAL_DATA_START_ROW; excelRow <= 300; excelRow++) {
    const aCell = totalSheet[totalCellAddr(excelRow, TOTAL_COL.PRODUCT)];
    if (!aCell || aCell.v === undefined || String(aCell.v).trim() === "") break;
    const prodName = String(aCell.v).trim();
    const readNum = (col) => {
      const cell = totalSheet[totalCellAddr(excelRow, col)];
      return cell?.v !== undefined ? Number(cell.v) || 0 : 0;
    };
    data.set(prodName, {
      pallet: readNum(TOTAL_COL.PALLET),
      box: readNum(TOTAL_COL.BOX),
    });
  }
  return data;
}

function resolveTotalCheckLabel(originalQty, warehouseQty, hadTotalSheet) {
  if (!hadTotalSheet) return "일치";
  const origP = originalQty?.pallet ?? 0;
  const origB = originalQty?.box ?? 0;
  const whP = warehouseQty?.pallet ?? 0;
  const whB = warehouseQty?.box ?? 0;
  return origP === whP && origB === whB ? "일치" : "불일치";
}

function updateTotalMatchSummary(workbook) {
  if (!totalMatchEl) return;
  const totalSheet = findTotalSheet(workbook);
  if (!totalSheet) {
    totalMatchEl.textContent = "-";
    totalMatchEl.className = "value";
    return;
  }

  let matchCount = 0;
  let mismatchCount = 0;
  for (let excelRow = TOTAL_DATA_START_ROW; excelRow <= 300; excelRow++) {
    const aCell = totalSheet[totalCellAddr(excelRow, TOTAL_COL.PRODUCT)];
    if (!aCell || aCell.v === undefined || String(aCell.v).trim() === "") break;
    const checkCell = totalSheet[totalCellAddr(excelRow, TOTAL_COL.CHECK)];
    const label = checkCell?.v !== undefined ? String(checkCell.v).trim() : "";
    if (label === "일치") matchCount += 1;
    else if (label === "불일치") mismatchCount += 1;
  }

  if (matchCount === 0 && mismatchCount === 0) {
    totalMatchEl.textContent = "-";
    totalMatchEl.className = "summary-value";
    return;
  }
  if (mismatchCount > 0) {
    totalMatchEl.textContent = `불일치 ${mismatchCount}건 / 일치 ${matchCount}건`;
    totalMatchEl.className = "summary-value total-match-mismatch";
  } else {
    totalMatchEl.textContent = `전체 일치 (${matchCount}건)`;
    totalMatchEl.className = "summary-value total-match-ok";
  }
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
  const keepStyle = totalSheet[addr]?.s;
  totalSheet[addr] = { t: type, v: val };
  if (keepStyle) totalSheet[addr].s = keepStyle;
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
    for (let col = TOTAL_COL.PRODUCT; col <= TOTAL_COL.CHECK; col++) {
      delete totalSheet[totalCellAddr(excelRow, col)];
    }
  }
}

function ensureTotalLocationHeader(totalSheet) {
  writeTotalCell(totalSheet, TOTAL_HEADER_ROW, TOTAL_COL.LOCATION, "창고위치", "s");
  ensureTotalCheckHeader(totalSheet);
}

function aggregateTotalSheet(workbook) {
  const hadTotalSheet = Boolean(findTotalSheet(workbook));
  const originalQuantities = hadTotalSheet
    ? readTotalSheetQuantities(findTotalSheet(workbook))
    : new Map();

  ensureTotalSheet(workbook);
  const totalSheet = findTotalSheet(workbook);
  if (!totalSheet) return;

  ensureTotalLocationHeader(totalSheet);

  const warehouseProducts = collectWarehouseProductData(workbook);
  const packUnitLookup = readPackUnitLookup(workbook);
  const preservedMeta = readTotalSheetRowMetadata(totalSheet);
  const oldLastRow = getTotalSheetLastDataRow(totalSheet);

  const allProductNames = new Set([
    ...warehouseProducts.keys(),
    ...originalQuantities.keys(),
  ]);
  const sortedNames = [...allProductNames].sort((a, b) =>
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
    const totals = warehouseProducts.get(prodName) || { pallet: 0, box: 0, locations: new Set() };
    const meta = preservedMeta.get(prodName) || {};
    const packInfo = packUnitLookup.get(prodName);
    const locationText = [...totals.locations]
      .sort((a, b) => a.localeCompare(b, "ko"))
      .join(", ");
    const originalQty = originalQuantities.get(prodName);
    const checkLabel = resolveTotalCheckLabel(originalQty, totals, hadTotalSheet);

    writeTotalCell(totalSheet, excelRow, TOTAL_COL.PRODUCT, prodName, "s");
    writeTotalCell(totalSheet, excelRow, TOTAL_COL.IPSUR, packInfo?.ipsuryang);
    writeTotalCell(totalSheet, excelRow, TOTAL_COL.BOX_UNIT, packInfo?.box_danyi);
    writeTotalCell(totalSheet, excelRow, TOTAL_COL.PALLET, totals.pallet);
    writeTotalCell(totalSheet, excelRow, TOTAL_COL.BOX, totals.box);
    writeTotalCell(totalSheet, excelRow, TOTAL_COL.JAN, meta.janryang);
    writeTotalCell(totalSheet, excelRow, TOTAL_COL.CHULGO, meta.chulgo);
    writeTotalCell(totalSheet, excelRow, TOTAL_COL.PANMAE, meta.panmae_ilbo);
    writeTotalCell(totalSheet, excelRow, TOTAL_COL.LOCATION, locationText, "s");
    writeTotalCell(totalSheet, excelRow, TOTAL_COL.CHECK, checkLabel, "s");

    const stockAddr = totalCellAddr(excelRow, TOTAL_COL.STOCK);
    const sumAddr = totalCellAddr(excelRow, TOTAL_COL.SUM);
    const diffAddr = totalCellAddr(excelRow, TOTAL_COL.DIFF);
    const stockStyle = totalSheet[stockAddr]?.s;
    const sumStyle = totalSheet[sumAddr]?.s;
    const diffStyle = totalSheet[diffAddr]?.s;
    totalSheet[stockAddr] = {
      t: "n",
      f: `((D${excelRow}*C${excelRow})+E${excelRow})*B${excelRow}+F${excelRow}`,
    };
    totalSheet[sumAddr] = {
      t: "n",
      f: `G${excelRow}+H${excelRow}`,
    };
    totalSheet[diffAddr] = {
      t: "n",
      f: `I${excelRow}-J${excelRow}`,
    };
    if (stockStyle) totalSheet[stockAddr].s = stockStyle;
    if (sumStyle) totalSheet[sumAddr].s = sumStyle;
    if (diffStyle) totalSheet[diffAddr].s = diffStyle;
  });

  if (totalSheet["!ref"]) {
    try {
      const range = XLSX.utils.decode_range(totalSheet["!ref"]);
      const newLastRow =
        sortedNames.length > 0
          ? TOTAL_DATA_START_ROW + sortedNames.length - 1
          : TOTAL_DATA_START_ROW - 1;
      range.e.r = Math.max(range.e.r, newLastRow - 1);
      range.e.c = Math.max(range.e.c, TOTAL_COL.CHECK);
      totalSheet["!ref"] = XLSX.utils.encode_range(range);
    } catch (e) {
      console.error("TOTAL 시트의 범위(!ref)를 갱신하는 데 실패했습니다.", e);
    }
  }

  updateTotalMatchSummary(workbook);
  console.log("✅ TOTAL 시트 집계 완료:", sortedNames.length, "품목");
}


/**
 * 창고 시트 통로(좌·우 랙 사이) 열 인덱스 집합
 */
function getAisleColumnSet(sheet) {
  const aisleCols = new Set();
  const { leftRacks, rightRacks } = detectRacks(sheet);
  if (!leftRacks.length || !rightRacks.length) return aisleCols;

  const leftEnd = leftRacks[leftRacks.length - 1].colIdx + 1;
  const rightStart = rightRacks[0].colIdx;
  for (let colIdx = leftEnd + 1; colIdx < rightStart; colIdx++) {
    aisleCols.add(colIdx);
  }
  return aisleCols;
}

/**
 * 한글·영문 혼합 텍스트의 Excel 열 너비(wch) 추정
 */
function measureDisplayWidth(text, fontSize = 8) {
  const str = String(text ?? "").trim();
  if (!str) return 0;
  let units = 0;
  for (const ch of str) {
    units += ch.charCodeAt(0) > 127 ? 2 : 1;
  }
  const sizeFactor = fontSize / 8;
  return Math.ceil(units * sizeFactor) + 1;
}

/** 랙 열에 재고가 하나라도 있는지 */
function rackColumnHasStock(sheet, colMapping, rackCol, maxRows) {
  const pCol = colMapping[rackCol];
  if (pCol === undefined) return false;
  const bCol = pCol + 1;

  for (let r = 1; r <= maxRows; r++) {
    const prodRowIdx = getWarehouseProdRowIdx(r);
    const qtyRowIdx = getWarehouseQtyRowIdx(r);
    const prodAddr = XLSX.utils.encode_cell({ r: prodRowIdx, c: pCol });
    const palletAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: pCol });
    const boxAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: bCol });

    const product = String(sheet[prodAddr]?.v ?? "").trim();
    const pallet = Number(sheet[palletAddr]?.v) || 0;
    const box = Number(sheet[boxAddr]?.v) || 0;
    if (product || pallet > 0 || box > 0) return true;
  }
  return false;
}

/** 재고가 있는 마지막 랙 번호 (없으면 1) */
function getLastStockRackIndex(sheet, colMapping, maxCols, maxRows) {
  let last = 0;
  for (let c = 1; c <= maxCols; c++) {
    if (rackColumnHasStock(sheet, colMapping, c, maxRows)) last = c;
  }
  return Math.max(last, 1);
}

/**
 * A4 가로 1장 인쇄용 열 너비 — 콘텐츠 길이 기반 자동 조절 + 빈 랙 축소
 */
function computeWarehousePrintColumnWidths(sheet, colMapping, maxCols, maxRows, aisleCols, rightLabelCol) {
  const widths = [];
  const lastStockRack = getLastStockRackIndex(sheet, colMapping, maxCols, maxRows);
  widths[0] = { wch: 3.5 };

  for (let c = 1; c <= maxCols; c++) {
    const pCol = colMapping[c];
    if (pCol === undefined) continue;
    const bCol = pCol + 1;
    const hasStock = rackColumnHasStock(sheet, colMapping, c, maxRows);
    const isTrailingEmpty = c > lastStockRack && !hasStock;

    if (isTrailingEmpty) {
      // 숨기지 않음 — 숨김/초협폭은 Excel에서 열이 사라진 것처럼 보임
      widths[pCol] = { wch: 6 };
      widths[bCol] = { wch: 5 };
      continue;
    }

    let maxProductLen = 4;
    let maxQtyLen = 2;
    for (let r = 1; r <= maxRows; r++) {
      const prodRowIdx = getWarehouseProdRowIdx(r);
      const qtyRowIdx = getWarehouseQtyRowIdx(r);
      const prodAddr = XLSX.utils.encode_cell({ r: prodRowIdx, c: pCol });
      const palletAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: pCol });
      const boxAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: bCol });

      const product = sheet[prodAddr]?.v;
      if (product !== undefined) {
        maxProductLen = Math.max(maxProductLen, measureDisplayWidth(product, 8));
      }
      const pallet = sheet[palletAddr]?.v;
      const box = sheet[boxAddr]?.v;
      if (pallet !== undefined && pallet !== "") {
        maxQtyLen = Math.max(maxQtyLen, measureDisplayWidth(pallet, 8));
      }
      if (box !== undefined && box !== "") {
        maxQtyLen = Math.max(maxQtyLen, measureDisplayWidth(box, 8));
      }
    }

    // P/B 열 동일 너비 — 제품명은 P+B 병합 셀 기준으로 쌍 너비 계산
    const pairTotalNeed = Math.max(maxProductLen, maxQtyLen + 1);
    const pairColWch = Math.min(12, Math.max(4.5, Math.ceil(pairTotalNeed / 2) + 0.5));
    widths[pCol] = { wch: pairColWch };
    widths[bCol] = { wch: pairColWch };
  }

  widths[rightLabelCol] = { wch: 3.5 };

  for (const colIdx of aisleCols) {
    widths[colIdx] = { wch: 1.8 };
  }

  // A4 가로 인쇄 가능 폭(wch)에 맞게 비율 축소
  const A4_LANDSCAPE_WCH_BUDGET = 105;
  let totalWch = 0;
  for (let i = 0; i <= rightLabelCol; i++) {
    if (widths[i] && !widths[i].hidden) totalWch += widths[i].wch || 0;
  }
  if (totalWch > A4_LANDSCAPE_WCH_BUDGET) {
    const scale = A4_LANDSCAPE_WCH_BUDGET / totalWch;
    for (let i = 0; i <= rightLabelCol; i++) {
      if (widths[i] && !widths[i].hidden) {
        widths[i] = { ...widths[i], wch: Math.max(1.8, Math.round(widths[i].wch * scale * 10) / 10) };
      }
    }
  }

  return widths;
}

/** @deprecated computeWarehousePrintColumnWidths 사용 */
function computeWarehouseColumnWidths(sheet, colMapping, maxCols, maxRows) {
  const widths = [];
  widths[0] = { wch: 4 };
  for (let c = 1; c <= maxCols; c++) {
    const pCol = colMapping[c];
    if (pCol === undefined) continue;
    let maxLen = 6;
    for (let r = 1; r <= maxRows; r++) {
      const prodAddr = XLSX.utils.encode_cell({ r: getWarehouseProdRowIdx(r), c: pCol });
      const cell = sheet[prodAddr];
      if (cell?.v !== undefined) maxLen = Math.max(maxLen, String(cell.v).trim().length);
    }
    widths[pCol] = { wch: Math.min(24, Math.max(12, maxLen + 2)) };
    widths[pCol + 1] = { wch: 7 };
  }
  return widths;
}

/** 시트에 실제 존재하는 셀 기준으로 !ref 범위를 재계산 (Excel 복구 경고 방지) */
function reconcileSheetRange(sheet) {
  if (!sheet) return;
  let maxR = 0;
  let maxC = 0;
  let hasCells = false;

  Object.keys(sheet).forEach((key) => {
    if (key[0] === "!") return;
    try {
      const { r, c } = XLSX.utils.decode_cell(key);
      maxR = Math.max(maxR, r);
      maxC = Math.max(maxC, c);
      hasCells = true;
    } catch (_) {
      /* ignore invalid keys */
    }
  });

  if (!hasCells) return;

  const range = sheet["!ref"]
    ? XLSX.utils.decode_range(sheet["!ref"])
    : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
  range.e.r = Math.max(range.e.r, maxR);
  range.e.c = Math.max(range.e.c, maxC);
  sheet["!ref"] = XLSX.utils.encode_range(range);
}

function buildDenseColumnWidths(widthMap, lastColIdx, aisleCols) {
  const cols = [];
  for (let i = 0; i <= lastColIdx; i++) {
    if (widthMap[i]) {
      const { hidden, ...rest } = widthMap[i];
      cols[i] = { ...rest, hidden: false };
      if (cols[i].wch != null) cols[i].wch = Math.max(3, cols[i].wch);
    } else {
      cols[i] = aisleCols.has(i) ? { wch: 3 } : { wch: 5 };
    }
  }
  return cols;
}

function ensureStyledCell(sheet, addr, fallback = { t: "s", v: "" }) {
  if (!sheet[addr]) {
    sheet[addr] = { ...fallback };
  }
  return sheet[addr];
}

const WAREHOUSE_LAYOUT = {
  SUBTITLE_ROW: 2,
  TITLE_ROW: 3,
  COL_NUM_ROW: 4,
  PRODUCT_LABEL_ROW: 5,
  PB_HEADER_ROW: 6,
  DATA_START_ROW: 7,
};

const WAREHOUSE_PRINT = {
  FONT: "맑은 고딕",
  FONT_TITLE: 11,
  FONT_HEADER: 8,
  FONT_DATA: 8,
  FONT_PRODUCT: 8,
  ROW_SUBTITLE: 14,
  ROW_TITLE: 20,
  ROW_HEADER: 13,
  ROW_PRODUCT: 15,
  ROW_QTY: 13,
};

const WAREHOUSE_THEME = {
  NAVY: "1E3A5F",
  LIGHT: "DDEBF7",
  WHITE: "FFFFFF",
  BLACK: "000000",
};

function getWarehouseProdRowIdx(logicalRow) {
  return WAREHOUSE_LAYOUT.DATA_START_ROW + 2 * (logicalRow - 1);
}

function getWarehouseQtyRowIdx(logicalRow) {
  return getWarehouseProdRowIdx(logicalRow) + 1;
}

function getWarehouseLastDataRow(maxRows) {
  return getWarehouseQtyRowIdx(maxRows);
}

/**
 * 창고 시트 export 전용 셀 병합 — 값은 좌상단 셀에 유지 (re-import 호환)
 */
function applyWarehouseMerges(sheet, { maxCols, maxRows, colMapping, aisleCols, rightLabelCol }) {
  const merges = [];

  for (let c = 1; c <= maxCols; c++) {
    const pCol = colMapping[c];
    if (pCol === undefined) continue;
    const bCol = pCol + 1;
    merges.push(
      { s: { r: WAREHOUSE_LAYOUT.COL_NUM_ROW, c: pCol }, e: { r: WAREHOUSE_LAYOUT.COL_NUM_ROW, c: bCol } },
      { s: { r: WAREHOUSE_LAYOUT.PRODUCT_LABEL_ROW, c: pCol }, e: { r: WAREHOUSE_LAYOUT.PRODUCT_LABEL_ROW, c: bCol } }
    );
  }

  for (let r = 1; r <= maxRows; r++) {
    const prodRowIdx = getWarehouseProdRowIdx(r);
    const qtyRowIdx = getWarehouseQtyRowIdx(r);

    // 제품명 행: P+B 가로 병합 (참고 이미지와 동일)
    for (let c = 1; c <= maxCols; c++) {
      const pCol = colMapping[c];
      if (pCol === undefined) continue;
      merges.push({ s: { r: prodRowIdx, c: pCol }, e: { r: prodRowIdx, c: pCol + 1 } });
    }

    const leftLabelAddr = XLSX.utils.encode_cell({ r: prodRowIdx, c: 0 });
    if (sheet[leftLabelAddr]?.v !== undefined) {
      merges.push({ s: { r: prodRowIdx, c: 0 }, e: { r: qtyRowIdx, c: 0 } });
    }

    const rightLabelAddr = XLSX.utils.encode_cell({ r: prodRowIdx, c: rightLabelCol });
    if (sheet[rightLabelAddr]?.v !== undefined) {
      merges.push({ s: { r: prodRowIdx, c: rightLabelCol }, e: { r: qtyRowIdx, c: rightLabelCol } });
    }
  }

  if (aisleCols.size > 0) {
    const aisleArr = [...aisleCols].sort((a, b) => a - b);
    const firstAisle = aisleArr[0];
    const lastAisle = aisleArr[aisleArr.length - 1];
    const lastDataRow = getWarehouseLastDataRow(maxRows);
    const aisleAddr = XLSX.utils.encode_cell({ r: WAREHOUSE_LAYOUT.COL_NUM_ROW, c: firstAisle });
    ensureStyledCell(sheet, aisleAddr, { t: "s", v: "통로" });
    merges.push({
      s: { r: WAREHOUSE_LAYOUT.COL_NUM_ROW, c: firstAisle },
      e: { r: lastDataRow, c: lastAisle },
    });
  }

  const legendEndCol = Math.min(rightLabelCol, 12);
  merges.push(
    { s: { r: WAREHOUSE_LAYOUT.SUBTITLE_ROW, c: 0 }, e: { r: WAREHOUSE_LAYOUT.SUBTITLE_ROW, c: legendEndCol } },
    { s: { r: WAREHOUSE_LAYOUT.TITLE_ROW, c: 0 }, e: { r: WAREHOUSE_LAYOUT.TITLE_ROW, c: legendEndCol } }
  );

  sheet["!merges"] = merges;
}

function applyWarehouseTitleBlock(sheet, sheetName, styles) {
  const titleAddr = XLSX.utils.encode_cell({ r: WAREHOUSE_LAYOUT.TITLE_ROW, c: 0 });
  const existingTitle = sheet["C4"]?.v || sheet[titleAddr]?.v;
  const titleText = existingTitle
    ? String(existingTitle).trim()
    : `<재고조사표 - ${sheetName}>`;
  const printDate = new Date().toISOString().slice(0, 10);

  sheet[titleAddr] = { t: "s", v: titleText };
  sheet[titleAddr].s = styles.sTitle;

  const subAddr = XLSX.utils.encode_cell({ r: WAREHOUSE_LAYOUT.SUBTITLE_ROW, c: 0 });
  sheet[subAddr] = { t: "s", v: `P = 파렛트(파란)   B = 박스(빨강)   ■ = 재고 있음   │   출력일: ${printDate}` };
  sheet[subAddr].s = styles.sSubtitle;
}

/**
 * A4 가로 1장 인쇄 설정
 */
function applyWarehousePrintSetup(sheet, rightLabelCol, lastDataRow) {
  sheet["!margins"] = {
    left: 0.2,
    right: 0.2,
    top: 0.3,
    bottom: 0.3,
    header: 0.15,
    footer: 0.15,
  };
  sheet["!pageSetup"] = {
    paperSize: 9,
    orientation: "landscape",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    scale: 100,
    horizontalDpi: 4294967295,
    verticalDpi: 4294967295,
  };
}

/** 입수량·박스단위 시트 헤더 행(0-based) 감지 */
function detectPackUnitHeaderRow(sheet) {
  for (let excelRow = 1; excelRow <= 4; excelRow++) {
    const aCell = sheet[XLSX.utils.encode_cell({ r: excelRow - 1, c: 0 })];
    const bCell = sheet[XLSX.utils.encode_cell({ r: excelRow - 1, c: 1 })];
    const aVal = String(aCell?.v ?? "").trim();
    const bVal = String(bCell?.v ?? "").trim();
    if (aVal.includes("제품") || bVal.includes("입수")) {
      return excelRow - 1;
    }
  }
  return 1;
}

/** 입수량·박스단위 시트 마지막 데이터 행(0-based) */
function detectPackUnitLastDataRow(sheet, dataStartRowIdx) {
  let last = dataStartRowIdx;
  for (let r = dataStartRowIdx; r <= 499; r++) {
    const aCell = sheet[XLSX.utils.encode_cell({ r, c: 0 })];
    if (!aCell?.v || String(aCell.v).trim() === "") break;
    last = r;
  }
  return last;
}

/**
 * 입수량·박스단위 시트 전용 인쇄용 서식
 */
function applyPackUnitSheetStyles(sheet) {
  const fontName = WAREHOUSE_PRINT.FONT;
  const headerRowIdx = detectPackUnitHeaderRow(sheet);
  const titleRowIdx = headerRowIdx > 0 ? headerRowIdx - 1 : 0;
  const dataStartRowIdx = headerRowIdx + 1;
  const lastDataRowIdx = detectPackUnitLastDataRow(sheet, dataStartRowIdx);

  const border = {
    top: { style: "thin", color: { rgb: "000000" } },
    bottom: { style: "thin", color: { rgb: "000000" } },
    left: { style: "thin", color: { rgb: "000000" } },
    right: { style: "thin", color: { rgb: "000000" } },
  };

  const sTitle = {
    font: { name: fontName, sz: 12, bold: true, color: { rgb: "000000" } },
    fill: { patternType: "solid", fgColor: { rgb: "FFFFFF" } },
    alignment: { vertical: "center", horizontal: "center" },
  };
  const sHead = {
    font: { name: fontName, sz: 9, bold: true, color: { rgb: "FFFFFF" } },
    fill: { patternType: "solid", fgColor: { rgb: "1E3A5F" } },
    alignment: { vertical: "center", horizontal: "center", wrapText: true },
    border,
  };
  const sProduct = (odd) => ({
    font: { name: fontName, sz: 9, bold: true, color: { rgb: "000000" } },
    fill: { patternType: "solid", fgColor: { rgb: odd ? "FFFFFF" : "F8FAFC" } },
    alignment: { vertical: "center", horizontal: "left", indent: 1 },
    border,
  });
  const sNum = (odd) => ({
    font: { name: fontName, sz: 9, color: { rgb: "000000" } },
    fill: { patternType: "solid", fgColor: { rgb: odd ? "FFFFFF" : "F8FAFC" } },
    alignment: { vertical: "center", horizontal: "center" },
    border,
  });

  let maxProductW = measureDisplayWidth("제품", 9);
  let maxIpsurW = measureDisplayWidth("입수량", 9);
  let maxBoxW = measureDisplayWidth("박스단위", 9);

  for (let r = dataStartRowIdx; r <= lastDataRowIdx; r++) {
    const aVal = sheet[XLSX.utils.encode_cell({ r, c: 0 })]?.v;
    const bVal = sheet[XLSX.utils.encode_cell({ r, c: 1 })]?.v;
    const cVal = sheet[XLSX.utils.encode_cell({ r, c: 2 })]?.v;
    maxProductW = Math.max(maxProductW, measureDisplayWidth(aVal, 9));
    maxIpsurW = Math.max(maxIpsurW, measureDisplayWidth(bVal, 9));
    maxBoxW = Math.max(maxBoxW, measureDisplayWidth(cVal, 9));
  }

  sheet["!cols"] = [
    { wch: Math.min(24, Math.max(10, maxProductW + 1)) },
    { wch: Math.max(8, maxIpsurW + 2) },
    { wch: Math.max(10, maxBoxW + 2) },
  ];

  if (!sheet["!rows"]) sheet["!rows"] = [];
  sheet["!rows"][titleRowIdx] = { hpt: 22 };
  sheet["!rows"][headerRowIdx] = { hpt: 18 };
  for (let r = dataStartRowIdx; r <= lastDataRowIdx; r++) {
    sheet["!rows"][r] = { hpt: 16 };
  }

  const titleAddr = XLSX.utils.encode_cell({ r: titleRowIdx, c: 0 });
  const titleText = sheet[titleAddr]?.v
    ? String(sheet[titleAddr].v).trim()
    : "입수량 · 박스단위 기준표";
  sheet[titleAddr] = { t: "s", v: titleText };
  sheet[titleAddr].s = sTitle;

  for (let c = 0; c <= 2; c++) {
    const addr = XLSX.utils.encode_cell({ r: headerRowIdx, c });
    ensureStyledCell(sheet, addr, { t: "s", v: ["제품", "입수량", "박스단위"][c] });
    sheet[addr].s = sHead;
  }

  let rowNum = 0;
  for (let r = dataStartRowIdx; r <= lastDataRowIdx; r++) {
    const odd = rowNum % 2 === 0;
    rowNum++;

    const prodAddr = XLSX.utils.encode_cell({ r, c: 0 });
    const ipsurAddr = XLSX.utils.encode_cell({ r, c: 1 });
    const boxAddr = XLSX.utils.encode_cell({ r, c: 2 });

    if (sheet[prodAddr]) sheet[prodAddr].s = sProduct(odd);
    if (sheet[ipsurAddr]) {
      sheet[ipsurAddr].s = sNum(odd);
      if (sheet[ipsurAddr].v !== undefined && sheet[ipsurAddr].v !== "") {
        sheet[ipsurAddr].t = "n";
        sheet[ipsurAddr].v = Number(sheet[ipsurAddr].v) || sheet[ipsurAddr].v;
      }
    }
    if (sheet[boxAddr]) {
      sheet[boxAddr].s = sNum(odd);
      if (sheet[boxAddr].v !== undefined && sheet[boxAddr].v !== "") {
        sheet[boxAddr].t = "n";
        sheet[boxAddr].v = Number(sheet[boxAddr].v) || sheet[boxAddr].v;
      }
    }
  }

  sheet["!merges"] = [{ s: { r: titleRowIdx, c: 0 }, e: { r: titleRowIdx, c: 2 } }];
  sheet["!margins"] = {
    left: 0.4,
    right: 0.4,
    top: 0.5,
    bottom: 0.5,
    header: 0.2,
    footer: 0.2,
  };
  sheet["!pageSetup"] = {
    paperSize: 9,
    orientation: "portrait",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    scale: 100,
    horizontalDpi: 4294967295,
    verticalDpi: 4294967295,
  };

  reconcileSheetRange(sheet);
}

/**
 * xlsx-js-style 라이브러리를 활용하여 다운로드 엑셀 파일에 격자 및 색상 서식을 입힙니다.
 */
function applyExcelStyles(workbook) {
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;

    if (sheetName.toUpperCase() === "TOTAL" || sheetName.includes("합계")) {
      applyTotalSheetStyles(sheet);
      continue;
    }

    if (isPackUnitSheetName(sheetName)) {
      applyPackUnitSheetStyles(sheet);
      continue;
    }

    if (!hasWarehouseLayout(sheet)) {
      continue;
    }

    const maxCols = detectMaxCols(sheet);
    const maxRows = detectMaxRows(sheet);
    const colMapping = detectColumnMapping(sheet, maxCols);
    const aisleCols = getAisleColumnSet(sheet);

    const lastColIdx = colMapping[maxCols] || (2 * maxCols + 1);
    // lastColIdx = 마지막 랙 P열, +1 = B열 → 우측 행번호는 B열 다음(+2)이어야 병합 충돌 없음
    const rightLabelCol = lastColIdx + 2;

    const colWidthMap = computeWarehousePrintColumnWidths(
      sheet, colMapping, maxCols, maxRows, aisleCols, rightLabelCol
    );
    sheet["!cols"] = buildDenseColumnWidths(colWidthMap, rightLabelCol, aisleCols);

    if (!sheet["!rows"]) sheet["!rows"] = [];
    sheet["!rows"][WAREHOUSE_LAYOUT.SUBTITLE_ROW] = { hpt: WAREHOUSE_PRINT.ROW_SUBTITLE };
    sheet["!rows"][WAREHOUSE_LAYOUT.TITLE_ROW] = { hpt: WAREHOUSE_PRINT.ROW_TITLE };
    sheet["!rows"][WAREHOUSE_LAYOUT.COL_NUM_ROW] = { hpt: WAREHOUSE_PRINT.ROW_HEADER };
    sheet["!rows"][WAREHOUSE_LAYOUT.PRODUCT_LABEL_ROW] = { hpt: WAREHOUSE_PRINT.ROW_HEADER };
    sheet["!rows"][WAREHOUSE_LAYOUT.PB_HEADER_ROW] = { hpt: WAREHOUSE_PRINT.ROW_HEADER };

    const fontName = WAREHOUSE_PRINT.FONT;
    const borderPrint = {
      top: { style: "thin", color: { rgb: WAREHOUSE_THEME.BLACK } },
      bottom: { style: "thin", color: { rgb: WAREHOUSE_THEME.BLACK } },
      left: { style: "thin", color: { rgb: WAREHOUSE_THEME.BLACK } },
      right: { style: "thin", color: { rgb: WAREHOUSE_THEME.BLACK } },
    };

    const sSubtitle = {
      font: { name: fontName, sz: 7, color: { rgb: "475569" } },
      fill: { patternType: "solid", fgColor: { rgb: WAREHOUSE_THEME.WHITE } },
      alignment: { vertical: "center", horizontal: "center" },
    };
    const sTitle = {
      font: { name: fontName, sz: WAREHOUSE_PRINT.FONT_TITLE, bold: true, color: { rgb: WAREHOUSE_THEME.BLACK } },
      fill: { patternType: "solid", fgColor: { rgb: WAREHOUSE_THEME.WHITE } },
      alignment: { vertical: "center", horizontal: "center" },
    };
    const sHeaderNavy = {
      font: { name: fontName, sz: WAREHOUSE_PRINT.FONT_HEADER, bold: true, color: { rgb: WAREHOUSE_THEME.WHITE } },
      fill: { patternType: "solid", fgColor: { rgb: WAREHOUSE_THEME.NAVY } },
      alignment: { vertical: "center", horizontal: "center" },
      border: borderPrint,
    };
    const sHeaderLight = {
      font: { name: fontName, sz: WAREHOUSE_PRINT.FONT_HEADER, bold: true, color: { rgb: WAREHOUSE_THEME.BLACK } },
      fill: { patternType: "solid", fgColor: { rgb: WAREHOUSE_THEME.LIGHT } },
      alignment: { vertical: "center", horizontal: "center" },
      border: borderPrint,
    };
    const sProductStock = {
      font: { name: fontName, sz: WAREHOUSE_PRINT.FONT_PRODUCT, bold: true, color: { rgb: WAREHOUSE_THEME.BLACK } },
      fill: { patternType: "solid", fgColor: { rgb: WAREHOUSE_THEME.WHITE } },
      alignment: { vertical: "center", horizontal: "center", wrapText: false, shrinkToFit: true },
      border: borderPrint,
    };
    const sQtyStock = {
      font: { name: fontName, sz: WAREHOUSE_PRINT.FONT_DATA, color: { rgb: WAREHOUSE_THEME.BLACK } },
      fill: { patternType: "solid", fgColor: { rgb: WAREHOUSE_THEME.WHITE } },
      alignment: { vertical: "center", horizontal: "center" },
      border: borderPrint,
    };
    const sEmptyGrid = {
      font: { name: fontName, sz: 7, color: { rgb: WAREHOUSE_THEME.BLACK } },
      fill: { patternType: "solid", fgColor: { rgb: WAREHOUSE_THEME.WHITE } },
      border: borderPrint,
    };
    const sAisle = {
      font: { name: fontName, sz: 7, bold: true, color: { rgb: "64748B" } },
      fill: { patternType: "solid", fgColor: { rgb: WAREHOUSE_THEME.LIGHT } },
      alignment: { vertical: "center", horizontal: "center", textRotation: 90 },
      border: borderPrint,
    };

    applyWarehouseTitleBlock(sheet, sheetName, { sTitle, sSubtitle });

    const colNumCorner = XLSX.utils.encode_cell({ r: WAREHOUSE_LAYOUT.COL_NUM_ROW, c: 0 });
    const prodLabelCorner = XLSX.utils.encode_cell({ r: WAREHOUSE_LAYOUT.PRODUCT_LABEL_ROW, c: 0 });
    const pbHeaderCorner = XLSX.utils.encode_cell({ r: WAREHOUSE_LAYOUT.PB_HEADER_ROW, c: 0 });
    ensureStyledCell(sheet, colNumCorner, { t: "s", v: "" });
    ensureStyledCell(sheet, prodLabelCorner, { t: "s", v: "" });
    ensureStyledCell(sheet, pbHeaderCorner, { t: "s", v: "" });
    sheet[colNumCorner].s = sHeaderNavy;
    sheet[prodLabelCorner].s = sHeaderLight;
    sheet[pbHeaderCorner].s = sHeaderNavy;

    const lastDataRow = getWarehouseLastDataRow(maxRows);
    for (let colIdx of aisleCols) {
      for (let rowIdx = WAREHOUSE_LAYOUT.COL_NUM_ROW; rowIdx <= lastDataRow; rowIdx++) {
        const addr = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
        ensureStyledCell(sheet, addr, { t: "s", v: "" });
        sheet[addr].s = sAisle;
      }
    }

    for (let r = 1; r <= maxRows; r++) {
      const prodRowIdx = getWarehouseProdRowIdx(r);
      const qtyRowIdx = getWarehouseQtyRowIdx(r);

      if (!sheet["!rows"][prodRowIdx]) sheet["!rows"][prodRowIdx] = { hpt: WAREHOUSE_PRINT.ROW_PRODUCT };
      if (!sheet["!rows"][qtyRowIdx]) sheet["!rows"][qtyRowIdx] = { hpt: WAREHOUSE_PRINT.ROW_QTY };

      const leftLabelAddr = XLSX.utils.encode_cell({ r: prodRowIdx, c: 0 });
      if (sheet[leftLabelAddr]) sheet[leftLabelAddr].s = sHeaderLight;

      const rightLabelAddr = XLSX.utils.encode_cell({ r: prodRowIdx, c: rightLabelCol });
      if (sheet[rightLabelAddr]) sheet[rightLabelAddr].s = sHeaderLight;

      for (let c = 1; c <= maxCols; c++) {
        const pCol = colMapping[c];
        if (pCol === undefined) continue;
        const bCol = pCol + 1;

        const prodAddr = XLSX.utils.encode_cell({ r: prodRowIdx, c: pCol });
        const prodBAddr = XLSX.utils.encode_cell({ r: prodRowIdx, c: bCol });
        const palletAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: pCol });
        const boxAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: bCol });

        ensureStyledCell(sheet, prodAddr, { t: "s", v: "" });
        ensureStyledCell(sheet, prodBAddr, { t: "s", v: "" });
        ensureStyledCell(sheet, palletAddr, { t: "n", v: 0 });
        ensureStyledCell(sheet, boxAddr, { t: "n", v: 0 });

        const productVal = String(sheet[prodAddr].v || "").trim();
        const palletVal = sheet[palletAddr].v;
        const boxVal = sheet[boxAddr].v;
        const hasStock = productVal !== "" || (Number(palletVal) || 0) > 0 || (Number(boxVal) || 0) > 0;

        sheet[prodAddr].s = hasStock ? sProductStock : sEmptyGrid;
        sheet[palletAddr].s = hasStock ? sQtyStock : sEmptyGrid;
        sheet[boxAddr].s = hasStock ? sQtyStock : sEmptyGrid;
        sheet[prodBAddr].s = sEmptyGrid;

        if (hasStock) {
          if (sheet[palletAddr].v !== undefined && sheet[palletAddr].v !== "") {
            sheet[palletAddr].t = "n";
            sheet[palletAddr].v = Number(sheet[palletAddr].v) || 0;
          }
          if (sheet[boxAddr].v !== undefined && sheet[boxAddr].v !== "") {
            sheet[boxAddr].t = "n";
            sheet[boxAddr].v = Number(sheet[boxAddr].v) || 0;
          }
        }
      }
    }

    for (let c = 1; c <= maxCols; c++) {
      const pCol = colMapping[c];
      if (pCol === undefined) continue;
      const bCol = pCol + 1;

      const colNumAddr = XLSX.utils.encode_cell({ r: WAREHOUSE_LAYOUT.COL_NUM_ROW, c: pCol });
      if (sheet[colNumAddr]) sheet[colNumAddr].s = sHeaderNavy;

      const prodLabelAddr = XLSX.utils.encode_cell({ r: WAREHOUSE_LAYOUT.PRODUCT_LABEL_ROW, c: pCol });
      if (sheet[prodLabelAddr]) sheet[prodLabelAddr].s = sHeaderLight;

      const pAddr = XLSX.utils.encode_cell({ r: WAREHOUSE_LAYOUT.PB_HEADER_ROW, c: pCol });
      if (sheet[pAddr]) sheet[pAddr].s = sHeaderNavy;
      const bAddr = XLSX.utils.encode_cell({ r: WAREHOUSE_LAYOUT.PB_HEADER_ROW, c: bCol });
      if (sheet[bAddr]) sheet[bAddr].s = sHeaderNavy;
    }

    applyWarehouseMerges(sheet, { maxCols, maxRows, colMapping, aisleCols, rightLabelCol });
    applyWarehousePrintSetup(sheet, rightLabelCol, lastDataRow);
    reconcileSheetRange(sheet);
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
  // A:제품명, B:입수량, C:박스단위, D:P, E:B, F:잔량, G:재고, H:출고, I:합계, J:판매일보, K:차이, L:창고위치, M:창고일치
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
    { wch: 10 }, // M: 창고일치
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

  const sCheckMatch = (isOdd) => ({
    font: { name: fontName, sz: 9, bold: true, color: { rgb: "166534" } },
    fill: { patternType: "solid", fgColor: { rgb: isOdd ? "F0FDF4" : "DCFCE7" } },
    alignment: { vertical: "center", horizontal: "center" },
    border
  });

  const sCheckMismatch = (isOdd) => ({
    font: { name: fontName, sz: 9, bold: true, color: { rgb: "B91C1C" } },
    fill: { patternType: "solid", fgColor: { rgb: isOdd ? "FEF2F2" : "FEE2E2" } },
    alignment: { vertical: "center", horizontal: "center" },
    border
  });

  // ── 4행 헤더 스타일 적용 (A~M 열) ──────────
  for (let col = 0; col <= TOTAL_COL.CHECK; col++) {
    const addr = XLSX.utils.encode_cell({ r: 3, c: col }); // 4행 = r:3
    if (!sheet[addr]) continue;
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
      null,             // M: 창고일치 (col 12) — 값에 따라 아래에서 개별 적용
    ];

    colStyles.forEach((style, colIdx) => {
      if (colIdx === TOTAL_COL.CHECK || !style) return;
      const addr = XLSX.utils.encode_cell({ r: excelRow - 1, c: colIdx });

      if (sheet[addr]) {
        sheet[addr].s = style;
      }
    });

    const checkAddr = XLSX.utils.encode_cell({ r: excelRow - 1, c: TOTAL_COL.CHECK });
    const checkCell = sheet[checkAddr];
    if (checkCell) {
      const label = String(checkCell.v || "").trim();
      checkCell.s = label === "불일치" ? sCheckMismatch(isOdd) : sCheckMatch(isOdd);
    }
  }

  // ── 행 높이 설정 ─────────────────────────────────────────
  if (!sheet["!rows"]) sheet["!rows"] = [];
  sheet["!rows"][3] = { hpt: 22 }; // 헤더 행(4행) 높이
  reconcileSheetRange(sheet);
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
      clearSyncFailed();
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
      // 깨진 DB도면이 캐시/스냅샷에 들어있으면 원본으로 쓰지 않음
      if (isWorkbookLayoutHealthy(workbookBytes)) {
        rememberPristineLayoutBytes(workbookBytes);
      } else {
        console.warn("복원 엑셀 레이아웃이 불완전하여 원본 레이아웃으로 사용하지 않습니다.");
        workbookBytes = null;
      }
    }

    if (!workbookBytes) {
      // 깨진 캐시/아카이브 대신 번들 템플릿(B1·B2·B3·W2 전체 도면) 우선
      try {
        workbookBytes = await loadBundledWorkbookBytes();
        rememberPristineLayoutBytes(workbookBytes);
        templateSource = "bundled";
      } catch (e) {
        if (inventoryRacks.length > 0) {
          currentWorkbook = buildWorkbookFromDbInventory(dbRacks, layoutRows);
          templateSource = layoutRows.length > 0 ? "db-layout" : "db-generated";
        } else {
          throw new Error("엑셀 템플릿 파일(재고조사표.xlsx)을 불러오는 데 실패했습니다 (네트워크 오류).");
        }
      }
    }

    if (loadToken !== supabaseLoadToken || isUserLocalWorkbook || isDbImportInProgress) return;

    if (workbookBytes) {
      currentWorkbook = XLSX.read(workbookBytes, { type: "array", cellStyles: true });
    }

    // DB에만 있는 시트는 원본 워크북에 추가만 하고, 전체 도면을 DB생으로 교체하지 않음
    let sheetMismatch = getSheetMismatch(dbRacks, currentWorkbook);
    if (sheetMismatch.length > 0 && inventoryRacks.length > 0 && currentWorkbook) {
      const layoutMap = getLayoutMapFromDb(layoutRows);
      sheetMismatch.forEach((sheetName) => {
        if (currentWorkbook.Sheets[sheetName]) return;
        const racks = inventoryRacks.filter((r) => r.sheet_name === sheetName);
        const layout = layoutMap[sheetName];
        const maxRow =
          layout?.maxRows ||
          (racks.length ? Math.max(18, ...racks.map((r) => r.rack_row)) : 18);
        let totalCols = 11;
        if (layout?.leftCols || layout?.rightCols) {
          totalCols = (layout.leftCols || 0) + (layout.rightCols || 0) || 11;
        } else if (racks.length) {
          totalCols = Math.max(11, ...racks.map((r) => r.rack_col));
        }
        XLSX.utils.book_append_sheet(
          currentWorkbook,
          createWarehouseLayoutSheet(sheetName, totalCols, maxRow),
          sheetName
        );
      });
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
  if (exportHtmlBtn) exportHtmlBtn.style.display = "inline-block";
    showWorkbookControls();

    const defaultSheet = resolveDefaultSheetName(
      currentWorkbook,
      getRememberedSheet(),
      dbSheetNames
    );
    sheetSelect.value = defaultSheet;
    rebuildProductCatalog();
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
      // 이 PC에만 있는 엑셀이면 B1/B2/B3 등 빈 구역이 다른 기기에 안 보임 → 클라우드에 재게시
      persistUploadedWorkbook(currentWorkbook).then((ok) => {
        if (ok) {
          syncStatusEl.textContent = "🟢 실시간 동기화 중 (엑셀 클라우드 재게시됨)";
        }
      }).catch((e) => console.warn("로컬 엑셀 클라우드 재게시 실패:", e));
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
    clearSyncFailed();
  } catch (err) {
    if (loadToken !== supabaseLoadToken) return;
    console.error("데이터베이스 로드 실패:", err);
    markSyncFailed(`🔴 동기화 실패 — 클릭하여 재시도`);
    updateSyncDiagnostics(0, false, "bundled");
    showToast(err.message || "데이터베이스 로드에 실패했습니다.", "error", 5000);
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
  setGlobalLoading(true, "🟡 DB 동기화 중...");

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
    clearSyncFailed();
    showToast("엑셀 데이터를 DB에 동기화했습니다.", "success");

    const workbookSaved = await persistUploadedWorkbook(currentWorkbook);
    if (!workbookSaved) {
      syncStatusEl.textContent = "🟡 재고 동기화됨 — 엑셀 파일 저장 실패 (모바일은 DB 도면 복원)";
      syncStatusEl.style.background = "rgba(245, 158, 11, 0.12)";
      syncStatusEl.style.border = "1px solid rgba(245, 158, 11, 0.4)";
      syncStatusEl.style.color = "#fbbf24";
      showToast("재고·도면은 DB에 저장됐으나 엑셀 원본 저장만 실패했습니다.", "warning", 5000);
    }
  } catch (err) {
    markSyncFailed("🔴 업로드 동기화 실패 — 클릭하여 재시도");
    showToast("서버에 데이터를 업로드하는 중 오류가 발생했습니다.", "error");
    console.error(err);
    setOfflineMode();
  } finally {
    setGlobalLoading(false);
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
    showToast("URL과 Anon Key를 모두 입력해야 저장됩니다.", "warning");
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
    showToast("연결 성공! 데이터를 불러옵니다.", "success");
  });
}

function disconnectDb() {
  if (confirm("Supabase 실시간 연동을 해제하시겠습니까? (로컬 오프라인 엑셀 방식으로 돌아갑니다.)")) {
    localStorage.removeItem("supabase_url");
    localStorage.removeItem("supabase_anon_key");
    hideDbModal();
    showToast("연동 해제 완료.", "info");
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

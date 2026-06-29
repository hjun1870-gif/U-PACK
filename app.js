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

// 모달 관련 요소
const editModal = document.getElementById("editModal");
const closeModal = document.getElementById("closeModal");
const modalCellTitle = document.getElementById("modalCellTitle");
const inputProduct = document.getElementById("inputProduct");
const inputPallet = document.getElementById("inputPallet");
const inputBox = document.getElementById("inputBox");
const cancelBtn = document.getElementById("cancelBtn");
const saveBtn = document.getElementById("saveBtn");

/**
 * 1. 이벤트 리스너 등록
 */
fileInput.addEventListener("change", handleFileSelect);
demoBtn.addEventListener("click", loadDemoWorkbook);
sheetSelect.addEventListener("change", (e) => {
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

// 검색창 외부 영역 클릭 시 검색 결과 래퍼 닫기
document.addEventListener("click", (e) => {
  if (!searchInput.contains(e.target) && !searchResultsWrap.contains(e.target)) {
    searchResultsWrap.style.display = "none";
  }
});

/**
 * 2. 엑셀 파일 로드 및 파싱 처리
 */
function handleFileSelect(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    try {
      // SheetJS로 바이너리 데이터 로드
      const workbook = XLSX.read(data, { type: "array" });
      currentWorkbook = workbook;
      
      // 시트 선택 드롭다운 갱신
      updateSheetDropdown(workbook.SheetNames);
      
      // 가이드 숨기고 그리드 표시 영역 세팅
      emptyGuide.style.display = "none";
      gridBoard.style.display = "grid";
      exportBtn.style.display = "inline-block";
      
      // 기본적으로 'B2' 시트를 렌더링하고, 없으면 첫 번째 시트 렌더링
      const defaultSheet = workbook.SheetNames.includes("B2") ? "B2" : workbook.SheetNames[0];
      sheetSelect.value = defaultSheet;
      renderActiveSheet(defaultSheet);
    } catch (err) {
      alert("엑셀 파일을 파싱하는 데 실패했습니다. 파일이 손상되었거나 유효한 엑셀 형식이 아닙니다.");
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
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

  // UI 요소 갱신 및 표시 전환
  updateSheetDropdown(wb.SheetNames);
  emptyGuide.style.display = "none";
  gridBoard.style.display = "grid";
  exportBtn.style.display = "inline-block";

  sheetSelect.value = "B2";
  renderActiveSheet("B2");
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
    // 15행 18열: SM300, P:10, B:17
    setMockStock(sheet, 15, 18, "SM300", 10, 17);
    // 16행 18열: 새콤초300, P:18
    setMockStock(sheet, 16, 18, "새콤초300", 18, 0);
    // 18행 18열: M800(신), P:10, B:5
    setMockStock(sheet, 18, 18, "M800(신)", 10, 5);
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
 * 3. 선택한 시트의 레이아웃 렌더링 및 매핑 알고리즘
 */
function renderActiveSheet(sheetName) {
  currentSheetName = sheetName;
  const sheet = currentWorkbook.Sheets[sheetName];
  currentSheetData = sheet;

  // 그리드 보드 및 검색 필터 초기화
  gridBoard.innerHTML = "";
  searchInput.value = "";

  // 18행 x 18열 격자판 뼈대 생성
  // CSS Grid 구조: [행라벨] [1~9열] [중앙통로] [10~18열] [행라벨]
  // 18열 헤더(가로축) 상단 인덱스 렌더링
  renderGridHeaders();

  // 엑셀 시트 4번째 행(인덱스 4)에서 1~18번 열 번호가 실제로 몇 번째 컬럼(Col)에 매핑되는지 동적으로 검출합니다.
  // 시트별 템플릿 차이(예: B1, B2 열 차이)에 유연하게 대응하기 위함입니다.
  const colMapping = detectColumnMapping(sheet);

  let totalProducts = 0;
  let totalPallets = 0;
  let totalBoxes = 0;

  // 1행부터 18행까지 루프 돌며 물리적 랙 셀 생성
  for (let r = 1; r <= 18; r++) {
    // 0-based 엑셀 행 인덱스 계산
    const prodRowIdx = 7 + 2 * (r - 1); // 제품명이 들어가는 행 (8번째 행)
    const qtyRowIdx = 8 + 2 * (r - 1);  // 수량이 들어가는 행 (9번째 행)

    // 왼쪽 행 번호 라벨 생성
    const leftRowLabel = document.createElement("div");
    leftRowLabel.className = "row-label";
    leftRowLabel.textContent = r;
    gridBoard.appendChild(leftRowLabel);

    // 1~9열 (좌측 영역) 렌더링
    for (let c = 1; c <= 9; c++) {
      const cellData = getRackCellData(sheet, colMapping, r, c, prodRowIdx, qtyRowIdx);
      const cellEl = createRackCellElement(r, c, cellData);
      gridBoard.appendChild(cellEl);

      if (cellData.product) {
        totalProducts++;
        totalPallets += cellData.pallet;
        totalBoxes += cellData.box;
      }
    }

    // 중앙 통로 셀 생성 (11번째 컬럼)
    const aisle = document.createElement("div");
    if (r === 18) {
      // 18행에는 입구 게이트와 화살표 렌더링
      aisle.className = "aisle-cell entrance-gate";
      const arrow = document.createElement("div");
      arrow.className = "entrance-arrow";
      // SVG 위쪽 화살표와 "B2 입구" 텍스트
      arrow.innerHTML = `
        <svg viewBox="0 0 24 24">
          <path d="M4 12l1.41 1.41L11 7.83V20h2V7.83l5.58 5.59L20 12l-8-8-8 8z"/>
        </svg>
        <span>입구</span>
      `;
      aisle.appendChild(arrow);
    } else {
      aisle.className = "aisle-cell entrance-path";
    }
    gridBoard.appendChild(aisle);

    // 10~18열 (우측 영역) 렌더링
    for (let c = 10; c <= 18; c++) {
      const cellData = getRackCellData(sheet, colMapping, r, c, prodRowIdx, qtyRowIdx);
      const cellEl = createRackCellElement(r, c, cellData);
      gridBoard.appendChild(cellEl);

      if (cellData.product) {
        totalProducts++;
        totalPallets += cellData.pallet;
        totalBoxes += cellData.box;
      }
    }

    // 오른쪽 행 번호 라벨 생성
    const rightRowLabel = document.createElement("div");
    rightRowLabel.className = "row-label";
    rightRowLabel.textContent = r;
    gridBoard.appendChild(rightRowLabel);
  }

  // 하단 헤더(가로축 대칭형 인덱스) 렌더링
  renderGridFooters();

  // 통계 요약판 수치 반영
  totalProductsEl.textContent = totalProducts;
  totalPalletsEl.textContent = totalPallets;
  totalBoxesEl.textContent = totalBoxes;
}

/**
 * 엑셀 시트의 4행에서 1~18번 열 번호가 기록된 컬럼 인덱스를 동적으로 찾아 매핑 맵(Map)을 반환합니다.
 */
function detectColumnMapping(sheet) {
  const mapping = {};
  // 시트가 유효하지 않으면 기본 매핑 생성
  if (!sheet) return createDefaultMapping();

  // A1 형태로 전체 범위를 파싱하여 행과 열의 범위를 계산
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:AM44");
  const row4Idx = 4; // 5번째 행 (0-based 4)

  // 0열부터 최대 열 범위까지 검사하여 1~18 숫자가 적힌 곳을 찾음
  for (let colIdx = range.s.c; colIdx <= range.e.c; colIdx++) {
    const cellAddress = XLSX.utils.encode_cell({ r: row4Idx, c: colIdx });
    const cell = sheet[cellAddress];
    if (cell && cell.v !== undefined) {
      const numVal = parseInt(cell.v, 10);
      if (numVal >= 1 && numVal <= 18) {
        // 이미 맵에 존재하지 않는 경우에만 설정 (B1, B2 시트의 병합 셀 등으로 동일 숫자가 인접해 나오는 것 처리)
        if (mapping[numVal] === undefined) {
          mapping[numVal] = colIdx;
        }
      }
    }
  }

  // 누락된 열 번호가 있다면 기본 매핑을 보완책으로 사용
  for (let c = 1; c <= 18; c++) {
    if (mapping[c] === undefined) {
      if (c <= 9) {
        mapping[c] = 2 * c - 1;
      } else {
        mapping[c] = 2 * c + 1;
      }
    }
  }

  return mapping;
}

function createDefaultMapping() {
  const mapping = {};
  for (let c = 1; c <= 18; c++) {
    mapping[c] = c <= 9 ? 2 * c - 1 : 2 * c + 1;
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

  // 파레트(P)가 저장되는 컬럼 인덱스
  const pColIdx = colMapping[col];
  // 박스(B)가 저장되는 컬럼은 항상 P의 바로 우측 컬럼
  const bColIdx = pColIdx + 1;

  // 엑셀 셀 주소(A1 표기법) 계산
  data.prodCellAddress = XLSX.utils.encode_cell({ r: prodRowIdx, c: pColIdx });
  data.palletCellAddress = XLSX.utils.encode_cell({ r: qtyRowIdx, c: pColIdx });
  data.boxCellAddress = XLSX.utils.encode_cell({ r: qtyRowIdx, c: bColIdx });

  // 엑셀에서 실제 셀 객체 추출
  const prodCell = sheet[data.prodCellAddress];
  const palletCell = sheet[data.palletCellAddress];
  const boxCell = sheet[data.boxCellAddress];

  // 값 추출
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
function createRackCellElement(row, col, cellData) {
  const cell = document.createElement("div");
  // 제품명이 존재하거나 수량이 하나라도 있을 때 active-stock 클래스 부여
  if (cellData.product || cellData.pallet > 0 || cellData.box > 0) {
    cell.className = "rack-cell active-stock";
  } else {
    cell.className = "rack-cell";
  }

  // 검색 시 필터링을 쉽게 하기 위해 제품명을 data 속성으로 주입
  cell.dataset.product = cellData.product.toLowerCase();
  cell.dataset.row = row;
  cell.dataset.col = col;

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

  // 셀 클릭 이벤트 바인딩 -> 수정 모달 띄우기
  cell.addEventListener("click", () => {
    showEditModal(row, col, cellData);
  });

  return cell;
}

/**
 * 격자 상/하단 인덱스 라벨들 생성 함수
 */
function renderGridHeaders() {
  // 첫 칸 빈자리 (행 라벨 공간)
  const emptyLeft = document.createElement("div");
  emptyLeft.className = "col-label empty-header";
  gridBoard.appendChild(emptyLeft);

  // 1~9열 상단 라벨
  for (let c = 1; c <= 9; c++) {
    const label = document.createElement("div");
    label.className = "col-label";
    label.textContent = c;
    gridBoard.appendChild(label);
  }

  // 중앙 통로용 공간
  const emptyCenter = document.createElement("div");
  emptyCenter.className = "col-label empty-header";
  gridBoard.appendChild(emptyCenter);

  // 10~18열 상단 라벨
  for (let c = 10; c <= 18; c++) {
    const label = document.createElement("div");
    label.className = "col-label";
    label.textContent = c;
    gridBoard.appendChild(label);
  }

  // 마지막 칸 빈자리 (오른쪽 행 라벨 공간)
  const emptyRight = document.createElement("div");
  emptyRight.className = "col-label empty-header";
  gridBoard.appendChild(emptyRight);
}

function renderGridFooters() {
  const emptyLeft = document.createElement("div");
  emptyLeft.className = "col-label empty-footer";
  gridBoard.appendChild(emptyLeft);

  // 좌측 하단 대칭용 인덱스 (1~9)
  for (let c = 1; c <= 9; c++) {
    const label = document.createElement("div");
    label.className = "col-label";
    label.textContent = c;
    gridBoard.appendChild(label);
  }

  // 중앙 통로 공간
  const emptyCenter = document.createElement("div");
  emptyCenter.className = "col-label empty-footer";
  gridBoard.appendChild(emptyCenter);

  // 우측 하단 대칭용 인덱스 (9, 8, 7, 6, 5, 4, 3, 2, 1) - 이미지와 동일
  const rightIndices = [9, 8, 7, 6, 5, 4, 3, 2, 1];
  rightIndices.forEach((val) => {
    const label = document.createElement("div");
    label.className = "col-label";
    label.textContent = val;
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

  // 3. 전체 시트(구역)를 순회하며 일치하는 재고 데이터 및 수량(P, B) 수집 (글로벌 스캔)
  currentWorkbook.SheetNames.forEach((sheetName) => {
    const sheet = currentWorkbook.Sheets[sheetName];
    if (!sheet) return;

    // 해당 시트의 열 좌표 매핑 정보 검출
    const colMapping = detectColumnMapping(sheet);

    // 1행부터 18행까지 순회
    for (let r = 1; r <= 18; r++) {
      const prodRowIdx = 7 + 2 * (r - 1);
      const qtyRowIdx = 8 + 2 * (r - 1);

      // 1열부터 18열까지 순회
      for (let c = 1; c <= 18; c++) {
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
function showEditModal(row, col, cellData) {
  selectedCellInfo = {
    row: row,
    col: col,
    data: cellData
  };

  modalCellTitle.textContent = `${row}행 - ${col}열 재고 수정`;
  inputProduct.value = cellData.product || "";
  inputPallet.value = cellData.pallet || 0;
  inputBox.value = cellData.box || 0;

  editModal.classList.add("show");
  inputProduct.focus();
}

function hideModal() {
  editModal.classList.remove("show");
  selectedCellInfo = null;
}

/**
 * 사용자가 수정한 재고를 엑셀 데이터 구조에 반영하고 화면을 새로고침합니다.
 */
function saveCellChanges() {
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

  // 2. 화면 갱신 (전체 렌더링을 다시 호출하여 합계 및 레이아웃을 일괄 동기화)
  renderActiveSheet(currentSheetName);

  // 3. 모달 닫기
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
 *  - TOTAL 시트의 A열 제품명을 읽어 키로 삼아
 *  - 모든 구역 시트에서 동일 제품명의 P(파렛트)/B(박스) 합계를 산출
 *  - TOTAL 시트 D열(P)과 E열(B)에 결과를 기입합니다.
 */
function aggregateTotalSheet(workbook) {
  // TOTAL 시트 찾기 (대소문자 무관)
  const totalSheetName = workbook.SheetNames.find(
    (n) => n.toUpperCase() === "TOTAL" || n.includes("합계")
  );
  if (!totalSheetName) return; // TOTAL 시트 없으면 스킵

  const totalSheet = workbook.Sheets[totalSheetName];
  if (!totalSheet) return;

  // ① TOTAL 시트의 제품 목록 파악 (A열, 5행부터 마지막 행까지)
  //    엑셀 행 5 = SheetJS r 인덱스 4
  const productRows = {}; // { "M300": 5행(엑셀 행번호), "M500": 6, ... }

  for (let excelRow = 5; excelRow <= 300; excelRow++) {
    const addr = XLSX.utils.encode_cell({ r: excelRow - 1, c: 0 }); // A열 = c:0
    const cell = totalSheet[addr];
    if (!cell || cell.v === undefined || String(cell.v).trim() === "") break;
    const prodName = String(cell.v).trim();
    productRows[prodName] = excelRow;
  }

  // ② 각 구역 시트(TOTAL 시트 제외) 순회 → 제품별 P/B 합계 누적
  const aggregated = {}; // { "M300": { pallet: 합계, box: 합계 }, ... }

  workbook.SheetNames.forEach((sheetName) => {
    if (sheetName === totalSheetName) return; // TOTAL 자기 자신 스킵

    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;

    const colMapping = detectColumnMapping(sheet);

    // 각 랙 셀 (1행 ~ 18행, 1열 ~ 18열) 스캔
    for (let r = 1; r <= 18; r++) {
      const prodRowIdx = 7 + 2 * (r - 1);
      const qtyRowIdx  = 8 + 2 * (r - 1);

      for (let c = 1; c <= 18; c++) {
        const pColIdx = colMapping[c];
        if (pColIdx === undefined) continue;
        const bColIdx = pColIdx + 1;

        const prodAddr   = XLSX.utils.encode_cell({ r: prodRowIdx, c: pColIdx });
        const palletAddr = XLSX.utils.encode_cell({ r: qtyRowIdx,  c: pColIdx });
        const boxAddr    = XLSX.utils.encode_cell({ r: qtyRowIdx,  c: bColIdx });

        const prodCell   = sheet[prodAddr];
        const palletCell = sheet[palletAddr];
        const boxCell    = sheet[boxAddr];

        if (prodCell && prodCell.v !== undefined) {
          const prodName = String(prodCell.v).trim();
          if (!prodName) continue;

          const pQty = (palletCell && palletCell.v !== undefined) ? (Number(palletCell.v) || 0) : 0;
          const bQty = (boxCell    && boxCell.v    !== undefined) ? (Number(boxCell.v)    || 0) : 0;

          if (pQty === 0 && bQty === 0) continue; // 수량 없으면 스킵

          if (!aggregated[prodName]) {
            aggregated[prodName] = { pallet: 0, box: 0 };
          }
          aggregated[prodName].pallet += pQty;
          aggregated[prodName].box    += bQty;
        }
      }
    }
  });

  // ③ 집계 결과를 TOTAL 시트의 D열(P)과 E열(B)에 기입
  //    엑셀 D열 = 인덱스 3, E열 = 인덱스 4
  for (const [prodName, excelRow] of Object.entries(productRows)) {
    const totals = aggregated[prodName] || { pallet: 0, box: 0 };

    const pAddr = XLSX.utils.encode_cell({ r: excelRow - 1, c: 3 }); // D열
    const bAddr = XLSX.utils.encode_cell({ r: excelRow - 1, c: 4 }); // E열

    // 파렛트(P) 기입
    if (!totalSheet[pAddr]) {
      totalSheet[pAddr] = { t: "n", v: totals.pallet };
    } else {
      totalSheet[pAddr].t = "n";
      totalSheet[pAddr].v = totals.pallet;
      delete totalSheet[pAddr].w; // 캐시된 표시값 초기화
    }

    // 박스(B) 기입
    if (!totalSheet[bAddr]) {
      totalSheet[bAddr] = { t: "n", v: totals.box };
    } else {
      totalSheet[bAddr].t = "n";
      totalSheet[bAddr].v = totals.box;
      delete totalSheet[bAddr].w;
    }
  }

  console.log("✅ TOTAL 시트 집계 완료:", aggregated);
}


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

    // 1. 열 너비 지정 (각 시트의 폭을 최적화)
    const colsWidths = [];
    colsWidths[0] = { wch: 6 }; // Row label 열 너비
    
    // 시트에서 1~18번 열의 Col 인덱스 맵 가져옴
    const colMapping = detectColumnMapping(sheet);
    
    // 1~18번 열의 Col 인덱스를 기준으로 열 너비 할당 (P: 14, B: 6)
    for (let c = 1; c <= 18; c++) {
      const pCol = colMapping[c];
      if (pCol !== undefined) {
        colsWidths[pCol] = { wch: 14 }; // 제품명 및 파렛트(P) 열
        colsWidths[pCol + 1] = { wch: 6 };  // 박스(B) 열
      }
    }
    
    // 통로 열(Col 19, 20) 및 기타 미지정 열 너비 기본값 채우기
    for (let i = 0; i < 39; i++) {
      if (!colsWidths[i]) {
        colsWidths[i] = { wch: 4 }; // 콤팩트 통로
      }
    }
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

    // 1행부터 18행까지 순회하며 격자 셀 스타일링
    for (let r = 1; r <= 18; r++) {
      const prodRowIdx = 7 + 2 * (r - 1);
      const qtyRowIdx = 8 + 2 * (r - 1);

      // 행 라벨 (좌/우) 스타일링
      const leftLabelAddr = XLSX.utils.encode_cell({ r: prodRowIdx, c: 0 });
      const leftQtyLabelAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: 0 });
      const rightLabelAddr = XLSX.utils.encode_cell({ r: prodRowIdx, c: 38 });
      const rightQtyLabelAddr = XLSX.utils.encode_cell({ r: qtyRowIdx, c: 38 });

      [leftLabelAddr, leftQtyLabelAddr, rightLabelAddr, rightQtyLabelAddr].forEach((addr) => {
        if (!sheet[addr]) sheet[addr] = { t: "n", v: r };
        sheet[addr].s = sHeader;
      });

      // 1~18열 격자 루프
      for (let c = 1; c <= 18; c++) {
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
    for (let c = 1; c <= 18; c++) {
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
  // A:제품명, B:입수량, C:박스단위, D:P, E:B, F:잔량, G:재고, H:출고, I:합계, J:판매일보, K:차이
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

  // ── 4행 헤더 스타일 적용 (A~K 열, 인덱스 0~10) ──────────
  for (let col = 0; col <= 10; col++) {
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

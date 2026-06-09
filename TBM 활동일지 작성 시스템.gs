/**
 * TBM 활동기록일지 자동 작성 시스템
 * 
 * [스프레드시트 ID 설정]
 * - PROD_PLAN_SS_ID : 2026년 제품생산 계획.xlsx 의 Google Sheets ID
 * - WORK_SCHEDULE_SS_ID : 생수공장 근무계획 스프레드시트 ID
 * 
 * [사용법]
 * 1. 아래 PROD_PLAN_SS_ID, WORK_SCHEDULE_SS_ID 값을 실제 스프레드시트 ID로 설정하세요.
 * 2. 메뉴 [TBM 자동화] > [오늘 TBM 일지 작성] 을 사용하세요.
 * 3. 특정 날짜 테스트는 메뉴 [TBM 자동화] > [날짜 지정 테스트] 를 사용하세요.
 */

// ============================================================
// ★ 스프레드시트 ID 설정 (여기를 수정하세요)
// ============================================================
const PROD_PLAN_SS_ID = '1ibZaOX0NKni6vRTVE0i2Wb5AXNHn9D72F5CFWSxH3qQ';
const WORK_SCHEDULE_SS_ID = '12yP7AxfC_27B6LhSEJ4lFH5ZYnymyQIlcbPXXhpDUBY';

// ============================================================
// 작업전 교육내용 레파토리 (순서 고정)
// ============================================================
const EDUCATION_REPERTOIRE = [
  {
    title: 'LOTO(잠금장치,표지판) 작업절차',
    items: ['1.LOTO(잠금장치,표지판)란?', '2.LOTO 작업절차'],
    requiresBtlCap: false
  },
  {
    title: '작업 시 감전 재해예방',
    items: ['1.작업별 주요 감전 요인', '2.감전 예방을 위한 공통 안전수칙'],
    requiresBtlCap: false
  },
  {
    title: '끼임예방',
    items: ['1.끼임사고사망발생사례', '2.현장에서반드시준수할사항'],
    requiresBtlCap: false
  },
  {
    title: '와이어로프 사용 및 폐기 결정가이드',
    items: ['1.와이어로프 사용 모습', '2.와이어로프 폐기 기준'],
    requiresBtlCap: false
  },
  {
    title: '크레인 재해예방',
    items: ['1.중대재해 사례', '2.크레인 위험요인 및 안전대책'],
    requiresBtlCap: false
  },
  {
    title: '작은설비 끼임',
    items: ['1.재해 사례', '2.사고사망 예방을 위해 반드시 준수할 사항'],
    requiresBtlCap: false
  },
  {
    title: '사출성형기 재해예방',
    items: ['1.중대재해 사례', '2.사출성형기 위험요인 및 안전대책'],
    requiresBtlCap: true  // BTL 또는 CAP 생산 시에만 사용
  },
  {
    title: '안전난간 안전작업 가이드',
    items: ['1.안전난간 (Safety Guard Rail)', '2.안전난간의 설치 기준', '3.안전대책'],
    requiresBtlCap: false
  },
  {
    title: '컨베이어 재해예방',
    items: ['1.중대재해 사례', '2.컨베이어 위험요인 및 안전대책'],
    requiresBtlCap: false
  },
  {
    title: '벨트 슬링 안전작업 가이드',
    items: ['1.벨트 슬링이란?', '2.벨트 슬링 폐기기준'],
    requiresBtlCap: false
  },
  {
    title: '산업용로봇 재해예방',
    items: ['1.중대재해 사례', '2.산업용로봇 위험요인 및 안전대'],
    requiresBtlCap: false
  }
];

// 위험요인/조치방안 레파토리 (18.9L도 아니고 LINE CHANGE도 아닐 때 사용)
const HAZARD_REPERTOIRE = [
  {
    hazard: '컨베이어벨트 회전부 끼임 위험',
    action: '방호장치(커버,비상정지장치)설치',
    onlyFor012L: false
  },
  {
    hazard: 'ROBOTIZER 장비와 충돌 및 끼임사고',
    action: '출입문 연동장치 설치\n안전센서 설치',
    onlyFor012L: true  // 0.12L 생산 날에만 사용
  }
];

// ============================================================
// 메뉴 생성
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('TBM 자동화')
    .addItem('오늘 TBM 일지 작성', 'createTodayTBM')
    .addSeparator()
    .addItem('날짜 지정 테스트 (YYYY-MM-DD)', 'createTBMByDateInput')
    .addToUi();
}

// ============================================================
// 메인 진입점 - 오늘 날짜로 실행
// ============================================================
function createTodayTBM() {
  const today = new Date();
  createTBMForDate(today);
}

// ============================================================
// 날짜 지정 테스트
// ============================================================
function createTBMByDateInput() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '날짜 지정 테스트',
    '날짜를 YYYY-MM-DD 형식으로 입력하세요:\n(예: 2026-05-22)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const inputText = response.getResponseText().trim();
  const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
  
  if (!dateRegex.test(inputText)) {
    ui.alert('오류', 'YYYY-MM-DD 형식으로 입력해 주세요.', ui.ButtonSet.OK);
    return;
  }
  
  const parts = inputText.split('-');
  const targetDate = new Date(
    parseInt(parts[0]),
    parseInt(parts[1]) - 1,
    parseInt(parts[2])
  );
  
  if (isNaN(targetDate.getTime())) {
    ui.alert('오류', '유효하지 않은 날짜입니다.', ui.ButtonSet.OK);
    return;
  }
  
  createTBMForDate(targetDate);
}

// ============================================================
// 핵심 함수 - 특정 날짜에 대한 TBM 일지 작성
// ============================================================
function createTBMForDate(targetDate) {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const tbmSS = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. 시트 이름 형식 결정 (MMDD)
    const month = targetDate.getMonth() + 1;
    const day = targetDate.getDate();
    const newSheetName = String(month).padStart(2, '0') + String(day).padStart(2, '0');
    
    // 이미 존재하는지 확인
    if (tbmSS.getSheetByName(newSheetName)) {
      const overwrite = ui.alert(
        '시트 이미 존재',
        `'${newSheetName}' 시트가 이미 존재합니다. 덮어쓰시겠습니까?`,
        ui.ButtonSet.YES_NO
      );
      if (overwrite !== ui.Button.YES) return;
      tbmSS.deleteSheet(tbmSS.getSheetByName(newSheetName));
    }
    
    // 2. 복사할 이전 시트 찾기
    const sourceSheet = findPreviousSheet(tbmSS, targetDate);
    if (!sourceSheet) {
      ui.alert('오류', '복사할 이전 날짜 시트를 찾을 수 없습니다.\n최소 하나의 MMDD 형식 시트가 필요합니다.', ui.ButtonSet.OK);
      return;
    }
    
    // 3. 시트 복사
    const newSheet = sourceSheet.copyTo(tbmSS);
    newSheet.setName(newSheetName);
    
    // 복사된 시트를 맨 뒤로 이동
    tbmSS.moveActiveSheet(tbmSS.getSheets().length);
    tbmSS.setActiveSheet(newSheet);
    
    // 4. 생산계획 스프레드시트에서 오늘 생산 정보 가져오기
    const prodInfo = getProductionInfo(targetDate);
    
    // 5. 근무계획 스프레드시트에서 오늘 TBM 대상자 가져오기
    const { targets: tbmTargets, excluded: tbmExcluded } = getTBMTargets(targetDate);
    
    // 6. 교육 레파토리 인덱스 (스크립트 속성에서 관리)
    const eduIndex = getNextEducationIndex(prodInfo.hasBtlOrCap);
    
    // 7. 위험요인/조치방안 인덱스
    const hazardContent = getHazardContent(prodInfo, eduIndex);
    
    // ============================================================
    // 셀 값 입력
    // ============================================================
    
    // C3: 날짜 (YYYY-MM-DD)
    const year = targetDate.getFullYear();
    const dateString = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    setCellValue(newSheet, 'C3', dateString);
    
    // B15:J15: 조업 특이사항
    const specialNote = buildSpecialNote(prodInfo);
    setMergedCellValue(newSheet, 'B15', specialNote);
    
    // B9:D13 / E9:G13: 위험요인 / 조치방안
    setMergedCellValue(newSheet, 'B9', hazardContent.hazard);
    setMergedCellValue(newSheet, 'E9', hazardContent.action);
    
    // B17:J23: 작업전 교육내용
    const eduContent = EDUCATION_REPERTOIRE[eduIndex];
    setMergedCellValue(newSheet, 'B17', buildEducationContent(eduContent));
    
    // B26:J27: 금일 TBM 대상 / 제외
    const tbmCellText = buildTBMCellText(tbmTargets, tbmExcluded);
    setMergedCellValue(newSheet, 'B26', tbmCellText);
    
    // ============================================================
    // 완료 메시지
    // ============================================================
    const summary = buildCompletionSummary(newSheetName, dateString, prodInfo, hazardContent, eduContent, tbmTargets, tbmExcluded);
    ui.alert('TBM 일지 작성 완료', summary, ui.ButtonSet.OK);
    
  } catch (e) {
    ui.alert('오류 발생', `작업 중 오류가 발생했습니다:\n${e.message}\n\n스택:\n${e.stack}`, ui.ButtonSet.OK);
    Logger.log(e);
  }
}

// ============================================================
// 이전 날짜 시트 찾기 (MMDD 형식 기준)
// ============================================================
function findPreviousSheet(ss, targetDate) {
  const sheets = ss.getSheets();
  const mmddPattern = /^\d{4}$/;
  
  // MMDD 형식 시트들만 추출하고 정렬
  const dateSheetsRaw = sheets.filter(s => mmddPattern.test(s.getName()));
  
  if (dateSheetsRaw.length === 0) return null;
  
  // MMDD를 숫자로 변환해 정렬 (연도를 붙여서 비교)
  // 주의: 연도 경계를 넘는 경우도 고려 (12월 → 1월)
  const targetMMDD = (targetDate.getMonth() + 1) * 100 + targetDate.getDate();
  
  const dateSheets = dateSheetsRaw.map(s => ({
    sheet: s,
    mmdd: parseInt(s.getName())
  }));
  
  // targetDate보다 이전인 시트들 중 가장 최근 시트
  // 단순히 MMDD 값으로 비교 (같은 연도 기준)
  const previousSheets = dateSheets
    .filter(s => s.mmdd < targetMMDD)
    .sort((a, b) => b.mmdd - a.mmdd);
  
  if (previousSheets.length > 0) {
    return previousSheets[0].sheet;
  }
  
  // 없으면 전체 중 가장 최근 시트 (연도를 넘어가는 경우)
  const allSorted = dateSheets.sort((a, b) => b.mmdd - a.mmdd);
  return allSorted[0].sheet;
}

// ============================================================
// 생산계획 정보 가져오기
// ============================================================
function getProductionInfo(targetDate) {
  const prodPlanSS = SpreadsheetApp.openById(PROD_PLAN_SS_ID);
  const month = targetDate.getMonth() + 1;
  const day = targetDate.getDate();
  
  // 해당 월의 시트 찾기 ("N월 제품생산 계획" 포함하는 시트)
  const allSheets = prodPlanSS.getSheets();
  let planSheet = null;
  for (const s of allSheets) {
    const name = s.getName();
    if (name.includes(`${month}월`) && name.includes('제품생산 계획')) {
      planSheet = s;
      break;
    }
  }
  
  if (!planSheet) {
    throw new Error(`${month}월 제품생산 계획 시트를 찾을 수 없습니다.`);
  }
  
  // A3:L35 데이터 읽기
  // 컬럼 구조 (0-based index):
  //  A=0: 일자, B=1: 요일(무시), C=2: 18.9L, D=3: 1.5L, E=4: 0.5L, F=5: 0.33L
  //  G=6: 0.12L, H=7: 0.19L캔, I=8: 1.5L BTL, J=9: 0.5L BTL, K=10: 0.33L BTL, L=11: CAP
  const data = planSheet.getRange('A3:L35').getValues();
  
  // 오늘 일자에 해당하는 행 찾기 (A열 = 일자)
  let todayRow = null;
  for (let i = 0; i < data.length; i++) {
    const cellVal = data[i][0]; // A열
    if (String(cellVal).trim() === String(day) || Number(cellVal) === day) {
      todayRow = data[i];
      break;
    }
  }
  
  if (!todayRow) {
    throw new Error(`${month}월 ${day}일에 해당하는 생산계획 행을 찾을 수 없습니다.`);
  }
  
  // 각 열 값 추출
  const colC = String(todayRow[2]).trim();   // C: 18.9L 제품
  const colD = String(todayRow[3]).trim();   // D: 1.5L 제품  (L/C 가능: 포장실 LINE CHANGE)
  const colE = String(todayRow[4]).trim();   // E: 0.5L 제품  (L/C 가능: 포장실 LINE CHANGE)
  const colF = String(todayRow[5]).trim();   // F: 0.33L 제품 (L/C 가능: 포장실 LINE CHANGE)
  const colG = String(todayRow[6]).trim();   // G: 0.12L 제품
  const colH = String(todayRow[7]).trim();   // H: 0.19L 캔
  const colI = String(todayRow[8]).trim();   // I: 1.5L BTL
  const colJ = String(todayRow[9]).trim();   // J: 0.5L BTL
  const colK = String(todayRow[10]).trim();  // K: 0.33L BTL
  const colL = String(todayRow[11]).trim();  // L: CAP
  
  // 헬퍼 함수
  // 숫자 값(생산 수치)인지 확인 - 텍스트 특이사항은 false 반환
  const isNumericValue = (val) => {
    if (!val || val === '' || val === '-') return false;
    const num = Number(String(val).replace(/,/g, ''));
    return !isNaN(num) && num > 0;
  };
  const isLc = (val) => val.toUpperCase().includes('L/C');
  
  // ─────────────────────────────────────────
  // 제품 생산 분석 (C~H열)
  // ─────────────────────────────────────────
  // 포장실 LINE CHANGE: D, E, F열에만 L/C가 들어갈 수 있음
  // C열(18.9L), G열(0.12L), H열(0.19L캔)은 L/C 대상 아님
  //
  // 출력 순서: 용량 높은 순 → 18.9L > 1.5L > 0.5L > 0.33L > 0.19L캔 > 0.12L
  
  let has189L = false;
  let hasPackagingLineChange = false;  // 포장실 LINE CHANGE
  let has012L = false;
  const activeProducts = [];
  
  // C열: 18.9L (L/C 불가, 숫자값만 인정)
  if (isNumericValue(colC)) {
    activeProducts.push('18.9L');
    has189L = true;
  }
  
  // D열: 1.5L (L/C = 포장실 LINE CHANGE, 숫자값만 생산으로 인정)
  if (isLc(colD)) {
    hasPackagingLineChange = true;
  } else if (isNumericValue(colD)) {
    activeProducts.push('1.5L');
  }
  
  // E열: 0.5L (L/C = 포장실 LINE CHANGE, 숫자값만 생산으로 인정)
  if (isLc(colE)) {
    hasPackagingLineChange = true;
  } else if (isNumericValue(colE)) {
    activeProducts.push('0.5L');
  }
  
  // F열: 0.33L (L/C = 포장실 LINE CHANGE, 숫자값만 생산으로 인정)
  if (isLc(colF)) {
    hasPackagingLineChange = true;
  } else if (isNumericValue(colF)) {
    activeProducts.push('0.33L');
  }
  
  // H열: 0.19L (L/C 불가, 순서상 0.33L 다음, 숫자값만 생산으로 인정)
  if (isNumericValue(colH)) {
    activeProducts.push('0.19L');
  }
  
  // G열: 0.12L (L/C 불가, 마지막 순서, 숫자값만 생산으로 인정)
  if (isNumericValue(colG)) {
    activeProducts.push('0.12L');
    has012L = true;
  }
  
  // ─────────────────────────────────────────
  // BTL 분석 (I~K열)
  // ─────────────────────────────────────────
  const btlMap = [
    { col: colI, label: '1.5L BTL' },
    { col: colJ, label: '0.5L BTL' },
    { col: colK, label: '0.33L BTL' }
  ];
  
  let hasBtlLineChange = false;  // 병제작기 LINE CHANGE
  const activeBtl = [];
  
  for (const b of btlMap) {
    if (isLc(b.col)) {
      hasBtlLineChange = true;
    } else if (isNumericValue(b.col)) {
      activeBtl.push(b.label);
    }
  }
  
  // ─────────────────────────────────────────
  // CAP 분석 (L열)
  // ─────────────────────────────────────────
  let hasCap = false;
  if (isNumericValue(colL) && !isLc(colL)) {
    hasCap = true;
  }
  
  const hasBtlOrCap = activeBtl.length > 0 || hasCap;
  
  return {
    month,
    day,
    activeProducts,     // ['18.9L', '1.5L', '0.33L', '0.19L캔', '0.12L'] 등 (용량 높은 순)
    has189L,
    has012L,
    hasPackagingLineChange,
    activeBtl,          // ['1.5L BTL'] 등
    hasCap,
    hasBtlLineChange,
    hasBtlOrCap
  };
}

// ============================================================
// 조업 특이사항 문자열 생성
// ============================================================
function buildSpecialNote(prodInfo) {
  const lines = [];
  
  // 1. 제품 생산
  if (prodInfo.activeProducts.length > 0) {
    lines.push(`1.${prodInfo.activeProducts.join('/')} 제품생산`);
  }
  
  // 2. BTL/CAP 또는 병제작기 LINE CHANGE
  if (prodInfo.hasBtlLineChange) {
    lines.push('2.병제작기 LINE CHANGE');
  } else if (prodInfo.activeBtl.length > 0 || prodInfo.hasCap) {
    const btlCapParts = [...prodInfo.activeBtl];
    if (prodInfo.hasCap) btlCapParts.push('CAP');
    lines.push(`2.${btlCapParts.join('/')} 생산`);
  }
  
  // 3. 포장실 LINE CHANGE
  if (prodInfo.hasPackagingLineChange) {
    lines.push('3.포장실 LINE CHANGE');
  }
  
  return lines.join('\n');
}

// ============================================================
// 위험요인/조치방안 컨텐츠 결정
// ============================================================
function getHazardContent(prodInfo, eduIndex) {
  // 우선순위 1: 병제작기 LINE CHANGE
  if (prodInfo.hasBtlLineChange) {
    return {
      hazard: '1.금형교체 작업 중 불시작동으로 인한 끼임 위험\n2.1,2,3호기 작업발판 승강시 추락위험\n3.줄걸이용구를 사용하여 중량물 운반간 중량물 낙하위험',
      action: '1.작업 중 전원부 잠금장치 실시 / 리미트 스위치 정상작동 여부확인 / 잠금장치 키 작업자 보관\n2.안전모 착용\n3.작업 전 줄걸이용구 상태 확인 / 작업간 줄걸이용구 체결상태 지적확인 실시'
    };
  }
  
  // 우선순위 2: 18.9L 생산
  if (prodInfo.has189L) {
    return {
      hazard: '지게차 운행시 추돌 주의',
      action: '지게차 안전거리확보 LED라인등 작동/지적확인 철저'
    };
  }
  
  // 우선순위 3: 랜덤 레파토리 (스크립트 속성에서 인덱스 관리)
  const hazardIndex = getNextHazardIndex(prodInfo.has012L);
  return {
    hazard: HAZARD_REPERTOIRE[hazardIndex].hazard,
    action: HAZARD_REPERTOIRE[hazardIndex].action
  };
}

// ============================================================
// 스크립트 속성 관리 - 교육 레파토리 인덱스
// ============================================================
function getNextEducationIndex(hasBtlOrCap) {
  const props = PropertiesService.getScriptProperties();
  let currentIndex = parseInt(props.getProperty('EDU_INDEX') || '0');
  
  // 현재 인덱스의 항목이 requiresBtlCap인데 BTL/CAP이 없으면 스킵
  let safetyCount = 0;
  while (EDUCATION_REPERTOIRE[currentIndex].requiresBtlCap && !hasBtlOrCap) {
    currentIndex = (currentIndex + 1) % EDUCATION_REPERTOIRE.length;
    safetyCount++;
    if (safetyCount > EDUCATION_REPERTOIRE.length) break; // 무한루프 방지
  }
  
  const usedIndex = currentIndex;
  const nextIndex = (currentIndex + 1) % EDUCATION_REPERTOIRE.length;
  props.setProperty('EDU_INDEX', String(nextIndex));
  
  return usedIndex;
}

// ============================================================
// 스크립트 속성 관리 - 위험요인 레파토리 인덱스
// ============================================================
function getNextHazardIndex(has012L) {
  const props = PropertiesService.getScriptProperties();
  let currentIndex = parseInt(props.getProperty('HAZARD_INDEX') || '0');
  
  // 0.12L 생산이 없는데 onlyFor012L인 항목이라면 스킵
  let safetyCount = 0;
  while (HAZARD_REPERTOIRE[currentIndex].onlyFor012L && !has012L) {
    currentIndex = (currentIndex + 1) % HAZARD_REPERTOIRE.length;
    safetyCount++;
    if (safetyCount > HAZARD_REPERTOIRE.length) break;
  }
  
  const usedIndex = currentIndex;
  const nextIndex = (currentIndex + 1) % HAZARD_REPERTOIRE.length;
  props.setProperty('HAZARD_INDEX', String(nextIndex));
  
  return usedIndex;
}

// ============================================================
// 교육내용 문자열 생성
// ============================================================
function buildEducationContent(eduContent) {
  const lines = [];
  lines.push(eduContent.title);
  for (const item of eduContent.items) {
    lines.push(item);
  }
  return lines.join('\n');
}

// ============================================================
// 근무계획에서 TBM 대상자 가져오기
// ============================================================
function getTBMTargets(targetDate) {
  const workSS = SpreadsheetApp.openById(WORK_SCHEDULE_SS_ID);
  const month = targetDate.getMonth() + 1;
  const day = targetDate.getDate();
  
  // 해당 월 시트 찾기 (예: "5월" 또는 "05월")
  const allSheets = workSS.getSheets();
  let scheduleSheet = null;
  for (const s of allSheets) {
    const name = s.getName().trim();
    if (name === `${month}월` || name === `${String(month).padStart(2, '0')}월`) {
      scheduleSheet = s;
      break;
    }
  }
  
  if (!scheduleSheet) {
    throw new Error(`${month}월 근무계획 시트를 찾을 수 없습니다.`);
  }
  
  // B3:AK20 범위를 읽기
  //  - B열(index 0): 직급, C열(index 1): 성명
  //  - G열(index 5)~AK열(index 35): 일자별 근무현황 (1일~31일)
  //  행 구조:
  //    data[0] = 3번째 행 (일자 헤더 행): B3~AK3
  //    data[1] = 4번째 행 (필요시 무시)
  //    data[2]~data[17] = 5번째~20번째 행 (직원 목록)
  const dataRange = scheduleSheet.getRange('B3:AK20');
  const data = dataRange.getValues();
  
  // 일자 행(3번째 행)에서 오늘 일자에 해당하는 열 찾기
  // 날짜는 G열부터 시작 → B3 기준 index 5부터 탐색
  const G_COL_OFFSET = 5; // G열 = B3 기준 index 5
  
  const dateRow = data[0]; // 3번째 행 (B3~AK3)
  let dayColIndex = -1;
  
  for (let c = G_COL_OFFSET; c < dateRow.length; c++) {
    const val = dateRow[c];
    if (Number(val) === day || String(val).trim() === String(day)) {
      dayColIndex = c;
      break;
    }
  }
  
  if (dayColIndex === -1) {
    throw new Error(`${month}월 ${day}일의 일자 열을 찾을 수 없습니다. (G열 이후에서 탐색)`);
  }
  
  // 직원 정보 수집 (5번째 행 ~ 20번째 행)
  // 즉 data[2] ~ data[17]
  // B열 (index 0) = 직급, C열 (index 1) = 성명
  
  // 휴무 코드: 공, 휴, 무, 년, 야, 퇴, 상, 근
  const ABSENT_CODES = ['공', '휴', '무', '년', '야', '퇴', '상', '근'];
  
  const targets = [];   // TBM 대상 (출근)
  const excluded = [];  // TBM 제외 (결근/휴무 등)
  
  for (let r = 2; r <= 17; r++) { // data index 2~17 = 5번째~20번째 행
    const row = data[r];
    if (!row) continue;
    
    const rank = String(row[0]).trim();    // B열: 직급
    const name = String(row[1]).trim();    // C열: 성명
    
    if (!name || name === '' || name === 'undefined') continue;
    
    const attendance = String(row[dayColIndex]).trim();
    
    // 결근 코드에 해당하면 TBM 제외 목록에 추가
    const isAbsent = ABSENT_CODES.some(code => attendance.includes(code));
    const displayName = (rank && rank !== '') ? `${name} ${rank}` : name;
    
    if (isAbsent) {
      excluded.push(displayName);
    } else {
      targets.push(displayName);
    }
  }
  
  return { targets, excluded };
}

// ============================================================
// 셀 값 설정 헬퍼 함수들
// ============================================================
function setCellValue(sheet, a1Notation, value) {
  sheet.getRange(a1Notation).setValue(value);
}

function setMergedCellValue(sheet, startCellA1, value) {
  // 병합된 셀의 경우 시작 셀에 값을 설정
  sheet.getRange(startCellA1).setValue(value);
}

// ============================================================
// 완료 요약 메시지 생성
// ============================================================
function buildTBMCellText(targets, excluded) {
  const lines = [];
  lines.push(`금일 TBM 대상 : ${targets.join(' / ') || '없음'}`);
  lines.push(`금일 TBM 제외 : ${excluded.join(' / ') || '없음'}`);
  return lines.join('\n');
}

function buildCompletionSummary(sheetName, dateString, prodInfo, hazardContent, eduContent, tbmTargets, tbmExcluded) {
  const lines = [];
  lines.push(`✅ [${sheetName}] 시트 작성 완료`);
  lines.push(`📅 날짜: ${dateString}`);
  lines.push('');
  lines.push('📋 조업 특이사항:');
  lines.push(buildSpecialNote(prodInfo));
  lines.push('');
  lines.push('⚠️ 교육 주제: ' + eduContent.title);
  lines.push('');
  lines.push(`👥 금일 TBM 대상 (${tbmTargets.length}명): ${tbmTargets.join(', ') || '없음'}`);
  lines.push(`🚫 금일 TBM 제외 (${tbmExcluded.length}명): ${tbmExcluded.join(', ') || '없음'}`);
  
  return lines.join('\n');
}

// ============================================================
// 유틸: 스크립트 속성 초기화 (최초 설정 또는 리셋 시 사용)
// ============================================================
function resetIndices() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '인덱스 초기화',
    '교육내용 및 위험요인 순번을 처음으로 초기화하시겠습니까?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('EDU_INDEX', '0');
    props.setProperty('HAZARD_INDEX', '0');
    ui.alert('초기화 완료', '순번이 초기화되었습니다.', ui.ButtonSet.OK);
  }
}

// ============================================================
// 유틸: 현재 순번 확인
// ============================================================
function checkCurrentIndices() {
  const props = PropertiesService.getScriptProperties();
  const eduIdx = parseInt(props.getProperty('EDU_INDEX') || '0');
  const hazardIdx = parseInt(props.getProperty('HAZARD_INDEX') || '0');
  
  SpreadsheetApp.getUi().alert(
    '현재 순번 확인',
    `다음에 사용될 교육내용: [${eduIdx + 1}] ${EDUCATION_REPERTOIRE[eduIdx].title}\n다음에 사용될 위험요인: ${HAZARD_REPERTOIRE[hazardIdx].hazard.substring(0, 30)}...`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

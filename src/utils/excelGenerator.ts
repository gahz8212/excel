import ExcelJS from "exceljs";
import { ReceiptData, TransferItem } from "../types/excel";
import { EXCEL_STYLES } from "../constants/styles";
interface ColumnConfig {
  header: [string, string?];
  key: keyof TransferItem;
  subKey?: keyof TransferItem;
  rowSpan?: number;
  colSpan?: number;
}
interface SummaryConfig {
  label: string;
  // ReceiptData의 키들 중 'list'만 제외한 키들만 허용
  key: Exclude<keyof ReceiptData, 'list'>; 
  unit: string;
  colSpan: number;
}

export const generateTransferReceipt = async (
  data: ReceiptData,
): Promise<ExcelJS.Workbook> => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("이체확인증");

  // 1. 컬럼 너비 설정
  sheet.columns = [
    { width: 15 }, // 이체일시
    { width: 15 }, // 출금은행
    { width: 15 }, // 출금계좌
    { width: 15 }, // 입금은행/계좌
    { width: 15 }, // 입금은행/계좌
    { width: 15 }, // 이체금액/수수료
    { width: 15 }, // 이체금액/수수료
    { width: 15 }, // 받는분/보내는분
  ];

  // 1. 메인 제목
  sheet.mergeCells("A1:H2");
  const titleCell = sheet.getCell("A1");
  titleCell.value = "확 인 증";
  titleCell.style = EXCEL_STYLES.MAIN_TITLE;

  // 1. 출력 일시 (H3 셀 부근)
  // 가로 2셀 병합해서 오른쪽 끝에 배치 (G3:H3)
  sheet.mergeCells("G5:H5");
  const printDateCell = sheet.getCell("G5");
  printDateCell.value = `출력일시: ${data.printDate}`;
  printDateCell.alignment = { horizontal: "right" };
  printDateCell.font = { size: 10 };

  // 2. 요약 정보 테이블 (4행 ~ 5행)
  const summaryConfig: SummaryConfig[] = [
    { label: "총 이체건수", key: "totalCount", unit: "건", colSpan: 2 },
    { label: "총 이체금액", key: "totalAmount", unit: "원", colSpan: 2 },
    {
      label: "수수료 총 발생건수",
      key: "totalFeeCount",
      unit: "건",
      colSpan: 2,
    },
    { label: "총 이체수수료", key: "totalFeeAmount", unit: "원", colSpan: 2 },
  ];

  renderSummary(sheet, 6, summaryConfig, data);

  const columnsConfig: ColumnConfig[] = [
    { header: ["이체일시"], key: "date", rowSpan: 2 }, // rowSpan:2 이므로 세로 병합
    { header: ["출금은행"], key: "outBank", rowSpan: 2 }, // 세로 병합
    { header: ["출금계좌"], key: "outAccount", rowSpan: 2 }, // 세로 병합
    {
      header: ["입금은행", "입금계좌번호"],
      key: "inBank",
      subKey: "inAccount",
      colSpan: 2,
    }, // 상하 분리
    {
      header: ["이체금액(원)", "수수료(원)"],
      key: "amount",
      subKey: "fee",
      colSpan: 2,
    }, // 상하 분리
    {
      header: ["받는분", "보내는분"],
      key: "receiver",
      subKey: "sender",
    }, // 상하 분리
  ];

  renderTableWithMerge(sheet, 9, columnsConfig, data.list);

  return workbook;
};

const renderTableWithMerge = (
  sheet: ExcelJS.Worksheet,
  startRow: number,
  config: ColumnConfig[],
  data: TransferItem[],
) => {
  let currentRow = startRow;

  // --- 1. 헤더 생성 로직 (상단 2행 점유) ---
  let headerCol = 1;
  config.forEach((col) => {
    const colSpan = col.colSpan || 1;
    const isRowSpan2 = col.rowSpan === 2; // rowspan: 2일 때만 아래 셀과 병합

    const r1 = currentRow;
    const c1 = headerCol;
    const r2 = currentRow + 1;
    const c2 = headerCol + colSpan - 1;

    if (isRowSpan2) {
      // 케이스 1: 세로 병합 헤더 (rowSpan: 2)
      if (r1 !== r2 || c1 !== c2) {
        sheet.mergeCells(r1, c1, r2, c2);
      }
      sheet.getCell(r1, c1).value = col.header[0];
    } else {
      // 케이스 2: 상하 분리 헤더 (rowSpan이 2가 아님)
      // 상단 헤더 가로 병합
      if (c1 !== c2) sheet.mergeCells(r1, c1, r1, c2);
      sheet.getCell(r1, c1).value = col.header[0];

      // 하단 헤더 가로 병합 (두 번째 헤더 텍스트가 있을 때만)
      if (c1 !== c2) sheet.mergeCells(r2, c1, r2, c2);
      if (col.header[1]) {
        sheet.getCell(r2, c1).value = col.header[1];
      }
    }

    // 헤더 스타일 적용 (전체 영역)
    applyStyleToRange(sheet, r1, c1, r2, c2, EXCEL_STYLES.TABLE_HEADER);
    headerCol += colSpan;
  });

  // 데이터 시작점: 헤더가 2개 행(currentRow, currentRow + 1)을 썼으므로 +2
  currentRow += 2;

  // --- 2. 데이터 주입 로직 (동일 원리) ---
  data.forEach((item) => {
    let currentCol = 1;
    config.forEach((col) => {
      const colSpan = col.colSpan || 1;
      const isRowSpan2 = col.rowSpan === 2;

      const r1 = currentRow;
      const c1 = currentCol;
      const r2 = currentRow + 1;
      const c2 = currentCol + colSpan - 1;

      if (isRowSpan2) {
        // 케이스 1: 세로 병합 데이터
        if (r1 !== r2 || c1 !== c2) {
          sheet.mergeCells(r1, c1, r2, c2);
        }
        sheet.getCell(r1, c1).value = item[col.key];
        applyStyleToRange(sheet, r1, c1, r2, c2, EXCEL_STYLES.TABLE_BODY);
      } else {
        // 케이스 2: 상하 분리 데이터
        if (c1 !== c2) sheet.mergeCells(r1, c1, r1, c2);
        sheet.getCell(r1, c1).value = item[col.key];

        if (c1 !== c2) sheet.mergeCells(r2, c1, r2, c2);
        if (col.subKey) {
          sheet.getCell(r2, c1).value = item[col.subKey as keyof TransferItem];
        }

        const style =
          col.key === "amount" || (col.subKey && col.subKey === "fee")
            ? EXCEL_STYLES.TABLE_BODY_RIGHT
            : EXCEL_STYLES.TABLE_BODY;

        applyStyleToRange(sheet, r1, c1, r1, c2, style);
        applyStyleToRange(sheet, r2, c1, r2, c2, style);
      }
      currentCol += colSpan;
    });
    currentRow += 2; // 다음 항목을 위해 2행 건너뜀
  });
};

const renderSummary = (
  sheet: ExcelJS.Worksheet,
  startRow: number,
  config: SummaryConfig[],
  data: ReceiptData,
) => {
  let currentCol = 1;

  config.forEach((item) => {
    const colSpan = item.colSpan || 1;
    const r1 = startRow;
    const r2 = startRow + 1;
    const c1 = currentCol;
    const c2 = currentCol + colSpan - 1;

    // 1. 병합
    sheet.mergeCells(r1, c1, r1, c2);
    sheet.mergeCells(r2, c1, r2, c2);

    // 2. 값 가져오기 (이제 rawValue는 절대 배열일 수 없습니다)
    const rawValue = data[item.key]; 
    
    // 라벨 입력
    sheet.getCell(r1, c1).value = item.label;
    
    // 값 입력 (숫자나 문자열만 남았으므로 안전함)
    const valueCell = sheet.getCell(r2, c1);
    valueCell.value = rawValue;

    // 3. 스타일 및 포맷 (테두리 포함)
    const labelStyle = EXCEL_STYLES.TABLE_HEADER;
    const valueStyle = typeof rawValue === 'number' 
      ? { ...EXCEL_STYLES.TABLE_BODY_RIGHT, numFmt: `#,##0"${item.unit}"` }
      : EXCEL_STYLES.TABLE_BODY;

    applyStyleToRange(sheet, r1, c1, r1, c2, labelStyle);
    applyStyleToRange(sheet, r2, c1, r2, c2, valueStyle as any);

    currentCol += colSpan;
  });
};
// 헬퍼 함수: 병합 범위 전체에 테두리 스타일을 입히기 위함
const applyStyleToRange = (
  sheet: any,
  r1: number,
  c1: number,
  r2: number,
  c2: number,
  style: any,
) => {
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      sheet.getCell(r, c).style = style;
    }
  }
};

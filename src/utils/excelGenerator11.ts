import ExcelJS from "exceljs";
import { EXCEL_STYLES } from "../constants/styles";

interface ColumnConfig {
  header: string;
  dataKeys?: string[]; // n단 데이터를 위한 키 배열
  children?: ColumnConfig[];
  colSpan?: number; // 사용자가 직접 지정하는 colSpan (선택적)
}



/**
 * 2. 최하단(Leaf) 컬럼들만 추출 (너비 설정 및 데이터 매핑용)
 */
const getLeafColumns = (config: ColumnConfig[]): ColumnConfig[] => {
  let leaves: ColumnConfig[] = [];
  config.forEach(col => {
    if (col.children && col.children.length > 0) {
      leaves = [...leaves, ...getLeafColumns(col.children)];
    } else {
      leaves.push(col);
    }
  });
  return leaves;
};



const getColSpan = (col: any): number => {
  if (col.colSpan) return col.colSpan; // 이 부분이 명시되어야 '이체일시'가 2칸이 됩니다.
  if (col.children && col.children.length > 0) {
    return col.children.reduce((acc: number, child: any) => acc + getColSpan(child), 0);
  }
  return 1;
};

/**
 * 2. 데이터 영역의 최대 높이 계산 (n단 데이터 대응)
 */
const getMaxDataHeight = (config: any[]): number => {
  let max = 1;
  const findMax = (cols: any[]) => {
    cols.forEach(col => {
      if (col.children && col.children.length > 0) {
        findMax(col.children);
      } else {
        const keyCount = col.dataKeys?.length || col.datakeys?.length || 0;
        if (keyCount > max) max = keyCount;
      }
    });
  };
  findMax(config);
  return max;
};

/**
 * 4. 헤더의 총 깊이 계산
 */
const getMaxHeaderDepth = (config: ColumnConfig[]): number => {
  let max = 0;
  config.forEach(col => {
    if (col.children && col.children.length > 0) {
      max = Math.max(max, 1 + getMaxHeaderDepth(col.children));
    } else {
      max = Math.max(max, 1);
    }
  });
  return max;
};

/**
 * 5. [재귀] 멀티 레벨 헤더 렌더링
 */
const renderMultiLevelHeader = (
  sheet: ExcelJS.Worksheet,
  config: ColumnConfig[],
  currentRow: number,
  totalHeaderDepth: number,
  startCol: number = 1,
) => {
  let currentCol = startCol;

  config.forEach((col) => {
    const colSpan = getColSpan(col); // 우리가 만든 colSpan 반영 함수
    const endCol = currentCol + colSpan - 1;

    if (col.children && col.children.length > 0) {
      // 상위 헤더 가로 병합
      sheet.mergeCells(currentRow, currentCol, currentRow, endCol);
      sheet.getCell(currentRow, currentCol).value = col.header;
      applyStyleToRange(
        sheet,
        currentRow,
        currentCol,
        currentRow,
        endCol,
        EXCEL_STYLES.TABLE_HEADER,
      );

      // 자식으로 내려감 (시작 위치는 그대로 currentCol)
      renderMultiLevelHeader(
        sheet,
        col.children,
        currentRow + 1,
        totalHeaderDepth - 1,
        currentCol,
      );
    } else {
      // 최하단 헤더 (가로 colSpan + 세로 totalHeaderDepth 병합)
      const endRow = currentRow + totalHeaderDepth - 1;
      sheet.mergeCells(currentRow, currentCol, endRow, endCol);
      sheet.getCell(currentRow, currentCol).value = col.header;
      applyStyleToRange(
        sheet,
        currentRow,
        currentCol,
        endRow,
        endCol,
        EXCEL_STYLES.TABLE_HEADER,
      );
    }

    // [중요] 사용한 너비만큼 확실히 점프!
    currentCol += colSpan;
  });
};

/**
 * 6. [재귀] 가변 n단 데이터 렌더링
 */
const renderFlexibleData = (
  sheet: ExcelJS.Worksheet,
  config: ColumnConfig[],
  item: any,
  currentRow: number,
  maxDataHeight: number,
  startCol: number = 1
) => {
  let currentCol = startCol;

  config.forEach((col) => {
    const colSpan = getColSpan(col);
    const endCol = currentCol + colSpan - 1;

    if (col.children && col.children.length > 0) {
      renderFlexibleData(sheet, col.children, item, currentRow, maxDataHeight, currentCol);
    } else {
      // 대소문자 구분 없이 dataKeys를 가져옴
      const keys = col.dataKeys || col.dataKeys || [];
      const keyCount = keys.length;

      if (keyCount === 0) {
        // 데이터 키가 없어도 칸은 확보 (테두리 유지)
        const r2 = currentRow + maxDataHeight - 1;
        sheet.mergeCells(currentRow, currentCol, r2, endCol);
        applyStyleToRange(sheet, currentRow, currentCol, r2, endCol, EXCEL_STYLES.TABLE_BODY);
      } else {
        keys.forEach((key: string, idx: number) => {
          const r1 = currentRow + idx;
          const isLastKey = idx === keyCount - 1;
          const val = item[key] ?? ""; // myData["date"] 등을 가져옴
          console.log('val:', val, 'key:', key);
          if (isLastKey && keyCount < maxDataHeight) {
            const r2 = currentRow + maxDataHeight - 1;
            sheet.mergeCells(r1, currentCol, r2, endCol);
            sheet.getCell(r1, currentCol).value = val;
            applyStyleToRange(sheet, r1, currentCol, r2, endCol, EXCEL_STYLES.TABLE_BODY);
          } else {
            sheet.mergeCells(r1, currentCol, r1, endCol);
            sheet.getCell(r1, currentCol).value = val;
            applyStyleToRange(sheet, r1, currentCol, r1, endCol, EXCEL_STYLES.TABLE_BODY);
          }
        });
      }
    }
    // 사용한 colSpan만큼 옆으로 정확히 이동
    currentCol += colSpan;
  });
};

/**
 * 7. [메인 실행 함수]
 */
export const generateSmartExcel = async (
  data: any[],
  config: ColumnConfig[],
) => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Report");

  const headerDepth = getMaxHeaderDepth(config);
  const dataHeight = getMaxDataHeight(config);

  // 헤더 그리기
  renderMultiLevelHeader(sheet, config, 1, headerDepth);

  // 데이터 그리기
  let currentDataRow = 1 + headerDepth;
  data.forEach((item) => {
    renderFlexibleData(sheet, config, item, currentDataRow, dataHeight);
    currentDataRow += dataHeight; // 한 항목이 차지하는 n단 높이만큼 이동
  });

  return workbook;
};

const applyStyleToRange = (
  sheet: ExcelJS.Worksheet,
  r1: number,
  c1: number,
  r2: number,
  c2: number,
  style: any,
) => {
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const cell = sheet.getCell(r, c);
      cell.fill = style.fill || cell.fill;
      cell.font = style.font || cell.font;
      cell.alignment = style.alignment || cell.alignment;
      cell.border = style.border || cell.border;
    }
  }
};

import { ExcelStyle } from '../types/excel';

export const EXCEL_STYLES: Record<string, ExcelStyle> = {
  MAIN_TITLE: {
    font: { name: '나눔고딕', size: 20, bold: true },
    alignment: { horizontal: 'center', vertical: 'middle' }
  },
  TABLE_HEADER: {
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } },
    font: { name: '나눔고딕', size: 10, bold: true },
    alignment: { horizontal: 'center', vertical: 'middle' },
    border: {
      top: { style: 'thin' }, bottom: { style: 'thin' },
      left: { style: 'thin' }, right: { style: 'thin' }
    }
  },
  TABLE_BODY: {
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } },
    font: { name: '나눔고딕', size: 10 },
    alignment: { horizontal: 'center', vertical: 'middle' },
    border: {
      top: { style: 'thin' }, bottom: { style: 'thin' },
      left: { style: 'thin' }, right: { style: 'thin' }
    }
  },
  TABLE_BODY_RIGHT: {
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } },
    font: { name: '나눔고딕', size: 10 },
    alignment: { horizontal: 'right', vertical: 'middle' },
    numFmt: '#,##0 "원"',
    border: {
      top: { style: 'thin' }, bottom: { style: 'thin' },
      left: { style: 'thin' }, right: { style: 'thin' }
    }
  }
};
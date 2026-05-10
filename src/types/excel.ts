import {  Style } from "exceljs";

export type ExcelStyle = Partial<Style>;
export interface TransferItem {
  date: string;
  outBank: string;
  outAccount: string;
  inBank: string;
  inAccount: string;
  amount: number;
  fee: number;
  receiver: string;
  sender: string;
}

// 전체 데이터 구조
export interface ReceiptData {
  printDate: string;
  totalCount: number;
  totalAmount: number;
  totalFeeCount: number;
  totalFeeAmount: number;
  list: TransferItem[];
}

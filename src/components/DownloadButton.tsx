// components/DownloadButton.tsx
"use client";

import { generateTransferReceipt } from '../utils/excelGenerator';
import type{ ReceiptData } from '../types/excel';
const mockData: ReceiptData = {
  printDate: "2026-04-17 14:30:22", // 출력일시
  totalCount: 996,                // 총 이체건수
  totalAmount: 11103211118,       // 총 이체금액
  totalFeeCount: 0,               // 이체수수료 총 발생건수
  totalFeeAmount: 0,              // 총 이체수수료
  
  // 상세 리스트 (이미지 하단부 데이터)
  list: [
    {
      date: "2025-09-12",
      outBank: "기업은행",
      outAccount: "예외출금(52000233501276)",
      inBank: "기업은행",
      inAccount: "52000233501010",
      amount: 2901000,
      fee: 0,
      receiver: "주식회사 세창전기",
      sender: "(주)세창전기"
    },
    // ... 추가 데이터들
  ]
};
export default function DownloadButton() {
  const handleDownload = async () => {
    // 1. 서버나 상태값으로부터 데이터 가져오기
    // const data = await fetch('/api/get-receipt-data').then(res => res.json());

    // 2. 엑셀 워크북 생성
    const workbook = await generateTransferReceipt(mockData);

    // 3. 브라우저에서 다운로드 실행
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `이체확인증_${mockData.printDate.split(' ')[0]}.xlsx`;
    anchor.click();
    window.URL.revokeObjectURL(url);
  };

  return <button onClick={handleDownload}>엑셀 다운로드</button>;
}
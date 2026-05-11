// components/DownloadButton.tsx
"use client";

import { generateSmartExcel } from '../utils/excelGenerator11';

import { data11 } from './data';
const myConfig = [
  { header: "이체일시", dataKeys: ["date"] },
  { header: "출금은행", dataKeys: ["bank"] },
  { header: "출금계좌", dataKeys: ["account"] },
  { header: "입금은행", children: [{ header: "입금계좌번호", dataKeys: ["inBank", "inAccount"] }] },


];

const myData = [
  { date: "오늘", bank: "국민은행", account: "123-456-789", inBank: "신한은행", inAccount: "987-654-321" },
];
export default function DownloadButton() {
  const handleDownload = async () => {
    // 1. 서버나 상태값으로부터 데이터 가져오기
    // const data = await fetch('/api/get-receipt-data').then(res => res.json());

    // 2. 엑셀 워크북 생성
    const workbook = await generateSmartExcel(myData, myConfig);

    // 3. 브라우저에서 다운로드 실행
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `이체확인증_${data11.printDate.split(' ')[0]}.xlsx`;
    anchor.click();
    window.URL.revokeObjectURL(url);
  };

  return <button onClick={handleDownload}>이체 확인증 다운로드</button>;
}
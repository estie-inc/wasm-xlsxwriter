"use client";
import initModule, { Workbook } from "wasm-xlsxwriter/web";

export default function Page() {
  return (
    <button
      onClick={async () => {
        const workbook = await generateXlsx();
        download(workbook);
      }}
    >
      Generate a xlsx file
    </button>
  );
}

const generateXlsx = async () => {
  await initModule();
  const workbook = new Workbook();
  const worksheet = workbook.addWorksheet();
  worksheet.write(0, 0, "Hello, world!");
  return workbook;
};

const download = (workbook: Workbook) => {
  const buffer = workbook.saveToBufferSync();
  const blob = new Blob([buffer]);
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "generated.xlsx";
  a.click();
  a.remove();
};

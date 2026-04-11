import xlsxInit, { Workbook } from "wasm-xlsxwriter/web";

const button = document.getElementById("generate") as HTMLButtonElement;
const status = document.getElementById("status") as HTMLParagraphElement;

button.addEventListener("click", async () => {
  status.textContent = "Generating...";

  await xlsxInit();

  const workbook = new Workbook();
  const worksheet = workbook.addWorksheet();
  worksheet.write(0, 0, "Hello from Vite!");
  worksheet.write(1, 0, new Date().toISOString());

  const buffer = workbook.saveToBufferSync();
  const blob = new Blob([buffer.slice()], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = "generated.xlsx";
  a.click();
  a.remove();
  URL.revokeObjectURL(url);

  status.textContent = "Done!";
});

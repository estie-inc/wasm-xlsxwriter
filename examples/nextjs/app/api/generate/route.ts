import { Workbook } from "wasm-xlsxwriter";

export async function GET() {
  const workbook = new Workbook();
  const worksheet = workbook.addWorksheet();
  worksheet.write(0, 0, "Hello from App Router Server");
  const buffer = workbook.saveToBufferSync();

  return new Response(buffer.slice().buffer, {
    headers: {
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": "attachment; filename=generated.xlsx",
    },
  });
}

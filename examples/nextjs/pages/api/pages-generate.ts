import type { NextApiRequest, NextApiResponse } from "next";
import { Workbook } from "wasm-xlsxwriter";

export default function handler(_req: NextApiRequest, res: NextApiResponse) {
  const workbook = new Workbook();
  const worksheet = workbook.addWorksheet();
  worksheet.write(0, 0, "Hello from Pages Router Server");
  const buffer = workbook.saveToBufferSync();

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.setHeader(
    "Content-Disposition",
    "attachment; filename=generated.xlsx"
  );
  res.send(Buffer.from(buffer));
}

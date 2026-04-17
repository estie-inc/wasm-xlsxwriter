"use client";

import { useState } from "react";
import initModule, { Workbook } from "wasm-xlsxwriter/web";

export default function Page() {
  const [status, setStatus] = useState("");

  const generate = async () => {
    try {
      await initModule();
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      worksheet.write(0, 0, "Hello from App Router Client");
      const buffer = workbook.saveToBufferSync();
      setStatus(`Generated ${buffer.length} bytes`);
    } catch (e) {
      setStatus(`Error: ${e}`);
    }
  };

  return (
    <div>
      <h1>App Router - Client Side</h1>
      <button onClick={generate} data-testid="generate">
        Generate
      </button>
      <p data-testid="status">{status}</p>
    </div>
  );
}

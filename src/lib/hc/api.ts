import { createServerFn } from "@tanstack/react-start";
import type { HcRecord } from "./types";

export const loadHc = createServerFn({ method: "GET" }).handler(async () => {
  const { loadOrSeed } = await import("./excel.server");
  return loadOrSeed();
});

export const saveHc = createServerFn({ method: "POST" })
  .validator((data: { records: HcRecord[] }) => data)
  .handler(async ({ data }) => {
    const { writeLibro } = await import("./excel.server");
    return writeLibro(data.records);
  });

export const downloadHcExcel = createServerFn({ method: "POST" })
  .validator((data: { records: HcRecord[] }) => data)
  .handler(async ({ data }) => {
    const { workbookBase64 } = await import("./excel.server");
    return workbookBase64(data.records);
  });

export const importHcExcel = createServerFn({ method: "POST" })
  .validator((data: { base64: string }) => data)
  .handler(async ({ data }) => {
    const { importFromBuffer } = await import("./excel.server");
    const buf = Buffer.from(data.base64, "base64");
    return importFromBuffer(buf);
  });

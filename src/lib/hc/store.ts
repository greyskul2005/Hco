import { create } from "zustand";
import { downloadHcExcel, importHcExcel, loadHc, saveHc } from "./api";
import { computeKpis, emptyRecord, impactoDe, qualityOf } from "./calc";
import { RESPONSABLES } from "./columns";
import type { HcKpis, HcRecord } from "./types";

function triggerDownload(base64: string, name: string) {
  const bin = atob(base64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  URL.revokeObjectURL(url);
}

interface HcState {
  records: HcRecord[];
  kpis: HcKpis | null;
  source: string;
  loading: boolean;
  saving: boolean;
  error: string | null;
  loaded: boolean;
  load: () => Promise<void>;
  persist: (next: HcRecord[]) => Promise<void>;
  addRecord: (partial: Partial<HcRecord>) => Promise<HcRecord>;
  updateRecord: (matricula: string, patch: Partial<HcRecord>) => Promise<void>;
  download: () => Promise<void>;
  importFile: (file: File) => Promise<void>;
}

export const useHcStore = create<HcState>((set, get) => ({
  records: [],
  kpis: null,
  source: "",
  loading: false,
  saving: false,
  error: null,
  loaded: false,
  load: async () => {
    set({ loading: true, error: null });
    try {
      const data = await loadHc();
      set({
        records: data.records,
        kpis: data.kpis,
        source: data.source,
        loading: false,
        loaded: true,
      });
    } catch (e) {
      set({
        error: e instanceof Error ? e.message : "No se pudo leer el libro Excel",
        loading: false,
      });
    }
  },
  persist: async (next) => {
    const normalized = next.map((r) => ({
      ...r,
      ImpactoHC: impactoDe(r.TipoRegistro, r.Estado),
      EstadoCalidad: qualityOf(r, next),
      ResponsableUnidad: r.UnidadOrganizativa
        ? (RESPONSABLES[r.UnidadOrganizativa] ?? r.ResponsableUnidad)
        : r.ResponsableUnidad,
    }));
    set({ records: normalized, kpis: computeKpis(normalized), saving: true, error: null });
    try {
      const res = await saveHc({ data: { records: normalized } });
      set({ kpis: res.kpis, saving: false, source: "excel" });
    } catch (e) {
      set({
        saving: false,
        error: e instanceof Error ? e.message : "No se pudo guardar el Excel",
      });
    }
  },
  addRecord: async (partial) => {
    const rec: HcRecord = {
      ...emptyRecord(),
      ...partial,
      DatoSintetico: "Si",
      Fuente: partial.Fuente ?? "Manual",
    };
    rec.ImpactoHC = impactoDe(rec.TipoRegistro, rec.Estado);
    if (!rec.Matricula) rec.Matricula = `SYN-W${String(Date.now()).slice(-6)}`;
    if (!rec.IDPlaza && rec.TipoRegistro === "Incremento") rec.IDPlaza = `PLZ-WEB-${rec.Matricula}`;
    if (!rec.Correo && rec.Matricula) rec.Correo = `${rec.Matricula.toLowerCase()}@demo.hccontrol.local`;
    await get().persist([...get().records, rec]);
    return rec;
  },
  updateRecord: async (matricula, patch) => {
    const next = get().records.map((r) =>
      r.Matricula === matricula ? { ...r, ...patch } : r,
    );
    await get().persist(next);
  },
  download: async () => {
    const b64 = await downloadHcExcel({ data: { records: get().records } });
    triggerDownload(b64, "HC_Control_Operativo.xlsx");
  },
  importFile: async (file) => {
    const buf = await file.arrayBuffer();
    const bytes = new Uint8Array(buf);
    let binary = "";
    bytes.forEach((b) => {
      binary += String.fromCharCode(b);
    });
    const base64 = btoa(binary);
    set({ loading: true, error: null });
    try {
      const data = await importHcExcel({ data: { base64 } });
      set({
        records: data.records,
        kpis: data.kpis,
        source: data.source,
        loading: false,
        loaded: true,
      });
    } catch (e) {
      set({
        loading: false,
        error: e instanceof Error ? e.message : "Importación no válida",
      });
      throw e;
    }
  },
}));

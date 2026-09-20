import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";
import type { HcPayload, HcRecord } from "./types";
import { COLUMNAS, COLUMN_ORDER, HEADER_BY_VISIBLE, TEXT_KEYS } from "./columns";
import { computeKpis, ECUACION, impactoDe, MARCA } from "./calc";

const NAVY = "FF0B1F3A";
const BLUE = "FF1D4ED8";
const CYAN = "FF0E7490";
const WHITE = "FFFFFFFF";
const INK = "FF0F172A";
const PAPER = "FFF4F7FB";
const INPUT = "FFFFF4CC";
const CALC = "FFE7F0FA";
const OK = "FFD1FAE5";
const ERR = "FFFEE2E2";
const INC = "FFDBEAFE";
const SUS = "FFCFFAFE";
const FUERA = "FFE2E8F0";
const DUP = "FFFDE8E8";
const INCOMP = "FFFEF3C7";
const MUTED = "FF64748B";

export const LIBRO_NAME = "HC_Control_Operativo.xlsx";

function libroPaths(): string[] {
  const cwd = process.cwd();
  return [
    path.join(cwd, "data", "libros", LIBRO_NAME),
    path.join(cwd, "public", "libros", LIBRO_NAME),
    path.join("/tmp", LIBRO_NAME),
  ];
}

function argbFill(argb: string): ExcelJS.FillPattern {
  return { type: "pattern", pattern: "solid", fgColor: { argb } };
}

function font(opts: Partial<ExcelJS.Font> = {}): Partial<ExcelJS.Font> {
  return { name: "Calibri", size: 11, color: { argb: INK }, ...opts };
}

function thinBorder(): Partial<ExcelJS.Borders> {
  const s: ExcelJS.Border = { style: "thin", color: { argb: "FFD0D7E2" } };
  return { top: s, left: s, bottom: s, right: s };
}

function paint(cell: ExcelJS.Cell, fill: string, opts?: { bold?: boolean; color?: string; size?: number; align?: Partial<ExcelJS.Alignment> }) {
  cell.fill = argbFill(fill);
  cell.font = font({
    bold: opts?.bold,
    color: { argb: opts?.color ?? INK },
    size: opts?.size ?? 11,
  });
  cell.alignment = opts?.align ?? { vertical: "middle", wrapText: true };
  cell.border = thinBorder();
}

function banner(ws: ExcelJS.Worksheet, lastCol: number, title: string, subtitle: string) {
  ws.mergeCells(1, 1, 1, lastCol);
  ws.mergeCells(2, 1, 2, lastCol);
  ws.mergeCells(3, 1, 3, lastCol);
  const a1 = ws.getCell(1, 1);
  a1.value = title;
  paint(a1, NAVY, { bold: true, color: WHITE, size: 18, align: { vertical: "middle", horizontal: "left", indent: 1 } });
  ws.getRow(1).height = 28;
  const a2 = ws.getCell(2, 1);
  a2.value = subtitle;
  paint(a2, BLUE, { color: WHITE, size: 11, align: { vertical: "middle", indent: 1 } });
  const a3 = ws.getCell(3, 1);
  a3.value = MARCA;
  paint(a3, "FFCFFAFE", { color: CYAN, size: 10, align: { vertical: "middle", indent: 1 } });
  ws.views = [{ state: "frozen", ySplit: 3, showGridLines: false }];
  ws.pageSetup = {
    orientation: "landscape",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    paperSize: 9,
  };
}

function rowFill(r: HcRecord): string {
  if (r.EstadoCalidad === "Duplicado") return DUP;
  if (r.EstadoCalidad === "Incompleto") return INCOMP;
  if (r.TipoRegistro === "Incremento") return INC;
  if (r.TipoRegistro === "Sustitucion") return SUS;
  if (r.TipoRegistro === "Fuera_HC") return FUERA;
  return WHITE;
}

export async function buildWorkbook(records: HcRecord[]): Promise<ExcelJS.Workbook> {
  const k = computeKpis(records);
  const wb = new ExcelJS.Workbook();
  wb.creator = "HC Control";
  wb.created = new Date("2026-09-20");
  wb.modified = new Date();
  wb.title = "HC Control · Libro operativo";
  wb.description = MARCA;

  buildPortada(wb, k, records);
  buildControl(wb, records, k);
  buildMovimientos(wb, records);
  buildReconciliacion(wb, k);
  buildCierre(wb, k);
  buildCalidad(wb, records);
  buildParametros(wb, k);
  buildLeyenda(wb);
  return wb;
}

function buildPortada(wb: ExcelJS.Workbook, k: ReturnType<typeof computeKpis>, records: HcRecord[]) {
  const ws = wb.addWorksheet("PORTADA", { properties: { tabColor: { argb: NAVY } } });
  banner(ws, 10, "HC CONTROL · LIBRO OPERATIVO", "Aplicativo web · almacenamiento en Excel · escenario 514 + 13 − 9 + 9 = 527");
  ws.getRow(1).height = 32;
  const kpis: [string, string | number, string][] = [
    ["HC actual", k.hc_actual, NAVY],
    ["Incrementos", k.incrementos, BLUE],
    ["Salidas CEV", k.salidas_cev, CYAN],
    ["Sustituciones", k.sustituciones, "FF155E75"],
    ["HC previsto", k.hc_previsto_final, k.ecuacion_ok ? "FF065F46" : "FF9F1239"],
  ];
  kpis.forEach((item, i) => {
    const c = 1 + i * 2;
    ws.mergeCells(5, c, 5, c + 1);
    ws.mergeCells(6, c, 7, c + 1);
    const lab = ws.getCell(5, c);
    lab.value = item[0];
    paint(lab, item[2], { bold: true, color: WHITE, size: 10, align: { vertical: "middle", horizontal: "center" } });
    const val = ws.getCell(6, c);
    val.value = item[1];
    paint(val, item[2], { bold: true, color: WHITE, size: 28, align: { vertical: "middle", horizontal: "center" } });
    ws.getCell(6, c + 1).fill = argbFill(item[2]);
    ws.getCell(7, c).fill = argbFill(item[2]);
    ws.getCell(7, c + 1).fill = argbFill(item[2]);
    ws.getCell(5, c + 1).fill = argbFill(item[2]);
  });
  ws.getRow(6).height = 36;
  ws.getRow(7).height = 12;

  ws.mergeCells("A9:J9");
  const eq = ws.getCell("A9");
  eq.value = `Ecuación oficial: ${ECUACION}   ·   Resultado: ${k.ecuacion_ok ? "OK" : "ERROR"}   ·   CEV neto: ${k.impacto_cev_neto}`;
  paint(eq, k.ecuacion_ok ? OK : ERR, { bold: true, size: 14, align: { vertical: "middle", horizontal: "center" } });
  ws.getRow(9).height = 26;

  ws.getCell("A11").value = "Mes";
  ws.getCell("B11").value = "Etiqueta";
  ws.getCell("C11").value = "HC calculado";
  ws.getCell("D11").value = "HC esperado";
  ws.getCell("E11").value = "Estado";
  for (let c = 1; c <= 5; c++) paint(ws.getCell(11, c), NAVY, { bold: true, color: WHITE, align: { horizontal: "center" } });
  k.evolucion.forEach((e, i) => {
    const r = 12 + i;
    ws.getCell(r, 1).value = e.mes;
    ws.getCell(r, 1).numFmt = "@";
    ws.getCell(r, 2).value = e.etiqueta;
    ws.getCell(r, 3).value = e.hc;
    ws.getCell(r, 4).value = e.esperado;
    ws.getCell(r, 5).value = e.ok ? "OK" : "ERROR";
    for (let c = 1; c <= 5; c++) {
      paint(ws.getCell(r, c), e.ok ? OK : ERR, { bold: c >= 3, align: { horizontal: c >= 3 ? "center" : "left" } });
    }
  });

  ws.mergeCells("A18:J18");
  ws.getCell("A18").value = "Distribución base por unidad (fotografía agosto = 514). Colores idénticos a los de la hoja CONTROL_HC y del aplicativo web.";
  paint(ws.getCell("A18"), PAPER, { size: 10 });

  const byUnit = new Map<string, number>();
  for (const rec of records) {
    if (rec.TipoRegistro === "Base" && rec.ComputaHC === "Si") {
      byUnit.set(rec.UnidadOrganizativa, (byUnit.get(rec.UnidadOrganizativa) ?? 0) + 1);
    }
  }
  ws.getCell("A19").value = "Unidad";
  ws.getCell("B19").value = "HC base";
  paint(ws.getCell("A19"), CYAN, { bold: true, color: WHITE });
  paint(ws.getCell("B19"), CYAN, { bold: true, color: WHITE });
  let ur = 20;
  for (const [u, n] of [...byUnit.entries()].sort((a, b) => b[1] - a[1])) {
    ws.getCell(ur, 1).value = u;
    ws.getCell(ur, 2).value = n;
    paint(ws.getCell(ur, 1), WHITE);
    paint(ws.getCell(ur, 2), CALC, { bold: true, align: { horizontal: "center" } });
    ur += 1;
  }

  ws.mergeCells("D19:J22");
  ws.getCell("D19").value =
    "Este archivo ES el almacén del sistema. Cada alta, baja o sustitución hecha en el aplicativo web se escribe aquí con el mismo color, la misma matrícula (texto) y las mismas reglas. Ábralo en Excel o LibreOffice: lo que ve es lo que gobierna el headcount.";
  paint(ws.getCell("D19"), "FFECFEFF", { size: 12, align: { wrapText: true, vertical: "middle", indent: 1 } });

  for (let c = 1; c <= 10; c++) ws.getColumn(c).width = 16;
  ws.getColumn(2).width = 36;
}

function buildControl(wb: ExcelJS.Workbook, records: HcRecord[], k: ReturnType<typeof computeKpis>) {
  const ws = wb.addWorksheet("CONTROL_HC", { properties: { tabColor: { argb: BLUE } } });
  const last = COLUMN_ORDER.length;
  banner(
    ws,
    last,
    "CONTROL_HC · registro maestro (almacén Excel)",
    `${records.length} filas  ·  HC actual ${k.hc_actual}  ·  HC previsto ${k.hc_previsto_final}  ·  celdas amarillas = entrada  ·  azules = calculadas`,
  );
  const headerRow = 5;
  COLUMNAS.forEach((col, i) => {
    const cell = ws.getCell(headerRow, i + 1);
    cell.value = col.visible;
    paint(cell, NAVY, { bold: true, color: WHITE, size: 10, align: { wrapText: true, vertical: "middle", horizontal: "center" } });
  });
  ws.getRow(headerRow).height = 32;
  records.forEach((rec, idx) => {
    const r = headerRow + 1 + idx;
    const fill = rowFill(rec);
    COLUMN_ORDER.forEach((key, i) => {
      const cell = ws.getCell(r, i + 1);
      const raw = rec[key as keyof HcRecord];
      cell.value = raw as ExcelJS.CellValue;
      const isCalc = key === "ImpactoHC";
      paint(cell, isCalc ? CALC : fill, { size: 10 });
      if (TEXT_KEYS.has(key)) cell.numFmt = "@";
      if (key.startsWith("Fecha") && raw) cell.numFmt = "YYYY-MM-DD";
    });
  });
  const end = headerRow + records.length;
  ws.autoFilter = { from: { row: headerRow, column: 1 }, to: { row: end, column: last } };
  ws.views = [{ state: "frozen", xSplit: 2, ySplit: headerRow, showGridLines: false }];
  ws.getColumn(1).width = 14;
  ws.getColumn(2).width = 28;
  for (let c = 3; c <= last; c++) ws.getColumn(c).width = 16;
  ws.getColumn(25).width = 42;
  ws.getColumn(26).width = 42;
}

function buildMovimientos(wb: ExcelJS.Workbook, records: HcRecord[]) {
  const ws = wb.addWorksheet("MOVIMIENTOS", { properties: { tabColor: { argb: CYAN } } });
  const headers = [
    "Matrícula",
    "Nombre",
    "Tipo",
    "Programa",
    "Unidad",
    "Estado",
    "Fecha prevista alta",
    "Fecha prevista baja",
    "Mes efecto",
    "Impacto HC",
    "ID plaza",
    "Matrícula sustituida",
    "Motivo",
  ];
  banner(ws, 13, "MOVIMIENTOS PREVISTOS", "Azul = incremento (crea plaza). Cian = CEV (reutiliza plaza, impacto 0).");
  headers.forEach((h, i) => paint(ws.getCell(5, i + 1), NAVY, { bold: true, color: WHITE, align: { horizontal: "center", wrapText: true } }));
  ws.getRow(5).height = 28;
  const movs = records.filter(
    (r) =>
      r.TipoRegistro === "Incremento" ||
      r.TipoRegistro === "Sustitucion" ||
      (r.TipoRegistro === "Base" && r.Programa === "CEV"),
  );
  movs.forEach((r, i) => {
    const row = 6 + i;
    const vals = [
      r.Matricula,
      r.NombreCompleto,
      r.TipoRegistro,
      r.Programa,
      r.UnidadOrganizativa,
      r.Estado,
      r.FechaPrevistaAlta,
      r.FechaPrevistaBaja,
      r.MesEfecto,
      r.ImpactoHC,
      r.IDPlaza,
      r.MatriculaSustituida,
      r.Motivo,
    ];
    const fill = r.TipoRegistro === "Incremento" ? INC : SUS;
    vals.forEach((v, c) => {
      const cell = ws.getCell(row, c + 1);
      cell.value = v as ExcelJS.CellValue;
      paint(cell, c === 9 ? CALC : fill);
      if (c === 0 || c === 10 || c === 11) cell.numFmt = "@";
    });
  });
  ws.autoFilter = { from: "A5", to: `M${5 + movs.length}` };
  ws.views = [{ state: "frozen", ySplit: 5, showGridLines: false }];
  headers.forEach((_, i) => {
    ws.getColumn(i + 1).width = i === 12 || i === 1 ? 36 : 18;
  });
}

function buildReconciliacion(wb: ExcelJS.Workbook, k: ReturnType<typeof computeKpis>) {
  const ws = wb.addWorksheet("RECONCILIACION", { properties: { tabColor: { argb: "FF065F46" } } });
  banner(ws, 5, "RECONCILIACIÓN OFICIAL 514 → 527", "No reinterpretar estas cifras. Confirmadas por el escenario de negocio.");
  ["Paso", "Descripción", "Signo", "Valor", "Acumulado"].forEach((h, i) =>
    paint(ws.getCell(5, i + 1), NAVY, { bold: true, color: WHITE, align: { horizontal: "center" } }),
  );
  const rows: [string, string, string, number, number, string][] = [
    ["1", "HC base a cierre de agosto de 2026", "=", 514, 514, CALC],
    ["2", "Incrementos iniciales (6 becarios + 1 RRLL)", "+", 7, 521, INC],
    ["3", "Posiciones adicionales octubre", "+", 2, 523, INC],
    ["4", "Posiciones adicionales noviembre", "+", 3, 526, INC],
    ["5", "Posición adicional diciembre", "+", 1, 527, INC],
    ["6", "Salidas programa CEV", "−", 9, 518, SUS],
    ["7", "Sustituciones CEV (reutilizan plaza)", "+", 9, 527, SUS],
  ];
  rows.forEach((row, i) => {
    row.slice(0, 5).forEach((v, c) => {
      const cell = ws.getCell(6 + i, c + 1);
      cell.value = v as ExcelJS.CellValue;
      paint(cell, row[5], { bold: c >= 3, size: c >= 3 ? 14 : 11, align: { horizontal: c >= 2 ? "center" : "left" } });
    });
  });
  ws.mergeCells("A14:C14");
  ws.getCell("A14").value = "Ecuación oficial";
  paint(ws.getCell("A14"), NAVY, { bold: true, color: WHITE, size: 14 });
  ws.getCell("D14").value = ECUACION;
  paint(ws.getCell("D14"), k.ecuacion_ok ? OK : ERR, { bold: true, size: 14, align: { horizontal: "center" } });
  ws.getCell("E14").value = k.ecuacion_ok ? "OK" : "ERROR";
  paint(ws.getCell("E14"), k.ecuacion_ok ? OK : ERR, { bold: true, size: 16, align: { horizontal: "center" } });
  ws.mergeCells("A16:E16");
  ws.getCell("A16").value =
    "El acumulado 518 del paso 6 es aritmético intermedio. El programa CEV se ejecuta pareado (DEC-006): impacto neto 0. Incremento y sustitución no se mezclan.";
  paint(ws.getCell("A16"), PAPER, { size: 10 });
  ws.getColumn(1).width = 10;
  ws.getColumn(2).width = 56;
  ws.getColumn(3).width = 12;
  ws.getColumn(4).width = 36;
  ws.getColumn(5).width = 16;
}

function buildCierre(wb: ExcelJS.Workbook, k: ReturnType<typeof computeKpis>) {
  const ws = wb.addWorksheet("CIERRE", { properties: { tabColor: { argb: "FF1E3A5F" } } });
  const headers = ["Mes", "Etiqueta", "Incrementos del mes", "Salidas CEV", "Sustituciones", "HC cierre", "Esperado", "Estado"];
  banner(ws, 8, "CIERRE MENSUAL", "Agosto 2026 cerrado en 514 (confirmado). El resto es previsión. No cerrar si PORTADA = ERROR.");
  headers.forEach((h, i) => paint(ws.getCell(5, i + 1), NAVY, { bold: true, color: WHITE, align: { horizontal: "center" } }));
  const incM: Record<string, number> = { "2026-08": 0, "2026-09": 7, "2026-10": 2, "2026-11": 3, "2026-12": 1 };
  k.evolucion.forEach((e, i) => {
    const vals = [
      e.mes,
      e.etiqueta,
      incM[e.mes] ?? 0,
      e.mes === "2026-11" ? 9 : 0,
      e.mes === "2026-11" ? 9 : 0,
      e.hc,
      e.esperado,
      e.ok ? (e.mes === "2026-08" ? "Cerrado" : "Previsto") : "ERROR",
    ];
    vals.forEach((v, c) => {
      const cell = ws.getCell(6 + i, c + 1);
      cell.value = v as ExcelJS.CellValue;
      paint(cell, e.ok ? (e.mes === "2026-08" ? OK : CALC) : ERR, { bold: c >= 5 });
      if (c === 0) cell.numFmt = "@";
    });
  });
  headers.forEach((_, i) => {
    ws.getColumn(i + 1).width = i === 1 ? 36 : 18;
  });
}

function buildCalidad(wb: ExcelJS.Workbook, records: HcRecord[]) {
  const ws = wb.addWorksheet("CALIDAD", { properties: { tabColor: { argb: "FFB45309" } } });
  banner(ws, 6, "INCIDENCIAS DE CALIDAD", "Filas rosa = duplicado. Ámbar = incompleto. Operar por excepción.");
  ["Matrícula", "Tipo", "Severidad", "Campo", "Estado", "Detalle"].forEach((h, i) =>
    paint(ws.getCell(5, i + 1), NAVY, { bold: true, color: WHITE, align: { horizontal: "center" } }),
  );
  const issues = records.filter((r) => r.EstadoCalidad !== "Completo");
  issues.forEach((r, i) => {
    const campo = !r.UnidadOrganizativa ? "Unidad organizativa" : !r.Correo ? "Correo" : "Matrícula";
    const vals = [
      r.Matricula,
      r.EstadoCalidad,
      r.EstadoCalidad === "Duplicado" ? "Alta" : "Media",
      campo,
      "Abierta",
      r.Observaciones,
    ];
    const fill = r.EstadoCalidad === "Duplicado" ? DUP : INCOMP;
    vals.forEach((v, c) => {
      const cell = ws.getCell(6 + i, c + 1);
      cell.value = v;
      paint(cell, fill);
      if (c === 0) cell.numFmt = "@";
    });
  });
  ws.getColumn(1).width = 16;
  ws.getColumn(6).width = 70;
}

function buildParametros(wb: ExcelJS.Workbook, k: ReturnType<typeof computeKpis>) {
  const ws = wb.addWorksheet("PARAMETROS", { properties: { tabColor: { argb: MUTED } } });
  banner(ws, 4, "PARÁMETROS CONFIRMADOS Y CATÁLOGOS", "Cifras 514 / 13 / 9 / 9 / 527 = dato confirmado. Catálogos de unidad = supuesto de diseño.");
  [
    ["HC_Base_Cierre_Ago2026", 514, "Confirmado"],
    ["Incrementos_totales", 13, "Confirmado"],
    ["Salidas_CEV", 9, "Confirmado"],
    ["Sustituciones_CEV", 9, "Confirmado"],
    ["HC_Estructural_Final", 527, "Confirmado"],
    ["Ecuacion", ECUACION, "Confirmado"],
    ["HC_actual_libro", k.hc_actual, "Calculado"],
    ["HC_previsto_libro", k.hc_previsto_final, "Calculado"],
  ].forEach((row, i) => {
    paint(ws.getCell(5 + i, 1), PAPER, { bold: true });
    ws.getCell(5 + i, 1).value = row[0];
    const v = ws.getCell(5 + i, 2);
    v.value = row[1] as ExcelJS.CellValue;
    paint(v, CALC, { bold: true, size: 14, align: { horizontal: "center" } });
    ws.getCell(5 + i, 3).value = row[2];
    paint(ws.getCell(5 + i, 3), INPUT);
  });
  ws.getColumn(1).width = 32;
  ws.getColumn(2).width = 36;
  ws.getColumn(3).width = 18;
}

function buildLeyenda(wb: ExcelJS.Workbook) {
  const ws = wb.addWorksheet("LEYENDA", { properties: { tabColor: { argb: CYAN } } });
  banner(ws, 3, "LEYENDA VISUAL · Excel = aplicativo", "Los mismos colores se usan en el libro y en la web. Así el dato es reconocible al abrir el archivo.");
  const items: [string, string, string][] = [
    ["Incremento", INC, "Crea plaza nueva. Impacto HC = +1. No usar para CEV."],
    ["Sustitución / CEV", SUS, "Reutiliza plaza. Impacto HC = 0. Neto del programa CEV = 0."],
    ["Fuera de HC", FUERA, "No computa. No altera 514 ni 527."],
    ["Duplicado", DUP, "Matrícula repetida. Bloquea el cierre si computa dos veces."],
    ["Incompleto", INCOMP, "Falta un campo obligatorio. No debe computar hasta completar."],
    ["Celda de entrada", INPUT, "Amarillo: el operador puede escribir."],
    ["Celda calculada", CALC, "Azul: la escribe el sistema (Impacto HC, KPIs)."],
    ["OK", OK, "Control en verde: la cifra cuadra."],
    ["ERROR", ERR, "Control en rojo: no cerrar el mes."],
    ["Encabezado", NAVY, "Azul oscuro corporativo del sistema HC Control."],
  ];
  ["Color", "Significado", "Regla"].forEach((h, i) =>
    paint(ws.getCell(5, i + 1), NAVY, { bold: true, color: WHITE, align: { horizontal: "center" } }),
  );
  items.forEach((it, i) => {
    const cell = ws.getCell(6 + i, 1);
    cell.value = it[0];
    paint(cell, it[1], { bold: true, color: it[1] === NAVY ? WHITE : INK, align: { horizontal: "center" } });
    ws.getCell(6 + i, 2).value = it[0];
    paint(ws.getCell(6 + i, 2), WHITE);
    ws.getCell(6 + i, 3).value = it[2];
    paint(ws.getCell(6 + i, 3), WHITE);
    ws.getRow(6 + i).height = 22;
  });
  ws.getColumn(1).width = 24;
  ws.getColumn(2).width = 22;
  ws.getColumn(3).width = 70;
}

function recordFromRow(ws: ExcelJS.Worksheet, rowNumber: number, headerMap: string[]): HcRecord | null {
  const rec = {} as Record<string, string | number>;
  let has = false;
  headerMap.forEach((key, i) => {
    if (!key) return;
    const cell = ws.getCell(rowNumber, i + 1);
    let v: string | number = "";
    const raw = cell.value;
    if (raw == null || raw === "") v = "";
    else if (typeof raw === "object" && "text" in (raw as object)) v = String((raw as { text: string }).text);
    else if (raw instanceof Date) v = raw.toISOString().slice(0, 10);
    else v = typeof raw === "number" ? raw : String(raw);
    rec[key] = v;
    if (v !== "") has = true;
  });
  if (!has || !rec.Matricula) return null;
  rec.ImpactoHC = impactoDe(String(rec.TipoRegistro ?? ""), String(rec.Estado ?? ""));
  return rec as unknown as HcRecord;
}

export function parseWorkbook(wb: ExcelJS.Workbook): HcRecord[] {
  const ws = wb.getWorksheet("CONTROL_HC") ?? wb.worksheets.find((s) => s.name.toLowerCase().includes("control")) ?? wb.worksheets[1];
  if (!ws) return [];
  let headerRow = 1;
  for (let r = 1; r <= 12; r++) {
    const v = String(ws.getCell(r, 1).value ?? "");
    if (v === "Matrícula" || v === "Matricula") {
      headerRow = r;
      break;
    }
  }
  const headerMap: string[] = [];
  for (let c = 1; c <= COLUMN_ORDER.length + 2; c++) {
    const label = String(ws.getCell(headerRow, c).value ?? "").trim();
    headerMap.push(HEADER_BY_VISIBLE[label] ?? (COLUMN_ORDER.includes(label as (typeof COLUMN_ORDER)[number]) ? label : ""));
  }
  const records: HcRecord[] = [];
  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const rec = recordFromRow(ws, r, headerMap);
    if (rec) records.push(rec);
  }
  return records;
}

async function persist(buf: Uint8Array) {
  const data = Buffer.from(buf);
  for (const p of libroPaths()) {
    try {
      await mkdir(path.dirname(p), { recursive: true });
      await writeFile(p, data);
    } catch {
      /* deploy filesystem may be read-only */
    }
  }
}

export async function writeLibro(records: HcRecord[]): Promise<{ kpis: ReturnType<typeof computeKpis>; bytes: number }> {
  const wb = await buildWorkbook(records);
  const buf = new Uint8Array(await wb.xlsx.writeBuffer());
  await persist(buf);
  return { kpis: computeKpis(records), bytes: buf.byteLength };
}

export async function readLibroBuffer(): Promise<Buffer | null> {
  for (const p of libroPaths()) {
    try {
      return await readFile(p);
    } catch {
      continue;
    }
  }
  return null;
}

export async function loadOrSeed(): Promise<{ records: HcRecord[]; kpis: ReturnType<typeof computeKpis>; source: string }> {
  const existing = await readLibroBuffer();
  if (existing) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(existing);
    const records = parseWorkbook(wb);
    if (records.length > 0) {
      return { records, kpis: computeKpis(records), source: "excel" };
    }
  }
  const seedPath = path.join(process.cwd(), "src", "data", "hc_control.json");
  const raw = await readFile(seedPath, "utf8");
  const payload = JSON.parse(raw) as HcPayload;
  const records = payload.registros;
  await writeLibro(records);
  return { records, kpis: computeKpis(records), source: "seed" };
}

export async function importFromBuffer(buf: Buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const records = parseWorkbook(wb);
  if (records.length === 0) throw new Error("El libro no contiene una hoja CONTROL_HC reconocible.");
  const written = await writeLibro(records);
  return { records, kpis: written.kpis, source: "import" as const };
}

export async function workbookBase64(records?: HcRecord[]): Promise<string> {
  if (records) {
    const wb = await buildWorkbook(records);
    const buf = Buffer.from(await wb.xlsx.writeBuffer());
    await persist(buf);
    return buf.toString("base64");
  }
  const existing = await readLibroBuffer();
  if (existing) return existing.toString("base64");
  const seeded = await loadOrSeed();
  const wb = await buildWorkbook(seeded.records);
  const buf = Buffer.from(await wb.xlsx.writeBuffer());
  return buf.toString("base64");
}

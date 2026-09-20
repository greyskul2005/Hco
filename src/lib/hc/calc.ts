import type { HcKpis, HcRecord } from "./types";

export const HC_BASE = 514;
export const ECUACION = "514 + 13 − 9 + 9 = 527";
export const MARCA =
  "SINTÉTICO · demostración · no utilizar como dato real de RRHH";

export function impactoDe(tipo: string, estado: string): number {
  if (estado === "Cancelado") return 0;
  if (tipo === "Incremento") return 1;
  return 0;
}

export function computeKpis(records: HcRecord[]): HcKpis {
  const vigente = (r: HcRecord) =>
    r.ComputaHC === "Si" &&
    (r.Estado === "Activo" || r.Estado === "Salida_Prevista");
  const hc_actual = records.filter(vigente).length;
  const incrementos = records.filter(
    (r) =>
      r.TipoRegistro === "Incremento" &&
      r.ComputaHC === "Si" &&
      r.Estado !== "Cancelado",
  ).length;
  const salidas_cev = records.filter(
    (r) => r.TipoRegistro === "Base" && r.Programa === "CEV",
  ).length;
  const sustituciones = records.filter(
    (r) => r.TipoRegistro === "Sustitucion" && r.ComputaHC === "Si",
  ).length;
  const fuera_hc = records.filter((r) => r.TipoRegistro === "Fuera_HC").length;
  const incompletos = records.filter((r) => r.EstadoCalidad === "Incompleto").length;
  const duplicados_registros = records.filter(
    (r) => r.EstadoCalidad === "Duplicado",
  ).length;
  const counts = new Map<string, number>();
  for (const r of records) {
    counts.set(r.Matricula, (counts.get(r.Matricula) ?? 0) + 1);
  }
  const matriculas_duplicadas = [...counts.entries()]
    .filter(([, n]) => n > 1)
    .map(([m]) => m);
  const hc_estructural = HC_BASE + incrementos;
  const hc_previsto_final = HC_BASE + incrementos - salidas_cev + sustituciones;

  const plan: { mes: string; etiqueta: string; esperado: number }[] = [
    { mes: "2026-08", etiqueta: "Agosto 2026 (cierre)", esperado: 514 },
    { mes: "2026-09", etiqueta: "Tras incrementos iniciales", esperado: 521 },
    { mes: "2026-10", etiqueta: "Octubre 2026", esperado: 523 },
    { mes: "2026-11", etiqueta: "Noviembre 2026", esperado: 526 },
    { mes: "2026-12", etiqueta: "Diciembre 2026", esperado: 527 },
  ];
  let acc = HC_BASE;
  const evolucion = plan.map((p) => {
    if (p.mes === "2026-08") {
      return { ...p, hc: HC_BASE, ok: HC_BASE === p.esperado };
    }
    const altas = records.filter(
      (r) =>
        r.TipoRegistro === "Incremento" &&
        r.ComputaHC === "Si" &&
        r.Estado !== "Cancelado" &&
        r.MesEfecto === p.mes,
    ).length;
    const salidas = records.filter(
      (r) =>
        r.TipoRegistro === "Base" &&
        r.Programa === "CEV" &&
        r.MesEfecto === p.mes,
    ).length;
    const sust = records.filter(
      (r) =>
        r.TipoRegistro === "Sustitucion" &&
        r.ComputaHC === "Si" &&
        r.MesEfecto === p.mes,
    ).length;
    acc = acc + altas - salidas + sust;
    return { ...p, hc: acc, ok: acc === p.esperado };
  });

  return {
    registros: records.length,
    hc_base_agosto: HC_BASE,
    hc_actual,
    incrementos,
    salidas_cev,
    sustituciones,
    fuera_hc,
    incompletos,
    duplicados_registros,
    matriculas_duplicadas,
    hc_estructural,
    hc_previsto_final,
    ecuacion: ECUACION,
    ecuacion_ok:
      hc_previsto_final === 527 &&
      incrementos === 13 &&
      salidas_cev === 9 &&
      sustituciones === 9 &&
      hc_actual === HC_BASE,
    evolucion,
    impacto_cev_neto: sustituciones - salidas_cev,
  };
}

export function emptyRecord(): HcRecord {
  return {
    Matricula: "",
    NombreCompleto: "",
    Empresa: "Empresa Demo Corporativa",
    Correo: "",
    Puesto: "",
    Responsable: "",
    UnidadOrganizativa: "",
    ResponsableUnidad: "",
    FechaAltaRRHH: "",
    IndicadorEstructura: "Si",
    SuccessID: "",
    TipoRegistro: "Incremento",
    Estado: "Incorporacion_Prevista",
    ComputaHC: "Si",
    OrigenPlaza: "Incremento_Aprobado",
    Programa: "Incremento_Estructural",
    FechaPrevistaAlta: "",
    FechaRealAlta: "",
    FechaPrevistaBaja: "",
    FechaRealBaja: "",
    MesEfecto: "2026-09",
    ImpactoHC: 1,
    IDPlaza: "",
    MatriculaSustituida: "",
    Motivo: "",
    Observaciones: "Alta desde el aplicativo web. Dato sintético.",
    Fuente: "Manual",
    EstadoCalidad: "Completo",
    DatoSintetico: "Si",
  };
}

export function qualityOf(r: HcRecord, all: HcRecord[]): string {
  if (!r.Matricula) return "Incompleto";
  if (all.filter((x) => x.Matricula === r.Matricula).length > 1) return "Duplicado";
  if (!r.TipoRegistro) return "Incompleto";
  if (r.TipoRegistro === "Sustitucion" && !r.MatriculaSustituida) return "Incompleto";
  if (r.Estado === "Activo" && r.ComputaHC === "Si" && !r.UnidadOrganizativa)
    return "Incompleto";
  return r.EstadoCalidad || "Completo";
}

export function rowTone(r: HcRecord): "inc" | "sus" | "fuera" | "dup" | "incomp" | "base" {
  if (r.EstadoCalidad === "Duplicado") return "dup";
  if (r.EstadoCalidad === "Incompleto") return "incomp";
  if (r.TipoRegistro === "Incremento") return "inc";
  if (r.TipoRegistro === "Sustitucion") return "sus";
  if (r.TipoRegistro === "Fuera_HC") return "fuera";
  return "base";
}

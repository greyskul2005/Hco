import type { ReactNode } from "react";
import { useMemo, useState } from "react";
import { ECUACION } from "@/lib/hc/calc";
import type { HcKpis, HcRecord } from "@/lib/hc/types";
import { cn } from "@/lib/utils";

const SHEETS = [
  "PORTADA",
  "CONTROL_HC",
  "MOVIMIENTOS",
  "RECONCILIACION",
  "CIERRE",
  "CALIDAD",
  "LEYENDA",
] as const;

function Cell({
  children,
  className,
  mono,
}: {
  children: ReactNode;
  className?: string;
  mono?: boolean;
}) {
  return (
    <div
      className={cn(
        "border border-border px-2 py-1.5 text-[11px] leading-tight",
        mono && "font-mono",
        className,
      )}
    >
      {children}
    </div>
  );
}

function Head({ children, className }: { children: ReactNode; className?: string }) {
  return <Cell className={cn("bg-primary text-center font-medium text-primary-foreground", className)}>{children}</Cell>;
}

export function ExcelWorkbook({ records, kpis }: { records: HcRecord[]; kpis: HcKpis | null }) {
  const [sheet, setSheet] = useState<(typeof SHEETS)[number]>("PORTADA");
  const movs = useMemo(
    () =>
      records.filter(
        (r) =>
          r.TipoRegistro === "Incremento" ||
          r.TipoRegistro === "Sustitucion" ||
          (r.TipoRegistro === "Base" && r.Programa === "CEV"),
      ),
    [records],
  );
  const issues = useMemo(
    () => records.filter((r) => r.EstadoCalidad !== "Completo"),
    [records],
  );
  const units = useMemo(() => {
    const m = new Map<string, number>();
    for (const r of records) {
      if (r.TipoRegistro === "Base" && r.ComputaHC === "Si" && r.UnidadOrganizativa) {
        m.set(r.UnidadOrganizativa, (m.get(r.UnidadOrganizativa) ?? 0) + 1);
      }
    }
    return [...m.entries()].sort((a, b) => b[1] - a[1]);
  }, [records]);

  return (
    <div className="overflow-hidden rounded-2xl shadow-border">
      <div className="flex items-center justify-between bg-[#217346] px-3 py-2 text-xs text-white">
        <span className="font-medium">HC_Control_Operativo.xlsx</span>
        <span className="opacity-80">Vista idéntica al libro que descarga</span>
      </div>
      <div className="bg-[#f3f3f3] px-3 py-2 text-[11px] text-muted-foreground">
        Inicio · Insertar · Datos · Revisión · Vista · El archivo real usa estos mismos colores de celda
      </div>
      <div className="min-h-[520px] bg-white p-3">
        {sheet === "PORTADA" && kpis ? (
          <div className="space-y-3">
            <div className="bg-primary px-3 py-2 font-display text-lg text-primary-foreground">
              HC CONTROL · LIBRO OPERATIVO
            </div>
            <div className="bg-ring px-3 py-1.5 text-xs text-primary-foreground">
              Aplicativo web · almacenamiento en Excel · {ECUACION}
            </div>
            <div className="bg-cyan-soft px-3 py-1 text-[11px] text-cyan">
              SINTÉTICO · demostración · no utilizar como dato real de RRHH
            </div>
            <div className="grid grid-cols-5 gap-2">
              {[
                ["HC actual", kpis.hc_actual, "bg-primary"],
                ["Incrementos", kpis.incrementos, "bg-ring"],
                ["Salidas CEV", kpis.salidas_cev, "bg-cyan"],
                ["Sustituciones", kpis.sustituciones, "bg-cyan"],
                ["HC previsto", kpis.hc_previsto_final, "bg-ok-fg"],
              ].map(([l, v, c]) => (
                <div key={String(l)} className={cn("rounded-sm p-2 text-center text-white", c)}>
                  <div className="text-[10px] uppercase tracking-wide opacity-80">{l}</div>
                  <div className="font-display text-2xl tabular-nums">{v}</div>
                </div>
              ))}
            </div>
            <div className={cn("p-2 text-center text-sm font-medium", kpis.ecuacion_ok ? "bg-ok text-ok-fg" : "bg-err")}>
              Ecuación oficial: {ECUACION} · {kpis.ecuacion_ok ? "OK" : "ERROR"}
            </div>
            <div className="grid grid-cols-5 text-xs">
              {["Mes", "Etiqueta", "HC calculado", "HC esperado", "Estado"].map((h) => (
                <Head key={h}>{h}</Head>
              ))}
              {kpis.evolucion.map((e) => (
                <div key={e.mes} className="contents">
                  <Cell className={e.ok ? "bg-ok" : "bg-err"} mono>
                    {e.mes}
                  </Cell>
                  <Cell className={e.ok ? "bg-ok" : "bg-err"}>{e.etiqueta}</Cell>
                  <Cell className={e.ok ? "bg-ok" : "bg-err"}>{e.hc}</Cell>
                  <Cell className={e.ok ? "bg-ok" : "bg-err"}>{e.esperado}</Cell>
                  <Cell className={e.ok ? "bg-ok" : "bg-err"}>{e.ok ? "OK" : "ERROR"}</Cell>
                </div>
              ))}
            </div>
            <div className="grid max-w-md grid-cols-2 text-xs">
              <Head>Unidad</Head>
              <Head>HC base</Head>
              {units.map(([u, n]) => (
                <div key={u} className="contents">
                  <Cell>{u}</Cell>
                  <Cell className="bg-calc">{n}</Cell>
                </div>
              ))}
            </div>
          </div>
        ) : null}

        {sheet === "CONTROL_HC" ? (
          <div className="overflow-x-auto">
            <div className="mb-2 bg-primary px-2 py-1 text-xs text-primary-foreground">
              CONTROL_HC · {records.length} filas · amarillo entrada / azul calculado · color de fila = tipo
            </div>
            <div className="grid min-w-[1100px] grid-cols-9 text-[11px]">
              {["Matrícula", "Nombre", "Unidad", "Tipo", "Estado", "Programa", "Impacto", "Plaza", "Calidad"].map(
                (h) => (
                  <Head key={h}>{h}</Head>
                ),
              )}
              {records.slice(0, 40).map((r, i) => {
                const fill =
                  r.EstadoCalidad === "Duplicado"
                    ? "bg-dup"
                    : r.EstadoCalidad === "Incompleto"
                      ? "bg-incomp"
                      : r.TipoRegistro === "Incremento"
                        ? "bg-inc"
                        : r.TipoRegistro === "Sustitucion"
                          ? "bg-sus"
                          : r.TipoRegistro === "Fuera_HC"
                            ? "bg-fuera"
                            : "bg-white";
                return (
                  <div key={`${r.Matricula}-${i}`} className="contents">
                    <Cell className={fill} mono>
                      {r.Matricula}
                    </Cell>
                    <Cell className={fill}>{r.NombreCompleto}</Cell>
                    <Cell className={fill}>{r.UnidadOrganizativa}</Cell>
                    <Cell className={fill}>{r.TipoRegistro}</Cell>
                    <Cell className={fill}>{r.Estado}</Cell>
                    <Cell className={fill}>{r.Programa}</Cell>
                    <Cell className="bg-calc">{r.ImpactoHC}</Cell>
                    <Cell className={fill} mono>
                      {r.IDPlaza}
                    </Cell>
                    <Cell className={fill}>{r.EstadoCalidad}</Cell>
                  </div>
                );
              })}
            </div>
            <p className="mt-2 text-[11px] text-muted-foreground">
              Mostrando 40 de {records.length} filas. El archivo descargado contiene el registro completo.
            </p>
          </div>
        ) : null}

        {sheet === "MOVIMIENTOS" ? (
          <div className="overflow-x-auto">
            <div className="grid min-w-[900px] grid-cols-6 text-[11px]">
              {["Matrícula", "Tipo", "Programa", "Mes", "Impacto", "Plaza"].map((h) => (
                <Head key={h}>{h}</Head>
              ))}
              {movs.map((r) => {
                const fill = r.TipoRegistro === "Incremento" ? "bg-inc" : "bg-sus";
                return (
                  <div key={r.Matricula} className="contents">
                    <Cell className={fill} mono>
                      {r.Matricula}
                    </Cell>
                    <Cell className={fill}>{r.TipoRegistro}</Cell>
                    <Cell className={fill}>{r.Programa}</Cell>
                    <Cell className={fill}>{r.MesEfecto}</Cell>
                    <Cell className="bg-calc">{r.ImpactoHC}</Cell>
                    <Cell className={fill} mono>
                      {r.IDPlaza}
                    </Cell>
                  </div>
                );
              })}
            </div>
          </div>
        ) : null}

        {sheet === "RECONCILIACION" && kpis ? (
          <div className="space-y-2 text-sm">
            {[
              ["1", "HC base agosto 2026", "514"],
              ["2", "Incrementos iniciales", "+7 → 521"],
              ["3", "Octubre +2", "523"],
              ["4", "Noviembre +3", "526"],
              ["5", "Diciembre +1", "527"],
              ["6", "Salidas CEV", "−9"],
              ["7", "Sustituciones CEV", "+9"],
            ].map((row) => (
              <div key={row[0]} className="grid grid-cols-[40px_1fr_120px] text-xs">
                <Cell className={Number(row[0]) >= 6 ? "bg-sus" : Number(row[0]) >= 2 ? "bg-inc" : "bg-calc"}>
                  {row[0]}
                </Cell>
                <Cell>{row[1]}</Cell>
                <Cell className="bg-calc text-center font-medium">{row[2]}</Cell>
              </div>
            ))}
            <div className="bg-ok p-3 text-center font-display text-xl text-ok-fg">{ECUACION} · OK</div>
          </div>
        ) : null}

        {sheet === "CIERRE" && kpis ? (
          <div className="grid grid-cols-4 text-xs">
            {["Mes", "HC cierre", "Esperado", "Estado"].map((h) => (
              <Head key={h}>{h}</Head>
            ))}
            {kpis.evolucion.map((e) => (
              <div key={e.mes} className="contents">
                <Cell className={e.ok ? "bg-ok" : "bg-err"}>{e.mes}</Cell>
                <Cell className="bg-calc">{e.hc}</Cell>
                <Cell className="bg-calc">{e.esperado}</Cell>
                <Cell className={e.ok ? "bg-ok" : "bg-err"}>
                  {e.mes === "2026-08" ? "Cerrado" : "Previsto"}
                </Cell>
              </div>
            ))}
          </div>
        ) : null}

        {sheet === "CALIDAD" ? (
          <div className="grid grid-cols-4 text-xs">
            {["Matrícula", "Tipo", "Severidad", "Detalle"].map((h) => (
              <Head key={h}>{h}</Head>
            ))}
            {issues.map((r) => (
              <div key={r.Matricula + r.NombreCompleto} className="contents">
                <Cell className={r.EstadoCalidad === "Duplicado" ? "bg-dup" : "bg-incomp"} mono>
                  {r.Matricula}
                </Cell>
                <Cell className={r.EstadoCalidad === "Duplicado" ? "bg-dup" : "bg-incomp"}>{r.EstadoCalidad}</Cell>
                <Cell className={r.EstadoCalidad === "Duplicado" ? "bg-dup" : "bg-incomp"}>
                  {r.EstadoCalidad === "Duplicado" ? "Alta" : "Media"}
                </Cell>
                <Cell className={r.EstadoCalidad === "Duplicado" ? "bg-dup" : "bg-incomp"}>{r.Observaciones}</Cell>
              </div>
            ))}
          </div>
        ) : null}

        {sheet === "LEYENDA" ? (
          <div className="grid max-w-xl grid-cols-2 text-xs">
            <Head>Color en Excel y en la web</Head>
            <Head>Significado</Head>
            {[
              ["bg-inc", "Incremento · crea plaza · impacto +1"],
              ["bg-sus", "Sustitución / CEV · reutiliza plaza · impacto 0"],
              ["bg-fuera", "Fuera de HC · no computa"],
              ["bg-dup", "Duplicado · bloquea cierre si computa"],
              ["bg-incomp", "Incompleto · no debe computar"],
              ["bg-calc", "Celda calculada (sistema)"],
              ["bg-ok", "Control OK"],
              ["bg-err", "Control ERROR"],
            ].map(([c, t]) => (
              <div key={t} className="contents">
                <Cell className={c}>&nbsp;</Cell>
                <Cell>{t}</Cell>
              </div>
            ))}
          </div>
        ) : null}
      </div>
      <div className="flex gap-1 overflow-x-auto bg-[#f3f3f3] px-2 py-1.5">
        {SHEETS.map((s) => (
          <button
            key={s}
            onClick={() => setSheet(s)}
            className={cn(
              "rounded-t-md px-3 py-1 text-[11px]",
              sheet === s ? "bg-white font-medium text-foreground shadow-sm" : "text-muted-foreground hover:bg-white/60",
            )}
          >
            {s}
          </button>
        ))}
      </div>
    </div>
  );
}

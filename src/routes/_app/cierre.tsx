import { createFileRoute } from "@tanstack/react-router";
import { KpiStrip } from "@/components/hc/kpi-strip";
import { Badge } from "@/components/ui/badge";
import { ECUACION } from "@/lib/hc/calc";
import { useHcStore } from "@/lib/hc/store";

export const Route = createFileRoute("/_app/cierre")({ component: CierrePage });

function CierrePage() {
  const { kpis } = useHcStore();
  return (
    <div className="mx-auto max-w-4xl space-y-6">
      <div>
        <h1 className="font-display text-4xl font-medium">Cierre mensual</h1>
        <p className="mt-2 text-sm text-muted-foreground">
          Hoja CIERRE del Excel. Agosto 2026 está cerrado en 514 (dato confirmado). No cierre un mes si la
          portada del libro muestra ERROR.
        </p>
      </div>
      <KpiStrip kpis={kpis} />
      <ol className="space-y-2 text-sm">
        {[
          "Fotografía RRHH conciliada con CONTROL_HC",
          "Duplicados revisados",
          "Incompletos que computan = 0",
          "Incrementos y sustituciones no mezclados",
          `Ecuación ${ECUACION} verificada`,
        ].map((s, i) => (
          <li key={s} className="flex items-center gap-3 rounded-xl bg-card px-4 py-3 shadow-border">
            <span className="flex size-7 items-center justify-center rounded-md bg-primary text-xs text-primary-foreground">
              {i + 1}
            </span>
            {s}
          </li>
        ))}
      </ol>
      <div className="overflow-hidden rounded-2xl shadow-border">
        <table className="w-full text-left text-sm">
          <thead className="bg-primary text-primary-foreground">
            <tr>
              {["Mes", "HC cierre", "Esperado", "Estado"].map((h) => (
                <th key={h} className="px-3 py-2.5 font-medium">
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {(kpis?.evolucion ?? []).map((e) => (
              <tr key={e.mes} className={e.ok ? "bg-ok" : "bg-err"}>
                <td className="px-3 py-2 font-mono text-xs">{e.mes}</td>
                <td className="px-3 py-2 tabular-nums">{e.hc}</td>
                <td className="px-3 py-2 tabular-nums">{e.esperado}</td>
                <td className="px-3 py-2">
                  <Badge tone={e.ok ? "ok" : "err"}>
                    {e.mes === "2026-08" ? "Cerrado" : e.ok ? "Previsto" : "ERROR"}
                  </Badge>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

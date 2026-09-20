import { createFileRoute } from "@tanstack/react-router";
import { Badge } from "@/components/ui/badge";
import { useHcStore } from "@/lib/hc/store";

export const Route = createFileRoute("/_app/calidad")({ component: CalidadPage });

function CalidadPage() {
  const { records, kpis } = useHcStore();
  const issues = records.filter((r) => r.EstadoCalidad !== "Completo");
  return (
    <div className="mx-auto max-w-6xl space-y-6">
      <div>
        <h1 className="font-display text-4xl font-medium">Calidad del dato</h1>
        <p className="mt-2 text-sm text-muted-foreground">
          Hoja CALIDAD del Excel. Rosa = duplicado. Ámbar = incompleto. Los incompletos de demostración no computan
          para no distorsionar 514.
        </p>
      </div>
      <div className="grid gap-3 sm:grid-cols-4">
        <Box label="Incompletos" value={kpis?.incompletos ?? 0} className="bg-incomp" />
        <Box label="Duplicados (filas)" value={kpis?.duplicados_registros ?? 0} className="bg-dup" />
        <Box label="Matrículas repetidas" value={kpis?.matriculas_duplicadas.length ?? 0} className="bg-dup" />
        <Box label="Fuera de HC" value={kpis?.fuera_hc ?? 0} className="bg-fuera" />
      </div>
      <div className="overflow-x-auto rounded-2xl shadow-border">
        <table className="w-full min-w-[720px] text-left text-sm">
          <thead className="bg-primary text-primary-foreground">
            <tr>
              {["Matrícula", "Incidencia", "Nombre", "Campo", "Detalle"].map((h) => (
                <th key={h} className="px-3 py-2.5 font-medium">
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {issues.map((r) => (
              <tr key={r.Matricula + r.NombreCompleto} className={r.EstadoCalidad === "Duplicado" ? "bg-dup" : "bg-incomp"}>
                <td className="px-3 py-2 font-mono text-xs">{r.Matricula}</td>
                <td className="px-3 py-2">
                  <Badge tone={r.EstadoCalidad === "Duplicado" ? "err" : "warn"}>{r.EstadoCalidad}</Badge>
                </td>
                <td className="px-3 py-2">{r.NombreCompleto}</td>
                <td className="px-3 py-2">{!r.UnidadOrganizativa ? "Unidad" : !r.Correo ? "Correo" : "Matrícula"}</td>
                <td className="px-3 py-2 text-xs">{r.Observaciones}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function Box({ label, value, className }: { label: string; value: number; className: string }) {
  return (
    <div className={`rounded-2xl px-4 py-3 shadow-border ${className}`}>
      <p className="text-xs uppercase tracking-wide">{label}</p>
      <p className="font-display text-3xl tabular-nums">{value}</p>
    </div>
  );
}

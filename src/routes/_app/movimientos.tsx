import { createFileRoute } from "@tanstack/react-router";
import { Badge } from "@/components/ui/badge";
import { useHcStore } from "@/lib/hc/store";

export const Route = createFileRoute("/_app/movimientos")({ component: MovPage });

function MovPage() {
  const { records, kpis } = useHcStore();
  const movs = records.filter(
    (r) =>
      r.TipoRegistro === "Incremento" ||
      r.TipoRegistro === "Sustitucion" ||
      (r.TipoRegistro === "Base" && r.Programa === "CEV"),
  );
  return (
    <div className="mx-auto max-w-6xl space-y-6">
      <div>
        <h1 className="font-display text-4xl font-medium">Movimientos HC</h1>
        <p className="mt-2 text-sm text-muted-foreground">
          Azul: incrementos que crean plaza. Cian: CEV que reutiliza plaza. Esta vista es la hoja MOVIMIENTOS del
          Excel.
        </p>
      </div>
      <div className="grid gap-3 sm:grid-cols-3">
        <Stat label="Incrementos" value={kpis?.incrementos ?? 0} className="bg-inc" />
        <Stat label="Salidas CEV" value={kpis?.salidas_cev ?? 0} className="bg-sus" />
        <Stat label="Sustituciones" value={kpis?.sustituciones ?? 0} className="bg-sus" />
      </div>
      <div className="overflow-x-auto rounded-2xl shadow-border">
        <table className="w-full min-w-[860px] text-left text-sm">
          <thead className="bg-primary text-primary-foreground">
            <tr>
              {["Matrícula", "Tipo", "Programa", "Unidad", "Mes", "Impacto", "Plaza", "Sustituye"].map((h) => (
                <th key={h} className="px-3 py-2.5 font-medium">
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {movs.map((r) => (
              <tr
                key={r.Matricula + r.TipoRegistro}
                className={r.TipoRegistro === "Incremento" ? "bg-inc" : "bg-sus"}
              >
                <td className="px-3 py-2 font-mono text-xs">{r.Matricula}</td>
                <td className="px-3 py-2">
                  <Badge tone={r.TipoRegistro === "Incremento" ? "inc" : "sus"}>{r.TipoRegistro}</Badge>
                </td>
                <td className="px-3 py-2">{r.Programa}</td>
                <td className="px-3 py-2">{r.UnidadOrganizativa}</td>
                <td className="px-3 py-2 font-mono text-xs">{r.MesEfecto}</td>
                <td className="bg-calc px-3 py-2 tabular-nums">{r.ImpactoHC}</td>
                <td className="px-3 py-2 font-mono text-xs">{r.IDPlaza}</td>
                <td className="px-3 py-2 font-mono text-xs">{r.MatriculaSustituida || "—"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function Stat({ label, value, className }: { label: string; value: number; className: string }) {
  return (
    <div className={`rounded-2xl px-4 py-3 shadow-border ${className}`}>
      <p className="text-xs uppercase tracking-wide text-muted-foreground">{label}</p>
      <p className="font-display text-3xl tabular-nums">{value}</p>
    </div>
  );
}

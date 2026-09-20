import { createFileRoute } from "@tanstack/react-router";
import { Bar, BarChart, CartesianGrid, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import { RESPONSABLES } from "@/lib/hc/columns";
import { useHcStore } from "@/lib/hc/store";

export const Route = createFileRoute("/_app/organizacion")({ component: OrgPage });

function OrgPage() {
  const { records } = useHcStore();
  const rows = Object.keys(RESPONSABLES).map((u) => {
    const base = records.filter(
      (r) => r.UnidadOrganizativa === u && r.TipoRegistro === "Base" && r.ComputaHC === "Si",
    ).length;
    const inc = records.filter(
      (r) => r.UnidadOrganizativa === u && r.TipoRegistro === "Incremento" && r.ComputaHC === "Si",
    ).length;
    const sus = records.filter(
      (r) => r.UnidadOrganizativa === u && r.TipoRegistro === "Sustitucion" && r.ComputaHC === "Si",
    ).length;
    return { unidad: u, responsable: RESPONSABLES[u], base, inc, sus };
  });
  return (
    <div className="mx-auto max-w-6xl space-y-6">
      <div>
        <h1 className="font-display text-4xl font-medium">Distribución organizativa</h1>
        <p className="mt-2 text-sm text-muted-foreground">
          Unidades sintéticas (pendiente catálogo real). La hoja PORTADA del Excel incluye la misma tabla de HC
          base por unidad.
        </p>
      </div>
      <div className="h-72 rounded-2xl bg-card p-4 shadow-border">
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={rows} layout="vertical" margin={{ left: 16, right: 8 }}>
            <CartesianGrid stroke="var(--color-border)" horizontal={false} />
            <XAxis type="number" tick={{ fill: "var(--color-muted-foreground)", fontSize: 12 }} />
            <YAxis
              type="category"
              dataKey="unidad"
              width={140}
              tick={{ fill: "var(--color-muted-foreground)", fontSize: 11 }}
            />
            <Tooltip />
            <Bar dataKey="base" fill="var(--color-primary)" name="Base" stackId="a" />
            <Bar dataKey="inc" fill="var(--color-ring)" name="Incremento" stackId="a" />
            <Bar dataKey="sus" fill="var(--color-cyan)" name="Sustitución" stackId="a" />
          </BarChart>
        </ResponsiveContainer>
      </div>
      <div className="overflow-x-auto rounded-2xl shadow-border">
        <table className="w-full text-left text-sm">
          <thead className="bg-primary text-primary-foreground">
            <tr>
              {["Unidad", "Responsable (sintético)", "Base", "Incrementos", "Sustituciones"].map((h) => (
                <th key={h} className="px-3 py-2.5 font-medium">
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.unidad} className="border-t border-border bg-card">
                <td className="px-3 py-2">{r.unidad}</td>
                <td className="px-3 py-2">{r.responsable}</td>
                <td className="bg-calc px-3 py-2 tabular-nums">{r.base}</td>
                <td className="bg-inc px-3 py-2 tabular-nums">{r.inc}</td>
                <td className="bg-sus px-3 py-2 tabular-nums">{r.sus}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

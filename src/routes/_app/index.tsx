import { createFileRoute, Link } from "@tanstack/react-router";
import { BookOpen, Download } from "lucide-react";
import {
  Bar,
  BarChart,
  CartesianGrid,
  Line,
  LineChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import { ExcelWorkbook } from "@/components/hc/excel-workbook";
import { KpiStrip } from "@/components/hc/kpi-strip";
import { Button } from "@/components/ui/button";
import { useHcStore } from "@/lib/hc/store";

export const Route = createFileRoute("/_app/")({ component: Home });

function Home() {
  const { kpis, records, download } = useHcStore();
  const evo = (kpis?.evolucion ?? []).map((e) => ({ mes: e.mes.slice(5), hc: e.hc, esperado: e.esperado }));
  const byTipo = [
    { tipo: "Base", n: records.filter((r) => r.TipoRegistro === "Base").length },
    { tipo: "Incremento", n: records.filter((r) => r.TipoRegistro === "Incremento").length },
    { tipo: "Sustitución", n: records.filter((r) => r.TipoRegistro === "Sustitucion").length },
    { tipo: "Fuera HC", n: records.filter((r) => r.TipoRegistro === "Fuera_HC").length },
  ];

  return (
    <div className="mx-auto max-w-6xl space-y-8">
      <div className="flex flex-wrap items-end justify-between gap-3">
        <div>
          <p className="text-xs uppercase tracking-[0.2em] text-cyan">Aplicativo web · almacén Excel</p>
          <h1 className="mt-1 font-display text-4xl font-medium tracking-tight">Resumen de headcount</h1>
          <p className="mt-2 max-w-2xl text-sm text-muted-foreground">
            Cada cambio se escribe en HC_Control_Operativo.xlsx con los mismos colores que ve aquí.
            Abra el libro: las cifras 514, +13, CEV −9/+9 y 527 son reconocibles de un vistazo.
          </p>
        </div>
        <div className="flex gap-2">
          <Button variant="outline" asChild>
            <Link to="/libro">
              <BookOpen /> Ver libro
            </Link>
          </Button>
          <Button onClick={() => void download()}>
            <Download /> Descargar Excel
          </Button>
        </div>
      </div>

      <KpiStrip kpis={kpis} />

      <div className="grid gap-4 lg:grid-cols-2">
        <section className="rounded-2xl bg-card p-5 shadow-border">
          <h2 className="font-display text-lg">Evolución prevista</h2>
          <p className="mb-4 text-xs text-muted-foreground">Agosto 514 → diciembre 527. CEV neto 0.</p>
          <div className="h-56">
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={evo}>
                <CartesianGrid stroke="var(--color-border)" vertical={false} />
                <XAxis dataKey="mes" tick={{ fill: "var(--color-muted-foreground)", fontSize: 12 }} />
                <YAxis domain={[510, 530]} tick={{ fill: "var(--color-muted-foreground)", fontSize: 12 }} />
                <Tooltip />
                <Line type="monotone" dataKey="hc" stroke="var(--color-primary)" strokeWidth={2.5} dot />
                <Line type="monotone" dataKey="esperado" stroke="var(--color-cyan)" strokeDasharray="4 4" />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </section>
        <section className="rounded-2xl bg-card p-5 shadow-border">
          <h2 className="font-display text-lg">Composición del libro</h2>
          <p className="mb-4 text-xs text-muted-foreground">Mismos colores de fila que CONTROL_HC.</p>
          <div className="h-56">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={byTipo}>
                <CartesianGrid stroke="var(--color-border)" vertical={false} />
                <XAxis dataKey="tipo" tick={{ fill: "var(--color-muted-foreground)", fontSize: 12 }} />
                <YAxis tick={{ fill: "var(--color-muted-foreground)", fontSize: 12 }} />
                <Tooltip />
                <Bar dataKey="n" fill="var(--color-primary)" radius={[6, 6, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </section>
      </div>

      <section>
        <h2 className="mb-3 font-display text-lg">El libro, como se abre en Excel</h2>
        <ExcelWorkbook records={records} kpis={kpis} />
      </section>
    </div>
  );
}

import { createFileRoute } from "@tanstack/react-router";
import { RecordTable } from "@/components/hc/record-table";
import { useHcStore } from "@/lib/hc/store";

export const Route = createFileRoute("/_app/registro")({ component: RegistroPage });

function RegistroPage() {
  const { records, kpis } = useHcStore();
  return (
    <div className="mx-auto max-w-6xl space-y-5">
      <div>
        <h1 className="font-display text-4xl font-medium">CONTROL_HC</h1>
        <p className="mt-2 text-sm text-muted-foreground">
          Registro maestro. {records.length} filas en el libro · HC actual {kpis?.hc_actual ?? "—"} ·
          previsto {kpis?.hc_previsto_final ?? "—"}. El color de cada fila es el mismo que en Excel.
        </p>
      </div>
      <RecordTable records={records} />
    </div>
  );
}

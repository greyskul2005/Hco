import { createFileRoute } from "@tanstack/react-router";
import { Download } from "lucide-react";
import { Button } from "@/components/ui/button";
import { useHcStore } from "@/lib/hc/store";

export const Route = createFileRoute("/_app/equipo")({ component: EquipoPage });

function EquipoPage() {
  const { download, kpis } = useHcStore();
  return (
    <div className="mx-auto max-w-3xl space-y-6">
      <div>
        <h1 className="font-display text-4xl font-medium">Equipo de operación</h1>
        <p className="mt-2 text-sm text-muted-foreground">
          Vista de colaboración. El artefacto que se comparte es el libro Excel, no una captura. Datos sintéticos.
        </p>
      </div>
      <div className="rounded-2xl bg-card p-5 shadow-border">
        <p className="text-xs uppercase tracking-wide text-muted-foreground">Canal Cierre mensual</p>
        <p className="mt-3 rounded-xl bg-paper px-4 py-3 text-sm">
          <span className="font-medium">Operador Demo</span>
          <span className="ml-2 text-xs text-muted-foreground">31/08/2026</span>
          <br />
          Cierre agosto 2026 confirmado: HC 514. Libro actualizado. Escenario {kpis?.ecuacion ?? "514 + 13 − 9 + 9 = 527"}.
        </p>
        <p className="mt-3 rounded-xl bg-paper px-4 py-3 text-sm">
          <span className="font-medium">Validador Demo</span>
          <span className="ml-2 text-xs text-muted-foreground">01/09/2026</span>
          <br />
          Portada del Excel en verde. CEV neto 0. No mezclar incrementos y sustituciones.
        </p>
        <div className="mt-4 flex items-center justify-between rounded-xl border border-border px-4 py-3">
          <div>
            <p className="text-sm font-medium">HC_Control_Operativo.xlsx</p>
            <p className="text-xs text-muted-foreground">Almacén oficial · PORTADA + CONTROL_HC + movimientos</p>
          </div>
          <Button size="sm" onClick={() => void download()}>
            <Download /> Descargar
          </Button>
        </div>
      </div>
    </div>
  );
}

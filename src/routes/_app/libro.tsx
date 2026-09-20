import { createFileRoute } from "@tanstack/react-router";
import { Download, Upload } from "lucide-react";
import { useRef } from "react";
import { toast } from "sonner";
import { ExcelWorkbook } from "@/components/hc/excel-workbook";
import { Button } from "@/components/ui/button";
import { useHcStore } from "@/lib/hc/store";

export const Route = createFileRoute("/_app/libro")({ component: LibroPage });

function LibroPage() {
  const { records, kpis, download, importFile, source, saving } = useHcStore();
  const inputRef = useRef<HTMLInputElement>(null);

  return (
    <div className="mx-auto max-w-6xl space-y-6">
      <div>
        <p className="text-xs uppercase tracking-[0.2em] text-cyan">Sistema de registro</p>
        <h1 className="mt-1 font-display text-4xl font-medium">Libro Excel</h1>
        <p className="mt-2 max-w-2xl text-sm text-muted-foreground">
          El aplicativo no guarda en una base de datos. Guarda en un libro .xlsx con portada, CONTROL_HC,
          movimientos, reconciliación 514→527, cierre, calidad y leyenda. Los colores de celda son los mismos
          que en esta pantalla: azul incremento, cian CEV, ámbar incompleto, rosa duplicado.
        </p>
      </div>
      <div className="flex flex-wrap gap-2">
        <Button onClick={() => void download()}>
          <Download /> Descargar HC_Control_Operativo.xlsx
        </Button>
        <Button variant="outline" onClick={() => inputRef.current?.click()}>
          <Upload /> Importar libro
        </Button>
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx"
          className="hidden"
          onChange={async (e) => {
            const file = e.target.files?.[0];
            e.target.value = "";
            if (!file) return;
            try {
              await importFile(file);
              toast.success("Libro importado y reescrito en el almacén Excel");
            } catch (err) {
              toast.error(err instanceof Error ? err.message : "No se pudo importar");
            }
          }}
        />
        <span className="self-center text-xs text-muted-foreground">
          Origen {source || "—"} {saving ? "· escribiendo…" : ""}
        </span>
      </div>
      <div className="grid gap-3 rounded-2xl bg-card p-4 text-sm shadow-border md:grid-cols-3">
        <p>
          <span className="block text-xs uppercase tracking-wide text-muted-foreground">Hojas</span>
          PORTADA, CONTROL_HC, MOVIMIENTOS, RECONCILIACION, CIERRE, CALIDAD, PARAMETROS, LEYENDA
        </p>
        <p>
          <span className="block text-xs uppercase tracking-wide text-muted-foreground">Identificadores</span>
          Matrícula, SuccessID e ID de plaza siempre como texto. Nunca se convierten a número.
        </p>
        <p>
          <span className="block text-xs uppercase tracking-wide text-muted-foreground">Reconocimiento visual</span>
          Abra el archivo: la portada muestra 514 / +13 / 9 / 9 / 527 y la ecuación en verde si cuadra.
        </p>
      </div>
      <ExcelWorkbook records={records} kpis={kpis} />
    </div>
  );
}

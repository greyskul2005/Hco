import { createFileRoute } from "@tanstack/react-router";
import { useMemo, useState } from "react";
import { toast } from "sonner";
import { RecordTable } from "@/components/hc/record-table";
import { Button } from "@/components/ui/button";
import { Input, Label, Select } from "@/components/ui/input";
import { useHcStore } from "@/lib/hc/store";
import type { HcRecord } from "@/lib/hc/types";

export const Route = createFileRoute("/_app/baja")({ component: BajaPage });

function BajaPage() {
  const { records, updateRecord } = useHcStore();
  const [picked, setPicked] = useState<HcRecord | null>(null);
  const [programa, setPrograma] = useState("Regular");
  const [fecha, setFecha] = useState("2026-11-30");
  const [mes, setMes] = useState("2026-11");
  const activos = useMemo(
    () => records.filter((r) => r.Estado === "Activo" && r.TipoRegistro === "Base" && r.ComputaHC === "Si"),
    [records],
  );

  return (
    <div className="mx-auto max-w-6xl space-y-6">
      <div>
        <h1 className="font-display text-4xl font-medium">Registrar baja</h1>
        <p className="mt-2 max-w-2xl text-sm text-muted-foreground">
          No se borra la fila. Pasa a Salida prevista y se reescribe el Excel. Si es CEV, la plaza se reutiliza:
          el alta del sustituto debe usar el mismo ID de plaza.
        </p>
      </div>
      {picked ? (
        <form
          className="space-y-4 rounded-2xl bg-card p-5 shadow-border"
          onSubmit={async (e) => {
            e.preventDefault();
            await updateRecord(picked.Matricula, {
              Estado: "Salida_Prevista",
              Programa: programa,
              FechaPrevistaBaja: fecha,
              MesEfecto: mes,
              Motivo:
                programa === "CEV"
                  ? "Salida prevista al programa CEV. La plaza se reutiliza."
                  : "Salida prevista. No eliminar el registro.",
            });
            toast.success(`Baja prevista escrita en Excel: ${picked.Matricula}`);
            setPicked(null);
          }}
        >
          <p className="font-display text-xl">
            {picked.NombreCompleto}{" "}
            <span className="font-mono text-sm text-muted-foreground">{picked.Matricula}</span>
          </p>
          <p className="text-sm text-muted-foreground">
            {picked.UnidadOrganizativa} · plaza {picked.IDPlaza}
          </p>
          <div className="grid gap-4 md:grid-cols-3">
            <div className="space-y-1.5">
              <Label>Programa</Label>
              <Select value={programa} onChange={(e) => setPrograma(e.target.value)}>
                <option>Regular</option>
                <option>CEV</option>
              </Select>
            </div>
            <div className="space-y-1.5">
              <Label>Fecha prevista baja</Label>
              <Input type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} />
            </div>
            <div className="space-y-1.5">
              <Label>Mes de efecto</Label>
              <Select value={mes} onChange={(e) => setMes(e.target.value)}>
                <option>2026-09</option>
                <option>2026-10</option>
                <option>2026-11</option>
                <option>2026-12</option>
              </Select>
            </div>
          </div>
          {programa === "CEV" ? (
            <p className="rounded-xl bg-sus px-4 py-3 text-sm text-cyan">
              Programa CEV: no cree un incremento. El sustituto se da de alta como Sustitución con esta misma plaza.
            </p>
          ) : null}
          <div className="flex gap-2">
            <Button type="submit">Guardar baja en Excel</Button>
            <Button type="button" variant="ghost" onClick={() => setPicked(null)}>
              Cancelar
            </Button>
          </div>
        </form>
      ) : (
        <RecordTable records={activos} onPick={setPicked} />
      )}
    </div>
  );
}

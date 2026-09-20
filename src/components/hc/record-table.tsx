import { useMemo, useState } from "react";
import { Badge } from "@/components/ui/badge";
import { Input, Select } from "@/components/ui/input";
import { rowTone } from "@/lib/hc/calc";
import type { HcRecord } from "@/lib/hc/types";
import { cn } from "@/lib/utils";

const PAGE = 20;

function toneClass(t: ReturnType<typeof rowTone>) {
  return {
    inc: "bg-inc",
    sus: "bg-sus",
    fuera: "bg-fuera",
    dup: "bg-dup",
    incomp: "bg-incomp",
    base: "bg-card",
  }[t];
}

function tipoTone(t: string) {
  if (t === "Incremento") return "inc" as const;
  if (t === "Sustitucion") return "sus" as const;
  if (t === "Fuera_HC") return "muted" as const;
  return "navy" as const;
}

export function RecordTable({
  records,
  onPick,
}: {
  records: HcRecord[];
  onPick?: (r: HcRecord) => void;
}) {
  const [q, setQ] = useState("");
  const [tipo, setTipo] = useState("Todos");
  const [estado, setEstado] = useState("Todos");
  const [unidad, setUnidad] = useState("Todas");
  const [page, setPage] = useState(0);

  const unidades = useMemo(
    () => [...new Set(records.map((r) => r.UnidadOrganizativa).filter(Boolean))].sort(),
    [records],
  );

  const filtered = useMemo(() => {
    const query = q.trim().toLowerCase();
    return records.filter((r) => {
      if (tipo !== "Todos" && r.TipoRegistro !== tipo) return false;
      if (estado !== "Todos" && r.Estado !== estado) return false;
      if (unidad !== "Todas" && r.UnidadOrganizativa !== unidad) return false;
      if (!query) return true;
      return (
        r.Matricula.toLowerCase().includes(query) ||
        r.NombreCompleto.toLowerCase().includes(query) ||
        r.Puesto.toLowerCase().includes(query) ||
        r.UnidadOrganizativa.toLowerCase().includes(query)
      );
    });
  }, [records, q, tipo, estado, unidad]);

  const pages = Math.max(1, Math.ceil(filtered.length / PAGE));
  const safePage = Math.min(page, pages - 1);
  const slice = filtered.slice(safePage * PAGE, safePage * PAGE + PAGE);

  return (
    <div className="space-y-3">
      <div className="grid gap-2 md:grid-cols-4">
        <Input
          placeholder="Buscar matrícula, nombre, puesto…"
          value={q}
          onChange={(e) => {
            setQ(e.target.value);
            setPage(0);
          }}
        />
        <Select
          value={tipo}
          onChange={(e) => {
            setTipo(e.target.value);
            setPage(0);
          }}
        >
          <option>Todos</option>
          <option>Base</option>
          <option>Incremento</option>
          <option>Sustitucion</option>
          <option>Fuera_HC</option>
        </Select>
        <Select
          value={estado}
          onChange={(e) => {
            setEstado(e.target.value);
            setPage(0);
          }}
        >
          <option>Todos</option>
          <option>Activo</option>
          <option>Incorporacion_Prevista</option>
          <option>Salida_Prevista</option>
          <option>Baja_Efectiva</option>
          <option>Cancelado</option>
        </Select>
        <Select
          value={unidad}
          onChange={(e) => {
            setUnidad(e.target.value);
            setPage(0);
          }}
        >
          <option>Todas</option>
          {unidades.map((u) => (
            <option key={u}>{u}</option>
          ))}
        </Select>
      </div>
      <p className="text-xs text-muted-foreground">
        {filtered.length} registros visibles · los colores coinciden con el libro Excel
      </p>
      <div className="overflow-x-auto rounded-xl shadow-border">
        <table className="w-full min-w-[980px] text-left text-sm">
          <thead className="bg-primary text-primary-foreground">
            <tr>
              {[
                "Matrícula",
                "Nombre",
                "Unidad",
                "Tipo",
                "Estado",
                "Programa",
                "Impacto",
                "Plaza",
                "Calidad",
              ].map((h) => (
                <th key={h} className="px-3 py-2.5 font-medium">
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {slice.map((r, i) => (
              <tr
                key={`${r.Matricula}-${i}`}
                className={cn("border-t border-border/70", toneClass(rowTone(r)), onPick && "cursor-pointer")}
                onClick={() => onPick?.(r)}
              >
                <td className="px-3 py-2 font-mono text-xs">{r.Matricula}</td>
                <td className="px-3 py-2">{r.NombreCompleto}</td>
                <td className="px-3 py-2">{r.UnidadOrganizativa}</td>
                <td className="px-3 py-2">
                  <Badge tone={tipoTone(r.TipoRegistro)}>{r.TipoRegistro}</Badge>
                </td>
                <td className="px-3 py-2 text-xs">{r.Estado}</td>
                <td className="px-3 py-2 text-xs">{r.Programa}</td>
                <td className="px-3 py-2 tabular-nums">{r.ImpactoHC}</td>
                <td className="px-3 py-2 font-mono text-xs">{r.IDPlaza}</td>
                <td className="px-3 py-2">
                  <Badge
                    tone={
                      r.EstadoCalidad === "Completo"
                        ? "ok"
                        : r.EstadoCalidad === "Duplicado"
                          ? "err"
                          : "warn"
                    }
                  >
                    {r.EstadoCalidad}
                  </Badge>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="flex items-center justify-between text-sm">
        <button
          className="text-muted-foreground hover:text-foreground"
          disabled={safePage === 0}
          onClick={() => setPage((p) => Math.max(0, p - 1))}
        >
          Anterior
        </button>
        <span className="tabular-nums text-muted-foreground">
          Página {safePage + 1} / {pages}
        </span>
        <button
          className="text-muted-foreground hover:text-foreground"
          disabled={safePage >= pages - 1}
          onClick={() => setPage((p) => p + 1)}
        >
          Siguiente
        </button>
      </div>
    </div>
  );
}

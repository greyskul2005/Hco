import { ECUACION } from "@/lib/hc/calc";
import type { HcKpis } from "@/lib/hc/types";
import { cn } from "@/lib/utils";

export function KpiStrip({ kpis }: { kpis: HcKpis | null }) {
  const items = [
    { label: "HC actual", value: kpis?.hc_actual ?? "—", tone: "bg-primary text-primary-foreground" },
    { label: "Incrementos", value: kpis ? `+${kpis.incrementos}` : "—", tone: "bg-ring text-primary-foreground" },
    { label: "Salidas CEV", value: kpis?.salidas_cev ?? "—", tone: "bg-cyan text-cyan-foreground" },
    { label: "Sustituciones", value: kpis?.sustituciones ?? "—", tone: "bg-cyan text-cyan-foreground" },
    { label: "HC previsto", value: kpis?.hc_previsto_final ?? "—", tone: "bg-ok-fg text-primary-foreground" },
  ];
  return (
    <div className="space-y-3">
      <div className="grid grid-cols-2 gap-3 md:grid-cols-5">
        {items.map((it) => (
          <div key={it.label} className={cn("rounded-xl px-4 py-3 shadow-border", it.tone)}>
            <p className="text-[11px] uppercase tracking-wider opacity-80">{it.label}</p>
            <p className="mt-1 font-display text-3xl tabular-nums leading-none">{it.value}</p>
          </div>
        ))}
      </div>
      <div
        className={cn(
          "rounded-xl px-4 py-3 text-center text-sm font-medium shadow-border",
          kpis?.ecuacion_ok ? "bg-ok text-ok-fg" : "bg-err text-err-fg",
        )}
      >
        Ecuación oficial {ECUACION}
        <span className="ml-2 tabular-nums">{kpis?.ecuacion_ok ? "OK" : "ERROR"}</span>
        <span className="ml-2">· CEV neto {kpis?.impacto_cev_neto ?? "—"}</span>
      </div>
    </div>
  );
}

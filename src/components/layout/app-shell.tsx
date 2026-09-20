import type { ReactNode } from "react";
import { Link, useRouterState } from "@tanstack/react-router";
import {
  AlertTriangle,
  BookOpen,
  Building2,
  CalendarCheck,
  Download,
  GitBranch,
  LayoutDashboard,
  Menu,
  Plus,
  Table2,
  UserMinus,
  Users,
  X,
} from "lucide-react";
import { useEffect, useState } from "react";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";
import { useHcStore } from "@/lib/hc/store";

const NAV = [
  { to: "/", label: "Resumen", icon: LayoutDashboard },
  { to: "/libro", label: "Libro Excel", icon: BookOpen },
  { to: "/registro", label: "Registro", icon: Table2 },
  { to: "/alta", label: "Alta", icon: Plus },
  { to: "/baja", label: "Baja", icon: UserMinus },
  { to: "/movimientos", label: "Movimientos", icon: GitBranch },
  { to: "/organizacion", label: "Organización", icon: Building2 },
  { to: "/calidad", label: "Calidad", icon: AlertTriangle },
  { to: "/cierre", label: "Cierre", icon: CalendarCheck },
  { to: "/equipo", label: "Equipo", icon: Users },
];

export function AppShell({ children }: { children: ReactNode }) {
  const pathname = useRouterState({ select: (s) => s.location.pathname });
  const { load, loaded, loading, saving, kpis, source, download } = useHcStore();
  const [open, setOpen] = useState(false);

  useEffect(() => {
    if (!loaded && !loading) void load();
  }, [load, loaded, loading]);

  const rail = (
    <div className="flex h-full flex-col bg-primary text-primary-foreground">
      <div className="px-5 pb-4 pt-6">
        <p className="font-display text-2xl font-medium tracking-tight">HC Control</p>
        <p className="mt-1 text-[11px] uppercase tracking-[0.18em] text-primary-foreground/55">
          Libro Excel operativo
        </p>
      </div>
      <nav className="flex-1 space-y-0.5 px-3">
        {NAV.map((item) => {
          const active = pathname === item.to;
          const Icon = item.icon;
          return (
            <Link
              key={item.to}
              to={item.to}
              onClick={() => setOpen(false)}
              className={cn(
                "flex h-11 items-center gap-3 rounded-lg px-3 text-sm transition-colors",
                active
                  ? "bg-white/12 text-white"
                  : "text-primary-foreground/70 hover:bg-white/6 hover:text-white",
              )}
            >
              <Icon className="size-4 shrink-0" />
              {item.label}
            </Link>
          );
        })}
      </nav>
      <div className="m-3 rounded-xl bg-white/8 p-3 text-xs text-primary-foreground/80">
        <p className="font-medium text-white">514 + 13 − 9 + 9 = 527</p>
        <p className="mt-1">
          Origen: {source === "excel" ? "libro Excel" : source === "seed" ? "semilla → Excel" : source || "—"}
        </p>
        <p className="mt-1 tabular-nums">HC actual {kpis?.hc_actual ?? "—"}</p>
      </div>
    </div>
  );

  return (
    <div className="min-h-dvh bg-bg text-foreground">
      <aside className="fixed inset-y-0 left-0 z-30 hidden w-60 lg:block">{rail}</aside>
      {open ? (
        <div className="fixed inset-0 z-40 lg:hidden">
          <button className="absolute inset-0 bg-primary/40" aria-label="Cerrar menú" onClick={() => setOpen(false)} />
          <div className="relative h-full w-64 max-w-[85vw]">{rail}</div>
        </div>
      ) : null}
      <div className="lg:pl-60">
        <header className="sticky top-0 z-20 flex h-14 items-center gap-3 border-b border-border bg-paper/90 px-4 backdrop-blur">
          <Button variant="ghost" size="icon" className="lg:hidden" onClick={() => setOpen(true)} aria-label="Menú">
            {open ? <X /> : <Menu />}
          </Button>
          <div className="min-w-0 flex-1">
            <p className="truncate text-sm text-muted-foreground">
              Organización Demo · datos sintéticos · el Excel es el almacén
            </p>
          </div>
          <span className="hidden text-xs text-muted-foreground sm:inline tabular-nums">
            {saving ? "Guardando Excel…" : loading ? "Leyendo Excel…" : "Libro sincronizado"}
          </span>
          <Button size="sm" variant="outline" onClick={() => void download()}>
            <Download />
            Descargar Excel
          </Button>
        </header>
        {loading ? (
          <div className="border-b border-border bg-cyan-soft px-4 py-2 text-sm text-cyan">
            Leyendo el libro Excel y materializando el almacén…
          </div>
        ) : null}
        <div className="px-4 py-6 sm:px-6 lg:px-8">{children}</div>
      </div>
    </div>
  );
}

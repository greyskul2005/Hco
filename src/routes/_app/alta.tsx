import type { ReactNode } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { useMemo, useState } from "react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Input, Label, Select, Textarea } from "@/components/ui/input";
import { impactoDe } from "@/lib/hc/calc";
import { EMPRESAS, PROGRAMAS, RESPONSABLES, UNIDADES } from "@/lib/hc/columns";
import { useHcStore } from "@/lib/hc/store";
import type { HcRecord } from "@/lib/hc/types";

export const Route = createFileRoute("/_app/alta")({ component: AltaPage });

function AltaPage() {
  const { addRecord, records } = useHcStore();
  const [form, setForm] = useState<Partial<HcRecord>>({
    TipoRegistro: "Incremento",
    Estado: "Incorporacion_Prevista",
    ComputaHC: "Si",
    OrigenPlaza: "Incremento_Aprobado",
    Programa: "Incremento_Estructural",
    Empresa: "Empresa Demo Corporativa",
    MesEfecto: "2026-09",
    IndicadorEstructura: "Si",
    DatoSintetico: "Si",
  });
  const impacto = impactoDe(form.TipoRegistro ?? "Incremento", form.Estado ?? "Incorporacion_Prevista");
  const salidas = useMemo(
    () => records.filter((r) => r.Estado === "Salida_Prevista" && r.TipoRegistro === "Base"),
    [records],
  );

  function set<K extends keyof HcRecord>(key: K, value: HcRecord[K]) {
    setForm((f) => {
      const next = { ...f, [key]: value };
      if (key === "TipoRegistro") {
        if (value === "Incremento") {
          next.OrigenPlaza = "Incremento_Aprobado";
          next.ComputaHC = "Si";
          next.Estado = "Incorporacion_Prevista";
        } else if (value === "Sustitucion") {
          next.OrigenPlaza = "CEV_Sustitucion";
          next.Programa = "CEV";
          next.ComputaHC = "Si";
          next.Estado = "Incorporacion_Prevista";
        } else if (value === "Fuera_HC") {
          next.OrigenPlaza = "Fuera_Estructura";
          next.ComputaHC = "No";
          next.IndicadorEstructura = "No";
        }
      }
      if (key === "UnidadOrganizativa") {
        next.ResponsableUnidad = RESPONSABLES[String(value)] ?? "";
        next.Responsable = next.ResponsableUnidad;
      }
      if (key === "MatriculaSustituida") {
        const origin = records.find((r) => r.Matricula === value);
        if (origin) {
          next.IDPlaza = origin.IDPlaza;
          next.UnidadOrganizativa = origin.UnidadOrganizativa;
          next.Responsable = origin.Responsable;
          next.ResponsableUnidad = origin.ResponsableUnidad;
          next.Empresa = origin.Empresa;
        }
      }
      return next;
    });
  }

  return (
    <div className="mx-auto max-w-3xl space-y-6">
      <div>
        <h1 className="font-display text-4xl font-medium">Alta en el libro</h1>
        <p className="mt-2 text-sm text-muted-foreground">
          El registro se añade a CONTROL_HC y se reescribe el Excel con el color de su tipo. Incremento = plaza
          nueva (+1). Sustitución = reutiliza plaza (0). No mezclar.
        </p>
      </div>
      <form
        className="space-y-4 rounded-2xl bg-card p-5 shadow-border"
        onSubmit={async (e) => {
          e.preventDefault();
          if (!form.NombreCompleto || !form.UnidadOrganizativa) {
            toast.error("Nombre y unidad son obligatorios");
            return;
          }
          if (form.TipoRegistro === "Sustitucion" && !form.MatriculaSustituida) {
            toast.error("Una sustitución requiere la matrícula sustituida");
            return;
          }
          const rec = await addRecord({
            ...form,
            ImpactoHC: impacto,
            Motivo:
              form.TipoRegistro === "Sustitucion"
                ? "Sustitución. Reutiliza plaza. No crea HC."
                : form.TipoRegistro === "Incremento"
                  ? "Incremento estructural. Crea plaza nueva."
                  : form.Motivo,
          });
          toast.success(`Escrito en Excel: ${rec.Matricula}`);
        }}
      >
        <div className="grid gap-4 md:grid-cols-2">
          <Field label="Tipo de registro">
            <Select value={form.TipoRegistro} onChange={(e) => set("TipoRegistro", e.target.value)}>
              <option>Incremento</option>
              <option>Sustitucion</option>
              <option>Base</option>
              <option>Fuera_HC</option>
            </Select>
          </Field>
          <Field label="Estado">
            <Select value={form.Estado} onChange={(e) => set("Estado", e.target.value)}>
              <option>Incorporacion_Prevista</option>
              <option>Activo</option>
            </Select>
          </Field>
          <Field label="Matrícula (texto, vacía = automática)">
            <Input value={form.Matricula ?? ""} onChange={(e) => set("Matricula", e.target.value)} className="font-mono" />
          </Field>
          <Field label="Nombre y apellidos">
            <Input value={form.NombreCompleto ?? ""} onChange={(e) => set("NombreCompleto", e.target.value)} />
          </Field>
          <Field label="Empresa">
            <Select value={form.Empresa} onChange={(e) => set("Empresa", e.target.value)}>
              {EMPRESAS.map((x) => (
                <option key={x}>{x}</option>
              ))}
            </Select>
          </Field>
          <Field label="Unidad organizativa">
            <Select value={form.UnidadOrganizativa ?? ""} onChange={(e) => set("UnidadOrganizativa", e.target.value)}>
              <option value="">Seleccione</option>
              {UNIDADES.map((x) => (
                <option key={x}>{x}</option>
              ))}
            </Select>
          </Field>
          <Field label="Puesto">
            <Input value={form.Puesto ?? ""} onChange={(e) => set("Puesto", e.target.value)} />
          </Field>
          <Field label="Programa">
            <Select value={form.Programa} onChange={(e) => set("Programa", e.target.value)}>
              {PROGRAMAS.map((x) => (
                <option key={x}>{x}</option>
              ))}
            </Select>
          </Field>
          <Field label="Fecha prevista alta">
            <Input type="date" value={form.FechaPrevistaAlta ?? ""} onChange={(e) => set("FechaPrevistaAlta", e.target.value)} />
          </Field>
          <Field label="Mes de efecto">
            <Select value={form.MesEfecto} onChange={(e) => set("MesEfecto", e.target.value)}>
              <option>2026-09</option>
              <option>2026-10</option>
              <option>2026-11</option>
              <option>2026-12</option>
            </Select>
          </Field>
          {form.TipoRegistro === "Sustitucion" ? (
            <Field label="Matrícula sustituida">
              <Select
                value={form.MatriculaSustituida ?? ""}
                onChange={(e) => set("MatriculaSustituida", e.target.value)}
              >
                <option value="">Seleccione baja prevista</option>
                {salidas.map((r) => (
                  <option key={r.Matricula} value={r.Matricula}>
                    {r.Matricula} · {r.NombreCompleto}
                  </option>
                ))}
              </Select>
            </Field>
          ) : null}
          <Field label="Motivo">
            <Textarea value={form.Motivo ?? ""} onChange={(e) => set("Motivo", e.target.value)} />
          </Field>
        </div>
        <div className="flex items-center justify-between rounded-xl bg-calc px-4 py-3">
          <p className="text-sm">
            Impacto HC <span className="font-display text-2xl tabular-nums">{impacto > 0 ? "+1" : "0"}</span>
            <span className="ml-2 text-muted-foreground">
              {form.TipoRegistro === "Sustitucion" ? "reutiliza plaza" : form.TipoRegistro === "Incremento" ? "crea plaza" : "sin alta estructural"}
            </span>
          </p>
          <Button type="submit">Guardar en Excel</Button>
        </div>
      </form>
    </div>
  );
}

function Field({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="space-y-1.5">
      <Label>{label}</Label>
      {children}
    </div>
  );
}

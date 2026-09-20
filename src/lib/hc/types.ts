export type TipoRegistro = "Base" | "Incremento" | "Sustitucion" | "Fuera_HC";
export type Estado =
  | "Activo"
  | "Incorporacion_Prevista"
  | "Salida_Prevista"
  | "Baja_Efectiva"
  | "Cancelado";
export type SiNo = "Si" | "No";

export interface HcRecord {
  Matricula: string;
  NombreCompleto: string;
  Empresa: string;
  Correo: string;
  Puesto: string;
  Responsable: string;
  UnidadOrganizativa: string;
  ResponsableUnidad: string;
  FechaAltaRRHH: string;
  IndicadorEstructura: SiNo | string;
  SuccessID: string;
  TipoRegistro: TipoRegistro | string;
  Estado: Estado | string;
  ComputaHC: SiNo | string;
  OrigenPlaza: string;
  Programa: string;
  FechaPrevistaAlta: string;
  FechaRealAlta: string;
  FechaPrevistaBaja: string;
  FechaRealBaja: string;
  MesEfecto: string;
  ImpactoHC: number;
  IDPlaza: string;
  MatriculaSustituida: string;
  Motivo: string;
  Observaciones: string;
  Fuente: string;
  EstadoCalidad: string;
  DatoSintetico: string;
}

export interface DecisionPendiente {
  id: string;
  descripcion: string;
  importancia: string;
  responsable_por_confirmar: string;
  impacto: string;
  estado: string;
}

export interface EvolucionMes {
  mes: string;
  etiqueta: string;
  hc: number;
  esperado: number;
  ok: boolean;
}

export interface HcKpis {
  registros: number;
  hc_base_agosto: number;
  hc_actual: number;
  incrementos: number;
  salidas_cev: number;
  sustituciones: number;
  fuera_hc: number;
  incompletos: number;
  duplicados_registros: number;
  matriculas_duplicadas: string[];
  hc_estructural: number;
  hc_previsto_final: number;
  ecuacion: string;
  ecuacion_ok: boolean;
  evolucion: EvolucionMes[];
  impacto_cev_neto: number;
}

export interface HcPayload {
  meta: { version: string; generado: string; marca: string; escenario: string };
  parametros: {
    hc_base_agosto_2026: number;
    incrementos_total: number;
    salidas_cev: number;
    sustituciones_cev: number;
    hc_estructural_final: number;
    empresas: string[];
    unidades: {
      codigo: string;
      nombre: string;
      responsable: string;
      empresa: string;
      base: number;
    }[];
    tipos_registro: string[];
    estados: string[];
    programas: string[];
    origen_plaza: string[];
  };
  kpis: HcKpis;
  decisiones_pendientes: DecisionPendiente[];
  columnas: {
    tecnico: string;
    visible: string;
    tipo_excel: string;
    tipo_lists: string;
    obligatorio: boolean;
    default: string;
    valores: string;
    descripcion: string;
    responsable: string;
  }[];
  registros: HcRecord[];
}

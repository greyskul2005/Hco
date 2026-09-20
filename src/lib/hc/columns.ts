export const COLUMNAS = [
  { tecnico: "Matricula", visible: "Matrícula", excel: "texto" },
  { tecnico: "NombreCompleto", visible: "Nombre y apellidos", excel: "texto" },
  { tecnico: "Empresa", visible: "Empresa", excel: "texto" },
  { tecnico: "Correo", visible: "Correo electrónico", excel: "texto" },
  { tecnico: "Puesto", visible: "Puesto", excel: "texto" },
  { tecnico: "Responsable", visible: "Responsable", excel: "texto" },
  { tecnico: "UnidadOrganizativa", visible: "Unidad organizativa", excel: "texto" },
  { tecnico: "ResponsableUnidad", visible: "Responsable de unidad", excel: "texto" },
  { tecnico: "FechaAltaRRHH", visible: "Fecha de alta RRHH", excel: "fecha" },
  { tecnico: "IndicadorEstructura", visible: "Indicador de estructura", excel: "texto" },
  { tecnico: "SuccessID", visible: "SuccessID", excel: "texto" },
  { tecnico: "TipoRegistro", visible: "Tipo de registro", excel: "texto" },
  { tecnico: "Estado", visible: "Estado", excel: "texto" },
  { tecnico: "ComputaHC", visible: "Computa HC", excel: "texto" },
  { tecnico: "OrigenPlaza", visible: "Origen de plaza", excel: "texto" },
  { tecnico: "Programa", visible: "Programa", excel: "texto" },
  { tecnico: "FechaPrevistaAlta", visible: "Fecha prevista alta", excel: "fecha" },
  { tecnico: "FechaRealAlta", visible: "Fecha real alta", excel: "fecha" },
  { tecnico: "FechaPrevistaBaja", visible: "Fecha prevista baja", excel: "fecha" },
  { tecnico: "FechaRealBaja", visible: "Fecha real baja", excel: "fecha" },
  { tecnico: "MesEfecto", visible: "Mes de efecto", excel: "texto" },
  { tecnico: "ImpactoHC", visible: "Impacto HC", excel: "numero" },
  { tecnico: "IDPlaza", visible: "ID de plaza", excel: "texto" },
  { tecnico: "MatriculaSustituida", visible: "Matrícula sustituida", excel: "texto" },
  { tecnico: "Motivo", visible: "Motivo", excel: "texto" },
  { tecnico: "Observaciones", visible: "Observaciones", excel: "texto" },
  { tecnico: "Fuente", visible: "Fuente", excel: "texto" },
  { tecnico: "EstadoCalidad", visible: "Estado de calidad", excel: "texto" },
  { tecnico: "DatoSintetico", visible: "Dato sintético", excel: "texto" },
] as const;

export const COLUMN_ORDER = COLUMNAS.map((c) => c.tecnico);
export const TEXT_KEYS = new Set([
  "Matricula",
  "SuccessID",
  "IDPlaza",
  "MatriculaSustituida",
  "Correo",
]);

export const TIPOS = ["Base", "Incremento", "Sustitucion", "Fuera_HC"] as const;
export const ESTADOS = [
  "Activo",
  "Incorporacion_Prevista",
  "Salida_Prevista",
  "Baja_Efectiva",
  "Cancelado",
] as const;
export const PROGRAMAS = [
  "Regular",
  "Becarios",
  "RRLL",
  "CEV",
  "Incremento_Estructural",
] as const;
export const ORIGENES = [
  "Estructural",
  "Incremento_Aprobado",
  "CEV_Sustitucion",
  "Fuera_Estructura",
] as const;
export const SI_NO = ["Si", "No"] as const;
export const EMPRESAS = ["Empresa Demo Corporativa", "Empresa Demo Servicios"] as const;
export const UNIDADES = [
  "Dirección General",
  "Finanzas",
  "Operaciones",
  "Tecnología",
  "Comercial",
  "Recursos Humanos",
  "Relaciones Laborales",
  "Legal y Cumplimiento",
  "Calidad",
  "Proyectos",
] as const;

export const RESPONSABLES: Record<string, string> = {
  "Dirección General": "Nuria Demo Valls",
  Finanzas: "Hugo Demo Serra",
  Operaciones: "Marta Demo Rius",
  Tecnología: "Pol Demo Ferrer",
  Comercial: "Elena Demo Pujol",
  "Recursos Humanos": "Irene Demo Casals",
  "Relaciones Laborales": "Joan Demo Vidal",
  "Legal y Cumplimiento": "Clara Demo Bosch",
  Calidad: "Paula Demo Rovira",
  Proyectos: "Oriol Demo Font",
};

export const HEADER_BY_VISIBLE: Record<string, string> = Object.fromEntries(
  COLUMNAS.map((c) => [c.visible, c.tecnico]),
);
export const HEADER_BY_TECH = Object.fromEntries(
  COLUMNAS.map((c) => [c.tecnico, c.tecnico]),
);

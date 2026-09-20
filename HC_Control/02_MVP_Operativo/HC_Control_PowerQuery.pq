// HC Control · Power Query (M)
// Origen preferente: Microsoft Lists CONTROL_HC
// Alternativa: Excel HC_Control_MVP_Operativo.xlsx hoja CONTROL_HC_IMPORTAR
// Identificadores SIEMPRE como texto.

let
    Origen = SharePoint.Tables("https://TENANT.sharepoint.com/sites/SITIO_HC", [ApiVersion = 15]),
    CONTROL_HC = Origen{[Title="CONTROL_HC"]}[Items],
    Tipos = Table.TransformColumnTypes(
        CONTROL_HC,
        {
            {"Matricula", type text},
            {"NombreCompleto", type text},
            {"Empresa", type text},
            {"Correo", type text},
            {"Puesto", type text},
            {"Responsable", type text},
            {"UnidadOrganizativa", type text},
            {"ResponsableUnidad", type text},
            {"FechaAltaRRHH", type date},
            {"IndicadorEstructura", type text},
            {"SuccessID", type text},
            {"TipoRegistro", type text},
            {"Estado", type text},
            {"ComputaHC", type text},
            {"OrigenPlaza", type text},
            {"Programa", type text},
            {"FechaPrevistaAlta", type date},
            {"FechaRealAlta", type date},
            {"FechaPrevistaBaja", type date},
            {"FechaRealBaja", type date},
            {"MesEfecto", type text},
            {"ImpactoHC", Int64.Type},
            {"IDPlaza", type text},
            {"MatriculaSustituida", type text},
            {"Motivo", type text},
            {"Observaciones", type text},
            {"Fuente", type text},
            {"EstadoCalidad", type text},
            {"DatoSintetico", type text}
        }
    ),
    Calidad = Table.AddColumn(
        Tipos,
        "FlagCalidad",
        each
            if [Matricula] = null or [Matricula] = "" then "Incompleto"
            else if [TipoRegistro] = "" then "Sin clasificacion"
            else if [TipoRegistro] = "Sustitucion" and ([MatriculaSustituida] = null or [MatriculaSustituida] = "") then "Incompleto"
            else [EstadoCalidad]
    )
in
    Calidad

// TENANT y SITIO_HC son pendientes (DEC-008). No inventados.

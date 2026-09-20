"""Fuente canónica de datos sintéticos HC Control v2.0.

Todas las personas, matrículas, responsables y fechas operativas son SINTÉTICAS.
Las cifras 514 / +13 / -9 / +9 / 527 son DATOS CONFIRMADOS del escenario.
"""
from __future__ import annotations

import json
from copy import deepcopy
from datetime import date
from pathlib import Path

VERSION = "2.0"
GENERATED = "2026-09-20"
MARCA_SINTETICO = "SINTÉTICO · demostración · no utilizar como dato real de RRHH"

# --- Cifras confirmadas (no modificar) ---
HC_BASE_AGO_2026 = 514
INCREMENTOS_INICIALES = 7  # 6 becarios + 1 RRLL
INCREMENTOS_OCT = 2
INCREMENTOS_NOV = 3
INCREMENTOS_DIC = 1
INCREMENTOS_TOTAL = 13
SALIDAS_CEV = 9
SUSTITUCIONES_CEV = 9
HC_ESTRUCTURAL_FINAL = 527
ECUACION = "514 + 13 - 9 + 9 = 527"
EVOLUCION = [
    ("2026-08", "Agosto 2026 (cierre)", 514),
    ("2026-09", "Tras incrementos iniciales", 521),
    ("2026-10", "Octubre 2026", 523),
    ("2026-11", "Noviembre 2026", 526),
    ("2026-12", "Diciembre 2026", 527),
]

# --- Catálogos (supuesto de diseño: unidades sintéticas) ---
EMPRESAS = ["Empresa Demo Corporativa", "Empresa Demo Servicios"]

UNIDADES = [
    # codigo, nombre, responsable, empresa, n_base
    ("DG", "Dirección General", "Nuria Demo Valls", "Empresa Demo Corporativa", 4),
    ("FIN", "Finanzas", "Hugo Demo Serra", "Empresa Demo Corporativa", 52),
    ("OPE", "Operaciones", "Marta Demo Rius", "Empresa Demo Servicios", 140),
    ("TEC", "Tecnología", "Pol Demo Ferrer", "Empresa Demo Corporativa", 82),
    ("COM", "Comercial", "Elena Demo Pujol", "Empresa Demo Servicios", 108),
    ("RRHH", "Recursos Humanos", "Irene Demo Casals", "Empresa Demo Corporativa", 24),
    ("RRLL", "Relaciones Laborales", "Joan Demo Vidal", "Empresa Demo Corporativa", 18),
    ("LEG", "Legal y Cumplimiento", "Clara Demo Bosch", "Empresa Demo Corporativa", 16),
    ("CAL", "Calidad", "Paula Demo Rovira", "Empresa Demo Servicios", 22),
    ("PRO", "Proyectos", "Oriol Demo Font", "Empresa Demo Corporativa", 48),
]

PUESTOS = {
    "Dirección General": ["Dirección General", "Adjuntía de Dirección", "Secretaría de Dirección", "Staff de Dirección"],
    "Finanzas": ["Controller", "Analista financiero", "Tesorería", "Consolidación", "Fiscalidad", "Reporting"],
    "Operaciones": ["Coordinación de operaciones", "Supervisor/a de planta", "Técnico/a de operaciones", "Planificación", "Logística"],
    "Tecnología": ["Ingeniería de software", "Analista de sistemas", "Ciberseguridad", "Arquitectura IT", "Soporte corporativo"],
    "Comercial": ["Gestión comercial", "Key account", "Preventa", "Atención a cliente", "Desarrollo de negocio"],
    "Recursos Humanos": ["People partner", "Administración de personal", "Selección", "Desarrollo de talento"],
    "Relaciones Laborales": ["Técnico/a de RRLL", "Compensación", "Prevención", "Relaciones sindicales"],
    "Legal y Cumplimiento": ["Asesoría jurídica", "Cumplimiento", "Protección de datos", "Contratación"],
    "Calidad": ["Técnico/a de calidad", "Auditoría interna", "Sistemas de gestión", "Mejora continua"],
    "Proyectos": ["Jefatura de proyecto", "PMO", "Analista de proyectos", "Oficina de transformación"],
}

NOMBRES = [
    "Alex", "Aina", "Bruno", "Carla", "Duna", "Enric", "Ferran", "Gala", "Helena", "Iria",
    "Joel", "Laia", "Marc", "Noa", "Ona", "Pau", "Queralt", "Roc", "Salma", "Toni",
    "Unai", "Vera", "Xenia", "Yago", "Zoe", "Berta", "Cesc", "Dídac", "Emma", "Fran",
]
APELLIDOS = [
    "Alba", "Bosch", "Casals", "Duran", "Espasa", "Font", "Grau", "Homs", "Isern", "Jové",
    "Llobet", "Miró", "Nadal", "Olivé", "Pujol", "Rius", "Serra", "Tormo", "Ubach", "Vidal",
]

TIPOS_REGISTRO = ["Base", "Incremento", "Sustitucion", "Fuera_HC"]
ESTADOS = ["Activo", "Incorporacion_Prevista", "Salida_Prevista", "Baja_Efectiva", "Cancelado"]
COMPUTA = ["Si", "No"]
ORIGEN_PLAZA = ["Estructural", "Incremento_Aprobado", "CEV_Sustitucion", "Fuera_Estructura"]
PROGRAMAS = ["Regular", "Becarios", "RRLL", "CEV", "Incremento_Estructural"]
INDICADOR_ESTRUCTURA = ["Si", "No"]
FUENTES = ["RRHH", "Manual", "Cierre"]
ESTADOS_CALIDAD = ["Completo", "Incompleto", "Duplicado", "Incoherente"]
MESES_EFECTO = ["2026-08", "2026-09", "2026-10", "2026-11", "2026-12"]

COLUMNAS = [
    # (tecnico, visible, tipo_excel, tipo_lists, obligatorio, default, valores, descripcion, responsable)
    ("Matricula", "Matrícula", "texto", "Una línea de texto", True, "", "", "Clave funcional de persona. Siempre texto. No usar número.", "Operación HC"),
    ("NombreCompleto", "Nombre y apellidos", "texto", "Una línea de texto", True, "", "", "Nombre sintético o real procedente de RRHH. No almacenar otros datos personales.", "RRHH / Operación HC"),
    ("Empresa", "Empresa", "texto", "Opción", True, "Empresa Demo Corporativa", " | ".join(EMPRESAS), "Entidad jurídica. Catálogo pendiente de confirmar con RRHH.", "RRHH"),
    ("Correo", "Correo electrónico", "texto", "Una línea de texto", False, "", "", "Correo corporativo. No obligatorio en incorporaciones previstas sin alta RRHH.", "RRHH"),
    ("Puesto", "Puesto", "texto", "Una línea de texto", True, "", "", "Denominación de puesto. No es catálogo cerrado en el MVP.", "RRHH"),
    ("Responsable", "Responsable", "texto", "Una línea de texto", False, "", "", "Nombre del responsable jerárquico. Texto, no persona de Microsoft 365, para evitar dependencias.", "RRHH"),
    ("UnidadOrganizativa", "Unidad organizativa", "texto", "Opción", True, "", " | ".join(u[1] for u in UNIDADES), "Unidad de adscripción. Catálogo sintético pendiente de sustituir por el real.", "RRHH"),
    ("ResponsableUnidad", "Responsable de la unidad organizativa", "texto", "Una línea de texto", False, "", "", "Responsable de la unidad. Se informa por catálogo para evitar tecleo masivo.", "Operación HC"),
    ("FechaAltaRRHH", "Fecha de alta", "fecha", "Fecha", False, "", "", "Fecha de alta en el sistema de RRHH. Vacía si aún no existe persona real.", "RRHH"),
    ("IndicadorEstructura", "Indicador de estructura", "texto", "Opción", True, "Si", "Si | No", "Si la persona/plaza forma parte de estructura. Distinto de ComputaHC.", "RRHH"),
    ("SuccessID", "SuccessID", "texto", "Una línea de texto", False, "", "", "Identificador del sistema de talento. Texto. Puede estar vacío.", "RRHH"),
    ("TipoRegistro", "Tipo de registro", "texto", "Opción", True, "Base", " | ".join(TIPOS_REGISTRO), "Base, Incremento, Sustitucion o Fuera_HC. Nunca mezclar incremento y sustitución.", "Operación HC"),
    ("Estado", "Estado", "texto", "Opción", True, "Activo", " | ".join(ESTADOS), "Situación operativa del registro.", "Operación HC"),
    ("ComputaHC", "Computa HC", "texto", "Opción", True, "Si", "Si | No", "Si el registro entra en el recuento de headcount.", "Operación HC"),
    ("OrigenPlaza", "Origen de plaza", "texto", "Opción", True, "Estructural", " | ".join(ORIGEN_PLAZA), "De dónde nace la plaza. Las sustituciones CEV reutilizan plaza estructural.", "Operación HC"),
    ("Programa", "Programa", "texto", "Opción", False, "Regular", " | ".join(PROGRAMAS), "Programa de origen. CEV no crea plaza nueva.", "Operación HC"),
    ("FechaPrevistaAlta", "Fecha prevista de alta", "fecha", "Fecha", False, "", "", "Planificación. No equivale a alta real.", "Operación HC"),
    ("FechaRealAlta", "Fecha real de alta", "fecha", "Fecha", False, "", "", "Hecho. Solo cuando la persona está incorporada.", "Operación HC"),
    ("FechaPrevistaBaja", "Fecha prevista de baja", "fecha", "Fecha", False, "", "", "Planificación de salida.", "Operación HC"),
    ("FechaRealBaja", "Fecha real de baja", "fecha", "Fecha", False, "", "", "Hecho de salida.", "Operación HC"),
    ("MesEfecto", "Mes de efecto", "texto", "Opción", False, "", " | ".join(MESES_EFECTO), "Mes en el que el movimiento afecta a la serie mensual. Formato AAAA-MM.", "Operación HC"),
    ("ImpactoHC", "Impacto HC", "número", "Número", True, "0", "+1 | 0 | -1", "Impacto estructural del registro. Calculado por regla, no tecleado.", "Sistema"),
    ("IDPlaza", "ID de plaza", "texto", "Una línea de texto", False, "", "", "Puente al modelo objetivo PLAZAS_HC. En sustitución CEV coincide con la plaza de la persona que sale.", "Operación HC"),
    ("MatriculaSustituida", "Matrícula sustituida", "texto", "Una línea de texto", False, "", "", "Obligatorio si TipoRegistro = Sustitucion. Enlaza con la baja que libera la plaza.", "Operación HC"),
    ("Motivo", "Motivo", "texto largo", "Varias líneas de texto", False, "", "", "Causa de negocio del movimiento.", "Operación HC"),
    ("Observaciones", "Observaciones", "texto largo", "Varias líneas de texto", False, "", "", "Notas operativas. Sin datos personales innecesarios.", "Operación HC"),
    ("Fuente", "Fuente", "texto", "Opción", True, "Manual", " | ".join(FUENTES), "Origen del registro.", "Operación HC"),
    ("EstadoCalidad", "Estado de calidad", "texto", "Opción", True, "Completo", " | ".join(ESTADOS_CALIDAD), "Resultado del control de calidad. Recalculable.", "Operación HC"),
    ("DatoSintetico", "Dato sintético", "texto", "Opción", True, "Si", "Si | No", "Si = demostración. En producción todos deben pasar a No tras la carga real.", "Gobierno del dato"),
]

COLUMN_ORDER = [c[0] for c in COLUMNAS]


def _nombre(i: int) -> str:
    n = NOMBRES[i % len(NOMBRES)]
    a1 = APELLIDOS[i % len(APELLIDOS)]
    a2 = APELLIDOS[(i * 3) % len(APELLIDOS)]
    return f"{n} Demo {a1} {a2}"


def _correo(matricula: str) -> str:
    return f"{matricula.lower()}@demo.hccontrol.local"


def _impacto(tipo: str, estado: str) -> int:
    if estado == "Cancelado":
        return 0
    if tipo == "Incremento":
        return 1
    if tipo == "Sustitucion":
        return 0
    if tipo == "Fuera_HC":
        return 0
    if tipo == "Base":
        return 0
    return 0


def _rec(**kwargs) -> dict:
    row = {k: "" for k in COLUMN_ORDER}
    row.update(
        {
            "IndicadorEstructura": "Si",
            "ComputaHC": "Si",
            "Estado": "Activo",
            "TipoRegistro": "Base",
            "OrigenPlaza": "Estructural",
            "Programa": "Regular",
            "Fuente": "RRHH",
            "EstadoCalidad": "Completo",
            "DatoSintetico": "Si",
            "ImpactoHC": 0,
            "MesEfecto": "2026-08",
        }
    )
    row.update(kwargs)
    row["ImpactoHC"] = _impacto(row["TipoRegistro"], row["Estado"])
    return row


def build_records() -> list[dict]:
    records: list[dict] = []
    n = 1
    cev_ids: list[str] = []
    # 9 bajas CEV: 3 Operaciones, 3 Comercial, 3 Tecnología (supuesto de diseño)
    cev_slots = {
        "Operaciones": {20, 60, 100},
        "Comercial": {10, 40, 80},
        "Tecnología": {5, 35, 70},
    }

    for codigo, unidad, resp_u, empresa, n_base in UNIDADES:
        puestos = PUESTOS[unidad]
        for j in range(n_base):
            matricula = f"SYN-{n:04d}"
            is_cev = j in cev_slots.get(unidad, set())
            if is_cev:
                cev_ids.append(matricula)
            alta = date(2018 + (n % 8), ((n * 3) % 12) + 1, ((n * 5) % 27) + 1)
            rec = _rec(
                Matricula=matricula,
                NombreCompleto=_nombre(n),
                Empresa=empresa,
                Correo=_correo(matricula),
                Puesto=puestos[j % len(puestos)],
                Responsable=resp_u,
                UnidadOrganizativa=unidad,
                ResponsableUnidad=resp_u,
                FechaAltaRRHH=alta.isoformat(),
                SuccessID=f"SID-{n:04d}",
                TipoRegistro="Base",
                Estado="Salida_Prevista" if is_cev else "Activo",
                ComputaHC="Si",
                OrigenPlaza="Estructural",
                Programa="CEV" if is_cev else "Regular",
                FechaPrevistaBaja="2026-11-30" if is_cev else "",
                MesEfecto="2026-11" if is_cev else "2026-08",
                IDPlaza=f"PLZ-{n:04d}",
                Motivo="Salida prevista al programa CEV. La plaza se reutiliza." if is_cev else "Fotografía de cierre agosto 2026.",
                Observaciones="Dato sintético. " + ("Pendiente confirmar calendario real CEV. Supuesto: efecto en noviembre 2026 junto con la sustitución, neto 0." if is_cev else "Base autorizada."),
                Fuente="RRHH",
            )
            records.append(rec)
            n += 1

    assert len(records) == HC_BASE_AGO_2026, len(records)
    assert len(cev_ids) == SALIDAS_CEV, cev_ids

    # 13 incrementos
    incrementos_spec = (
        [("Becarios", "2026-09-15", "2026-09", "Becario/a con incremento de HC")] * 6
        + [("RRLL", "2026-09-22", "2026-09", "Incorporación RRLL con incremento de HC")]
        + [("Incremento_Estructural", "2026-10-01", "2026-10", "Posición adicional aprobada")] * 2
        + [("Incremento_Estructural", "2026-11-03", "2026-11", "Posición adicional aprobada")] * 3
        + [("Incremento_Estructural", "2026-12-01", "2026-12", "Posición adicional aprobada")]
    )
    dest_units = [
        "Operaciones", "Comercial", "Tecnología", "Finanzas", "Proyectos", "Calidad",
        "Relaciones Laborales",
        "Operaciones", "Comercial",
        "Tecnología", "Finanzas", "Proyectos",
        "Recursos Humanos",
    ]
    for i, ((programa, fprev, mes, puesto), unidad) in enumerate(zip(incrementos_spec, dest_units), start=1):
        codigo = next(u for u in UNIDADES if u[1] == unidad)
        matricula = f"SYN-I{i:02d}"
        records.append(
            _rec(
                Matricula=matricula,
                NombreCompleto=_nombre(800 + i),
                Empresa=codigo[3],
                Correo=_correo(matricula) if i <= 7 else "",
                Puesto=puesto,
                Responsable=codigo[2],
                UnidadOrganizativa=unidad,
                ResponsableUnidad=codigo[2],
                FechaAltaRRHH="",
                IndicadorEstructura="Si",
                SuccessID=f"SID-I{i:02d}",
                TipoRegistro="Incremento",
                Estado="Incorporacion_Prevista",
                ComputaHC="Si",
                OrigenPlaza="Incremento_Aprobado",
                Programa=programa,
                FechaPrevistaAlta=fprev,
                MesEfecto=mes,
                IDPlaza=f"PLZ-INC-{i:02d}",
                Motivo="Incremento estructural aprobado. Crea plaza nueva.",
                Observaciones="Dato sintético. Fecha prevista = supuesto de diseño. Referencia de aprobación PENDIENTE.",
                Fuente="Manual",
            )
        )

    # 9 sustituciones CEV (reutilizan plaza, impacto 0)
    for i, mat_sal in enumerate(cev_ids, start=1):
        sal = next(r for r in records if r["Matricula"] == mat_sal)
        matricula = f"SYN-S{i:02d}"
        records.append(
            _rec(
                Matricula=matricula,
                NombreCompleto=_nombre(900 + i),
                Empresa=sal["Empresa"],
                Correo=_correo(matricula),
                Puesto="Becario/a sustituto CEV",
                Responsable=sal["Responsable"],
                UnidadOrganizativa=sal["UnidadOrganizativa"],
                ResponsableUnidad=sal["ResponsableUnidad"],
                FechaAltaRRHH="",
                SuccessID=f"SID-S{i:02d}",
                TipoRegistro="Sustitucion",
                Estado="Incorporacion_Prevista",
                ComputaHC="Si",
                OrigenPlaza="CEV_Sustitucion",
                Programa="CEV",
                FechaPrevistaAlta="2026-11-03",
                FechaPrevistaBaja="",
                MesEfecto="2026-11",
                IDPlaza=sal["IDPlaza"],
                MatriculaSustituida=mat_sal,
                Motivo="Sustitución CEV. Reutiliza la plaza existente. No crea HC nuevo.",
                Observaciones="Dato sintético. Impacto neto del programa CEV = 0. Fecha = supuesto de diseño.",
                Fuente="Manual",
            )
        )

    # Fuera de HC (colaboradores que no computan)
    for i in range(1, 7):
        matricula = f"SYN-X{i:02d}"
        records.append(
            _rec(
                Matricula=matricula,
                NombreCompleto=_nombre(950 + i),
                Empresa="Empresa Demo Servicios",
                Correo=_correo(matricula),
                Puesto="Colaboración externa (no HC)",
                Responsable="Marta Demo Rius",
                UnidadOrganizativa="Operaciones",
                ResponsableUnidad="Marta Demo Rius",
                FechaAltaRRHH="2026-03-01",
                IndicadorEstructura="No",
                SuccessID=f"SID-X{i:02d}",
                TipoRegistro="Fuera_HC",
                Estado="Activo",
                ComputaHC="No",
                OrigenPlaza="Fuera_Estructura",
                Programa="Regular",
                MesEfecto="2026-08",
                IDPlaza="",
                Motivo="No computa en headcount estructural.",
                Observaciones="Dato sintético de control. Debe permanecer fuera del recuento 514/527.",
                Fuente="Manual",
            )
        )

    # Calidad: duplicado
    orig = next(r for r in records if r["Matricula"] == "SYN-0007")
    dup = deepcopy(orig)
    dup.update(
        {
            "NombreCompleto": orig["NombreCompleto"] + " (duplicado)",
            "EstadoCalidad": "Duplicado",
            "ComputaHC": "No",
            "TipoRegistro": "Fuera_HC",
            "OrigenPlaza": "Fuera_Estructura",
            "IndicadorEstructura": "No",
            "Motivo": "Registro duplicado de matrícula para prueba de calidad. No computa.",
            "Observaciones": "Incidencia de calidad sintética. Debe detectarse por matrícula repetida.",
            "Fuente": "Manual",
            "IDPlaza": "",
        }
    )
    records.append(dup)
    orig["EstadoCalidad"] = "Duplicado"
    orig["Observaciones"] = (orig["Observaciones"] + " Duplicidad detectada en registro paralelo de calidad.").strip()

    # Calidad: incompletos (no computan hasta completar)
    records.append(
        _rec(
            Matricula="SYN-INC01",
            NombreCompleto="Persona Demo Incompleta Uno",
            Empresa="Empresa Demo Corporativa",
            Correo="syn-inc01@demo.hccontrol.local",
            Puesto="Puesto no informado en unidad",
            Responsable="",
            UnidadOrganizativa="",
            ResponsableUnidad="",
            TipoRegistro="Incremento",
            Estado="Incorporacion_Prevista",
            ComputaHC="No",
            OrigenPlaza="Incremento_Aprobado",
            Programa="Incremento_Estructural",
            FechaPrevistaAlta="2026-10-15",
            MesEfecto="",
            IDPlaza="",
            Motivo="Caso de prueba: falta unidad organizativa.",
            Observaciones="Incompleto sintético. No forma parte del +13 aprobado.",
            Fuente="Manual",
            EstadoCalidad="Incompleto",
            IndicadorEstructura="Si",
        )
    )
    records.append(
        _rec(
            Matricula="SYN-INC02",
            NombreCompleto="Persona Demo Incompleta Dos",
            Empresa="Empresa Demo Corporativa",
            Correo="",
            Puesto="Analista",
            Responsable="Pol Demo Ferrer",
            UnidadOrganizativa="Tecnología",
            ResponsableUnidad="Pol Demo Ferrer",
            TipoRegistro="Base",
            Estado="Activo",
            ComputaHC="No",
            OrigenPlaza="Estructural",
            Programa="Regular",
            MesEfecto="2026-08",
            IDPlaza="",
            Motivo="Caso de prueba: falta correo y no computa hasta completar.",
            Observaciones="Incompleto sintético. Excluido del HC actual para no distorsionar 514.",
            Fuente="Manual",
            EstadoCalidad="Incompleto",
        )
    )

    return records


def kpis(records: list[dict]) -> dict:
    def cnt(**kwargs):
        n = 0
        for r in records:
            ok = True
            for k, v in kwargs.items():
                if r.get(k) != v:
                    ok = False
                    break
            if ok:
                n += 1
        return n

    base = cnt(TipoRegistro="Base", ComputaHC="Si")
    # 514 base computing, including CEV planned exits (still count until baja efectiva)
    hc_actual = cnt(Estado="Activo", ComputaHC="Si") + cnt(Estado="Salida_Prevista", ComputaHC="Si")
    incrementos = cnt(TipoRegistro="Incremento", Estado="Incorporacion_Prevista", ComputaHC="Si")
    # the incomplete incremento has ComputaHC No, so incrementos should be 13
    incrementos_total = sum(
        1
        for r in records
        if r["TipoRegistro"] == "Incremento" and r["Estado"] != "Cancelado" and r["ComputaHC"] == "Si"
    )
    salidas_cev = cnt(Programa="CEV", TipoRegistro="Base")
    sustituciones = cnt(TipoRegistro="Sustitucion", ComputaHC="Si")
    fuera = cnt(TipoRegistro="Fuera_HC")
    incompletos = cnt(EstadoCalidad="Incompleto")
    duplicados_reg = cnt(EstadoCalidad="Duplicado")
    matriculas = [r["Matricula"] for r in records]
    dup_keys = sorted({m for m in matriculas if matriculas.count(m) > 1})
    hc_estructural = HC_BASE_AGO_2026 + incrementos_total
    hc_previsto = HC_BASE_AGO_2026 + incrementos_total - salidas_cev + sustituciones
    evo = []
    acc = HC_BASE_AGO_2026
    for mes, etiqueta, esperado in EVOLUCION:
        if mes == "2026-08":
            valor = HC_BASE_AGO_2026
        else:
            altas = sum(
                1
                for r in records
                if r["TipoRegistro"] == "Incremento"
                and r["ComputaHC"] == "Si"
                and r["MesEfecto"] == mes
                and r["Estado"] != "Cancelado"
            )
            salidas = sum(
                1
                for r in records
                if r["TipoRegistro"] == "Base"
                and r["Programa"] == "CEV"
                and r["MesEfecto"] == mes
            )
            sust = sum(
                1
                for r in records
                if r["TipoRegistro"] == "Sustitucion"
                and r["ComputaHC"] == "Si"
                and r["MesEfecto"] == mes
            )
            acc = acc + altas - salidas + sust
            valor = acc
        evo.append(
            {
                "mes": mes,
                "etiqueta": etiqueta,
                "hc": valor,
                "esperado": esperado,
                "ok": valor == esperado,
            }
        )
    return {
        "version": VERSION,
        "marca_sintetico": MARCA_SINTETICO,
        "registros": len(records),
        "base_computa": base,
        "hc_base_agosto": HC_BASE_AGO_2026,
        "hc_actual": hc_actual,
        "incrementos": incrementos_total,
        "incrementos_iniciales": INCREMENTOS_INICIALES,
        "incrementos_oct": INCREMENTOS_OCT,
        "incrementos_nov": INCREMENTOS_NOV,
        "incrementos_dic": INCREMENTOS_DIC,
        "salidas_cev": salidas_cev,
        "sustituciones": sustituciones,
        "fuera_hc": fuera,
        "incompletos": incompletos,
        "duplicados_registros": duplicados_reg,
        "matriculas_duplicadas": dup_keys,
        "hc_estructural": hc_estructural,
        "hc_previsto_final": hc_previsto,
        "ecuacion": ECUACION,
        "ecuacion_ok": hc_previsto == HC_ESTRUCTURAL_FINAL
        and incrementos_total == INCREMENTOS_TOTAL
        and salidas_cev == SALIDAS_CEV
        and sustituciones == SUSTITUCIONES_CEV
        and hc_actual == HC_BASE_AGO_2026,
        "evolucion": evo,
        "impacto_cev_neto": sustituciones - salidas_cev,
    }


def decisiones_pendientes() -> list[dict]:
    return [
        {
            "id": "DEC-001",
            "descripcion": "Catálogo real de unidades organizativas y responsables de unidad.",
            "importancia": "Alta",
            "responsable_por_confirmar": "RRHH / Organización",
            "impacto": "Sustituye el catálogo sintético de 10 unidades. No cambia las cifras 514-527.",
            "estado": "Pendiente",
        },
        {
            "id": "DEC-002",
            "descripcion": "Catálogo real de empresas / entidades jurídicas del grupo.",
            "importancia": "Alta",
            "responsable_por_confirmar": "RRHH / Legal",
            "impacto": "Reemplaza Empresa Demo Corporativa / Servicios.",
            "estado": "Pendiente",
        },
        {
            "id": "DEC-003",
            "descripcion": "Fechas reales de efecto de los 7 incrementos iniciales (6 becarios + 1 RRLL).",
            "importancia": "Alta",
            "responsable_por_confirmar": "Dirección / RRHH",
            "impacto": "Hoy se usa 15/09/2026 y 22/09/2026 como supuesto de diseño para reproducir 521 en septiembre.",
            "estado": "Pendiente",
        },
        {
            "id": "DEC-004",
            "descripcion": "Naturaleza, puesto y unidad de las +2 (oct), +3 (nov) y +1 (dic) posiciones adicionales.",
            "importancia": "Alta",
            "responsable_por_confirmar": "Dirección",
            "impacto": "No se ha inventado el contenido de esas plazas. Constan como 'Posición adicional aprobada'.",
            "estado": "Pendiente",
        },
        {
            "id": "DEC-005",
            "descripcion": "Referencias de aprobación de los +13 incrementos estructurales.",
            "importancia": "Alta",
            "responsable_por_confirmar": "Dirección / Control de gestión",
            "impacto": "El campo Motivo no contiene números de acta. Hay que informarlos cuando existan.",
            "estado": "Pendiente",
        },
        {
            "id": "DEC-006",
            "descripcion": "Calendario real del programa CEV (fecha de salida de las 9 personas y de alta de los 9 becarios).",
            "importancia": "Alta",
            "responsable_por_confirmar": "RRHH / Programa CEV",
            "impacto": "Supuesto de diseño: salidas 30/11/2026 y sustituciones 03/11/2026, ambas con MesEfecto 2026-11, para no alterar la serie confirmada 523 (oct) / 526 (nov). Impacto neto 0.",
            "estado": "Pendiente",
        },
        {
            "id": "DEC-007",
            "descripcion": "Definición jurídica y operativa de CEV (centro, empresa, programa).",
            "importancia": "Media",
            "responsable_por_confirmar": "RRHH",
            "impacto": "No se ha interpretado CEV más allá de lo indicado: 9 salidas + 9 sustituciones, neto 0.",
            "estado": "Pendiente",
        },
        {
            "id": "DEC-008",
            "descripcion": "Sitio de SharePoint y equipo de Teams definitivos (nombre, URL, permisos).",
            "importancia": "Alta",
            "responsable_por_confirmar": "TI colaborativo / Operación HC",
            "impacto": "Bloquea la construcción en el tenant. El MVP no arranca sin sitio.",
            "estado": "Pendiente",
        },
        {
            "id": "DEC-009",
            "descripcion": "Propietario funcional, operador diario y validador del cierre mensual.",
            "importancia": "Alta",
            "responsable_por_confirmar": "Dirección",
            "impacto": "Roles descritos de forma genérica (Operador HC, Validador, Lector).",
            "estado": "Pendiente",
        },
        {
            "id": "DEC-010",
            "descripcion": "Estructura exacta de la descarga corporativa de RRHH (nombres de columnas reales).",
            "importancia": "Alta",
            "responsable_por_confirmar": "RRHH / Sistemas",
            "impacto": "La hoja CARGA_REAL usa los campos declarados. Habrá que mapear si los encabezados reales difieren.",
            "estado": "Pendiente",
        },
        {
            "id": "DEC-011",
            "descripcion": "Regla de cómputo de becarios: cuándo computan y cuándo no, fuera del caso +6 con incremento.",
            "importancia": "Media",
            "responsable_por_confirmar": "RRHH / Control de gestión",
            "impacto": "Los 6 becarios del escenario computan porque se ha confirmado incremento. Otros becarios no se han supuesto.",
            "estado": "Pendiente",
        },
        {
            "id": "DEC-012",
            "descripcion": "Umbral de calidad para bloquear el cierre mensual (nº máximo de incompletos / duplicados).",
            "importancia": "Media",
            "responsable_por_confirmar": "Gobierno del dato",
            "impacto": "Recomendación: 0 duplicados y 0 incompletos de registros que computan.",
            "estado": "Pendiente",
        },
        {
            "id": "DEC-013",
            "descripcion": "Workspace y capacidad de Power BI (Pro / PPU / Fabric) y si el informe será de importación o DirectQuery.",
            "importancia": "Media",
            "responsable_por_confirmar": "TI analítica",
            "impacto": "La especificación asume Import sobre Lists/Excel. No se ha creado un PBIX binario en este entorno.",
            "estado": "Pendiente",
        },
        {
            "id": "DEC-014",
            "descripcion": "Cuenta de servicio o identidad para los flujos de Power Automate y política de reintentos del tenant.",
            "importancia": "Media",
            "responsable_por_confirmar": "TI automatización",
            "impacto": "Los flujos están diseñados pero no pueden activarse sin entorno y conector.",
            "estado": "Pendiente",
        },
        {
            "id": "DEC-015",
            "descripcion": "Fecha de arranque de la migración Nivel A → Nivel B (tres listas).",
            "importancia": "Baja",
            "responsable_por_confirmar": "Arquitectura / Operación HC",
            "impacto": "El MVP ya lleva IDPlaza y MatriculaSustituida como puente. No hay fecha comprometida.",
            "estado": "Pendiente",
        },
    ]


def validate(records: list[dict], k: dict) -> list[str]:
    errors = []
    if k["hc_base_agosto"] != 514:
        errors.append("HC base distinto de 514")
    if k["hc_actual"] != 514:
        errors.append(f"HC actual {k['hc_actual']} != 514")
    if k["incrementos"] != 13:
        errors.append(f"Incrementos {k['incrementos']} != 13")
    if k["salidas_cev"] != 9:
        errors.append("Salidas CEV != 9")
    if k["sustituciones"] != 9:
        errors.append("Sustituciones != 9")
    if k["hc_previsto_final"] != 527:
        errors.append(f"HC previsto {k['hc_previsto_final']} != 527")
    if k["hc_estructural"] != 527:
        errors.append(f"HC estructural {k['hc_estructural']} != 527")
    if k["impacto_cev_neto"] != 0:
        errors.append("CEV neto != 0")
    for e in k["evolucion"]:
        if not e["ok"]:
            errors.append(f"Evolución {e['mes']}: {e['hc']} != {e['esperado']}")
    if not k["ecuacion_ok"]:
        errors.append("Ecuación oficial no OK")
    # identificadores texto
    for r in records:
        if not isinstance(r["Matricula"], str) or r["Matricula"] == "":
            errors.append("Matrícula vacía o no texto")
            break
        if r["TipoRegistro"] == "Sustitucion" and not r["MatriculaSustituida"]:
            errors.append(f"Sustitución sin matrícula sustituida: {r['Matricula']}")
        if r["TipoRegistro"] == "Sustitucion" and r["ImpactoHC"] != 0:
            errors.append(f"Sustitución con impacto no cero: {r['Matricula']}")
        if r["TipoRegistro"] == "Incremento" and r["ComputaHC"] == "Si" and r["ImpactoHC"] != 1:
            errors.append(f"Incremento sin impacto +1: {r['Matricula']}")
    return errors


def dump_json(path: Path, records: list[dict], k: dict) -> None:
    payload = {
        "meta": {
            "version": VERSION,
            "generado": GENERATED,
            "marca": MARCA_SINTETICO,
            "escenario": ECUACION,
        },
        "parametros": {
            "hc_base_agosto_2026": HC_BASE_AGO_2026,
            "incrementos_total": INCREMENTOS_TOTAL,
            "salidas_cev": SALIDAS_CEV,
            "sustituciones_cev": SUSTITUCIONES_CEV,
            "hc_estructural_final": HC_ESTRUCTURAL_FINAL,
            "empresas": EMPRESAS,
            "unidades": [
                {
                    "codigo": u[0],
                    "nombre": u[1],
                    "responsable": u[2],
                    "empresa": u[3],
                    "base": u[4],
                }
                for u in UNIDADES
            ],
            "tipos_registro": TIPOS_REGISTRO,
            "estados": ESTADOS,
            "programas": PROGRAMAS,
            "origen_plaza": ORIGEN_PLAZA,
        },
        "kpis": k,
        "decisiones_pendientes": decisiones_pendientes(),
        "columnas": [
            {
                "tecnico": c[0],
                "visible": c[1],
                "tipo_excel": c[2],
                "tipo_lists": c[3],
                "obligatorio": c[4],
                "default": c[5],
                "valores": c[6],
                "descripcion": c[7],
                "responsable": c[8],
            }
            for c in COLUMNAS
        ],
        "registros": records,
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    recs = build_records()
    k = kpis(recs)
    errs = validate(recs, k)
    print(json.dumps(k, ensure_ascii=False, indent=2))
    print("ERRORES", errs)
    dump_json(Path("/workspace/src/data/hc_control.json"), recs, k)
    dump_json(Path("/workspace/HC_Control/02_MVP_Operativo/hc_control.json"), recs, k)
    if errs:
        raise SystemExit(1)
    print("OK", len(recs), "registros")

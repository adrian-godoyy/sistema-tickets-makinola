"""
Funciones de conexión, guardado y lectura del historial de tickets en Supabase.
"""

from __future__ import annotations

from datetime import date, datetime
from typing import Any

import pandas as pd
import streamlit as st
from supabase import Client, create_client


NOMBRE_TABLA = "tickets"
TAMANO_LOTE = 400
TAMANO_PAGINA = 1000


def _obtener_secreto(nombre: str) -> str:
    try:
        valor = st.secrets[nombre]
    except KeyError as exc:
        raise RuntimeError(
            f"Falta configurar el secreto {nombre!r} en Streamlit Cloud."
        ) from exc

    valor = str(valor).strip()

    if not valor:
        raise RuntimeError(
            f"El secreto {nombre!r} está vacío en Streamlit Cloud."
        )

    return valor


@st.cache_resource(show_spinner=False)
def obtener_cliente_supabase() -> Client:
    url = _obtener_secreto("SUPABASE_URL")

    if "SUPABASE_SECRET_KEY" in st.secrets:
        key = _obtener_secreto("SUPABASE_SECRET_KEY")
    else:
        key = _obtener_secreto("SUPABASE_KEY")

    return create_client(url, key)


def probar_conexion() -> tuple[bool, str]:
    try:
        cliente = obtener_cliente_supabase()
        (
            cliente.table(NOMBRE_TABLA)
            .select("id_ticket")
            .limit(1)
            .execute()
        )
        return True, "Conexión con Supabase correcta."
    except Exception as error:
        return False, f"No se pudo conectar con Supabase: {error}"


def _valor_texto(valor: Any, predeterminado: str = "") -> str:
    if valor is None or pd.isna(valor):
        return predeterminado
    return str(valor).strip()


def _valor_numero(valor: Any) -> float | None:
    if valor is None or pd.isna(valor):
        return None
    try:
        return float(valor)
    except (TypeError, ValueError):
        return None


def _valor_booleano(valor: Any) -> bool:
    if valor is None or pd.isna(valor):
        return False
    if isinstance(valor, bool):
        return valor
    if isinstance(valor, (int, float)):
        return bool(valor)
    return str(valor).strip().lower() in {
        "true", "1", "sí", "si", "yes",
        "closed", "cerrado", "resolved", "resuelto",
    }


def _valor_fecha(valor: Any) -> str | None:
    if valor is None or pd.isna(valor):
        return None
    fecha = pd.to_datetime(valor, errors="coerce")
    if pd.isna(fecha):
        return None
    return fecha.isoformat()


def preparar_registros(
    df: pd.DataFrame,
    nombre_archivo: str,
) -> list[dict[str, Any]]:
    if df.empty:
        return []

    registros: list[dict[str, Any]] = []

    for _, fila in df.iterrows():
        id_ticket = fila.get("id")

        if id_ticket is None or pd.isna(id_ticket):
            continue

        id_ticket_texto = str(id_ticket).strip()

        if not id_ticket_texto:
            continue

        registros.append(
            {
                "id_ticket": id_ticket_texto,
                "fecha_apertura": _valor_fecha(fila.get("fecha_apertura")),
                "fecha_asignacion": _valor_fecha(fila.get("fecha_asignacion")),
                "fecha_cierre": _valor_fecha(fila.get("fecha_cierre")),
                "asunto": _valor_texto(fila.get("asunto"), "Sin asunto"),
                "tecnico": _valor_texto(fila.get("tecnico"), "Sin asignar"),
                "estado": _valor_texto(fila.get("estado"), "Sin estado"),
                "producto": _valor_texto(fila.get("categoria"), "Sin producto"),
                "cliente": _valor_texto(
                    fila.get("cliente"),
                    "Sin información",
                ),
                "descripcion": _valor_texto(fila.get("descripcion"), ""),
                "prioridad": _valor_texto(
                    fila.get("prioridad"),
                    "Sin definir",
                ),
                "tiempo_primera_respuesta_min": _valor_numero(
                    fila.get("tiempo_primera_respuesta_min")
                ),
                "tiempo_resolucion_horas": _valor_numero(
                    fila.get("tiempo_resolucion_horas")
                ),
                "tiempo_asignacion_cierre_min": _valor_numero(
                    fila.get("tiempo_asignacion_cierre_min")
                ),
                "cerrado": _valor_booleano(fila.get("cerrado")),
                "fecha_carga": datetime.now().isoformat(),
                "nombre_archivo": _valor_texto(
                    nombre_archivo,
                    "archivo_sin_nombre",
                ),
            }
        )

    return registros


def _dividir_lotes(
    registros: list[dict[str, Any]],
    tamano: int = TAMANO_LOTE,
):
    for inicio in range(0, len(registros), tamano):
        yield registros[inicio:inicio + tamano]


def guardar_tickets(
    df: pd.DataFrame,
    nombre_archivo: str,
) -> tuple[int, str]:
    registros = preparar_registros(
        df=df,
        nombre_archivo=nombre_archivo,
    )

    if not registros:
        return 0, "No se encontraron tickets válidos para guardar."

    try:
        cliente = obtener_cliente_supabase()
        procesados = 0

        for lote in _dividir_lotes(registros):
            (
                cliente.table(NOMBRE_TABLA)
                .upsert(
                    lote,
                    on_conflict="id_ticket",
                )
                .execute()
            )
            procesados += len(lote)

        return (
            procesados,
            (
                f"Se procesaron {procesados} tickets. "
                "Los IDs existentes fueron actualizados y no duplicados."
            ),
        )

    except Exception as error:
        return 0, f"Error al guardar los tickets: {error}"


def cargar_historial(
    fecha_inicio: date | datetime | str | None = None,
    fecha_fin: date | datetime | str | None = None,
) -> pd.DataFrame:
    try:
        cliente = obtener_cliente_supabase()
        filas: list[dict[str, Any]] = []
        inicio = 0

        while True:
            fin = inicio + TAMANO_PAGINA - 1

            consulta = (
                cliente.table(NOMBRE_TABLA)
                .select("*")
                .order("fecha_apertura")
            )

            if fecha_inicio is not None:
                consulta = consulta.gte(
                    "fecha_apertura",
                    pd.to_datetime(fecha_inicio).isoformat(),
                )

            if fecha_fin is not None:
                fecha_fin_timestamp = pd.to_datetime(fecha_fin)

                if fecha_fin_timestamp.hour == 0:
                    fecha_fin_timestamp = (
                        fecha_fin_timestamp
                        + pd.Timedelta(days=1)
                        - pd.Timedelta(microseconds=1)
                    )

                consulta = consulta.lte(
                    "fecha_apertura",
                    fecha_fin_timestamp.isoformat(),
                )

            respuesta = consulta.range(inicio, fin).execute()
            pagina = respuesta.data or []
            filas.extend(pagina)

            if len(pagina) < TAMANO_PAGINA:
                break

            inicio += TAMANO_PAGINA

        historial = pd.DataFrame(filas)

        if historial.empty:
            return historial

        historial = historial.rename(
            columns={
                "id_ticket": "id",
                "producto": "categoria",
            }
        )

        for columna in [
            "fecha_apertura",
            "fecha_asignacion",
            "fecha_cierre",
            "fecha_carga",
        ]:
            if columna in historial.columns:
                historial[columna] = pd.to_datetime(
                    historial[columna],
                    errors="coerce",
                )

        for columna in [
            "tiempo_primera_respuesta_min",
            "tiempo_resolucion_horas",
            "tiempo_asignacion_cierre_min",
        ]:
            if columna in historial.columns:
                historial[columna] = pd.to_numeric(
                    historial[columna],
                    errors="coerce",
                )

        if "cerrado" in historial.columns:
            historial["cerrado"] = (
                historial["cerrado"]
                .fillna(False)
                .astype(bool)
            )

        return historial

    except Exception as error:
        st.error(f"No se pudo leer el historial de Supabase: {error}")
        return pd.DataFrame()


def contar_tickets_guardados() -> int:
    try:
        cliente = obtener_cliente_supabase()

        respuesta = (
            cliente.table(NOMBRE_TABLA)
            .select("id_ticket", count="exact")
            .limit(1)
            .execute()
        )

        return int(respuesta.count or 0)

    except Exception:
        return 0

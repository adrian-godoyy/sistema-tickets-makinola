"""
Funciones para conectar la aplicación con Supabase,
guardar tickets y consultar el historial almacenado.
"""

from __future__ import annotations

from datetime import date, datetime
from typing import Any, Iterator

import pandas as pd
import streamlit as st
from supabase import Client, create_client


# =========================================================
# CONFIGURACIÓN
# =========================================================

NOMBRE_TABLA = "tickets"

# Cantidad de registros enviados en cada solicitud.
TAMANO_LOTE = 400

# Cantidad máxima de registros leídos en cada página.
TAMANO_PAGINA = 1000


# =========================================================
# CONEXIÓN CON SUPABASE
# =========================================================

def _obtener_secreto(nombre: str) -> str:
    """
    Obtiene un secreto configurado en Streamlit Cloud.
    """

    try:
        valor = st.secrets[nombre]
    except KeyError as error:
        raise RuntimeError(
            f"Falta configurar el secreto {nombre!r} "
            "en Streamlit Cloud."
        ) from error

    valor = str(valor).strip()

    if not valor:
        raise RuntimeError(
            f"El secreto {nombre!r} está vacío "
            "en Streamlit Cloud."
        )

    return valor


def obtener_cliente_supabase() -> Client:
    """
    Crea el cliente de Supabase para operaciones internas del servidor.
    La clave secreta solo debe estar guardada en Streamlit Secrets.
    """
    url = _obtener_secreto("SUPABASE_URL")
    secret_key = _obtener_secreto("SUPABASE_SECRET_KEY")

    return create_client(url, secret_key)


def probar_conexion() -> tuple[bool, str]:
    """
    Comprueba si la aplicación puede acceder a la tabla tickets.
    """

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
        return (
            False,
            f"No se pudo conectar con Supabase: {error}",
        )


# =========================================================
# CONVERSIÓN DE VALORES
# =========================================================

def _valor_texto(
    valor: Any,
    predeterminado: str = "",
) -> str:
    """
    Convierte un valor a texto evitando guardar nan o valores vacíos.
    """

    if valor is None:
        return predeterminado

    try:
        if pd.isna(valor):
            return predeterminado
    except (TypeError, ValueError):
        pass

    texto = str(valor).strip()

    if not texto:
        return predeterminado

    return texto


def _valor_numero(valor: Any) -> float | None:
    """
    Convierte un valor a número.
    Si no puede convertirse, devuelve None.
    """

    if valor is None:
        return None

    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass

    try:
        return float(valor)
    except (TypeError, ValueError):
        return None


def _valor_booleano(valor: Any) -> bool:
    """
    Convierte distintos formatos de estado a verdadero o falso.
    """

    if valor is None:
        return False

    try:
        if pd.isna(valor):
            return False
    except (TypeError, ValueError):
        pass

    if isinstance(valor, bool):
        return valor

    if isinstance(valor, (int, float)):
        return bool(valor)

    texto = str(valor).strip().lower()

    return texto in {
        "true",
        "1",
        "sí",
        "si",
        "yes",
        "closed",
        "cerrado",
        "resolved",
        "resuelto",
        "solved",
    }


def _valor_fecha(valor: Any) -> str | None:
    """
    Convierte una fecha a formato ISO 8601 para Supabase.
    """

    if valor is None:
        return None

    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass

    fecha = pd.to_datetime(
        valor,
        errors="coerce",
    )

    if pd.isna(fecha):
        return None

    return fecha.isoformat()


def _normalizar_id_ticket(valor: Any) -> str | None:
    """
    Convierte el ID del ticket a texto.

    Evita que un ID numérico entero quede guardado como 123.0.
    """

    if valor is None:
        return None

    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))

    texto = str(valor).strip()

    if not texto:
        return None

    return texto


# =========================================================
# PREPARACIÓN DE REGISTROS
# =========================================================

def preparar_registros(
    df: pd.DataFrame,
    nombre_archivo: str,
) -> list[dict[str, Any]]:
    """
    Convierte el DataFrame normalizado de la aplicación
    al formato de la tabla public.tickets.
    """

    if df is None or df.empty:
        return []

    registros: list[dict[str, Any]] = []

    fecha_carga = datetime.now().isoformat()

    for _, fila in df.iterrows():
        id_ticket = _normalizar_id_ticket(
            fila.get("id")
        )

        if not id_ticket:
            continue

        registro = {
            "id_ticket": id_ticket,

            "fecha_apertura": _valor_fecha(
                fila.get("fecha_apertura")
            ),

            "fecha_asignacion": _valor_fecha(
                fila.get("fecha_asignacion")
            ),

            "fecha_cierre": _valor_fecha(
                fila.get("fecha_cierre")
            ),

            "asunto": _valor_texto(
                fila.get("asunto"),
                "Sin asunto",
            ),

            "tecnico": _valor_texto(
                fila.get("tecnico"),
                "Sin asignar",
            ),

            "estado": _valor_texto(
                fila.get("estado"),
                "Sin estado",
            ),

            # En la aplicación el producto se llama categoria.
            "producto": _valor_texto(
                fila.get("categoria"),
                "Sin producto",
            ),

            "cliente": _valor_texto(
                fila.get("cliente"),
                "Sin información",
            ),

            "descripcion": _valor_texto(
                fila.get("descripcion"),
                "",
            ),

            "prioridad": _valor_texto(
                fila.get("prioridad"),
                "Sin definir",
            ),

            "tiempo_primera_respuesta_min": _valor_numero(
                fila.get(
                    "tiempo_primera_respuesta_min"
                )
            ),

            "tiempo_resolucion_horas": _valor_numero(
                fila.get(
                    "tiempo_resolucion_horas"
                )
            ),

            "tiempo_asignacion_cierre_min": _valor_numero(
                fila.get(
                    "tiempo_asignacion_cierre_min"
                )
            ),

            "cerrado": _valor_booleano(
                fila.get("cerrado")
            ),

            "fecha_carga": fecha_carga,

            "nombre_archivo": _valor_texto(
                nombre_archivo,
                "archivo_sin_nombre",
            ),
        }

        registros.append(registro)

    return registros


def _dividir_lotes(
    registros: list[dict[str, Any]],
    tamano: int = TAMANO_LOTE,
) -> Iterator[list[dict[str, Any]]]:
    """
    Divide los registros en grupos pequeños.
    """

    for inicio in range(
        0,
        len(registros),
        tamano,
    ):
        yield registros[
            inicio:inicio + tamano
        ]


# =========================================================
# GUARDAR TICKETS
# =========================================================

def guardar_tickets(
    df: pd.DataFrame,
    nombre_archivo: str,
) -> tuple[int, str]:
    """
    Guarda los tickets en Supabase.

    id_ticket es la clave principal:
    - Si el ticket no existe, se inserta.
    - Si ya existe, se actualiza.
    - No se crean duplicados.
    """

    registros = preparar_registros(
        df=df,
        nombre_archivo=nombre_archivo,
    )

    if not registros:
        return (
            0,
            "No se encontraron tickets válidos "
            "para guardar.",
        )

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

        mensaje = (
            f"Se procesaron {procesados} tickets. "
            "Los tickets existentes fueron actualizados "
            "y no se duplicaron."
        )

        return procesados, mensaje

    except Exception as error:
        return (
            0,
            f"Error al guardar los tickets: {error}",
        )


# =========================================================
# CARGAR HISTORIAL
# =========================================================

def cargar_historial(
    fecha_inicio: date | datetime | str | None = None,
    fecha_fin: date | datetime | str | None = None,
) -> pd.DataFrame:
    """
    Descarga los tickets almacenados en Supabase.

    Se puede consultar:
    - Todo el historial.
    - Un rango de fechas.
    """

    try:
        cliente = obtener_cliente_supabase()

        filas: list[dict[str, Any]] = []

        inicio = 0

        while True:
            fin = inicio + TAMANO_PAGINA - 1

            consulta = (
                cliente.table(NOMBRE_TABLA)
                .select("*")
                .order(
                    "fecha_apertura",
                    desc=False,
                )
            )

            if fecha_inicio is not None:
                fecha_inicio_iso = pd.to_datetime(
                    fecha_inicio,
                    errors="coerce",
                )

                if pd.notna(fecha_inicio_iso):
                    consulta = consulta.gte(
                        "fecha_apertura",
                        fecha_inicio_iso.isoformat(),
                    )

            if fecha_fin is not None:
                fecha_fin_timestamp = pd.to_datetime(
                    fecha_fin,
                    errors="coerce",
                )

                if pd.notna(fecha_fin_timestamp):
                    # Incluye todo el último día.
                    if (
                        fecha_fin_timestamp.hour == 0
                        and fecha_fin_timestamp.minute == 0
                        and fecha_fin_timestamp.second == 0
                    ):
                        fecha_fin_timestamp = (
                            fecha_fin_timestamp
                            + pd.Timedelta(days=1)
                            - pd.Timedelta(
                                microseconds=1
                            )
                        )

                    consulta = consulta.lte(
                        "fecha_apertura",
                        fecha_fin_timestamp.isoformat(),
                    )

            respuesta = (
                consulta.range(
                    inicio,
                    fin,
                )
                .execute()
            )

            pagina = respuesta.data or []

            filas.extend(pagina)

            if len(pagina) < TAMANO_PAGINA:
                break

            inicio += TAMANO_PAGINA

        historial = pd.DataFrame(filas)

        if historial.empty:
            return historial

        # Volver a los nombres usados por app.py.
        historial = historial.rename(
            columns={
                "id_ticket": "id",
                "producto": "categoria",
            }
        )

        columnas_fecha = [
            "fecha_apertura",
            "fecha_asignacion",
            "fecha_cierre",
            "fecha_carga",
        ]

        for columna in columnas_fecha:
            if columna in historial.columns:
                historial[columna] = pd.to_datetime(
                    historial[columna],
                    errors="coerce",
                )

        columnas_numericas = [
            "tiempo_primera_respuesta_min",
            "tiempo_resolucion_horas",
            "tiempo_asignacion_cierre_min",
        ]

        for columna in columnas_numericas:
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
        st.error(
            "No se pudo leer el historial "
            f"de Supabase: {error}"
        )

        return pd.DataFrame()


# =========================================================
# CONTAR TICKETS ALMACENADOS
# =========================================================

def contar_tickets_guardados() -> int:
    """
    Devuelve el total de tickets almacenados.
    """

    try:
        cliente = obtener_cliente_supabase()

        respuesta = (
            cliente.table(NOMBRE_TABLA)
            .select(
                "id_ticket",
                count="exact",
            )
            .limit(1)
            .execute()
        )

        return int(
            respuesta.count or 0
        )

    except Exception:
        return 0

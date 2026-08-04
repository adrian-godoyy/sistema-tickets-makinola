from pathlib import Path
from io import BytesIO
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        PageBreak,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )
    from reportlab.graphics.shapes import Drawing, String
    from reportlab.graphics.charts.barcharts import VerticalBarChart

    REPORTLAB_DISPONIBLE = True
except ImportError:
    REPORTLAB_DISPONIBLE = False


# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================

st.set_page_config(
    page_title="Sistema de Tickets",
    page_icon="🎫",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent
ARCHIVO_CSV = BASE_DIR / "datos" / "tickets.csv"

LOGO = BASE_DIR / "makinola3000.png"

CATEGORIAS_VISIBLES = [
    "Contabilidad",
    "Remuneraciones",
    "Renta",
    "Facturación",
]

ENTORNOS_VISIBLES = [
    "Nube",
    "Escritorio",
]

# David y Neil quedan totalmente excluidos.
TECNICOS_EXCLUIDOS = [
    "david",
    "david paredes castro",
    "neil",
    "neil torres",
]

ESTADOS_CERRADOS = [
    "closed",
    "cerrado",
    "resolved",
    "resuelto",
]


# =========================================================
# FUNCIONES GENERALES
# =========================================================

def limpiar_texto(valor):
    """Convierte un valor en texto limpio."""
    if pd.isna(valor):
        return ""

    return str(valor).strip()


def normalizar_nombre(nombre):
    """Normaliza un nombre para comparar sin importar mayúsculas."""
    return limpiar_texto(nombre).lower()


def tecnico_excluido(nombre):
    """
    Excluye cualquier nombre que corresponda a David o Neil.
    También cubre nombres completos y espacios adicionales.
    """
    nombre_normalizado = normalizar_nombre(nombre)

    return (
        nombre_normalizado in TECNICOS_EXCLUIDOS
        or nombre_normalizado.startswith("david ")
        or nombre_normalizado.startswith("neil ")
    )


def es_ticket_cerrado(estado):
    """Indica si un ticket está cerrado o resuelto."""
    return limpiar_texto(estado).lower() in ESTADOS_CERRADOS


def convertir_a_fecha(serie):
    """Convierte una serie a fecha y hora."""
    return pd.to_datetime(
        serie,
        errors="coerce",
    )


def formato_minutos(valor):
    """
    Muestra una cantidad de minutos de forma legible.

    Ejemplos:
    25 -> 25 min
    75 -> 1 h 15 min
    """
    if pd.isna(valor):
        return "Sin datos"

    minutos = max(0, int(round(float(valor))))

    if minutos < 60:
        return f"{minutos} min"

    horas = minutos // 60
    resto = minutos % 60

    return f"{horas} h {resto:02d} min"


def formato_horas(valor):
    """Convierte horas decimales a un texto legible."""
    if pd.isna(valor):
        return "Sin datos"

    return formato_minutos(float(valor) * 60)


def convertir_hora_a_decimal(serie):
    """
    Convierte timestamps a horas decimales.

    Ejemplo:
    10:30 = 10.5
    """
    return (
        serie.dt.hour
        + serie.dt.minute / 60
        + serie.dt.second / 3600
    )


def formato_hora_decimal(valor):
    """
    Convierte una hora decimal a formato HH:MM.

    Ejemplo:
    10.5 = 10:30
    """
    if pd.isna(valor):
        return "Sin datos"

    segundos = int(round(float(valor) * 3600))
    segundos %= 24 * 3600

    horas = segundos // 3600
    minutos = (segundos % 3600) // 60

    return f"{horas:02d}:{minutos:02d}"


# =========================================================
# CLASIFICACIÓN AUTOMÁTICA
# =========================================================

def clasificar_ticket(asunto, etiqueta_programa=None):
    """
    Clasifica únicamente en:
    Contabilidad, Remuneraciones, Renta o Facturación.

    Se prioriza E1 / Etiqueta 1 del modelo_v1.xlsx.
    No existe la categoría Otros.
    """
    texto = limpiar_texto(asunto).lower()
    etiqueta = limpiar_texto(etiqueta_programa).lower()

    combinado = f"{etiqueta} {texto}"

    # Etiquetas directas del modelo final.
    if etiqueta in {
        "eremuneraciones",
        "remuneraciones",
    }:
        return "Remuneraciones"

    if etiqueta in {
        "erenta",
        "renta",
    }:
        return "Renta"

    if etiqueta in {
        "efacturacionelectronica",
        "facturacion",
        "facturación",
        "facturacion electronica",
        "facturación electrónica",
    }:
        return "Facturación"

    if etiqueta in {
        "econtabilidad",
        "contabilidad",
        "contaplus",
        "conta plus",
    }:
        return "Contabilidad"

    palabras_remuneraciones = [
        "remuneracion",
        "remuneraciones",
        "sueldo",
        "liquidacion",
        "liquidación",
        "previred",
        "imposiciones",
        "afp",
        "fonasa",
        "isapre",
        "finiquito",
        "haberes",
        "gratificacion",
        "gratificación",
        "libro de remuneraciones",
        "lre",
        "contrato",
        "vacaciones",
        "licencia",
    ]

    palabras_renta = [
        "renta",
        "operacion renta",
        "operación renta",
        "declaracion de renta",
        "declaración de renta",
        "formulario 22",
        "f22",
        "declaracion jurada",
        "declaración jurada",
        "dj ",
        "dj1847",
        "dj 1847",
        "dj1866",
        "dj 1866",
        "dj1887",
        "dj 1887",
        "dj1945",
        "dj 1945",
        "dj1946",
        "dj 1946",
        "honorarios",
        "impuesto a la renta",
    ]

    palabras_facturacion = [
        "factura",
        "facturacion",
        "facturación",
        "factura electronica",
        "factura electrónica",
        "boleta electronica",
        "boleta electrónica",
        "nota de credito",
        "nota de crédito",
        "nota de debito",
        "nota de débito",
        "dte",
        "documento tributario",
    ]

    palabras_contabilidad = [
        "contabilidad",
        "contable",
        "contaplus",
        "conta plus",
        "balance",
        "libro diario",
        "libro mayor",
        "centralizacion",
        "centralización",
        "asiento",
        "cuenta contable",
        "plan de cuentas",
        "comprobante",
        "conciliacion",
        "conciliación",
        "activo fijo",
        "depreciacion",
        "depreciación",
        "cierre contable",
        "registro de compras",
    ]

    if any(palabra in combinado for palabra in palabras_remuneraciones):
        return "Remuneraciones"

    if any(palabra in combinado for palabra in palabras_renta):
        return "Renta"

    if any(palabra in combinado for palabra in palabras_facturacion):
        return "Facturación"

    if any(palabra in combinado for palabra in palabras_contabilidad):
        return "Contabilidad"

    # Respaldo solicitado: no crear categoría "Otros".
    return "Contabilidad"


def clasificar_entorno(valor_entorno, texto_auxiliar=""):
    """
    Clasifica únicamente en Nube o Escritorio.
    Se prioriza E2 / Etiqueta 2.
    """
    entorno = limpiar_texto(valor_entorno).lower()
    texto = limpiar_texto(texto_auxiliar).lower()
    combinado = f"{entorno} {texto}"

    if "nube" in combinado:
        return "Nube"

    if "escritorio" in combinado:
        return "Escritorio"

    # Respaldo cuando E2 viene vacío.
    return "Escritorio"


# =========================================================
# NORMALIZACIÓN DE DATOS
# =========================================================

def normalizar_datos(df):
    """
    Convierte el Excel real o un CSV antiguo a una estructura común.
    """
    df = df.copy()

    # -----------------------------------------------------
    # FORMATO DEL EXCEL REAL
    # -----------------------------------------------------
    if "ID del ticket" in df.columns:
        mapa_columnas = {
            "ID del ticket": "id",
            "Ticket creado - Marca de tiempo": "fecha_apertura",
            "Asunto del ticket": "asunto",
            "Nombre del agente asignado": "tecnico",
            "Ticket asignado - Marca de tiempo": "fecha_asignacion",
            "Primera asignación del ticket - Marca de tiempo": (
                "fecha_asignacion"
            ),
            "Ticket resuelto - Marca de tiempo": "fecha_cierre",
            "Estado del ticket": "estado",
            "Tiempo de primera respuesta (min)": (
                "tiempo_primera_respuesta_min"
            ),
            "Tiempo de resolución completa (min)": (
                "tiempo_resolucion_min"
            ),
            "Tiempo de resolución completa - Horas trabajo (h)": (
                "tiempo_resolucion_horas"
            ),
            "Etiqueta 1": "etiqueta_programa",
            "Etiqueta 2": "etiqueta_entorno",
            "E1": "etiqueta_programa",
            "E2": "etiqueta_entorno",
        }

        columnas_disponibles = {
            original: destino
            for original, destino in mapa_columnas.items()
            if original in df.columns
        }

        df = df.rename(columns=columnas_disponibles)

        if "cliente" not in df.columns:
            df["cliente"] = "Sin información"

        if "descripcion" not in df.columns:
            df["descripcion"] = df.get(
                "asunto",
                "Sin descripción",
            )

        if "prioridad" not in df.columns:
            df["prioridad"] = "Sin definir"

    # -----------------------------------------------------
    # FORMATO ANTIGUO O CSV
    # -----------------------------------------------------
    elif "fecha_apertura" not in df.columns:
        st.error("❌ No reconozco la estructura del archivo.")
        st.write("Columnas encontradas:")
        st.write(list(df.columns))
        st.stop()

    columnas_necesarias = [
        "id",
        "fecha_apertura",
        "fecha_asignacion",
        "fecha_cierre",
        "asunto",
        "tecnico",
        "estado",
    ]

    for columna in columnas_necesarias:
        if columna not in df.columns:
            df[columna] = pd.NA

    # Convertir fechas.
    for columna in [
        "fecha_apertura",
        "fecha_asignacion",
        "fecha_cierre",
    ]:
        df[columna] = convertir_a_fecha(df[columna])

    # Limpiar textos.
    df["asunto"] = (
        df["asunto"]
        .fillna("Sin asunto")
        .astype(str)
        .str.strip()
    )

    df["tecnico"] = (
        df["tecnico"]
        .fillna("Sin asignar")
        .astype(str)
        .str.strip()
    )

    df["estado"] = (
        df["estado"]
        .fillna("Sin estado")
        .astype(str)
        .str.strip()
    )

    # Clasificación definitiva del modelo_v1.xlsx.
    if "etiqueta_programa" not in df.columns:
        df["etiqueta_programa"] = pd.NA

    if "etiqueta_entorno" not in df.columns:
        df["etiqueta_entorno"] = pd.NA

    df["categoria"] = df.apply(
        lambda fila: clasificar_ticket(
            fila.get("asunto"),
            fila.get("etiqueta_programa"),
        ),
        axis=1,
    )

    df["entorno"] = df.apply(
        lambda fila: clasificar_entorno(
            fila.get("etiqueta_entorno"),
            " ".join(
                [
                    limpiar_texto(fila.get("asunto")),
                    limpiar_texto(fila.get("etiqueta_programa")),
                ]
            ),
        ),
        axis=1,
    )

    # Primera respuesta.
    if "tiempo_primera_respuesta_min" in df.columns:
        df["tiempo_primera_respuesta_min"] = pd.to_numeric(
            df["tiempo_primera_respuesta_min"],
            errors="coerce",
        )

        df["tiempo_primera_respuesta_horas"] = (
            df["tiempo_primera_respuesta_min"] / 60
        )

    elif "tiempo_primera_respuesta_horas" not in df.columns:
        df["tiempo_primera_respuesta_horas"] = pd.NA

    # Tiempo de resolución.
    if "tiempo_resolucion_min" in df.columns:
        df["tiempo_resolucion_min"] = pd.to_numeric(
            df["tiempo_resolucion_min"],
            errors="coerce",
        )

        df["tiempo_resolucion_horas"] = (
            df["tiempo_resolucion_min"] / 60
        )

    elif "tiempo_resolucion_horas" not in df.columns:
        df["tiempo_resolucion_horas"] = (
            df["fecha_cierre"]
            - df["fecha_apertura"]
        ).dt.total_seconds() / 3600

    # Cierre desde asignación del mismo ticket.
    df["tiempo_asignacion_cierre_min"] = (
        df["fecha_cierre"]
        - df["fecha_asignacion"]
    ).dt.total_seconds() / 60

    # Eliminar tiempos imposibles.
    df.loc[
        df["tiempo_asignacion_cierre_min"] < 0,
        "tiempo_asignacion_cierre_min",
    ] = pd.NA

    df["cerrado"] = df["estado"].apply(
        es_ticket_cerrado
    )

    # Excluir David y Neil antes de filtros, gráficos y PDF.
    df = df[
        ~df["tecnico"].apply(tecnico_excluido)
    ].copy()

    return df


# =========================================================
# PRODUCTIVIDAD DEL PRIMER TICKET DIARIO
# =========================================================

def calcular_primer_ticket_diario(df):
    """
    Para cada técnico y día:

    1. Considera tickets cerrados con asignación y cierre.
    2. Exige que asignación y cierre ocurran el mismo día.
    3. Ordena por hora de cierre.
    4. Elige el primer ticket cerrado.
    5. Calcula su demora usando la asignación y cierre
       DEL MISMO TICKET.

    Esto evita resultados incorrectos de 24, 44 o más horas.
    """
    datos = df[
        df["cerrado"]
        & df["fecha_asignacion"].notna()
        & df["fecha_cierre"].notna()
    ].copy()

    if datos.empty:
        return pd.DataFrame()

    # Solo tickets asignados y cerrados durante el mismo día.
    datos = datos[
        datos["fecha_asignacion"].dt.date
        == datos["fecha_cierre"].dt.date
    ].copy()

    # Solo diferencias válidas.
    datos = datos[
        datos["fecha_cierre"]
        >= datos["fecha_asignacion"]
    ].copy()

    datos["demora_primer_ticket_min"] = (
        datos["fecha_cierre"]
        - datos["fecha_asignacion"]
    ).dt.total_seconds() / 60

    # Seguridad adicional: máximo 24 horas.
    datos = datos[
        datos["demora_primer_ticket_min"].between(
            0,
            24 * 60,
            inclusive="both",
        )
    ].copy()

    if datos.empty:
        return pd.DataFrame()

    datos["fecha_trabajo"] = (
        datos["fecha_cierre"].dt.date
    )

    datos = datos.sort_values(
        by=[
            "tecnico",
            "fecha_trabajo",
            "fecha_cierre",
        ]
    )

    # Toma el primer ticket cerrado por técnico y día.
    primeros = (
        datos.groupby(
            [
                "tecnico",
                "fecha_trabajo",
            ],
            as_index=False,
        )
        .first()
    )

    primeros["hora_asignacion_decimal"] = (
        convertir_hora_a_decimal(
            primeros["fecha_asignacion"]
        )
    )

    primeros["hora_cierre_decimal"] = (
        convertir_hora_a_decimal(
            primeros["fecha_cierre"]
        )
    )

    return primeros


def calcular_ultimo_cierre_diario(df):
    """
    Obtiene el último ticket cerrado de cada técnico y día.
    """
    datos = df[
        df["cerrado"]
        & df["fecha_cierre"].notna()
    ].copy()

    if datos.empty:
        return pd.DataFrame()

    datos["fecha_trabajo"] = (
        datos["fecha_cierre"].dt.date
    )

    ultimos = (
        datos.groupby(
            [
                "tecnico",
                "fecha_trabajo",
            ],
            as_index=False,
        )
        .agg(
            ultimo_cierre=("fecha_cierre", "max"),
        )
    )

    ultimos["hora_ultimo_cierre_decimal"] = (
        convertir_hora_a_decimal(
            ultimos["ultimo_cierre"]
        )
    )

    return ultimos


def resumir_productividad(primeros_diarios, ultimos_diarios):
    """
    Resume los promedios diarios por técnico.
    """
    if primeros_diarios.empty:
        return pd.DataFrame()

    resumen = (
        primeros_diarios
        .groupby(
            "tecnico",
            as_index=False,
        )
        .agg(
            dias_analizados=(
                "fecha_trabajo",
                "nunique",
            ),
            primera_asignacion_promedio=(
                "hora_asignacion_decimal",
                "mean",
            ),
            primer_cierre_promedio=(
                "hora_cierre_decimal",
                "mean",
            ),
            demora_promedio_min=(
                "demora_primer_ticket_min",
                "mean",
            ),
        )
    )

    if not ultimos_diarios.empty:
        resumen_ultimo = (
            ultimos_diarios
            .groupby(
                "tecnico",
                as_index=False,
            )
            .agg(
                ultimo_cierre_promedio=(
                    "hora_ultimo_cierre_decimal",
                    "mean",
                )
            )
        )

        resumen = resumen.merge(
            resumen_ultimo,
            on="tecnico",
            how="left",
        )
    else:
        resumen["ultimo_cierre_promedio"] = pd.NA

    return resumen



# =========================================================
# ANALÍTICA AVANZADA
# =========================================================

def clasificar_sla_horas(valor_horas):
    """
    Clasifica el tiempo de resolución en tramos.
    """
    if pd.isna(valor_horas):
        return "Sin dato"

    valor = float(valor_horas)

    if valor < 1:
        return "Menos de 1 h"

    if valor < 4:
        return "1 a 4 h"

    if valor < 8:
        return "4 a 8 h"

    return "Más de 8 h"


def calcular_productividad_diaria_tecnico(df):
    """
    Calcula productividad diaria por técnico usando tickets cerrados.
    """
    datos = df[
        df["cerrado"]
        & df["fecha_cierre"].notna()
    ].copy()

    if datos.empty:
        return pd.DataFrame()

    datos["Fecha"] = datos["fecha_cierre"].dt.date

    diario = (
        datos.groupby(
            [
                "tecnico",
                "Fecha",
            ],
            as_index=False,
        )
        .size()
        .rename(
            columns={
                "size": "Tickets cerrados",
            }
        )
    )

    return diario


def resumir_productividad_diaria(diario):
    """
    Resume promedio, mejor día y peor día por técnico.
    """
    if diario.empty:
        return pd.DataFrame()

    resumen = (
        diario.groupby(
            "tecnico",
            as_index=False,
        )
        .agg(
            dias_activos=("Fecha", "nunique"),
            promedio_diario=("Tickets cerrados", "mean"),
            mejor_dia=("Tickets cerrados", "max"),
            peor_dia=("Tickets cerrados", "min"),
            total_cerrados=("Tickets cerrados", "sum"),
        )
    )

    return resumen


def calcular_indice_desempeno(df, resumen_productividad):
    """
    Índice simple y transparente de 1 a 7.

    Componentes:
    - 40% volumen de tickets cerrados
    - 25% rapidez de resolución
    - 20% rapidez del primer ticket
    - 15% regularidad diaria
    """
    if resumen_productividad.empty:
        return pd.DataFrame()

    base = resumen_productividad.copy()

    # Resolución promedio por técnico
    resolucion = (
        df.groupby(
            "tecnico",
            as_index=False,
        )
        .agg(
            resolucion_promedio=(
                "tiempo_resolucion_horas",
                "mean",
            )
        )
    )

    base = base.merge(
        resolucion,
        on="tecnico",
        how="left",
    )

    primeros = calcular_primer_ticket_diario(df)

    if not primeros.empty:
        demora = (
            primeros.groupby(
                "tecnico",
                as_index=False,
            )
            .agg(
                demora_primer_ticket=(
                    "demora_primer_ticket_min",
                    "mean",
                )
            )
        )

        base = base.merge(
            demora,
            on="tecnico",
            how="left",
        )
    else:
        base["demora_primer_ticket"] = pd.NA

    def escalar_mayor_mejor(serie):
        serie = pd.to_numeric(
            serie,
            errors="coerce",
        )

        if serie.notna().sum() == 0:
            return pd.Series(
                [0.5] * len(serie),
                index=serie.index,
            )

        minimo = serie.min()
        maximo = serie.max()

        if maximo == minimo:
            return pd.Series(
                [1.0] * len(serie),
                index=serie.index,
            )

        return (
            (serie - minimo)
            / (maximo - minimo)
        )

    def escalar_menor_mejor(serie):
        return 1 - escalar_mayor_mejor(serie)

    base["score_volumen"] = escalar_mayor_mejor(
        base["total_cerrados"]
    )

    base["score_rapidez"] = escalar_menor_mejor(
        base["resolucion_promedio"]
    )

    base["score_primer_ticket"] = escalar_menor_mejor(
        base["demora_primer_ticket"]
    )

    # Regularidad: menos dispersión entre mejor y peor día es mejor.
    base["rango_diario"] = (
        base["mejor_dia"]
        - base["peor_dia"]
    )

    base["score_regularidad"] = escalar_menor_mejor(
        base["rango_diario"]
    )

    base["indice_0_1"] = (
        base["score_volumen"] * 0.40
        + base["score_rapidez"] * 0.25
        + base["score_primer_ticket"] * 0.20
        + base["score_regularidad"] * 0.15
    )

    # Convertir de 0..1 a escala 1..7.
    base["Nota desempeño"] = (
        1 + base["indice_0_1"] * 6
    ).round(2)

    return base


# =========================================================
# GRÁFICOS PLOTLY
# =========================================================

def grafico_tickets_tecnico(datos):
    """
    Gráfico de tickets por técnico.
    El menor valor se destaca en rojo.
    """
    if datos.empty:
        return None

    datos = datos.copy()
    menor = datos["Cantidad"].min()

    datos["Grupo"] = datos["Cantidad"].apply(
        lambda valor: (
            "Menor cantidad"
            if valor == menor
            else "Resto"
        )
    )

    figura = px.bar(
        datos,
        x="Técnico",
        y="Cantidad",
        color="Grupo",
        text="Cantidad",
        title="Tickets por técnico",
        color_discrete_map={
            "Resto": "#4C78A8",
            "Menor cantidad": "#D62728",
        },
    )

    figura.update_traces(
        width=0.40,
        textposition="outside",
    )

    figura.update_layout(
        showlegend=False,
        bargap=0.45,
        xaxis_title="Técnico",
        yaxis_title="Cantidad",
    )

    return figura


def grafico_categorias(datos):
    """
    Muestra los cuatro programas divididos por entorno:
    Nube y Escritorio.
    """
    if datos.empty:
        return None

    figura = px.bar(
        datos,
        x="Categoría",
        y="Cantidad",
        color="Entorno",
        barmode="group",
        text="Cantidad",
        title="Tickets por categoría y entorno",
        category_orders={
            "Categoría": CATEGORIAS_VISIBLES,
            "Entorno": ENTORNOS_VISIBLES,
        },
    )

    figura.update_traces(
        width=0.32,
        textposition="outside",
    )

    figura.update_layout(
        bargap=0.35,
        bargroupgap=0.12,
        xaxis_title="Programa",
        yaxis_title="Cantidad",
        legend_title="Entorno",
    )

    return figura


def grafico_estados(datos):
    """Gráfico de tickets por estado."""
    if datos.empty:
        return None

    figura = px.bar(
        datos,
        x="Estado",
        y="Cantidad",
        text="Cantidad",
        title="Tickets por estado",
        color_discrete_sequence=["#4C78A8"],
    )

    figura.update_traces(
        width=0.40,
        textposition="outside",
    )

    figura.update_layout(
        showlegend=False,
        bargap=0.45,
        xaxis_title="Estado",
        yaxis_title="Cantidad",
    )

    return figura


def grafico_demora_primer_ticket(resumen):
    """
    Gráfico de demora promedio del primer ticket diario.

    El técnico que MÁS demora queda en rojo.
    Las barras son delgadas.
    """
    if resumen.empty:
        return None

    datos = resumen[
        [
            "tecnico",
            "demora_promedio_min",
        ]
    ].dropna().copy()

    if datos.empty:
        return None

    datos = datos.rename(
        columns={
            "tecnico": "Técnico",
            "demora_promedio_min": "Minutos",
        }
    )

    mayor_demora = datos["Minutos"].max()

    datos["Grupo"] = datos["Minutos"].apply(
        lambda valor: (
            "Mayor demora"
            if valor == mayor_demora
            else "Resto"
        )
    )

    figura = px.bar(
        datos,
        x="Técnico",
        y="Minutos",
        color="Grupo",
        text="Minutos",
        title=(
            "Demora promedio diaria hasta cerrar "
            "el primer ticket"
        ),
        color_discrete_map={
            "Resto": "#4C78A8",
            "Mayor demora": "#D62728",
        },
    )

    figura.update_traces(
        width=0.40,
        texttemplate="%{text:.1f} min",
        textposition="outside",
    )

    figura.update_layout(
        showlegend=False,
        bargap=0.45,
        xaxis_title="Técnico",
        yaxis_title="Minutos",
    )

    return figura


# =========================================================
# GRÁFICOS PARA EL PDF
# =========================================================

def crear_grafico_barras_pdf(
    titulo,
    etiquetas,
    valores,
    destacar="ninguno",
    nota=None,
):
    """
    Crea un gráfico de barras para el PDF.

    destacar:
    - "menor": marca en rojo el valor menor.
    - "mayor": marca en rojo el valor mayor.
    - "ninguno": todas las barras quedan azules.
    """
    ancho = 720
    alto = 270

    dibujo = Drawing(ancho, alto)

    dibujo.add(
        String(
            ancho / 2,
            alto - 18,
            titulo,
            textAnchor="middle",
            fontName="Helvetica-Bold",
            fontSize=13,
        )
    )

    etiquetas = [str(v)[:22] for v in etiquetas]
    valores = [
        float(v) if pd.notna(v) else 0.0
        for v in valores
    ]

    if not valores:
        dibujo.add(
            String(
                ancho / 2,
                alto / 2,
                "Sin datos para mostrar",
                textAnchor="middle",
                fontSize=10,
            )
        )
        return dibujo

    grafico = VerticalBarChart()
    grafico.x = 58
    grafico.y = 68
    grafico.width = ancho - 100
    grafico.height = alto - 120
    grafico.data = [valores]
    grafico.categoryAxis.categoryNames = etiquetas

    setattr(grafico.categoryAxis.labels, "angle", 28)
    setattr(grafico.categoryAxis.labels, "fontSize", 7)
    setattr(grafico.categoryAxis.labels, "dy", -12)

    grafico.valueAxis.valueMin = 0

    maximo = max(valores)
    grafico.valueAxis.valueMax = max(maximo * 1.20, 1)

    paso = grafico.valueAxis.valueMax / 5
    grafico.valueAxis.valueStep = max(round(paso, 1), 1)

    setattr(grafico.valueAxis.labels, "fontSize", 8)

    # Barras delgadas
    setattr(grafico, "barWidth", 12)
    setattr(grafico, "groupSpacing", 22)

    azul = colors.HexColor("#4C78A8")
    rojo = colors.HexColor("#D62728")

    objetivo = None

    if destacar == "menor":
        objetivo = min(valores)
    elif destacar == "mayor":
        objetivo = max(valores)

    for indice, valor in enumerate(valores):
        color_barra = (
            rojo
            if objetivo is not None and valor == objetivo
            else azul
        )

        try:
            grafico.bars[(0, indice)].fillColor = color_barra
            grafico.bars[(0, indice)].strokeColor = color_barra
        except Exception:
            pass

    dibujo.add(grafico)

    if nota:
        dibujo.add(
            String(
                ancho / 2,
                12,
                nota,
                textAnchor="middle",
                fontName="Helvetica-Oblique",
                fontSize=8,
                fillColor=colors.HexColor("#555555"),
            )
        )

    return dibujo


# =========================================================
# PDF
# =========================================================

def generar_pdf_informe(
    df,
    total,
    cerrados,
    pendientes,
    promedio_texto,
    tecnico_filtro,
    estado_filtro,
    categoria_filtro,
    rango_fechas,
    tabla_productividad,
):
    """Genera un PDF con resumen, productividad y detalle."""
    if not REPORTLAB_DISPONIBLE:
        return None

    buffer = BytesIO()

    documento = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=1.2 * cm,
        leftMargin=1.2 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.2 * cm,
        title="Informe de Tickets de Soporte",
    )

    estilos = getSampleStyleSheet()
    elementos = []

    elementos.append(
        Paragraph(
            "Informe de Análisis de Tickets de Soporte",
            estilos["Title"],
        )
    )

    elementos.append(
        Paragraph(
            "Generado: "
            + datetime.now().strftime("%d-%m-%Y %H:%M"),
            estilos["BodyText"],
        )
    )

    elementos.append(Spacer(1, 0.4 * cm))

    filtros = [
        f"Técnico: {tecnico_filtro}",
        f"Estado: {estado_filtro}",
        f"Categoría: {categoria_filtro}",
    ]

    if (
        isinstance(rango_fechas, (list, tuple))
        and len(rango_fechas) == 2
    ):
        filtros.append(
            f"Período: {rango_fechas[0]} al {rango_fechas[1]}"
        )

    elementos.append(
        Paragraph(
            "<b>Filtros aplicados</b>",
            estilos["Heading2"],
        )
    )

    elementos.append(
        Paragraph(
            " | ".join(filtros),
            estilos["BodyText"],
        )
    )

    elementos.append(Spacer(1, 0.4 * cm))

    mascara_mismo_dia_pdf = (
        df["fecha_apertura"].notna()
        & df["fecha_cierre"].notna()
        & (
            df["fecha_apertura"].dt.date
            == df["fecha_cierre"].dt.date
        )
    )

    resueltos_mismo_dia_pdf = int(
        mascara_mismo_dia_pdf.sum()
    )

    dias_analizados_pdf = (
        df.loc[
            mascara_mismo_dia_pdf,
            "fecha_cierre",
        ]
        .dropna()
        .dt.date
        .nunique()
    )

    tecnicos_pdf = int(
        df["tecnico"].nunique()
    )

    promedio_resueltos_dia_pdf = (
        resueltos_mismo_dia_pdf / dias_analizados_pdf
        if dias_analizados_pdf > 0
        else 0
    )

    promedio_diario_tecnico_pdf = (
        resueltos_mismo_dia_pdf
        / dias_analizados_pdf
        / tecnicos_pdf
        if dias_analizados_pdf > 0 and tecnicos_pdf > 0
        else 0
    )

    resumen_general = [
        ["Indicador", "Resultado"],
        ["Total de tickets", str(total)],
        ["Cerrados", str(cerrados)],
        [
            "Promedio resueltos por día",
            f"{promedio_resueltos_dia_pdf:.1f}",
        ],
        [
            "Técnicos considerados",
            str(tecnicos_pdf),
        ],
        [
            "Promedio diario por técnico",
            f"{promedio_diario_tecnico_pdf:.1f}",
        ],
        [
            "Días analizados",
            str(dias_analizados_pdf),
        ],
    ]

    tabla_resumen = Table(
        resumen_general,
        colWidths=[8 * cm, 6 * cm],
    )

    tabla_resumen.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.lightgrey,
                ),
                (
                    "FONTNAME",
                    (0, 0),
                    (-1, 0),
                    "Helvetica-Bold",
                ),
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.5,
                    colors.grey,
                ),
            ]
        )
    )

    elementos.append(tabla_resumen)
    elementos.append(Spacer(1, 0.5 * cm))

    # -----------------------------------------------------
    # VISUALIZACIONES PRINCIPALES DEL PDF
    # -----------------------------------------------------
    elementos.append(
        Paragraph(
            "Visualizaciones principales",
            estilos["Heading2"],
        )
    )

    # Tickets por técnico
    por_tecnico_pdf = (
        df["tecnico"]
        .value_counts()
        .rename_axis("Técnico")
        .reset_index(name="Cantidad")
    )

    elementos.append(
        crear_grafico_barras_pdf(
            titulo="Tickets por técnico",
            etiquetas=por_tecnico_pdf["Técnico"].tolist(),
            valores=por_tecnico_pdf["Cantidad"].tolist(),
            destacar="menor",
            nota=(
                "La barra roja identifica al técnico "
                "con menor cantidad de tickets."
            ),
        )
    )

    # Tickets por categoría y entorno
    for entorno_pdf in ENTORNOS_VISIBLES:
        datos_entorno_pdf = (
            df[
                (df["entorno"] == entorno_pdf)
                & df["categoria"].isin(CATEGORIAS_VISIBLES)
            ]["categoria"]
            .value_counts()
            .reindex(
                CATEGORIAS_VISIBLES,
                fill_value=0,
            )
            .rename_axis("Categoría")
            .reset_index(name="Cantidad")
        )

        elementos.append(
            crear_grafico_barras_pdf(
                titulo=f"Tickets por programa - {entorno_pdf}",
                etiquetas=datos_entorno_pdf["Categoría"].tolist(),
                valores=datos_entorno_pdf["Cantidad"].tolist(),
                destacar="ninguno",
                nota=(
                    "Programas: Contabilidad, Remuneraciones, "
                    "Renta y Facturación."
                ),
            )
        )

    # Tickets por estado
    por_estado_pdf = (
        df["estado"]
        .value_counts()
        .rename_axis("Estado")
        .reset_index(name="Cantidad")
    )

    elementos.append(
        crear_grafico_barras_pdf(
            titulo="Tickets por estado",
            etiquetas=por_estado_pdf["Estado"].tolist(),
            valores=por_estado_pdf["Cantidad"].tolist(),
            destacar="ninguno",
            nota="Distribución de tickets por estado.",
        )
    )

    # Demora promedio para cerrar primer ticket
    if (
        not tabla_productividad.empty
        and "Demora promedio (min)" in tabla_productividad.columns
    ):
        demora_pdf = tabla_productividad[
            [
                "Técnico",
                "Demora promedio (min)",
            ]
        ].dropna()

        elementos.append(
            crear_grafico_barras_pdf(
                titulo=(
                    "Demora promedio para cerrar "
                    "el primer ticket diario"
                ),
                etiquetas=demora_pdf["Técnico"].tolist(),
                valores=demora_pdf[
                    "Demora promedio (min)"
                ].tolist(),
                destacar="mayor",
                nota=(
                    "La barra roja identifica al técnico "
                    "con mayor demora promedio."
                ),
            )
        )

    elementos.append(PageBreak())

    elementos.append(
        Paragraph(
            "Productividad del primer ticket diario",
            estilos["Heading2"],
        )
    )

    if tabla_productividad.empty:
        elementos.append(
            Paragraph(
                "No hay datos suficientes.",
                estilos["BodyText"],
            )
        )
    else:
        columnas_pdf = [
            "Técnico",
            "Días analizados",
            "Asignación promedio",
            "Primer cierre promedio",
            "Último cierre promedio",
            "Demora promedio",
        ]

        datos_pdf = [columnas_pdf]

        for _, fila in tabla_productividad.iterrows():
            datos_pdf.append(
                [
                    str(fila["Técnico"]),
                    str(fila["Días analizados"]),
                    str(fila["Asignación promedio"]),
                    str(fila["Primer cierre promedio"]),
                    str(fila["Último cierre promedio"]),
                    str(fila["Demora promedio"]),
                ]
            )

        tabla_pdf = Table(
            datos_pdf,
            colWidths=[
                4.5 * cm,
                2.2 * cm,
                3.5 * cm,
                3.5 * cm,
                3.5 * cm,
                3.5 * cm,
            ],
            repeatRows=1,
        )

        estilos_tabla = [
            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.lightgrey,
            ),
            (
                "FONTNAME",
                (0, 0),
                (-1, 0),
                "Helvetica-Bold",
            ),
            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.4,
                colors.grey,
            ),
        ]

        # Marcar en rojo la fila del técnico que más demora.
        if "Demora promedio (min)" in tabla_productividad.columns:
            demora_numerica = pd.to_numeric(
                tabla_productividad[
                    "Demora promedio (min)"
                ],
                errors="coerce",
            )

            if demora_numerica.notna().any():
                posicion = demora_numerica.argmax() + 1

                estilos_tabla.append(
                    (
                        "BACKGROUND",
                        (0, posicion),
                        (-1, posicion),
                        colors.HexColor("#D62728"),
                    )
                )

        tabla_pdf.setStyle(
            TableStyle(estilos_tabla)
        )

        elementos.append(tabla_pdf)

        elementos.append(
            Paragraph(
                "La fila roja identifica al técnico con mayor demora "
                "promedio para cerrar su primer ticket diario.",
                estilos["BodyText"],
            )
        )

    # -----------------------------------------------------
    # DISTRIBUCIÓN DE PROGRAMAS POR TÉCNICO
    # -----------------------------------------------------
    elementos.append(
        Spacer(1, 0.4 * cm)
    )

    elementos.append(
        Paragraph(
            "Tickets atendidos por técnico y categoría",
            estilos["Heading2"],
        )
    )

    distribucion_pdf = (
        df[
            df["categoria"].isin(
                CATEGORIAS_VISIBLES
            )
        ]
        .groupby(
            [
                "tecnico",
                "categoria",
            ]
        )
        .size()
        .reset_index(
            name="Cantidad"
        )
    )

    if not distribucion_pdf.empty:
        tabla_distribucion_pdf = (
            distribucion_pdf
            .pivot_table(
                index="tecnico",
                columns="categoria",
                values="Cantidad",
                aggfunc="sum",
                fill_value=0,
            )
            .reindex(
                columns=CATEGORIAS_VISIBLES,
                fill_value=0,
            )
            .reset_index()
        )

        tabla_distribucion_pdf["Total"] = (
            tabla_distribucion_pdf[
                CATEGORIAS_VISIBLES
            ].sum(axis=1)
        )

        tabla_distribucion_pdf[
            "Especialidad principal"
        ] = tabla_distribucion_pdf[
            CATEGORIAS_VISIBLES
        ].idxmax(axis=1)

        columnas_programas_pdf = [
            "Técnico",
            "Contabilidad",
            "Remuneraciones",
            "Renta",
            "Facturación",
            "Total",
            "Especialidad principal",
        ]

        tabla_distribucion_pdf = (
            tabla_distribucion_pdf
            .rename(
                columns={
                    "tecnico": "Técnico",
                }
            )
        )

        datos_programas_pdf = [
            columnas_programas_pdf
        ]

        for _, fila in tabla_distribucion_pdf.iterrows():
            datos_programas_pdf.append(
                [
                    str(fila[columna])
                    for columna in columnas_programas_pdf
                ]
            )

        tabla_programas_pdf = Table(
            datos_programas_pdf,
            colWidths=[
                4.0 * cm,
                2.6 * cm,
                2.9 * cm,
                2.2 * cm,
                2.7 * cm,
                2.0 * cm,
                4.0 * cm,
            ],
            repeatRows=1,
        )

        tabla_programas_pdf.setStyle(
            TableStyle(
                [
                    (
                        "BACKGROUND",
                        (0, 0),
                        (-1, 0),
                        colors.lightgrey,
                    ),
                    (
                        "FONTNAME",
                        (0, 0),
                        (-1, 0),
                        "Helvetica-Bold",
                    ),
                    (
                        "GRID",
                        (0, 0),
                        (-1, -1),
                        0.4,
                        colors.grey,
                    ),
                    (
                        "ALIGN",
                        (1, 1),
                        (-2, -1),
                        "CENTER",
                    ),
                ]
            )
        )

        elementos.append(
            tabla_programas_pdf
        )

        elementos.append(
            Paragraph(
                "La columna Especialidad principal corresponde "
                "al programa en el que cada técnico atendió "
                "la mayor cantidad de tickets.",
                estilos["BodyText"],
            )
        )

    # -----------------------------------------------------
    # ANALÍTICA AVANZADA EN PDF
    # -----------------------------------------------------

    elementos.append(
        Paragraph(
            "Analítica avanzada",
            estilos["Heading2"],
        )
    )

    # =====================================================
    # SLA / TIEMPOS DE RESOLUCIÓN
    # =====================================================
    sla_pdf = df.copy()

    sla_pdf["Tramo SLA"] = (
        pd.to_numeric(
            sla_pdf["tiempo_resolucion_horas"],
            errors="coerce",
        )
        .apply(
            clasificar_sla_horas
        )
    )

    orden_sla_pdf = [
        "Menos de 1 h",
        "1 a 4 h",
        "4 a 8 h",
        "Más de 8 h",
    ]

    resumen_sla_pdf = (
        sla_pdf[
            sla_pdf["Tramo SLA"].isin(
                orden_sla_pdf
            )
        ]["Tramo SLA"]
        .value_counts()
        .reindex(
            orden_sla_pdf,
            fill_value=0,
        )
        .rename_axis("Tramo")
        .reset_index(
            name="Cantidad"
        )
    )

    total_sla_pdf = int(
        resumen_sla_pdf["Cantidad"].sum()
    )

    if total_sla_pdf > 0:
        resumen_sla_pdf["Porcentaje"] = (
            resumen_sla_pdf["Cantidad"]
            / total_sla_pdf
            * 100
        ).round(1)
    else:
        resumen_sla_pdf["Porcentaje"] = 0.0

    elementos.append(
        Paragraph(
            "Distribución de tiempos de resolución",
            estilos["Heading3"],
        )
    )

    datos_sla_pdf = [
        [
            "Tramo",
            "Cantidad",
            "Porcentaje",
        ]
    ]

    for _, fila in resumen_sla_pdf.iterrows():
        datos_sla_pdf.append(
            [
                str(fila["Tramo"]),
                str(int(fila["Cantidad"])),
                f"{float(fila['Porcentaje']):.1f}%",
            ]
        )

    tabla_sla_pdf = Table(
        datos_sla_pdf,
        colWidths=[
            6 * cm,
            3 * cm,
            3 * cm,
        ],
        repeatRows=1,
    )

    tabla_sla_pdf.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.lightgrey,
                ),
                (
                    "FONTNAME",
                    (0, 0),
                    (-1, 0),
                    "Helvetica-Bold",
                ),
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.4,
                    colors.grey,
                ),
            ]
        )
    )

    elementos.append(
        tabla_sla_pdf
    )

    # =====================================================
    # PRODUCTIVIDAD DIARIA REAL
    # =====================================================
    productividad_diaria_pdf = (
        calcular_productividad_diaria_tecnico(
            df
        )
    )

    resumen_diario_pdf = (
        resumir_productividad_diaria(
            productividad_diaria_pdf
        )
    )

    if not resumen_diario_pdf.empty:
        elementos.append(
            Spacer(
                1,
                0.35 * cm,
            )
        )

        elementos.append(
            Paragraph(
                "Productividad diaria por técnico",
                estilos["Heading3"],
            )
        )

        datos_diarios_pdf = [
            [
                "Técnico",
                "Días activos",
                "Promedio diario",
                "Mejor día",
                "Peor día",
                "Total cerrados",
            ]
        ]

        for _, fila in resumen_diario_pdf.iterrows():
            datos_diarios_pdf.append(
                [
                    str(fila["tecnico"]),
                    str(int(fila["dias_activos"])),
                    f"{float(fila['promedio_diario']):.2f}",
                    str(int(fila["mejor_dia"])),
                    str(int(fila["peor_dia"])),
                    str(int(fila["total_cerrados"])),
                ]
            )

        tabla_diaria_pdf = Table(
            datos_diarios_pdf,
            colWidths=[
                4.3 * cm,
                2.3 * cm,
                2.8 * cm,
                2.2 * cm,
                2.2 * cm,
                2.8 * cm,
            ],
            repeatRows=1,
        )

        tabla_diaria_pdf.setStyle(
            TableStyle(
                [
                    (
                        "BACKGROUND",
                        (0, 0),
                        (-1, 0),
                        colors.lightgrey,
                    ),
                    (
                        "FONTNAME",
                        (0, 0),
                        (-1, 0),
                        "Helvetica-Bold",
                    ),
                    (
                        "GRID",
                        (0, 0),
                        (-1, -1),
                        0.4,
                        colors.grey,
                    ),
                ]
            )
        )

        elementos.append(
            tabla_diaria_pdf
        )

    # =====================================================
    # MATRIZ TÉCNICO × PROGRAMA
    # =====================================================
    distribucion_programa_pdf = (
        df[
            df["categoria"].isin(
                CATEGORIAS_VISIBLES
            )
        ]
        .groupby(
            [
                "tecnico",
                "categoria",
            ]
        )
        .size()
        .reset_index(
            name="Cantidad"
        )
    )

    if not distribucion_programa_pdf.empty:
        matriz_programa_pdf = (
            distribucion_programa_pdf
            .pivot_table(
                index="tecnico",
                columns="categoria",
                values="Cantidad",
                aggfunc="sum",
                fill_value=0,
            )
            .reindex(
                columns=CATEGORIAS_VISIBLES,
                fill_value=0,
            )
            .reset_index()
        )

        matriz_programa_pdf["Total"] = (
            matriz_programa_pdf[
                CATEGORIAS_VISIBLES
            ].sum(
                axis=1
            )
        )

        matriz_programa_pdf[
            "Especialidad principal"
        ] = matriz_programa_pdf[
            CATEGORIAS_VISIBLES
        ].idxmax(
            axis=1
        )

        elementos.append(
            Spacer(
                1,
                0.35 * cm,
            )
        )

        elementos.append(
            Paragraph(
                "Distribución técnico por programa",
                estilos["Heading3"],
            )
        )

        columnas_programa_pdf = [
            "Técnico",
            "Contabilidad",
            "Remuneraciones",
            "Renta",
            "Facturación",
            "Total",
            "Especialidad principal",
        ]

        matriz_programa_pdf = (
            matriz_programa_pdf
            .rename(
                columns={
                    "tecnico": "Técnico",
                }
            )
        )

        datos_programa_pdf = [
            columnas_programa_pdf
        ]

        for _, fila in matriz_programa_pdf.iterrows():
            datos_programa_pdf.append(
                [
                    str(fila[columna])
                    for columna in columnas_programa_pdf
                ]
            )

        tabla_programa_pdf = Table(
            datos_programa_pdf,
            colWidths=[
                4.0 * cm,
                2.6 * cm,
                2.9 * cm,
                2.2 * cm,
                2.7 * cm,
                2.0 * cm,
                4.0 * cm,
            ],
            repeatRows=1,
        )

        tabla_programa_pdf.setStyle(
            TableStyle(
                [
                    (
                        "BACKGROUND",
                        (0, 0),
                        (-1, 0),
                        colors.lightgrey,
                    ),
                    (
                        "FONTNAME",
                        (0, 0),
                        (-1, 0),
                        "Helvetica-Bold",
                    ),
                    (
                        "GRID",
                        (0, 0),
                        (-1, -1),
                        0.4,
                        colors.grey,
                    ),
                ]
            )
        )

        elementos.append(
            tabla_programa_pdf
        )

    # =====================================================
    # NUBE VS ESCRITORIO POR TÉCNICO
    # =====================================================
    entorno_pdf = (
        df[
            df["entorno"].isin(
                ENTORNOS_VISIBLES
            )
        ]
        .groupby(
            [
                "tecnico",
                "entorno",
            ]
        )
        .size()
        .reset_index(
            name="Cantidad"
        )
    )

    if not entorno_pdf.empty:
        tabla_entorno_pdf = (
            entorno_pdf
            .pivot_table(
                index="tecnico",
                columns="entorno",
                values="Cantidad",
                aggfunc="sum",
                fill_value=0,
            )
            .reindex(
                columns=ENTORNOS_VISIBLES,
                fill_value=0,
            )
            .reset_index()
        )

        tabla_entorno_pdf["Total"] = (
            tabla_entorno_pdf[
                ENTORNOS_VISIBLES
            ].sum(
                axis=1
            )
        )

        for entorno_nombre in ENTORNOS_VISIBLES:
            tabla_entorno_pdf[
                f"% {entorno_nombre}"
            ] = (
                tabla_entorno_pdf[
                    entorno_nombre
                ]
                / tabla_entorno_pdf[
                    "Total"
                ].replace(
                    0,
                    pd.NA,
                )
                * 100
            ).round(
                1
            )

        elementos.append(
            Spacer(
                1,
                0.35 * cm,
            )
        )

        elementos.append(
            Paragraph(
                "Nube vs Escritorio por técnico",
                estilos["Heading3"],
            )
        )

        columnas_entorno_pdf = [
            "Técnico",
            "Nube",
            "% Nube",
            "Escritorio",
            "% Escritorio",
            "Total",
        ]

        tabla_entorno_pdf = (
            tabla_entorno_pdf
            .rename(
                columns={
                    "tecnico": "Técnico",
                }
            )
        )

        datos_entorno_pdf = [
            columnas_entorno_pdf
        ]

        for _, fila in tabla_entorno_pdf.iterrows():
            datos_entorno_pdf.append(
                [
                    str(fila["Técnico"]),
                    str(int(fila["Nube"])),
                    (
                        ""
                        if pd.isna(
                            fila["% Nube"]
                        )
                        else f"{float(fila['% Nube']):.1f}%"
                    ),
                    str(int(fila["Escritorio"])),
                    (
                        ""
                        if pd.isna(
                            fila["% Escritorio"]
                        )
                        else f"{float(fila['% Escritorio']):.1f}%"
                    ),
                    str(int(fila["Total"])),
                ]
            )

        tabla_entorno_pdf_obj = Table(
            datos_entorno_pdf,
            colWidths=[
                4.5 * cm,
                2.0 * cm,
                2.2 * cm,
                2.5 * cm,
                2.5 * cm,
                2.0 * cm,
            ],
            repeatRows=1,
        )

        tabla_entorno_pdf_obj.setStyle(
            TableStyle(
                [
                    (
                        "BACKGROUND",
                        (0, 0),
                        (-1, 0),
                        colors.lightgrey,
                    ),
                    (
                        "FONTNAME",
                        (0, 0),
                        (-1, 0),
                        "Helvetica-Bold",
                    ),
                    (
                        "GRID",
                        (0, 0),
                        (-1, -1),
                        0.4,
                        colors.grey,
                    ),
                ]
            )
        )

        elementos.append(
            tabla_entorno_pdf_obj
        )

    # =====================================================
    # TIEMPO CALENDARIO HASTA EL CIERRE
    # =====================================================
    cierres_pdf = df[
        df["fecha_apertura"].notna()
        & df["fecha_cierre"].notna()
    ].copy()

    if not cierres_pdf.empty:
        cierres_pdf["dias_hasta_cierre"] = (
            cierres_pdf["fecha_cierre"].dt.normalize()
            - cierres_pdf["fecha_apertura"].dt.normalize()
        ).dt.days

        def tramo_dias_pdf(valor):
            if valor <= 0:
                return "Mismo día"
            if valor == 1:
                return "1 día después"
            if valor <= 3:
                return "2 a 3 días"
            return "Más de 3 días"

        cierres_pdf[
            "Tramo cierre"
        ] = cierres_pdf[
            "dias_hasta_cierre"
        ].apply(
            tramo_dias_pdf
        )

        orden_cierre_pdf = [
            "Mismo día",
            "1 día después",
            "2 a 3 días",
            "Más de 3 días",
        ]

        resumen_cierre_pdf = (
            cierres_pdf[
                "Tramo cierre"
            ]
            .value_counts()
            .reindex(
                orden_cierre_pdf,
                fill_value=0,
            )
            .rename_axis(
                "Tramo"
            )
            .reset_index(
                name="Cantidad"
            )
        )

        elementos.append(
            Spacer(
                1,
                0.35 * cm,
            )
        )

        elementos.append(
            Paragraph(
                "Tiempo calendario hasta el cierre",
                estilos["Heading3"],
            )
        )

        datos_cierre_pdf = [
            [
                "Tramo",
                "Cantidad",
            ]
        ]

        for _, fila in resumen_cierre_pdf.iterrows():
            datos_cierre_pdf.append(
                [
                    str(fila["Tramo"]),
                    str(
                        int(
                            fila["Cantidad"]
                        )
                    ),
                ]
            )

        tabla_cierre_pdf = Table(
            datos_cierre_pdf,
            colWidths=[
                7 * cm,
                3 * cm,
            ],
            repeatRows=1,
        )

        tabla_cierre_pdf.setStyle(
            TableStyle(
                [
                    (
                        "BACKGROUND",
                        (0, 0),
                        (-1, 0),
                        colors.lightgrey,
                    ),
                    (
                        "FONTNAME",
                        (0, 0),
                        (-1, 0),
                        "Helvetica-Bold",
                    ),
                    (
                        "GRID",
                        (0, 0),
                        (-1, -1),
                        0.4,
                        colors.grey,
                    ),
                ]
            )
        )

        elementos.append(
            tabla_cierre_pdf
        )

    # =====================================================
    # EVOLUCIÓN SEMANAL
    # =====================================================
    semanal_pdf = df[
        df["cerrado"]
        & df["fecha_cierre"].notna()
    ].copy()

    if not semanal_pdf.empty:
        semanal_pdf[
            "Semana"
        ] = (
            semanal_pdf[
                "fecha_cierre"
            ]
            .dt.to_period(
                "W"
            )
            .astype(
                str
            )
        )

        semanal_resumen_pdf = (
            semanal_pdf
            .groupby(
                [
                    "Semana",
                    "tecnico",
                ],
                as_index=False,
            )
            .size()
            .rename(
                columns={
                    "size": "Tickets cerrados",
                }
            )
        )

        pivot_semanal_pdf = (
            semanal_resumen_pdf
            .pivot_table(
                index="Semana",
                columns="tecnico",
                values="Tickets cerrados",
                aggfunc="sum",
                fill_value=0,
            )
            .reset_index()
        )

        elementos.append(
            Spacer(
                1,
                0.35 * cm,
            )
        )

        elementos.append(
            Paragraph(
                "Evolución semanal por técnico",
                estilos["Heading3"],
            )
        )

        columnas_semana_pdf = (
            ["Semana"]
            + [
                columna
                for columna
                in pivot_semanal_pdf.columns
                if columna != "Semana"
            ]
        )

        datos_semana_pdf = [
            columnas_semana_pdf
        ]

        for _, fila in pivot_semanal_pdf.iterrows():
            datos_semana_pdf.append(
                [
                    str(fila[columna])
                    for columna in columnas_semana_pdf
                ]
            )

        tabla_semana_pdf = Table(
            datos_semana_pdf,
            repeatRows=1,
        )

        tabla_semana_pdf.setStyle(
            TableStyle(
                [
                    (
                        "BACKGROUND",
                        (0, 0),
                        (-1, 0),
                        colors.lightgrey,
                    ),
                    (
                        "FONTNAME",
                        (0, 0),
                        (-1, 0),
                        "Helvetica-Bold",
                    ),
                    (
                        "FONTSIZE",
                        (0, 0),
                        (-1, -1),
                        7,
                    ),
                    (
                        "GRID",
                        (0, 0),
                        (-1, -1),
                        0.3,
                        colors.grey,
                    ),
                ]
            )
        )

        elementos.append(
            tabla_semana_pdf
        )

    # =====================================================
    # ÍNDICE DE DESEMPEÑO
    # =====================================================
    if not resumen_diario_pdf.empty:
        indice_pdf = calcular_indice_desempeno(
            df,
            resumen_diario_pdf,
        )

        if not indice_pdf.empty:
            indice_pdf = indice_pdf.sort_values(
                by="Nota desempeño",
                ascending=False,
            )

            elementos.append(
                Spacer(
                    1,
                    0.35 * cm,
                )
            )

            elementos.append(
                Paragraph(
                    "Índice de desempeño",
                    estilos["Heading3"],
                )
            )

            datos_indice_pdf = [
                [
                    "Técnico",
                    "Tickets cerrados",
                    "Resolución prom. (h)",
                    "Primer ticket (min)",
                    "Promedio diario",
                    "Nota",
                ]
            ]

            for _, fila in indice_pdf.iterrows():
                datos_indice_pdf.append(
                    [
                        str(fila["tecnico"]),
                        str(int(fila["total_cerrados"])),
                        (
                            ""
                            if pd.isna(
                                fila["resolucion_promedio"]
                            )
                            else f"{float(fila['resolucion_promedio']):.2f}"
                        ),
                        (
                            ""
                            if pd.isna(
                                fila["demora_primer_ticket"]
                            )
                            else f"{float(fila['demora_primer_ticket']):.2f}"
                        ),
                        f"{float(fila['promedio_diario']):.2f}",
                        f"{float(fila['Nota desempeño']):.2f}",
                    ]
                )

            tabla_indice_pdf = Table(
                datos_indice_pdf,
                colWidths=[
                    4.2 * cm,
                    2.7 * cm,
                    3.2 * cm,
                    3.2 * cm,
                    2.8 * cm,
                    1.8 * cm,
                ],
                repeatRows=1,
            )

            tabla_indice_pdf.setStyle(
                TableStyle(
                    [
                        (
                            "BACKGROUND",
                            (0, 0),
                            (-1, 0),
                            colors.lightgrey,
                        ),
                        (
                            "FONTNAME",
                            (0, 0),
                            (-1, 0),
                            "Helvetica-Bold",
                        ),
                        (
                            "GRID",
                            (0, 0),
                            (-1, -1),
                            0.4,
                            colors.grey,
                        ),
                    ]
                )
            )

            elementos.append(
                tabla_indice_pdf
            )

            elementos.append(
                Paragraph(
                    "La nota combina 40% volumen de tickets, "
                    "25% rapidez de resolución, "
                    "20% rapidez del primer ticket y "
                    "15% regularidad diaria.",
                    estilos["BodyText"],
                )
            )

    # =====================================================
    # RESUMEN EJECUTIVO
    # =====================================================
    elementos.append(
        Spacer(
            1,
            0.35 * cm,
        )
    )

    elementos.append(
        Paragraph(
            "Resumen ejecutivo",
            estilos["Heading3"],
        )
    )

    mensajes_pdf = []

    mensajes_pdf.append(
        f"Se analizaron {len(df)} tickets y participaron "
        f"{df['tecnico'].nunique()} técnicos."
    )

    if not resumen_diario_pdf.empty:
        promedio_equipo_pdf = (
            resumen_diario_pdf[
                "total_cerrados"
            ].sum()
            / max(
                resumen_diario_pdf[
                    "dias_activos"
                ].max(),
                1,
            )
        )

        mensajes_pdf.append(
            f"El equipo cerró aproximadamente "
            f"{promedio_equipo_pdf:.1f} tickets por día."
        )

    categoria_counts_pdf = (
        df["categoria"]
        .value_counts()
    )

    if not categoria_counts_pdf.empty:
        categoria_top_pdf = (
            categoria_counts_pdf.idxmax()
        )

        mensajes_pdf.append(
            f"El programa con mayor volumen fue "
            f"{categoria_top_pdf}."
        )

    entorno_counts_pdf = (
        df["entorno"]
        .value_counts()
    )

    if not entorno_counts_pdf.empty:
        entorno_top_pdf = (
            entorno_counts_pdf.idxmax()
        )

        porcentaje_entorno_pdf = (
            entorno_counts_pdf.max()
            / entorno_counts_pdf.sum()
            * 100
        )

        mensajes_pdf.append(
            f"El entorno predominante fue "
            f"{entorno_top_pdf} con "
            f"{porcentaje_entorno_pdf:.1f}% "
            f"de los tickets."
        )

    for mensaje_pdf in mensajes_pdf:
        elementos.append(
            Paragraph(
                f"• {mensaje_pdf}",
                estilos["BodyText"],
            )
        )

    elementos.append(PageBreak())

    elementos.append(
        Paragraph(
            "Detalle de tickets",
            estilos["Heading2"],
        )
    )

    columnas_detalle = [
        "id",
        "fecha_apertura",
        "fecha_asignacion",
        "fecha_cierre",
        "asunto",
        "tecnico",
        "categoria",
        "entorno",
        "estado",
        "tiempo_asignacion_cierre_min",
    ]

    columnas_detalle = [
        columna
        for columna in columnas_detalle
        if columna in df.columns
    ]

    datos_detalle = [columnas_detalle]

    for _, fila in df[columnas_detalle].head(500).iterrows():
        datos_detalle.append(
            [
                ""
                if pd.isna(valor)
                else str(valor)[:80]
                for valor in fila.tolist()
            ]
        )

    anchos = {
        "id": 2.0 * cm,
        "fecha_apertura": 3.4 * cm,
        "fecha_asignacion": 3.4 * cm,
        "fecha_cierre": 3.4 * cm,
        "asunto": 7.0 * cm,
        "tecnico": 3.6 * cm,
        "categoria": 3.0 * cm,
        "entorno": 2.5 * cm,
        "estado": 2.7 * cm,
        "tiempo_asignacion_cierre_min": 3.0 * cm,
    }

    tabla_detalle = Table(
        datos_detalle,
        colWidths=[
            anchos.get(columna, 4 * cm)
            for columna in columnas_detalle
        ],
        repeatRows=1,
    )

    tabla_detalle.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.lightgrey,
                ),
                (
                    "FONTNAME",
                    (0, 0),
                    (-1, 0),
                    "Helvetica-Bold",
                ),
                (
                    "FONTSIZE",
                    (0, 0),
                    (-1, -1),
                    6.5,
                ),
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.3,
                    colors.grey,
                ),
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "TOP",
                ),
            ]
        )
    )

    elementos.append(tabla_detalle)

    documento.build(elementos)
    buffer.seek(0)

    return buffer.getvalue()

# =========================================================
# LOGO EN BARRA LATERAL
# =========================================================

if LOGO.exists():
    st.sidebar.image(
        str(LOGO),
        use_container_width=True,
    )
else:
    st.sidebar.warning("No se encontró logo.png")
# =========================================================
# CARGA DEL ARCHIVO
# =========================================================

st.sidebar.header("📂 Cargar datos")

archivo_subido = st.sidebar.file_uploader(
    "Selecciona un archivo CSV o Excel",
    type=["csv", "xlsx"],
)

if archivo_subido is None:
    if not ARCHIVO_CSV.exists():
        st.title(
            "🎫 Sistema Inteligente de Análisis de Tickets"
        )

        st.info(
            "Carga un archivo CSV o Excel desde la barra lateral."
        )

        st.stop()

    tickets = pd.read_csv(
        ARCHIVO_CSV
    )

    st.sidebar.info(
        "Usando archivo predeterminado: datos/tickets.csv"
    )

else:
    nombre_archivo = archivo_subido.name.lower()

    if nombre_archivo.endswith(".csv"):
        tickets = pd.read_csv(
            archivo_subido
        )
    else:
        # El modelo_v1.xlsx usa la hoja DATOS.
        try:
            libro_excel = pd.ExcelFile(
                archivo_subido
            )

            if "DATOS" in libro_excel.sheet_names:
                tickets = pd.read_excel(
                    libro_excel,
                    sheet_name="DATOS",
                )
            else:
                tickets = pd.read_excel(
                    libro_excel,
                    sheet_name=0,
                )

        except Exception:
            tickets = pd.read_excel(
                archivo_subido
            )

    st.sidebar.success(
        f"✅ Archivo cargado: {archivo_subido.name}"
    )

tickets = normalizar_datos(tickets)

# Segunda exclusión de seguridad.
tickets = tickets[
    ~tickets["tecnico"].apply(tecnico_excluido)
].copy()


# =========================================================
# TÍTULO
# =========================================================

st.title(
    "🎫 Sistema Inteligente de Análisis de Tickets"
)

st.write(
    "Panel estadístico para el análisis de soporte técnico."
)


# =========================================================
# FILTROS
# =========================================================

st.sidebar.header("🔎 Filtros")

tecnicos_disponibles = sorted(
    tickets["tecnico"]
    .dropna()
    .astype(str)
    .unique()
    .tolist()
)

tecnico_seleccionado = st.sidebar.selectbox(
    "Técnico",
    ["Todos"] + tecnicos_disponibles,
)

estados_disponibles = sorted(
    tickets["estado"]
    .dropna()
    .astype(str)
    .unique()
    .tolist()
)

estado_seleccionado = st.sidebar.selectbox(
    "Estado",
    ["Todos"] + estados_disponibles,
)

categoria_seleccionada = st.sidebar.selectbox(
    "Categoría",
    ["Todas"] + CATEGORIAS_VISIBLES,
)

entorno_seleccionado = st.sidebar.selectbox(
    "Entorno",
    ["Todos"] + ENTORNOS_VISIBLES,
)

st.sidebar.subheader(
    "📅 Rango de fechas"
)

fechas_validas = (
    tickets["fecha_apertura"]
    .dropna()
)

if not fechas_validas.empty:
    fecha_minima = fechas_validas.min().date()
    fecha_maxima = fechas_validas.max().date()

    rango_fechas = st.sidebar.date_input(
        "Selecciona el período",
        value=(
            fecha_minima,
            fecha_maxima,
        ),
        min_value=fecha_minima,
        max_value=fecha_maxima,
    )
else:
    rango_fechas = ()


# =========================================================
# APLICAR FILTROS
# =========================================================

tickets_filtrados = tickets.copy()

if tecnico_seleccionado != "Todos":
    tickets_filtrados = tickets_filtrados[
        tickets_filtrados["tecnico"]
        == tecnico_seleccionado
    ]

if estado_seleccionado != "Todos":
    tickets_filtrados = tickets_filtrados[
        tickets_filtrados["estado"]
        == estado_seleccionado
    ]

if categoria_seleccionada != "Todas":
    tickets_filtrados = tickets_filtrados[
        tickets_filtrados["categoria"]
        == categoria_seleccionada
    ]

if entorno_seleccionado != "Todos":
    tickets_filtrados = tickets_filtrados[
        tickets_filtrados["entorno"]
        == entorno_seleccionado
    ]

if (
    isinstance(rango_fechas, (list, tuple))
    and len(rango_fechas) == 2
):
    fecha_inicio, fecha_fin = rango_fechas

    tickets_filtrados = tickets_filtrados[
        tickets_filtrados[
            "fecha_apertura"
        ].dt.date.between(
            fecha_inicio,
            fecha_fin,
        )
    ]


# =========================================================
# ESTADÍSTICAS GENERALES
# =========================================================

total = len(tickets_filtrados)

cerrados = int(
    tickets_filtrados["cerrado"].sum()
)

pendientes = total - cerrados

promedio_resolucion_horas = pd.to_numeric(
    tickets_filtrados[
        "tiempo_resolucion_horas"
    ],
    errors="coerce",
).mean()

promedio_texto = formato_horas(
    promedio_resolucion_horas
)


# =========================================================
# PRODUCTIVIDAD
# =========================================================

primeros_diarios = calcular_primer_ticket_diario(
    tickets_filtrados
)

ultimos_diarios = calcular_ultimo_cierre_diario(
    tickets_filtrados
)

resumen_productividad = resumir_productividad(
    primeros_diarios,
    ultimos_diarios,
)

if resumen_productividad.empty:
    media_asignacion = pd.NA
    media_primer_cierre = pd.NA
    media_ultimo_cierre = pd.NA
    media_demora = pd.NA
else:
    # Media desde todos los primeros tickets diarios válidos.
    media_asignacion = (
        primeros_diarios[
            "hora_asignacion_decimal"
        ].mean()
    )

    media_primer_cierre = (
        primeros_diarios[
            "hora_cierre_decimal"
        ].mean()
    )

    media_ultimo_cierre = (
        ultimos_diarios[
            "hora_ultimo_cierre_decimal"
        ].mean()
        if not ultimos_diarios.empty
        else pd.NA
    )

    media_demora = (
        primeros_diarios[
            "demora_primer_ticket_min"
        ].mean()
    )


# =========================================================
# INDICADORES GENERALES
# =========================================================

# Tickets resueltos el mismo día en que fueron creados.
mascara_mismo_dia = (
    tickets_filtrados["fecha_apertura"].notna()
    & tickets_filtrados["fecha_cierre"].notna()
    & (
        tickets_filtrados["fecha_apertura"].dt.date
        == tickets_filtrados["fecha_cierre"].dt.date
    )
)

resueltos_mismo_dia = int(
    mascara_mismo_dia.sum()
)

# Días distintos con resoluciones del mismo día.
dias_analizados = (
    tickets_filtrados.loc[
        mascara_mismo_dia,
        "fecha_cierre",
    ]
    .dropna()
    .dt.date
    .nunique()
)

tecnicos_considerados = int(
    tickets_filtrados["tecnico"]
    .dropna()
    .nunique()
)

# Promedio total de tickets resueltos por día.
if dias_analizados > 0:
    promedio_resueltos_dia = (
        resueltos_mismo_dia
        / dias_analizados
    )
else:
    promedio_resueltos_dia = 0.0

# Promedio diario por técnico.
if (
    dias_analizados > 0
    and tecnicos_considerados > 0
):
    promedio_diario_tecnico = (
        resueltos_mismo_dia
        / dias_analizados
        / tecnicos_considerados
    )
else:
    promedio_diario_tecnico = 0.0

respuesta_promedio_horas = pd.to_numeric(
    tickets_filtrados[
        "tiempo_primera_respuesta_horas"
    ],
    errors="coerce",
).mean()

# Primera fila.
col1, col2, col3 = st.columns(3)

col1.metric(
    "🎫 Total de tickets",
    total,
)

col2.metric(
    "✅ Cerrados",
    cerrados,
)

col3.metric(
    "📊 Promedio resueltos por día",
    f"{promedio_resueltos_dia:.1f}",
)

st.divider()

# Segunda fila.
col5, col6, col7, col8 = st.columns(4)

col5.metric(
    "👨‍💻 Técnicos considerados",
    tecnicos_considerados,
)

col6.metric(
    "👤 Promedio diario por técnico",
    f"{promedio_diario_tecnico:.1f}",
)

col7.metric(
    "⚡ Primera respuesta promedio",
    formato_horas(
        respuesta_promedio_horas
    ),
)

col8.metric(
    "📆 Días analizados",
    dias_analizados,
)

st.caption(
    "El promedio diario por técnico se calcula con los tickets "
    "resueltos el mismo día, dividido por los días analizados "
    "y por la cantidad de técnicos considerados."
)

st.divider()


# =========================================================
# GRÁFICOS PRINCIPALES
# =========================================================

col_grafico1, col_grafico2 = st.columns(2)

with col_grafico1:
    por_tecnico = (
        tickets_filtrados["tecnico"]
        .value_counts()
        .rename_axis("Técnico")
        .reset_index(name="Cantidad")
    )

    figura_tecnico = grafico_tickets_tecnico(
        por_tecnico
    )

    if figura_tecnico is not None:
        st.plotly_chart(
            figura_tecnico,
            use_container_width=True,
        )

        st.caption(
            "La barra roja identifica al técnico con menor cantidad "
            "de tickets."
        )
    else:
        st.info(
            "No hay datos para el gráfico de técnicos."
        )

with col_grafico2:
    # Crear todas las combinaciones:
    # Programa x Nube/Escritorio.
    indice_completo = pd.MultiIndex.from_product(
        [
            CATEGORIAS_VISIBLES,
            ENTORNOS_VISIBLES,
        ],
        names=[
            "Categoría",
            "Entorno",
        ],
    )

    por_categoria = (
        tickets_filtrados[
            tickets_filtrados["categoria"].isin(
                CATEGORIAS_VISIBLES
            )
        ]
        .groupby(
            [
                "categoria",
                "entorno",
            ]
        )
        .size()
        .reindex(
            indice_completo,
            fill_value=0,
        )
        .rename("Cantidad")
        .reset_index()
    )

    figura_categoria = grafico_categorias(
        por_categoria
    )

    if figura_categoria is not None:
        st.plotly_chart(
            figura_categoria,
            use_container_width=True,
        )

    st.caption(
        "Contabilidad, Remuneraciones, Renta y Facturación "
        "se muestran separados entre Nube y Escritorio."
    )

# =========================================================
# TICKETS ATENDIDOS POR TÉCNICO Y CATEGORÍA
# =========================================================

st.subheader(
    "👨‍💻 Tickets atendidos por técnico y categoría"
)

tickets_tecnico_categoria = (
    tickets_filtrados[
        tickets_filtrados["categoria"].isin(
            CATEGORIAS_VISIBLES
        )
    ]
    .groupby(
        [
            "tecnico",
            "categoria",
        ]
    )
    .size()
    .reset_index(
        name="Cantidad"
    )
)

if not tickets_tecnico_categoria.empty:
    figura_tecnico_categoria = px.bar(
        tickets_tecnico_categoria,
        x="tecnico",
        y="Cantidad",
        color="categoria",
        barmode="group",
        text="Cantidad",
        title=(
            "Distribución de programas "
            "atendidos por técnico"
        ),
        category_orders={
            "categoria": CATEGORIAS_VISIBLES,
        },
        labels={
            "tecnico": "Técnico",
            "categoria": "Programa",
        },
    )

    figura_tecnico_categoria.update_traces(
        width=0.16,
        textposition="outside",
    )

    figura_tecnico_categoria.update_layout(
        bargap=0.30,
        bargroupgap=0.08,
        xaxis_title="Técnico",
        yaxis_title="Cantidad de tickets",
        legend_title="Programa",
    )

    st.plotly_chart(
        figura_tecnico_categoria,
        use_container_width=True,
    )

    st.caption(
        "Este gráfico muestra cuántos tickets atendió "
        "cada técnico en Contabilidad, Remuneraciones, "
        "Renta y Facturación."
    )

    # Tabla resumen por técnico y programa.
    tabla_tecnico_categoria = (
        tickets_tecnico_categoria
        .pivot_table(
            index="tecnico",
            columns="categoria",
            values="Cantidad",
            aggfunc="sum",
            fill_value=0,
        )
        .reindex(
            columns=CATEGORIAS_VISIBLES,
            fill_value=0,
        )
        .reset_index()
    )

    tabla_tecnico_categoria["Total"] = (
        tabla_tecnico_categoria[
            CATEGORIAS_VISIBLES
        ].sum(axis=1)
    )

    tabla_tecnico_categoria[
        "Especialidad principal"
    ] = tabla_tecnico_categoria[
        CATEGORIAS_VISIBLES
    ].idxmax(axis=1)

    tabla_tecnico_categoria = (
        tabla_tecnico_categoria
        .rename(
            columns={
                "tecnico": "Técnico",
            }
        )
        .sort_values(
            by="Total",
            ascending=False,
        )
    )

    st.dataframe(
        tabla_tecnico_categoria,
        use_container_width=True,
        hide_index=True,
    )

else:
    tabla_tecnico_categoria = pd.DataFrame()

    st.info(
        "No hay datos suficientes para mostrar "
        "tickets por técnico y categoría."
    )

st.divider()


por_estado = (
    tickets_filtrados["estado"]
    .value_counts()
    .rename_axis("Estado")
    .reset_index(name="Cantidad")
)

figura_estado = grafico_estados(
    por_estado
)

if figura_estado is not None:
    st.plotly_chart(
        figura_estado,
        use_container_width=True,
    )

st.divider()


# =========================================================
# EVOLUCIÓN
# =========================================================

st.subheader(
    "📈 Evolución de tickets por fecha"
)

tickets_por_fecha = (
    tickets_filtrados
    .dropna(
        subset=["fecha_apertura"]
    )
    .assign(
        Fecha=lambda datos: (
            datos["fecha_apertura"].dt.date
        )
    )
    .groupby(
        "Fecha",
        as_index=False,
    )
    .size()
    .rename(
        columns={"size": "Cantidad"}
    )
)

if not tickets_por_fecha.empty:
    figura_evolucion = px.line(
        tickets_por_fecha,
        x="Fecha",
        y="Cantidad",
        markers=True,
        title="Cantidad de tickets creados por día",
    )

    st.plotly_chart(
        figura_evolucion,
        use_container_width=True,
    )
else:
    st.info(
        "No hay fechas para mostrar la evolución."
    )

st.divider()



# =========================================================
# SLA / TIEMPOS DE RESOLUCIÓN
# =========================================================

st.subheader(
    "⏱️ Distribución de tiempos de resolución"
)

sla_datos = tickets_filtrados.copy()

sla_datos["Tramo SLA"] = (
    pd.to_numeric(
        sla_datos["tiempo_resolucion_horas"],
        errors="coerce",
    )
    .apply(
        clasificar_sla_horas
    )
)

orden_sla = [
    "Menos de 1 h",
    "1 a 4 h",
    "4 a 8 h",
    "Más de 8 h",
]

sla_resumen = (
    sla_datos[
        sla_datos["Tramo SLA"].isin(
            orden_sla
        )
    ]["Tramo SLA"]
    .value_counts()
    .reindex(
        orden_sla,
        fill_value=0,
    )
    .rename_axis("Tramo")
    .reset_index(name="Cantidad")
)

total_sla = int(
    sla_resumen["Cantidad"].sum()
)

if total_sla > 0:
    sla_resumen["Porcentaje"] = (
        sla_resumen["Cantidad"]
        / total_sla
        * 100
    ).round(1)
else:
    sla_resumen["Porcentaje"] = 0.0

figura_sla = px.bar(
    sla_resumen,
    x="Tramo",
    y="Cantidad",
    text="Porcentaje",
    title="Tickets por tramo de tiempo de resolución",
)

figura_sla.update_traces(
    width=0.45,
    texttemplate="%{text:.1f}%",
    textposition="outside",
)

figura_sla.update_layout(
    bargap=0.35,
    xaxis_title="Tiempo de resolución",
    yaxis_title="Cantidad de tickets",
)

st.plotly_chart(
    figura_sla,
    use_container_width=True,
)

st.dataframe(
    sla_resumen,
    use_container_width=True,
    hide_index=True,
)

st.caption(
    "Este indicador permite ver qué porcentaje de tickets se resuelve "
    "en menos de 1 hora, entre 1 y 4 horas, entre 4 y 8 horas o en más de 8 horas."
)

st.divider()


# =========================================================
# PRODUCTIVIDAD DIARIA REAL POR TÉCNICO
# =========================================================

st.subheader(
    "📅 Productividad diaria por técnico"
)

productividad_diaria = calcular_productividad_diaria_tecnico(
    tickets_filtrados
)

resumen_diario = resumir_productividad_diaria(
    productividad_diaria
)

if not resumen_diario.empty:
    tabla_diaria = resumen_diario.rename(
        columns={
            "tecnico": "Técnico",
            "dias_activos": "Días activos",
            "promedio_diario": "Promedio diario",
            "mejor_dia": "Mejor día",
            "peor_dia": "Peor día",
            "total_cerrados": "Total cerrados",
        }
    ).copy()

    tabla_diaria["Promedio diario"] = (
        tabla_diaria["Promedio diario"]
        .round(2)
    )

    figura_diaria = px.bar(
        tabla_diaria,
        x="Técnico",
        y="Promedio diario",
        text="Promedio diario",
        title="Promedio de tickets cerrados por día y técnico",
    )

    figura_diaria.update_traces(
        width=0.40,
        textposition="outside",
    )

    figura_diaria.update_layout(
        bargap=0.45,
        yaxis_title="Tickets promedio por día",
    )

    st.plotly_chart(
        figura_diaria,
        use_container_width=True,
    )

    st.dataframe(
        tabla_diaria.sort_values(
            by="Promedio diario",
            ascending=False,
        ),
        use_container_width=True,
        hide_index=True,
    )

else:
    st.info(
        "No hay suficientes datos para calcular productividad diaria."
    )

st.divider()


# =========================================================
# MAPA DE CALOR TÉCNICO × PROGRAMA
# =========================================================

st.subheader(
    "🔥 Mapa de calor técnico por programa"
)

heatmap_base = (
    tickets_filtrados[
        tickets_filtrados["categoria"].isin(
            CATEGORIAS_VISIBLES
        )
    ]
    .groupby(
        [
            "tecnico",
            "categoria",
        ]
    )
    .size()
    .reset_index(
        name="Cantidad"
    )
)

if not heatmap_base.empty:
    matriz_heatmap = (
        heatmap_base
        .pivot_table(
            index="tecnico",
            columns="categoria",
            values="Cantidad",
            aggfunc="sum",
            fill_value=0,
        )
        .reindex(
            columns=CATEGORIAS_VISIBLES,
            fill_value=0,
        )
    )

    figura_heatmap = px.imshow(
        matriz_heatmap,
        text_auto=True,
        aspect="auto",
        labels={
            "x": "Programa",
            "y": "Técnico",
            "color": "Tickets",
        },
        title="Intensidad de atención por técnico y programa",
    )

    st.plotly_chart(
        figura_heatmap,
        use_container_width=True,
    )

    st.caption(
        "Los valores más altos indican en qué programas concentra más atención cada técnico."
    )

else:
    st.info(
        "No hay datos para construir el mapa de calor."
    )

st.divider()


# =========================================================
# NUBE VS ESCRITORIO POR TÉCNICO
# =========================================================

st.subheader(
    "☁️🖥️ Nube vs Escritorio por técnico"
)

entorno_tecnico = (
    tickets_filtrados[
        tickets_filtrados["entorno"].isin(
            ENTORNOS_VISIBLES
        )
    ]
    .groupby(
        [
            "tecnico",
            "entorno",
        ]
    )
    .size()
    .reset_index(
        name="Cantidad"
    )
)

if not entorno_tecnico.empty:
    totales_entorno = (
        entorno_tecnico
        .groupby(
            "tecnico"
        )["Cantidad"]
        .transform("sum")
    )

    entorno_tecnico["Porcentaje"] = (
        entorno_tecnico["Cantidad"]
        / totales_entorno
        * 100
    ).round(1)

    figura_entorno_tecnico = px.bar(
        entorno_tecnico,
        x="tecnico",
        y="Porcentaje",
        color="entorno",
        barmode="stack",
        text="Porcentaje",
        title="Distribución porcentual de Nube y Escritorio por técnico",
        labels={
            "tecnico": "Técnico",
            "entorno": "Entorno",
        },
    )

    figura_entorno_tecnico.update_traces(
        texttemplate="%{text:.1f}%",
        textposition="inside",
    )

    figura_entorno_tecnico.update_layout(
        yaxis_title="Porcentaje",
        legend_title="Entorno",
    )

    st.plotly_chart(
        figura_entorno_tecnico,
        use_container_width=True,
    )

else:
    st.info(
        "No hay datos de entorno para mostrar."
    )

st.divider()


# =========================================================
# TICKETS CERRADOS FUERA DEL MISMO DÍA
# =========================================================

st.subheader(
    "🕒 Tiempo calendario hasta el cierre"
)

cierres_dias = tickets_filtrados[
    tickets_filtrados["fecha_apertura"].notna()
    & tickets_filtrados["fecha_cierre"].notna()
].copy()

if not cierres_dias.empty:
    cierres_dias["dias_hasta_cierre"] = (
        cierres_dias["fecha_cierre"].dt.normalize()
        - cierres_dias["fecha_apertura"].dt.normalize()
    ).dt.days

    def tramo_dias(valor):
        if valor <= 0:
            return "Mismo día"
        if valor == 1:
            return "1 día después"
        if valor <= 3:
            return "2 a 3 días"
        return "Más de 3 días"

    cierres_dias["Tramo cierre"] = (
        cierres_dias["dias_hasta_cierre"]
        .apply(
            tramo_dias
        )
    )

    orden_cierre = [
        "Mismo día",
        "1 día después",
        "2 a 3 días",
        "Más de 3 días",
    ]

    resumen_cierre = (
        cierres_dias["Tramo cierre"]
        .value_counts()
        .reindex(
            orden_cierre,
            fill_value=0,
        )
        .rename_axis("Tramo")
        .reset_index(
            name="Cantidad"
        )
    )

    figura_cierre_dias = px.bar(
        resumen_cierre,
        x="Tramo",
        y="Cantidad",
        text="Cantidad",
        title="Tiempo calendario entre creación y cierre",
    )

    figura_cierre_dias.update_traces(
        width=0.45,
        textposition="outside",
    )

    st.plotly_chart(
        figura_cierre_dias,
        use_container_width=True,
    )

st.divider()


# =========================================================
# EVOLUCIÓN SEMANAL POR TÉCNICO
# =========================================================

st.subheader(
    "📈 Evolución semanal por técnico"
)

semanal = tickets_filtrados[
    tickets_filtrados["cerrado"]
    & tickets_filtrados["fecha_cierre"].notna()
].copy()

if not semanal.empty:
    semanal["Semana"] = (
        semanal["fecha_cierre"]
        .dt.to_period("W")
        .astype(str)
    )

    semanal_resumen = (
        semanal.groupby(
            [
                "Semana",
                "tecnico",
            ],
            as_index=False,
        )
        .size()
        .rename(
            columns={
                "size": "Tickets cerrados",
                "tecnico": "Técnico",
            }
        )
    )

    figura_semanal = px.line(
        semanal_resumen,
        x="Semana",
        y="Tickets cerrados",
        color="Técnico",
        markers=True,
        title="Evolución semanal de tickets cerrados por técnico",
    )

    st.plotly_chart(
        figura_semanal,
        use_container_width=True,
    )

st.divider()


# =========================================================
# ÍNDICE DE DESEMPEÑO 1 A 7
# =========================================================

st.subheader(
    "⭐ Índice de desempeño"
)

indice_desempeno = calcular_indice_desempeno(
    tickets_filtrados,
    resumen_diario,
)

if not indice_desempeno.empty:
    tabla_indice = indice_desempeno[
        [
            "tecnico",
            "total_cerrados",
            "resolucion_promedio",
            "demora_primer_ticket",
            "promedio_diario",
            "Nota desempeño",
        ]
    ].copy()

    tabla_indice = tabla_indice.rename(
        columns={
            "tecnico": "Técnico",
            "total_cerrados": "Tickets cerrados",
            "resolucion_promedio": "Resolución promedio (h)",
            "demora_primer_ticket": "Primer ticket (min)",
            "promedio_diario": "Promedio diario",
        }
    )

    for columna in [
        "Resolución promedio (h)",
        "Primer ticket (min)",
        "Promedio diario",
    ]:
        tabla_indice[columna] = pd.to_numeric(
            tabla_indice[columna],
            errors="coerce",
        ).round(2)

    tabla_indice = tabla_indice.sort_values(
        by="Nota desempeño",
        ascending=False,
    )

    figura_indice = px.bar(
        tabla_indice,
        x="Técnico",
        y="Nota desempeño",
        text="Nota desempeño",
        title="Índice de desempeño por técnico (escala 1 a 7)",
    )

    figura_indice.update_traces(
        width=0.40,
        textposition="outside",
    )

    figura_indice.update_layout(
        yaxis=dict(
            range=[1, 7],
        ),
        yaxis_title="Nota",
    )

    st.plotly_chart(
        figura_indice,
        use_container_width=True,
    )

    st.dataframe(
        tabla_indice,
        use_container_width=True,
        hide_index=True,
    )

    st.caption(
        "La nota combina 40% volumen de tickets, 25% rapidez de resolución, "
        "20% rapidez del primer ticket y 15% regularidad diaria."
    )

st.divider()


# =========================================================
# RESUMEN EJECUTIVO AUTOMÁTICO
# =========================================================

st.subheader(
    "🧾 Resumen ejecutivo"
)

mensajes_resumen = []

mensajes_resumen.append(
    f"Se analizaron {total} tickets y participaron "
    f"{tecnicos_considerados} técnicos."
)

if promedio_resueltos_dia > 0:
    mensajes_resumen.append(
        f"El equipo cerró en promedio {promedio_resueltos_dia:.1f} tickets por día."
    )

if not tickets_tecnico_categoria.empty:
    categoria_top = (
        tickets_filtrados["categoria"]
        .value_counts()
        .idxmax()
    )

    mensajes_resumen.append(
        f"El programa con mayor volumen fue {categoria_top}."
    )

if "entorno" in tickets_filtrados.columns:
    entorno_counts = (
        tickets_filtrados["entorno"]
        .value_counts()
    )

    if not entorno_counts.empty:
        entorno_top = entorno_counts.idxmax()
        porcentaje_top = (
            entorno_counts.max()
            / entorno_counts.sum()
            * 100
        )

        mensajes_resumen.append(
            f"El entorno predominante fue {entorno_top} "
            f"con {porcentaje_top:.1f}% de los tickets."
        )

if not resumen_diario.empty:
    mejor_tecnico = (
        resumen_diario
        .sort_values(
            by="promedio_diario",
            ascending=False,
        )
        .iloc[0]
    )

    mensajes_resumen.append(
        f"{mejor_tecnico['tecnico']} tuvo el mayor promedio diario "
        f"con {mejor_tecnico['promedio_diario']:.1f} tickets cerrados por día."
    )

for mensaje in mensajes_resumen:
    st.write(
        f"• {mensaje}"
    )

st.divider()


# =========================================================
# PRODUCTIVIDAD Y PRIMER TICKET
# =========================================================

st.subheader(
    "⏱️ Productividad diaria y horarios promedio"
)

p1, p2, p3, p4 = st.columns(4)

p1.metric(
    "🕘 Asignación del primer ticket promedio",
    formato_hora_decimal(
        media_asignacion
    ),
)

p2.metric(
    "✅ Cierre del primer ticket promedio",
    formato_hora_decimal(
        media_primer_cierre
    ),
)

p3.metric(
    "🏁 Último cierre promedio",
    formato_hora_decimal(
        media_ultimo_cierre
    ),
)

p4.metric(
    "⏳ Demora media al primer cierre",
    formato_minutos(
        media_demora
    ),
)

st.caption(
    "La demora se calcula usando la asignación y el cierre del "
    "mismo ticket. Para cada técnico y día se selecciona el primer "
    "ticket que cerró y luego se promedian las demoras diarias."
)

figura_demora = grafico_demora_primer_ticket(
    resumen_productividad
)

if figura_demora is not None:
    st.plotly_chart(
        figura_demora,
        use_container_width=True,
    )

    st.caption(
        "🔴 La barra roja identifica al técnico que más demora, "
        "en promedio, en cerrar su primer ticket diario."
    )
else:
    st.info(
        "No hay suficientes tickets asignados y cerrados "
        "el mismo día para calcular la demora."
    )


# =========================================================
# TABLA DE PRODUCTIVIDAD
# =========================================================

st.subheader(
    "📊 Productividad por técnico"
)

if resumen_productividad.empty:
    tabla_productividad = pd.DataFrame()

    st.info(
        "No hay datos suficientes para calcular la productividad."
    )
else:
    tabla_productividad = pd.DataFrame(
        {
            "Técnico": resumen_productividad[
                "tecnico"
            ],
            "Días analizados": resumen_productividad[
                "dias_analizados"
            ],
            "Asignación promedio": resumen_productividad[
                "primera_asignacion_promedio"
            ].apply(
                formato_hora_decimal
            ),
            "Primer cierre promedio": resumen_productividad[
                "primer_cierre_promedio"
            ].apply(
                formato_hora_decimal
            ),
            "Último cierre promedio": resumen_productividad[
                "ultimo_cierre_promedio"
            ].apply(
                formato_hora_decimal
            ),
            "Demora promedio": resumen_productividad[
                "demora_promedio_min"
            ].apply(
                formato_minutos
            ),
            "Demora promedio (min)": resumen_productividad[
                "demora_promedio_min"
            ].round(1),
        }
    ).sort_values(
        by="Demora promedio (min)",
        ascending=True,
    )

    mayor_demora_tabla = (
        tabla_productividad[
            "Demora promedio (min)"
        ].max()
    )

    def resaltar_mayor_demora(fila):
        if (
            fila["Demora promedio (min)"]
            == mayor_demora_tabla
        ):
            return [
                (
                    "background-color: #d62728; "
                    "color: white; "
                    "font-weight: bold;"
                )
            ] * len(fila)

        return [""] * len(fila)

    st.dataframe(
        tabla_productividad.style.apply(
            resaltar_mayor_demora,
            axis=1,
        ),
        use_container_width=True,
        hide_index=True,
    )


# =========================================================
# DETALLE DIARIO DEL PRIMER TICKET
# =========================================================

with st.expander(
    "📅 Ver primer ticket diario utilizado en el cálculo"
):
    if primeros_diarios.empty:
        st.info(
            "No hay datos para mostrar."
        )
    else:
        detalle_diario = pd.DataFrame(
            {
                "Fecha": primeros_diarios[
                    "fecha_trabajo"
                ],
                "Técnico": primeros_diarios[
                    "tecnico"
                ],
                "ID ticket": primeros_diarios[
                    "id"
                ],
                "Asignación": primeros_diarios[
                    "fecha_asignacion"
                ].dt.strftime(
                    "%d-%m-%Y %H:%M:%S"
                ),
                "Cierre": primeros_diarios[
                    "fecha_cierre"
                ].dt.strftime(
                    "%d-%m-%Y %H:%M:%S"
                ),
                "Demora": primeros_diarios[
                    "demora_primer_ticket_min"
                ].apply(
                    formato_minutos
                ),
                "Demora (min)": primeros_diarios[
                    "demora_primer_ticket_min"
                ].round(1),
            }
        )

        st.dataframe(
            detalle_diario.sort_values(
                by=[
                    "Fecha",
                    "Técnico",
                ],
                ascending=[
                    False,
                    True,
                ],
            ),
            use_container_width=True,
            hide_index=True,
        )


# =========================================================
# PDF
# =========================================================

st.divider()

st.subheader(
    "📄 Informe PDF"
)

if REPORTLAB_DISPONIBLE:
    pdf_bytes = generar_pdf_informe(
        df=tickets_filtrados,
        total=total,
        cerrados=cerrados,
        pendientes=pendientes,
        promedio_texto=promedio_texto,
        tecnico_filtro=tecnico_seleccionado,
        estado_filtro=estado_seleccionado,
        categoria_filtro=(
            f"{categoria_seleccionada} | "
            f"Entorno: {entorno_seleccionado}"
        ),
        rango_fechas=rango_fechas,
        tabla_productividad=tabla_productividad,
    )

    st.download_button(
        label="📥 Descargar informe PDF",
        data=pdf_bytes,
        file_name="informe_tickets.pdf",
        mime="application/pdf",
    )

    st.caption(
        "El PDF respeta los filtros actuales."
    )
else:
    st.warning(
        "Para generar el PDF instala ReportLab con: "
        "python -m pip install reportlab"
    )


# =========================================================
# DETALLE GENERAL
# =========================================================

st.divider()

st.subheader(
    "📋 Detalle de tickets"
)

columnas_mostrar = [
    "id",
    "fecha_apertura",
    "fecha_asignacion",
    "fecha_cierre",
    "asunto",
    "tecnico",
    "categoria",
    "entorno",
    "estado",
    "tiempo_primera_respuesta_horas",
    "tiempo_resolucion_horas",
    "tiempo_asignacion_cierre_min",
]

columnas_existentes = [
    columna
    for columna in columnas_mostrar
    if columna in tickets_filtrados.columns
]

st.dataframe(
    tickets_filtrados[
        columnas_existentes
    ],
    use_container_width=True,
    hide_index=True,
)
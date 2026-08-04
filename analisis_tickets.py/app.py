from pathlib import Path
from io import BytesIO
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.lib.styles import (
        ParagraphStyle,
        getSampleStyleSheet,
    )
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        HRFlowable,
        Image,
        KeepTogether,
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
LOGO = BASE_DIR / "makinola3000.png"

CATEGORIAS_VISIBLES = [
    "ContaPlus",
    "eContabilidad",
    "eFacturacionElectronica",
    "eRemuneraciones",
    "eRenta",
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

def clasificar_ticket(asunto, etiqueta_producto=None):
    """
    Clasifica cada ticket en uno de estos productos:

    - ContaPlus
    - eContabilidad
    - eFacturacionElectronica
    - eRemuneraciones
    - eRenta

    Se prioriza E1 / Etiqueta 1 del archivo Excel.
    Si la etiqueta no viene informada, se intenta clasificar por el asunto.
    """
    texto = limpiar_texto(asunto).lower()
    etiqueta = limpiar_texto(etiqueta_producto).lower()

    # Normalización simple para reconocer variantes.
    etiqueta_compacta = (
        etiqueta
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
        .replace("ó", "o")
        .replace("í", "i")
        .replace("á", "a")
        .replace("é", "e")
        .replace("ú", "u")
    )

    texto_normalizado = (
        texto
        .replace("ó", "o")
        .replace("í", "i")
        .replace("á", "a")
        .replace("é", "e")
        .replace("ú", "u")
    )

    # -----------------------------------------------------
    # CLASIFICACIÓN DIRECTA POR E1 / ETIQUETA 1
    # -----------------------------------------------------
    if etiqueta_compacta in {
        "contaplus",
        "conta+",
    }:
        return "ContaPlus"

    if etiqueta_compacta in {
        "econtabilidad",
        "contabilidadelectronica",
    }:
        return "eContabilidad"

    if etiqueta_compacta in {
        "efacturacionelectronica",
        "facturacionelectronica",
        "efacturacion",
    }:
        return "eFacturacionElectronica"

    if etiqueta_compacta in {
        "eremuneraciones",
        "remuneracioneselectronicas",
        "eremuneracion",
    }:
        return "eRemuneraciones"

    if etiqueta_compacta in {
        "erenta",
        "rentaelectronica",
    }:
        return "eRenta"

    if etiqueta_compacta in {
        "otros",
        "otro",
    }:
        return "Otros"

    # -----------------------------------------------------
    # CLASIFICACIÓN AUXILIAR POR ASUNTO
    # -----------------------------------------------------
    palabras_contaplus = [
        "contaplus",
        "conta plus",
        "conta+",
        "plan de cuentas",
        "libro diario",
        "libro mayor",
        "asiento contable",
        "centralizacion",
        "balance",
    ]

    palabras_econtabilidad = [
        "econtabilidad",
        "contabilidad electronica",
        "registro de compras",
        "registro de ventas",
        "libro electronico",
        "conciliacion",
        "comprobante contable",
    ]

    palabras_facturacion = [
        "efacturacionelectronica",
        "facturacion electronica",
        "factura electronica",
        "boleta electronica",
        "nota de credito",
        "nota de debito",
        "dte",
        "documento tributario electronico",
    ]

    palabras_remuneraciones = [
        "eremuneraciones",
        "remuneraciones",
        "liquidacion",
        "sueldo",
        "previred",
        "afp",
        "fonasa",
        "isapre",
        "finiquito",
        "libro de remuneraciones",
        "lre",
        "haberes",
        "descuentos",
        "vacaciones",
        "licencia",
    ]

    palabras_renta = [
        "erenta",
        "operacion renta",
        "declaracion de renta",
        "formulario 22",
        "f22",
        "declaracion jurada",
        "dj ",
        "dj1847",
        "dj1866",
        "dj1887",
        "dj1945",
        "dj1946",
        "honorarios",
        "impuesto a la renta",
    ]

    if any(palabra in texto_normalizado for palabra in palabras_contaplus):
        return "ContaPlus"

    if any(palabra in texto_normalizado for palabra in palabras_econtabilidad):
        return "eContabilidad"

    if any(palabra in texto_normalizado for palabra in palabras_facturacion):
        return "eFacturacionElectronica"

    if any(palabra in texto_normalizado for palabra in palabras_remuneraciones):
        return "eRemuneraciones"

    if any(palabra in texto_normalizado for palabra in palabras_renta):
        return "eRenta"

    return "Otros"



def normalizar_producto_excel(valor):
    """
    Normaliza directamente el valor de la columna Producto del Excel.

    Productos permitidos:
    - ContaPlus
    - eContabilidad
    - eFacturacionElectronica
    - eRemuneraciones
    - eRenta

    Cualquier otro valor se descarta del análisis.
    """
    producto = limpiar_texto(valor)

    if not producto:
        return pd.NA

    compacto = (
        producto.lower()
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
        .replace("ó", "o")
        .replace("í", "i")
        .replace("á", "a")
        .replace("é", "e")
        .replace("ú", "u")
    )

    equivalencias = {
        "contaplus": "ContaPlus",
        "conta+": "ContaPlus",
        "econtabilidad": "eContabilidad",
        "efacturacionelectronica": "eFacturacionElectronica",
        "efacturacion": "eFacturacionElectronica",
        "eremuneraciones": "eRemuneraciones",
        "eremuneracion": "eRemuneraciones",
        "erenta": "eRenta",
    }

    return equivalencias.get(
        compacto,
        pd.NA,
    )


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
            "Etiqueta 1": "etiqueta_producto",
            "Etiqueta 2": "etiqueta_entorno",
            "E1": "etiqueta_producto",
            "E2": "etiqueta_entorno",
            "Producto": "producto_excel",
            "PRODUCTO": "producto_excel",
            "producto": "producto_excel",
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

    # Producto: usar directamente la columna Producto del Excel.
    if "producto_excel" not in df.columns:
        df["producto_excel"] = pd.NA

    if "etiqueta_producto" not in df.columns:
        df["etiqueta_producto"] = pd.NA

    if "etiqueta_entorno" not in df.columns:
        df["etiqueta_entorno"] = pd.NA

    df["categoria"] = df["producto_excel"].apply(
        normalizar_producto_excel
    )

    # Si el archivo antiguo no contiene Producto, usar E1 como respaldo.
    mascara_sin_producto = df["categoria"].isna()

    df.loc[
        mascara_sin_producto,
        "categoria",
    ] = df.loc[
        mascara_sin_producto,
        "etiqueta_producto",
    ].apply(
        normalizar_producto_excel
    )

    df["entorno"] = df.apply(
        lambda fila: clasificar_entorno(
            fila.get("etiqueta_entorno"),
            " ".join(
                [
                    limpiar_texto(fila.get("asunto")),
                    limpiar_texto(fila.get("etiqueta_producto")),
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

    # Mantener únicamente los productos permitidos.
    # Así se elimina completamente la categoría "Otros".
    df = df[
        df["categoria"].isin(
            CATEGORIAS_VISIBLES
        )
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
    - 35% volumen de tickets cerrados
    - 30% rapidez de resolución
    - 15% rapidez del primer ticket
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

    # Pesos solicitados:
    # - 35% volumen de tickets
    # - 30% rapidez de resolución
    # - 15% rapidez del primer ticket
    # - 20% regularidad diaria
    

     # sos suman 100%, por lo que no es necesario normalizar.
   

    base["indice_0_1"] = (
        base["score_volumen"] * 0.35
        + base["score_rapidez"] * 0.30
        + base["score_primer_ticket"] * 0.15
        + base["score_regularidad"] * 0.20
    ) / peso_total

    # Convertir de 0..1 a escala 1..7.
    base["Nota desempeño"] = (
        1 + base["indice_0_1"].clip(0, 1) * 6
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
    Muestra la cantidad total de tickets por producto.
    """
    if datos.empty:
        return None

    figura = px.bar(
        datos,
        x="Producto",
        y="Cantidad",
        text="Cantidad",
        title="Tickets por producto",
        category_orders={
            "Producto": CATEGORIAS_VISIBLES,
        },
        color_discrete_sequence=["#4C78A8"],
    )

    figura.update_traces(
        width=0.42,
        textposition="outside",
    )

    figura.update_layout(
        showlegend=False,
        bargap=0.38,
        xaxis_title="Producto",
        yaxis_title="Cantidad",
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

    # Colores corporativos del informe.
    COLOR_PRINCIPAL = colors.HexColor("#172A3A")
    COLOR_SECUNDARIO = colors.HexColor("#2E5D7B")
    COLOR_ACENTO = colors.HexColor("#D4A72C")
    COLOR_SUAVE = colors.HexColor("#EAF0F4")
    COLOR_TEXTO = colors.HexColor("#263238")

    fecha_generacion = datetime.now()

    estilo_portada_titulo = ParagraphStyle(
        "PortadaTitulo",
        parent=estilos["Title"],
        fontName="Helvetica-Bold",
        fontSize=26,
        leading=31,
        alignment=TA_CENTER,
        textColor=COLOR_PRINCIPAL,
        spaceAfter=14,
    )

    estilo_portada_subtitulo = ParagraphStyle(
        "PortadaSubtitulo",
        parent=estilos["BodyText"],
        fontName="Helvetica",
        fontSize=13,
        leading=18,
        alignment=TA_CENTER,
        textColor=COLOR_SECUNDARIO,
        spaceAfter=10,
    )

    estilo_seccion = ParagraphStyle(
        "SeccionCorporativa",
        parent=estilos["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=17,
        leading=21,
        textColor=COLOR_PRINCIPAL,
        spaceBefore=5,
        spaceAfter=10,
    )

    estilo_subseccion = ParagraphStyle(
        "SubseccionCorporativa",
        parent=estilos["Heading3"],
        fontName="Helvetica-Bold",
        fontSize=12,
        leading=15,
        textColor=COLOR_SECUNDARIO,
        spaceBefore=6,
        spaceAfter=7,
    )

    estilo_indice = ParagraphStyle(
        "Indice",
        parent=estilos["BodyText"],
        fontName="Helvetica",
        fontSize=11,
        leading=20,
        leftIndent=1.0 * cm,
        textColor=COLOR_TEXTO,
    )

    def agregar_pie_pagina(canvas, doc):
        """Agrega fecha, marca y número de página."""
        canvas.saveState()

        ancho_pagina, _ = landscape(A4)

        canvas.setStrokeColor(COLOR_SECUNDARIO)
        canvas.setLineWidth(0.5)
        canvas.line(
            doc.leftMargin,
            0.72 * cm,
            ancho_pagina - doc.rightMargin,
            0.72 * cm,
        )

        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(colors.HexColor("#5F6B73"))

        canvas.drawString(
            doc.leftMargin,
            0.38 * cm,
            "MAKINOLA 3000 - Sistema Inteligente de Análisis de Tickets",
        )

        canvas.drawCentredString(
            ancho_pagina / 2,
            0.38 * cm,
            fecha_generacion.strftime("Generado el %d-%m-%Y a las %H:%M"),
        )

        canvas.drawRightString(
            ancho_pagina - doc.rightMargin,
            0.38 * cm,
            f"Página {doc.page}",
        )

        canvas.restoreState()

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

    # -----------------------------------------------------
    # PORTADA
    # -----------------------------------------------------
    elementos.append(Spacer(1, 1.0 * cm))

    if LOGO.exists():
        logo_pdf = Image(
            str(LOGO),
            width=4.4 * cm,
            height=4.4 * cm,
        )
        logo_pdf.hAlign = "CENTER"
        elementos.append(logo_pdf)
        elementos.append(Spacer(1, 0.25 * cm))

    elementos.append(
        Paragraph(
            "MAKINOLA 3000",
            estilo_portada_subtitulo,
        )
    )

    elementos.append(
        Paragraph(
            "Informe de Análisis de Tickets de Soporte",
            estilo_portada_titulo,
        )
    )

    elementos.append(
        HRFlowable(
            width="72%",
            thickness=2,
            color=COLOR_ACENTO,
            spaceBefore=8,
            spaceAfter=16,
            hAlign="CENTER",
        )
    )

    elementos.append(
        Paragraph(
            "Dashboard de productividad, tiempos de respuesta, "
            "productos, entornos y cumplimiento operativo.",
            estilo_portada_subtitulo,
        )
    )

    elementos.append(Spacer(1, 0.6 * cm))

    portada_datos = [
        ["Fecha de generación", fecha_generacion.strftime("%d-%m-%Y %H:%M")],
        ["Total de registros analizados", str(len(df))],
        ["Filtros aplicados", " | ".join(filtros)],
    ]

    tabla_portada = Table(
        portada_datos,
        colWidths=[5.2 * cm, 16.5 * cm],
    )

    tabla_portada.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, -1), COLOR_PRINCIPAL),
                ("TEXTCOLOR", (0, 0), (0, -1), colors.white),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("BACKGROUND", (1, 0), (1, -1), COLOR_SUAVE),
                ("TEXTCOLOR", (1, 0), (1, -1), COLOR_TEXTO),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.white),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 9),
                ("RIGHTPADDING", (0, 0), (-1, -1), 9),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )

    elementos.append(tabla_portada)
    elementos.append(PageBreak())

    # -----------------------------------------------------
    # ÍNDICE
    # -----------------------------------------------------
    elementos.append(
        Paragraph(
            "Índice del informe",
            estilo_seccion,
        )
    )

    elementos.append(
        HRFlowable(
            width="100%",
            thickness=1,
            color=COLOR_ACENTO,
            spaceAfter=10,
        )
    )

    indice_items = [
        "1. Resumen general e indicadores principales",
        "2. Visualizaciones principales",
        "3. Productividad del primer ticket diario",
        "4. Distribución de productos por técnico",
        "5. Tiempo de resolución por producto",
        "6. Productividad diaria y evolución semanal",
        "7. Índice de desempeño",
        "8. Resumen ejecutivo",
        "9. Detalle completo de tickets",
    ]

    for item in indice_items:
        elementos.append(
            Paragraph(
                item,
                estilo_indice,
            )
        )

    elementos.append(PageBreak())

    # -----------------------------------------------------
    # RESUMEN GENERAL
    # -----------------------------------------------------
    elementos.append(
        Paragraph(
            "1. Resumen general",
            estilo_seccion,
        )
    )

    elementos.append(
        Paragraph(
            "Los resultados siguientes respetan todos los filtros "
            "seleccionados en la aplicación.",
            estilos["BodyText"],
        )
    )

    elementos.append(Spacer(1, 0.25 * cm))

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
                    COLOR_PRINCIPAL,
                ),
                (
                    "TEXTCOLOR",
                    (0, 0),
                    (-1, 0),
                    colors.white,
                ),
                (
                    "FONTNAME",
                    (0, 0),
                    (-1, 0),
                    "Helvetica-Bold",
                ),
                (
                    "BACKGROUND",
                    (0, 1),
                    (-1, -1),
                    COLOR_SUAVE,
                ),
                (
                    "TEXTCOLOR",
                    (0, 1),
                    (-1, -1),
                    COLOR_TEXTO,
                ),
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.5,
                    colors.white,
                ),
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE",
                ),
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    7,
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    7,
                ),
            ]
        )
    )

    elementos.append(tabla_resumen)
    elementos.append(PageBreak())

    # -----------------------------------------------------
    # VISUALIZACIONES PRINCIPALES DEL PDF
    # -----------------------------------------------------
    elementos.append(
        Paragraph(
            "2. Visualizaciones principales",
            estilo_seccion,
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

    # Tickets por producto
    por_producto_pdf = (
        df[
            df["categoria"].isin(
                CATEGORIAS_VISIBLES
            )
        ]["categoria"]
        .value_counts()
        .reindex(
            CATEGORIAS_VISIBLES,
            fill_value=0,
        )
        .rename_axis("Producto")
        .reset_index(name="Cantidad")
    )

    elementos.append(
        crear_grafico_barras_pdf(
            titulo="Tickets por producto",
            etiquetas=por_producto_pdf["Producto"].tolist(),
            valores=por_producto_pdf["Cantidad"].tolist(),
            destacar="ninguno",
            nota=(
                "Productos: ContaPlus, eContabilidad, "
                "eFacturacionElectronica, eRemuneraciones, "
                "eRenta y Otros."
            ),
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
            "3. Productividad del primer ticket diario",
            estilo_seccion,
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
            "4. Tickets atendidos por técnico y categoría",
            estilo_seccion,
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

        columnas_productos_pdf = [
            "Técnico",
            "ContaPlus",
            "eContabilidad",
            "eFacturacionElectronica",
            "eRemuneraciones",
            "eRenta",
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

        datos_productos_pdf = [
            columnas_productos_pdf
        ]

        for _, fila in tabla_distribucion_pdf.iterrows():
            datos_productos_pdf.append(
                [
                    str(fila[columna])
                    for columna in columnas_productos_pdf
                ]
            )

        tabla_productos_pdf = Table(
            datos_productos_pdf,
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

        tabla_productos_pdf.setStyle(
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
            tabla_productos_pdf
        )

        elementos.append(
            Paragraph(
                "La columna Especialidad principal corresponde "
                "al producto en el que cada técnico atendió "
                "la mayor cantidad de tickets.",
                estilos["BodyText"],
            )
        )

    # -----------------------------------------------------
    # ANALÍTICA AVANZADA EN PDF
    # -----------------------------------------------------

    elementos.append(
        Paragraph(
            "5. Analítica avanzada",
            estilo_seccion,
        )
    )

    # =====================================================
    # TIEMPO DE RESOLUCIÓN POR PRODUCTO
    # =====================================================
    elementos.append(
        Paragraph(
            "Tiempo de resolución por producto",
            estilo_subseccion,
        )
    )

    resolucion_producto_pdf = (
        df[
            df["categoria"].isin(
                CATEGORIAS_VISIBLES
            )
        ]
        .assign(
            tiempo_horas=lambda datos: pd.to_numeric(
                datos["tiempo_resolucion_horas"],
                errors="coerce",
            )
        )
        .dropna(
            subset=["tiempo_horas"]
        )
    )

    if not resolucion_producto_pdf.empty:
        resumen_producto_pdf = (
            resolucion_producto_pdf
            .groupby(
                "categoria",
                as_index=False,
            )
            .agg(
                Tickets=("id", "count"),
                Promedio_horas=("tiempo_horas", "mean"),
                Mediana_horas=("tiempo_horas", "median"),
                Maximo_horas=("tiempo_horas", "max"),
            )
            .rename(
                columns={
                    "categoria": "Producto",
                    "Promedio_horas": "Promedio (h)",
                    "Mediana_horas": "Mediana (h)",
                    "Maximo_horas": "Máximo (h)",
                }
            )
            .set_index("Producto")
            .reindex(
                CATEGORIAS_VISIBLES
            )
            .fillna(0)
            .reset_index()
        )

        elementos.append(
            crear_grafico_barras_pdf(
                titulo="Tiempo promedio de resolución por producto",
                etiquetas=resumen_producto_pdf["Producto"].tolist(),
                valores=resumen_producto_pdf["Promedio (h)"].tolist(),
                destacar="mayor",
                nota=(
                    "La barra roja identifica el producto con "
                    "mayor tiempo promedio de resolución."
                ),
            )
        )

        datos_resolucion_pdf = [
            [
                "Producto",
                "Tickets",
                "Promedio (h)",
                "Mediana (h)",
                "Máximo (h)",
            ]
        ]

        for _, fila in resumen_producto_pdf.iterrows():
            datos_resolucion_pdf.append(
                [
                    str(fila["Producto"]),
                    str(int(fila["Tickets"])),
                    f"{float(fila['Promedio (h)']):.2f}",
                    f"{float(fila['Mediana (h)']):.2f}",
                    f"{float(fila['Máximo (h)']):.2f}",
                ]
            )

        tabla_resolucion_pdf = Table(
            datos_resolucion_pdf,
            colWidths=[
                5.2 * cm,
                2.3 * cm,
                3.0 * cm,
                3.0 * cm,
                3.0 * cm,
            ],
            repeatRows=1,
        )

        tabla_resolucion_pdf.setStyle(
            TableStyle(
                [
                    (
                        "BACKGROUND",
                        (0, 0),
                        (-1, 0),
                        COLOR_PRINCIPAL,
                    ),
                    (
                        "TEXTCOLOR",
                        (0, 0),
                        (-1, 0),
                        colors.white,
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
                        (-1, -1),
                        "CENTER",
                    ),
                ]
            )
        )

        elementos.append(
            tabla_resolucion_pdf
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
    distribucion_producto_pdf = (
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

    if not distribucion_producto_pdf.empty:
        matriz_producto_pdf = (
            distribucion_producto_pdf
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

        matriz_producto_pdf["Total"] = (
            matriz_producto_pdf[
                CATEGORIAS_VISIBLES
            ].sum(
                axis=1
            )
        )

        matriz_producto_pdf[
            "Especialidad principal"
        ] = matriz_producto_pdf[
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
                "Distribución técnico por producto",
                estilos["Heading3"],
            )
        )

        columnas_producto_pdf = [
            "Técnico",
            "ContaPlus",
            "eContabilidad",
            "eFacturacionElectronica",
            "eRemuneraciones",
            "eRenta",
            "Total",
            "Especialidad principal",
        ]

        matriz_producto_pdf = (
            matriz_producto_pdf
            .rename(
                columns={
                    "tecnico": "Técnico",
                }
            )
        )

        datos_producto_pdf = [
            columnas_producto_pdf
        ]

        for _, fila in matriz_producto_pdf.iterrows():
            datos_producto_pdf.append(
                [
                    str(fila[columna])
                    for columna in columnas_producto_pdf
                ]
            )

        tabla_producto_pdf = Table(
            datos_producto_pdf,
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

        tabla_producto_pdf.setStyle(
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
            tabla_producto_pdf
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
                    "La nota combina 35% volumen de tickets, "
                    "30% rapidez de resolución, "
                    "15% rapidez del primer ticket y "
                    "15% regularidad diaria. "
                    "Los porcentajes se normalizan porque suman 95%.",
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
            f"El producto con mayor volumen fue "
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

    documento.build(
        elementos,
        onFirstPage=agregar_pie_pagina,
        onLaterPages=agregar_pie_pagina,
    )
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
    # No mostrar resultados de ejemplo al abrir la aplicación.
    st.title(
        "🎫 Sistema Inteligente de Análisis de Tickets"
    )

    st.markdown(
        """
        ### Bienvenido

        Para comenzar el análisis:

        1. Presiona **Upload** en la barra lateral.
        2. Selecciona el archivo CSV o Excel exportado.
        3. Espera unos segundos mientras se procesan los datos.
        4. Revisa los indicadores, gráficos y descarga el informe PDF.

        **La aplicación no guarda permanentemente el archivo cargado.**
        """
    )

    st.info(
        "📂 Carga un archivo CSV o Excel para visualizar los resultados."
    )

    st.stop()

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

tecnicos_seleccionados = st.sidebar.multiselect(
    "Comparar técnicos",
    options=tecnicos_disponibles,
    default=[],
    placeholder="Selecciona uno o más técnicos",
    help=(
        "Deja el campo vacío para visualizar a todos. "
        "Selecciona dos o más técnicos para compararlos."
    ),
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
    "Producto",
    ["Todos"] + CATEGORIAS_VISIBLES,
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

if tecnicos_seleccionados:
    tickets_filtrados = tickets_filtrados[
        tickets_filtrados["tecnico"].isin(
            tecnicos_seleccionados
        )
    ]

if estado_seleccionado != "Todos":
    tickets_filtrados = tickets_filtrados[
        tickets_filtrados["estado"]
        == estado_seleccionado
    ]

if categoria_seleccionada != "Todos":
    tickets_filtrados = tickets_filtrados[
        tickets_filtrados["categoria"]
        == categoria_seleccionada
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
    por_categoria = (
        tickets_filtrados[
            tickets_filtrados["categoria"].isin(
                CATEGORIAS_VISIBLES
            )
        ]["categoria"]
        .value_counts()
        .reindex(
            CATEGORIAS_VISIBLES,
            fill_value=0,
        )
        .rename_axis("Producto")
        .reset_index(name="Cantidad")
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
        "El gráfico muestra la cantidad de tickets de ContaPlus, "
        "eContabilidad, eFacturacionElectronica, eRemuneraciones, "
        "eRenta y Otros."
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
            "Distribución de productos "
            "atendidos por técnico"
        ),
        category_orders={
            "categoria": CATEGORIAS_VISIBLES,
        },
        labels={
            "tecnico": "Técnico",
            "categoria": "Producto",
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
        legend_title="Producto",
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

    # Tabla resumen por técnico y producto.
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
# TIEMPO DE RESOLUCIÓN POR PRODUCTO
# =========================================================

st.subheader(
    "⏱️ Tiempo de resolución por producto"
)

resolucion_producto = (
    tickets_filtrados[
        tickets_filtrados["categoria"].isin(
            CATEGORIAS_VISIBLES
        )
    ]
    .assign(
        tiempo_horas=lambda datos: pd.to_numeric(
            datos["tiempo_resolucion_horas"],
            errors="coerce",
        )
    )
    .dropna(
        subset=["tiempo_horas"]
    )
)

if not resolucion_producto.empty:
    resumen_resolucion_producto = (
        resolucion_producto
        .groupby(
            "categoria",
            as_index=False,
        )
        .agg(
            Tickets=("id", "count"),
            Promedio_horas=("tiempo_horas", "mean"),
            Mediana_horas=("tiempo_horas", "median"),
            Maximo_horas=("tiempo_horas", "max"),
        )
        .rename(
            columns={
                "categoria": "Producto",
                "Promedio_horas": "Promedio (h)",
                "Mediana_horas": "Mediana (h)",
                "Maximo_horas": "Máximo (h)",
            }
        )
        .set_index("Producto")
        .reindex(
            CATEGORIAS_VISIBLES
        )
        .fillna(0)
        .reset_index()
    )

    for columna in [
        "Promedio (h)",
        "Mediana (h)",
        "Máximo (h)",
    ]:
        resumen_resolucion_producto[columna] = (
            resumen_resolucion_producto[columna]
            .round(2)
        )

    figura_resolucion_producto = px.bar(
        resumen_resolucion_producto,
        x="Producto",
        y="Promedio (h)",
        text="Promedio (h)",
        title="Tiempo promedio de resolución por producto",
        category_orders={
            "Producto": CATEGORIAS_VISIBLES,
        },
        color_discrete_sequence=["#4C78A8"],
    )

    figura_resolucion_producto.update_traces(
        width=0.42,
        texttemplate="%{text:.2f} h",
        textposition="outside",
    )

    figura_resolucion_producto.update_layout(
        showlegend=False,
        bargap=0.38,
        xaxis_title="Producto",
        yaxis_title="Horas promedio",
    )

    st.plotly_chart(
        figura_resolucion_producto,
        use_container_width=True,
    )

    st.dataframe(
        resumen_resolucion_producto,
        use_container_width=True,
        hide_index=True,
    )

    st.caption(
        "El promedio y la mediana permiten comparar cuánto demora "
        "la resolución de tickets en cada producto."
    )

else:
    resumen_resolucion_producto = pd.DataFrame()

    st.info(
        "No hay tiempos de resolución válidos para comparar por producto."
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
    "🔥 Mapa de calor técnico por producto"
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
            "x": "Producto",
            "y": "Técnico",
            "color": "Tickets",
        },
        title="Intensidad de atención por técnico y producto",
    )

    st.plotly_chart(
        figura_heatmap,
        use_container_width=True,
    )

    st.caption(
        "Los valores más altos indican en qué productos concentra más atención cada técnico."
    )

else:
    st.info(
        "No hay datos para construir el mapa de calor."
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
        "La nota combina 35% volumen de tickets, 30% rapidez de resolución, "
        "15% rapidez del primer ticket y 15% regularidad diaria. "
        "Los porcentajes se normalizan porque suman 95%."
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
        f"El producto con mayor volumen fue {categoria_top}."
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
        tecnico_filtro=(
            ", ".join(tecnicos_seleccionados)
            if tecnicos_seleccionados
            else "Todos"
        ),
        estado_filtro=estado_seleccionado,
        categoria_filtro=categoria_seleccionada,
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
        "El PDF respeta los filtros actuales e incluye portada, "
        "logo, índice, gráficos, fecha de generación y numeración de páginas."
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

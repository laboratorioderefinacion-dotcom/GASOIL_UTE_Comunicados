#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# GASOIL UTE COMUNICADOS

import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO
from datetime import datetime
from mailmerge import MailMerge

st.set_page_config(page_title="GAS OIL UTE - Comunicados", layout="centered")

# -----------------------------
# Funciones auxiliares (igual lógica que tu script)
# -----------------------------
def extraer_valor(df, nombre_celda, n_muestra):
    """
    Busca en columna 1 el nombre_celda y devuelve valor en col (3+n_muestra).
    """
    col = 3 + n_muestra
    fila = df[df[1] == nombre_celda]
    if fila.empty:
        return "#########"  # El análisis no existe
    valor = fila.iloc[0, col]
    if pd.isna(valor) or str(valor).strip() == "":
        return "-----"  # Existe pero sin dato
    return str(valor)

def extraer_valor_float(df, nombre_celda, n_muestra):
    valor = extraer_valor(df, nombre_celda, n_muestra)
    if valor in ["-----", "#########"]:
        return valor
    if isinstance(valor, str) and (valor.startswith(">") or valor.startswith("<") or valor.startswith("&gt;") or valor.startswith("&lt;")):
        return valor
    try:
        return float(str(valor).replace(",", "."))
    except ValueError:
        return "#########"

def formatear(valor, decimales):
    if valor in ["#########", "-----"]:
        return valor
    if isinstance(valor, float):
        return f"{valor:.{decimales}f}".replace(".", ",")
    return str(valor)

def extraer_datos(df, nombre_celda, n_muestra):
    """
    Busca en columna 0 el nombre_celda y devuelve valor en col (3+n_muestra).
    """
    fila = df[df[0] == nombre_celda]
    if fila.empty:
        return "#########"  # El análisis no existe
    valor = fila.iloc[0, 3 + n_muestra]
    if pd.isna(valor) or str(valor).strip() == "":
        return "-----"  # Existe pero sin dato
    return str(valor)

def formatear_fecha_datos(fecha_str, incluir_hora=True):
    """
    Convierte fecha LIMS: "%d/%b/%Y %H:%M:%S" -> dd/mm/yyyy (con o sin hora)
    """
    if fecha_str in ["-----", "#########"]:
        return fecha_str
    try:
        dt = datetime.strptime(fecha_str, "%d/%b/%Y %H:%M:%S")
        return dt.strftime("%d/%m/%Y %H:%M") if incluir_hora else dt.strftime("%d/%m/%Y")
    except ValueError:
        return "#########"

def extraer_valor_norma(df, nombre):
    """
    Normas generales están en columna 2, con clave en columna 1.
    """
    resultado = df.loc[df[1] == nombre, 2]
    return resultado.values[0] if not resultado.empty else ""

def construir_numero_comunicado(nombre_base: str) -> str:
    """
    Usa el nombre del CSV (sin .csv) para formar el comunicado.
    Ej:
      "cm 003-26"  -> "C.M. 003-26"
      "CM003-26"   -> "C.M. 003-26"
      "cm_003-26"  -> "C.M. 003-26"
      "ac 120-26"  -> "A.C. 120-26"
    Si no matchea, devuelve el nombre tal cual.
    """
    if not nombre_base:
        return "#########"

    nb = nombre_base.strip()

    # Prefijo AC/CM al inicio, tolerando espacios/guiones/underscore/puntos
    m = re.match(r"^(ac|cm)\s*[-_\.]?\s*(.+)$", nb, flags=re.IGNORECASE)
    if not m:
        return nb

    pref = m.group(1).upper()
    resto = m.group(2).strip()

    if pref == "AC":
        return f"A.C. {resto}"
    else:
        return f"C.M. {resto}"

def sanitizar_nombre_archivo(nombre: str) -> str:
    """
    Evita caracteres inválidos en nombres de archivo Windows.
    """
    return re.sub(r'[\\/*?:"<>|]', "-", nombre).strip()


# -----------------------------
# UI
# -----------------------------
st.title("🛢️ GAS OIL UTE | Generador de informe Asistencia Comercial")

uploaded = st.file_uploader("📄 Cargar CSV de LIMS", type=["csv"], accept_multiple_files=False)


if uploaded is None:
    st.stop()

# Nombre base del archivo CSV (sin extensión)
nombre_archivo_LIMS = os.path.splitext(uploaded.name)[0]
numero_comunicado = construir_numero_comunicado(nombre_archivo_LIMS)

st.divider()

# ORDEN QUE PEDISTE:
cliente = st.text_input("Nombre del cliente", value="")

col1, col2 = st.columns(2)
with col1:
    n_muestras = st.number_input("¿Cuántas muestras desea procesar? (1 a 6)", min_value=1, max_value=6, value=1, step=1)
with col2:
    prioridad = st.number_input("Prioridad (días hábiles)", min_value=1, max_value=10, value=1, step=1)

st.divider()

# Leer CSV (con fallback de encoding por si acaso)
try:
    df_a = pd.read_csv(uploaded, encoding="latin1", sep=";", header=None)
except UnicodeDecodeError:
    df_a = pd.read_csv(uploaded, encoding="utf-8", sep=";", header=None)
except Exception as e:
    st.error(f"❌ No pude leer el CSV: {e}")
    st.stop()

# Plantillas en la raíz del repo (mismo nivel que este .py)
plantilla = os.path.join(os.getcwd(), f"GASOIL UTE {int(n_muestras)}M.docx")

if not os.path.isfile(plantilla):
    st.error(
        f"❌ No encuentro la plantilla **{os.path.basename(plantilla)}**.\n\n"
        "📌 Verificá que esté en el repositorio, en el mismo nivel que este archivo `.py`."
    )
    st.stop()

# -----------------------------
# Extracción por muestra
# -----------------------------
datos_muestras = {}

for i in range(1, int(n_muestras) + 1):

    datos_muestras[f"lims_m{i}"] = extraer_datos(
        df_a,
        "Número de Muestra",
        i
    )

    datos_muestras["motivo"] = extraer_datos(
        df_a,
        "Observaciones",
        1
    )

    datos_muestras[f"fecha_ext_m{i}"] = formatear_fecha_datos(
        extraer_datos(
            df_a,
            "Fecha de Extracción",
            i
        ),
        incluir_hora=False
    )

    datos_muestras[f"lug_ext_m{i}"] = extraer_datos(
        df_a,
        "Lugar de Extracción",
        i
    )

    datos_muestras[f"vol_m{i}"] = extraer_datos(
        df_a,
        "Volumen de Muestra",
        i
    ).replace(".", ",")

    datos_muestras[f"asp_GO_m{i}"] = extraer_valor(
        df_a,
        "Aspecto GO",
        i
    )

    datos_muestras[f"temp_m{i}"] = formatear(
        extraer_valor_float(
            df_a,
            "Temperatura",
            i
        ),
        1
    )

    datos_muestras[f"color_vis_m{i}"] = extraer_valor(
        df_a,
        "Color",
        i
    )

    datos_muestras[f"fase_m{i}"] = extraer_valor(
        df_a,
        "Fase No Miscible",
        i
    )

    datos_muestras[f"particulas_m{i}"] = extraer_valor(
        df_a,
        "Partículas",
        i
    )

    datos_muestras[f"color_m{i}"] = extraer_valor(
        df_a,
        "Color ASTM 1500",
        i
    )

    datos_muestras[f"est_ox_m{i}"] = extraer_valor(
        df_a,
        "Estabilidad a la Oxidación (g/m3)",
        i
    )

    datos_muestras[f"POFF_m{i}"] = extraer_valor(
        df_a,
        "Punto de Obturación de Filtro Frío",
        i
    )

    datos_muestras[f"FBT_m{i}"] = formatear(
        extraer_valor_float(
            df_a,
            "FBT",
            i
        ),
        2
    )

    # Densidad

    valor_densidad = extraer_valor_float(
        df_a,
        "Densidad a 15ºC",
        i
    )

    if isinstance(valor_densidad, float):

        datos_muestras[f"dens_15_m{i}"] = formatear(
            valor_densidad * 1000,
            1
        )

    else:

        datos_muestras[f"dens_15_m{i}"] = valor_densidad

    # Particulado

    valor_particulado = extraer_valor_float(
        df_a,
        "Particulado",
        i
    )

    valor_vol_filt = extraer_valor_float(
        df_a,
        "Volumen Filtrado",
        i
    )

    if isinstance(valor_particulado, float):

        datos_muestras[f"particulado_m{i}"] = (
            formatear(valor_particulado, 2)
            + "\n(Vol. Filtrado: "
            + formatear(valor_vol_filt, 2)
            + " L)"
        )

    else:

        datos_muestras[f"particulado_m{i}"] = "-----"

datos_muestras["fecha_rec"] = formatear_fecha_datos(
    extraer_datos(
        df_a,
        "Fecha de Recepción",
        1
    ),
    incluir_hora=True
)

producto = extraer_datos(df_a, "Producto", 1)

nombre_archivo_nuevo = sanitizar_nombre_archivo(
    f"{numero_comunicado} Gas Oil UTE - {cliente}"
)


# Datos para merge
datos_fusion = {

    "num_com": numero_comunicado,
    "cliente": cliente,
    "prioridad": prioridad,

    "n_dens": str(
        extraer_valor_norma(
            df_a,
            "Densidad a 15ºC"
        )
    ).replace("_", " ")
    if extraer_valor_norma(df_a, "Densidad a 15ºC")
    else "#########",

    "n_asp_GO": str(
        extraer_valor_norma(
            df_a,
            "Aspecto GO"
        )
    ).replace("_", " ")
    if extraer_valor_norma(df_a, "Aspecto GO")
    else "#########",

    "n_color": str(
        extraer_valor_norma(
            df_a,
            "Color ASTM 1500"
        )
    ).replace("_", " ")
    if extraer_valor_norma(df_a, "Color ASTM 1500")
    else "#########",

    "n_est_ox": str(
        extraer_valor_norma(
            df_a,
            "Estabilidad a la Oxidación (g/m3)"
        )
    ).replace("_", " ")
    if extraer_valor_norma(
        df_a,
        "Estabilidad a la Oxidación (g/m3)"
    )
    else "#########",

    "n_part": str(
        extraer_valor_norma(
            df_a,
            "Particulado"
        )
    ).replace("_", " ")
    if extraer_valor_norma(df_a, "Particulado")
    else "#########",

    "n_POFF": str(
        extraer_valor_norma(
            df_a,
            "Punto de Obturación de Filtro Frío"
        )
    ).replace("_", " ")
    if extraer_valor_norma(
        df_a,
        "Punto de Obturación de Filtro Frío"
    )
    else "#########",

    "n_FBT": str(
        extraer_valor_norma(
            df_a,
            "FBT"
        )
    ).replace("_", " ")
    if extraer_valor_norma(df_a, "FBT")
    else "#########",
}

datos_fusion.update(datos_muestras)

# -----------------------------
# Generación del Word
# -----------------------------

if st.button("📝 Generar informe Word", type="primary"):
    try:
        doc = MailMerge(plantilla)
        doc.merge(**{k: str(v) for k, v in datos_fusion.items()})

        buffer = BytesIO()
        doc.write(buffer)
        buffer.seek(0)
        
        st.download_button(
            "⬇️ Descargar informe .docx",
            data=buffer,
            file_name=f"{nombre_archivo_nuevo}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except Exception as e:
        st.error(f"❌ Error generando el Word: {e}")


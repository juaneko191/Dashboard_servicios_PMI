
import streamlit as st
import pandas as pd
from pathlib import Path

# =====================================================
# CONFIGURACIÓN DE PÁGINA
# =====================================================
st.set_page_config(
    page_title="Consulta de inversiones y servicios priorizados en el PMI",
    layout="wide"
)

# =====================================================
# RUTA DONDE ESTÁN LOS EXCEL
# =====================================================
ruta_base = Path(__file__).parent

archivo_opmi = ruta_base / "DATA_OPMI.xlsx"
archivo_pmi = ruta_base / "PMI_CARTERA_BRECHAS_2026.03.09_MOD.xlsx"

# =====================================================
# CARGA DE DATOS
# =====================================================
@st.cache_data
def cargar_datos():
    df_opmi = pd.read_excel(archivo_opmi, engine="openpyxl")
    df_pmi = pd.read_excel(archivo_pmi, engine="openpyxl")

    # Limpiar nombres de columnas
    df_opmi.columns = df_opmi.columns.str.strip()
    df_pmi.columns = df_pmi.columns.str.strip()

    # Limpiar textos vacíos
    for col in df_opmi.select_dtypes(include="object").columns:
        df_opmi[col] = df_opmi[col].astype(str).str.strip()
        df_opmi[col] = df_opmi[col].replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})

    for col in df_pmi.select_dtypes(include="object").columns:
        df_pmi[col] = df_pmi[col].astype(str).str.strip()
        df_pmi[col] = df_pmi[col].replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})

    # Asegurar tipos de datos para el enlace
    df_opmi["UNIDAD"] = pd.to_numeric(df_opmi["UNIDAD"], errors="coerce")
    df_pmi["ID OPMI"] = pd.to_numeric(df_pmi["ID OPMI"], errors="coerce")

    # Unir tablas
    df = df_pmi.merge(
        df_opmi,
        left_on="ID OPMI",
        right_on="UNIDAD",
        how="left",
        suffixes=("", "_OPMI")
    )

    # Convertir campos numéricos
    columnas_numericas = [
        "PMI2026",
        "PMI2027",
        "PMI2028",
        "PIM 2026",
        "DEVENGADO 2026",
        "ORDEN"
    ]

    for col in columnas_numericas:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df



with st.spinner("Cargando datos, por favor espere..."):
    df = cargar_datos()

st.success("Datos cargados correctamente")


# =====================================================
# FUNCIÓN PARA ORDENAR
# Primero ORDEN ascendente, luego CONDICIONINVERSION ascendente
# Los vacíos van al final
# =====================================================
def ordenar_tabla(dataframe):
    return (
        dataframe.assign(
            TIENE_VACIO=dataframe["ORDEN"].isna() | dataframe["CONDICIONINVERSION"].isna(),
            ORDEN_AUX=dataframe["ORDEN"].fillna(999999),
            CONDICION_AUX=dataframe["CONDICIONINVERSION"].fillna("ZZZ")
        )
        .sort_values(
            by=["TIENE_VACIO", "ORDEN_AUX", "CONDICION_AUX"],
            ascending=[True, True, True]
        )
        .drop(columns=["TIENE_VACIO", "ORDEN_AUX", "CONDICION_AUX"])
    )

# =====================================================
# FORMATO DE NÚMEROS
# Separador de miles con coma y sin decimales
# =====================================================
def formato_entero(valor):
    if pd.isna(valor):
        return "0"
    return f"{valor:,.0f}"

# =====================================================
# TÍTULO
# =====================================================
st.markdown(
    """
    <div style="background-color:#E6E6E6; padding:18px; border-radius:4px;">
        <h1 style="text-align:center; margin:0; font-size:30px;">
            Consulta de inversiones y servicios priorizados en el PMI
        </h1>
    </div>
    """,
    unsafe_allow_html=True
)

st.write("")

# =====================================================
# FILTROS
# =====================================================
# =====================================================
# FILTROS DEPENDIENTES + BOTÓN LIMPIAR FILTROS
# =====================================================

def opciones_filtro(dataframe, columna):
    valores = dataframe[columna].dropna().unique().tolist()
    valores = sorted(valores)
    return ["All"] + valores


def limpiar_filtros():
    st.session_state["filtro_nivel"] = "All"
    st.session_state["filtro_departamento"] = "All"
    st.session_state["filtro_provincia"] = "All"
    st.session_state["filtro_distrito"] = "All"
    st.session_state["filtro_opmi"] = "All"


# Inicializar valores por defecto
for key in [
    "filtro_nivel",
    "filtro_departamento",
    "filtro_provincia",
    "filtro_distrito",
    "filtro_opmi"
]:
    if key not in st.session_state:
        st.session_state[key] = "All"


# Botón limpiar filtros
st.button(
    "Limpiar filtros",
    on_click=limpiar_filtros
)


# =========================
# NIVEL DE GOBIERNO
# =========================
opciones_nivel = opciones_filtro(df, "NIVEL DE GOBIERNO")

if st.session_state["filtro_nivel"] not in opciones_nivel:
    st.session_state["filtro_nivel"] = "All"


# =========================
# DEPARTAMENTO depende de NIVEL
# =========================
df_base_departamento = df.copy()

if st.session_state["filtro_nivel"] != "All":
    df_base_departamento = df_base_departamento[
        df_base_departamento["NIVEL DE GOBIERNO"] == st.session_state["filtro_nivel"]
    ]

opciones_departamento = opciones_filtro(df_base_departamento, "DEPARTAMENTO")

if st.session_state["filtro_departamento"] not in opciones_departamento:
    st.session_state["filtro_departamento"] = "All"


# =========================
# PROVINCIA depende de NIVEL + DEPARTAMENTO
# =========================
df_base_provincia = df_base_departamento.copy()

if st.session_state["filtro_departamento"] != "All":
    df_base_provincia = df_base_provincia[
        df_base_provincia["DEPARTAMENTO"] == st.session_state["filtro_departamento"]
    ]

opciones_provincia = opciones_filtro(df_base_provincia, "PROVINCIA")

if st.session_state["filtro_provincia"] not in opciones_provincia:
    st.session_state["filtro_provincia"] = "All"


# =========================
# DISTRITO depende de NIVEL + DEPARTAMENTO + PROVINCIA
# =========================
df_base_distrito = df_base_provincia.copy()

if st.session_state["filtro_provincia"] != "All":
    df_base_distrito = df_base_distrito[
        df_base_distrito["PROVINCIA"] == st.session_state["filtro_provincia"]
    ]

opciones_distrito = opciones_filtro(df_base_distrito, "DISTRITO")

if st.session_state["filtro_distrito"] not in opciones_distrito:
    st.session_state["filtro_distrito"] = "All"


# =========================
# OPMI depende de NIVEL + DEPARTAMENTO + PROVINCIA + DISTRITO
# =========================
df_base_opmi = df_base_distrito.copy()

if st.session_state["filtro_distrito"] != "All":
    df_base_opmi = df_base_opmi[
        df_base_opmi["DISTRITO"] == st.session_state["filtro_distrito"]
    ]

opciones_opmi = opciones_filtro(df_base_opmi, "OPMI")

if st.session_state["filtro_opmi"] not in opciones_opmi:
    st.session_state["filtro_opmi"] = "All"


# =========================
# MOSTRAR FILTROS
# =========================
col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    st.selectbox(
        "NIVEL DE GOBIERNO",
        opciones_nivel,
        key="filtro_nivel"
    )

with col2:
    st.selectbox(
        "DEPARTAMENTO",
        opciones_departamento,
        key="filtro_departamento"
    )

with col3:
    st.selectbox(
        "PROVINCIA",
        opciones_provincia,
        key="filtro_provincia"
    )

with col4:
    st.selectbox(
        "DISTRITO",
        opciones_distrito,
        key="filtro_distrito"
    )

with col5:
    st.selectbox(
        "OPMI",
        opciones_opmi,
        key="filtro_opmi"
    )


# =====================================================
# APLICAR FILTROS
# =====================================================
df_filtrado = df.copy()

if st.session_state["filtro_nivel"] != "All":
    df_filtrado = df_filtrado[
        df_filtrado["NIVEL DE GOBIERNO"] == st.session_state["filtro_nivel"]
    ]

if st.session_state["filtro_departamento"] != "All":
    df_filtrado = df_filtrado[
        df_filtrado["DEPARTAMENTO"] == st.session_state["filtro_departamento"]
    ]

if st.session_state["filtro_provincia"] != "All":
    df_filtrado = df_filtrado[
        df_filtrado["PROVINCIA"] == st.session_state["filtro_provincia"]
    ]

if st.session_state["filtro_distrito"] != "All":
    df_filtrado = df_filtrado[
        df_filtrado["DISTRITO"] == st.session_state["filtro_distrito"]
    ]

if st.session_state["filtro_opmi"] != "All":
    df_filtrado = df_filtrado[
        df_filtrado["OPMI"] == st.session_state["filtro_opmi"]
    ]

# Ordenar tabla filtrada
df_filtrado = ordenar_tabla(df_filtrado)

# =====================================================
# KPIs
# =====================================================
total_pmi_2026 = df_filtrado["PMI2026"].sum()
total_pim_2026 = df_filtrado["PIM 2026"].sum()
total_devengado_2026 = df_filtrado["DEVENGADO 2026"].sum()

kpi1, kpi2, kpi3 = st.columns(3)

with kpi1:
    st.metric("PMI 2026", formato_entero(total_pmi_2026))

with kpi2:
    st.metric("PIM 2026", formato_entero(total_pim_2026))

with kpi3:
    st.metric("Devengado 2026", formato_entero(total_devengado_2026))

st.write("")

# =====================================================
# COLUMNAS A MOSTRAR EN TABLA
# =====================================================
columnas_tabla = [
    "ID OPMI",
    "CODIGOUNICO",
    "CODIGOIDEA",
    "NOMBREINVERSION",
    "PMI2026",
    "PMI2027",
    "PMI2028",
    "PMI2029",
    "PIM 2026",
    "DEVENGADO 2026",
    "ID_SERVICIO",
    "SERVICIO",
    "ORDEN",
    "CONDICIONINVERSION"
]

columnas_existentes = [col for col in columnas_tabla if col in df_filtrado.columns]

tabla_mostrar = df_filtrado[columnas_existentes].copy()

# =====================================================
# ALIAS DE COLUMNAS PARA MOSTRAR EN LA TABLA
# =====================================================
alias_columnas = {
    "ID OPMI": "ID OPMI",
    "CODIGOUNICO": "CUI",
    "CODIGOIDEA": "Código Idea",
    "NOMBREINVERSION": "Nombre de la inversión",
    "PMI2026": "PMI 2026",
    "PMI2027": "PMI 2027",
    "PMI2028": "PMI 2028",
    "PMI2029": "PMI 2029",
    "PIM 2026": "PIM 2026",
    "DEVENGADO 2026": "Devengado 2026",
    "ID_SERVICIO": "ID Servicio",
    "SERVICIO": "Servicio",
    "ORDEN": "Prioridad del servicio",
    "CONDICIONINVERSION": "Prelación"
}

tabla_mostrar = tabla_mostrar.rename(columns=alias_columnas)

# =====================================================
# MOSTRAR TABLA
# =====================================================
st.subheader("Cartera de inversiones")

st.dataframe(
    tabla_mostrar,
    use_container_width=True,
    hide_index=True
)

from io import BytesIO

# =====================================================
# BOTÓN DE DESCARGA EN EXCEL
# =====================================================
output = BytesIO()

with pd.ExcelWriter(output, engine="openpyxl") as writer:
    tabla_mostrar.to_excel(
        writer,
        index=False,
        sheet_name="Tabla filtrada"
    )

excel_data = output.getvalue()

st.download_button(
    label="Descargar tabla filtrada en Excel",
    data=excel_data,
    file_name="tabla_filtrada_pmi.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

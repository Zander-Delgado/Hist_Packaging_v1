import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

st.set_page_config(page_title="Strategic Procurement Analytics", page_icon="📊", layout="wide")

# =========================================================
# FUNCIONES AUXILIARES DE PROCESAMIENTO
# =========================================================

def asignar_temporada(fecha):
    """Asigna la temporada del 01 de Julio al 30 de Junio del siguiente año."""
    if pd.isna(fecha):
        return "Sin Fecha"
    anio = fecha.year
    mes = fecha.month
    if mes >= 7:
        return f"T{str(anio)[2:]}{str(anio + 1)[2:]}"
    else:
        return f"T{str(anio - 1)[2:]}{str(anio)[2:]}"

@st.cache_data
def cargar_y_limpiar_datos(archivo_subido):
    # Detectar y leer la hoja que contenga la columna 'Material'
    xls = pd.ExcelFile(archivo_subido)
    hoja_correcta = None
    for sheet in xls.sheet_names:
        df_temp = pd.read_excel(xls, sheet_name=sheet, nrows=5)
        if "Material" in df_temp.columns:
            hoja_correcta = sheet
            break
    
    if hoja_correcta is None:
        hoja_correcta = xls.sheet_names[0]
        
    df = pd.read_excel(xls, sheet_name=hoja_correcta)
    
    # 1. Limpieza de órdenes archivadas o no ejecutadas
    df = df[~df["Indicador de borrado"].astype(str).str.upper().isin(["S", "L"])].copy()
    
    # 2. Tipos de datos y fechas
    df["Fecha documento"] = pd.to_datetime(df["Fecha documento"], errors="coerce")
    df = df.dropna(subset=["Fecha documento"])
    
    # 3. Asignación de temporada
    df["Temporada"] = df["Fecha documento"].apply(asignar_temporada)
    
    # 4. Asegurar campos numéricos
    df["Cantidad de pedido"] = pd.to_numeric(df["Cantidad de pedido"], errors="coerce").fillna(0)
    df["Valor neto de pedido"] = pd.to_numeric(df["Valor neto de pedido"], errors="coerce").fillna(0)
    
    # 5. Precio unitario real considerando la cantidad base de SAP
    if "Cantidad base" in df.columns:
        df["Cantidad base"] = pd.to_numeric(df["Cantidad base"], errors="coerce").fillna(1)
        df["Cantidad base"] = np.where(df["Cantidad base"] <= 0, 1, df["Cantidad base"])
        if "Precio neto" in df.columns:
            df["Precio_Unitario_Real"] = df["Precio neto"] / df["Cantidad base"]
        else:
            df["Precio_Unitario_Real"] = np.where(df["Cantidad de pedido"] > 0, df["Valor neto de pedido"] / df["Cantidad de pedido"], 0)
    else:
        df["Precio_Unitario_Real"] = np.where(df["Cantidad de pedido"] > 0, df["Valor neto de pedido"] / df["Cantidad de pedido"], 0)
        
    return df

# =========================================================
# INTERFAZ Y BARRA LATERAL (CARGA Y FILTROS)
# =========================================================

st.title("📦 Strategic Sourcing Analytics - Compras Históricas")
st.markdown("Herramienta corporativa para análisis de demanda, clasificación ABC, precios históricos y benchmarking de licitaciones.")

st.sidebar.header("📁 Origen de Datos")
archivo = st.sidebar.file_uploader("Cargar reporte SAP (.xlsx)", type=["xlsx"])

if archivo is None:
    st.info("👈 Por favor, cargue un archivo Excel descargado de SAP en el menú lateral para iniciar el análisis.")
    st.stop()

df_clean = cargar_y_limpiar_datos(archivo)

st.sidebar.header("🔍 Filtros de Segmentación")

# Filtro: Organización de Compras (Filial)
filiales = sorted(df_clean["Organización compras"].dropna().unique().tolist())
filial_sel = st.sidebar.multiselect("Organización de Compras:", filiales, default=filiales)

# Filtro: Categoría
if "Categoría" in df_clean.columns:
    categorias = sorted(df_clean["Categoría"].dropna().unique().tolist())
    cat_sel = st.sidebar.multiselect("Categoría:", categorias, default=categorias)
else:
    cat_sel = []

# Filtro: Centro (Almacenes)
centros = sorted(df_clean["Centro"].dropna().unique().tolist())
centro_sel = st.sidebar.multiselect("Centro / Almacén:", centros, default=centros)

# Filtro: Temporada
temporadas_disp = sorted(df_clean["Temporada"].unique().tolist())
temp_sel = st.sidebar.multiselect("Temporada:", temporadas_disp, default=temporadas_disp)

# Aplicación de filtros
df_filtrado = df_clean[
    (df_clean["Organización compras"].isin(filial_sel)) &
    (df_clean["Centro"].isin(centro_sel)) &
    (df_clean["Temporada"].isin(temp_sel))
]

if "Categoría" in df_clean.columns and cat_sel:
    df_filtrado = df_filtrado[df_filtrado["Categoría"].isin(cat_sel)]

if df_filtrado.empty:
    st.warning("No se encontraron registros para los filtros seleccionados.")
    st.stop()

# =========================================================
# RESUMEN DE INDICADORES CLAVE
# =========================================================

st.markdown("### Resumen de la Selección")
m1, m2, m3, m4 = st.columns(4)
m1.metric("Órdenes Válidas (Líneas PO)", f"{len(df_filtrado):,}")
m2.metric("Volumen Total", f"{df_filtrado['Cantidad de pedido'].sum():,.0f} u")
m3.metric("Gasto Total", f"{df_filtrado['Valor neto de pedido'].sum():,.2f}")
m4.metric("Materiales Únicos", f"{df_filtrado['Material'].nunique():,}")

st.divider()

# =========================================================
# MÓDULOS DE ANÁLISIS DE NEGOCIO
# =========================================================

tabs = st.tabs([
    "1 & 2. Materiales, Volumen y Gasto",
    "3. Clasificación ABC (Pareto)",
    "4. Recurrencia de Compra",
    "5. Variación Interanual (YoY)",
    "6. Priorización para Licitación",
    "7. Análisis de Precios",
    "8. Ranking de Proveedores"
])

# ---------------------------------------------------------
# PESTAÑA 1 & 2: MATERIALES, VOLUMEN E IMPORTE
# ---------------------------------------------------------
with tabs[0]:
    st.subheader("1 & 2. ¿Qué materiales se compraron por temporada, en qué volumen e importe?")
    
    agrup_mat = df_filtrado.groupby(["Temporada", "Material", "Texto breve"]).agg(
        Volumen_Total=("Cantidad de pedido", "sum"),
        Gasto_Total=("Valor neto de pedido", "sum"),
        Nro_Pedidos=("Documento compras", "nunique")
    ).reset_index().sort_values(by=["Temporada", "Gasto_Total"], ascending=[True, False])
    
    st.dataframe(
        agrup_mat.style.format({
            "Volumen_Total": "{:,.0f}",
            "Gasto_Total": "{:,.2f}",
            "Nro_Pedidos": "{:,.0f}"
        }),
        use_container_width=True
    )

# ---------------------------------------------------------
# PESTAÑA 3: CLASIFICACIÓN ABC (PARETO)
# ---------------------------------------------------------
with tabs[1]:
    st.subheader("3. Clasificación ABC de Materiales por Gasto")
    
    abc_df = df_filtrado.groupby(["Material", "Texto breve"]).agg(
        Gasto_Total=("Valor neto de pedido", "sum"),
        Volumen_Total=("Cantidad de pedido", "sum")
    ).reset_index().sort_values(by="Gasto_Total", ascending=False)
    
    total_gasto = abc_df["Gasto_Total"].sum()
    abc_df["Participacion"] = (abc_df["Gasto_Total"] / total_gasto) * 100
    abc_df["Participacion_Acum"] = abc_df["Participacion"].cumsum()
    
    # Clasificación de Pareto estándar: A (80%), B (15%), C (5%)
    condiciones = [
        abc_df["Participacion_Acum"] <= 80,
        abc_df["Participacion_Acum"] <= 95
    ]
    abc_df["Clase_ABC"] = np.select(condiciones, ["A", "B"], default="C")
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Materiales Clase A (Gasto Crítico ~80%)", f"{(abc_df['Clase_ABC'] == 'A').sum()} items")
    c2.metric("Materiales Clase B (Gasto Medio ~15%)", f"{(abc_df['Clase_ABC'] == 'B').sum()} items")
    c3.metric("Materiales Clase C (Cola Larga ~5%)", f"{(abc_df['Clase_ABC'] == 'C').sum()} items")
    
    st.dataframe(
        abc_df.style.format({
            "Gasto_Total": "{:,.2f}",
            "Volumen_Total": "{:,.0f}",
            "Participacion": "{:.2f}%",
            "Participacion_Acum": "{:.2f}%"
        }),
        use_container_width=True
    )

# ---------------------------------------------------------
# PESTAÑA 4: RECURRENCIA DE COMPRAS
# ---------------------------------------------------------
with tabs[2]:
    st.subheader("4. Materiales Comprados con Mayor Frecuencia (Nro. de Órdenes)")
    
    recurrencia = df_filtrado.groupby(["Material", "Texto breve"]).agg(
        Frecuencia_PO=("Documento compras", "nunique"),
        Lineas_Registradas=("Documento compras", "count"),
        Volumen_Total=("Cantidad de pedido", "sum")
    ).reset_index().sort_values(by="Frecuencia_PO", ascending=False)
    
    top_10_rec = recurrencia.head(10)
    
    fig_rec, ax_rec = plt.subplots(figsize=(10, 4.5))
    sns.barplot(data=top_10_rec, y="Texto breve", x="Frecuencia_PO", palette="Blues_r", ax=ax_rec)
    ax_rec.set_title("Top 10 Materiales con Más Órdenes de Compra Emitidas")
    ax_rec.set_xlabel("Cantidad de Órdenes Únicas (PO)")
    ax_rec.set_ylabel("Material")
    st.pyplot(fig_rec)
    
    st.dataframe(
        recurrencia.style.format({
            "Frecuencia_PO": "{:,.0f}",
            "Lineas_Registradas": "{:,.0f}",
            "Volumen_Total": "{:,.0f}"
        }),
        use_container_width=True
    )

# ---------------------------------------------------------
# PESTAÑA 5: VARIACIÓN INTERANUAL (YoY)
# ---------------------------------------------------------
with tabs[3]:
    st.subheader("5. Variación Interanual de Compras (Volumen y Valor)")
    
    temps = sorted(df_filtrado["Temporada"].unique().tolist())
    if len(temps) < 2:
        st.info(f"Se requieren al menos 2 temporadas para calcular la variación interanual. Temporadas detectadas: {temps}")
    else:
        t1, t2 = st.columns(2)
        temp_base = t1.selectbox("Temporada Base (T1):", temps, index=0)
        temp_comp = t2.selectbox("Temporada a Comparar (T2):", temps, index=len(temps)-1)
        
        pivot_vol = df_filtrado.pivot_table(
            index=["Material", "Texto breve"],
            columns="Temporada",
            values="Cantidad de pedido",
            aggfunc="sum",
            fill_value=0
        )
        
        pivot_val = df_filtrado.pivot_table(
            index=["Material", "Texto breve"],
            columns="Temporada",
            values="Valor neto de pedido",
            aggfunc="sum",
            fill_value=0
        )
        
        df_yoy = pd.DataFrame(index=pivot_vol.index)
        df_yoy[f"Vol_{temp_base}"] = pivot_vol[temp_base]
        df_yoy[f"Vol_{temp_comp}"] = pivot_vol[temp_comp]
        df_yoy["Var_Vol_%"] = np.where(
            df_yoy[f"Vol_{temp_base}"] > 0,
            ((df_yoy[f"Vol_{temp_comp}"] - df_yoy[f"Vol_{temp_base}"]) / df_yoy[f"Vol_{temp_base}"]) * 100,
            np.nan
        )
        
        df_yoy[f"Gasto_{temp_base}"] = pivot_val[temp_base]
        df_yoy[f"Gasto_{temp_comp}"] = pivot_val[temp_comp]
        df_yoy["Var_Gasto_%"] = np.where(
            df_yoy[f"Gasto_{temp_base}"] > 0,
            ((df_yoy[f"Gasto_{temp_comp}"] - df_yoy[f"Gasto_{temp_base}"]) / df_yoy[f"Gasto_{temp_base}"]) * 100,
            np.nan
        )
        
        df_yoy = df_yoy.reset_index().sort_values(by=f"Gasto_{temp_comp}", ascending=False)
        
        st.dataframe(
            df_yoy.style.format({
                f"Vol_{temp_base}": "{:,.0f}",
                f"Vol_{temp_comp}": "{:,.0f}",
                "Var_Vol_%": "{:+.2f}%",
                f"Gasto_{temp_base}": "{:,.2f}",
                f"Gasto_{temp_comp}": "{:,.2f}",
                "Var_Gasto_%": "{:+.2f}%"
            }),
            use_container_width=True
        )

# ---------------------------------------------------------
# PESTAÑA 6: PRIORIZACIÓN PARA LICITACIÓN
# ---------------------------------------------------------
with tabs[4]:
    st.subheader("6. Matriz de Priorización de Materiales para Licitación")
    st.markdown("""
    Identificación de ítems candidatos a licitación considerando:
    * **Clasificación ABC:** Foco en ítems Clase A y B.
    * **Volumen Consolidado:** Concentración de demanda suficiente para negociar economías de escala.
    * **Multi-Filial:** Demanda compartida en más de una Organización de Compras.
    """)
    
    licit_df = df_filtrado.groupby(["Material", "Texto breve"]).agg(
        Gasto_Total=("Valor neto de pedido", "sum"),
        Volumen_Total=("Cantidad de pedido", "sum"),
        Organizaciones_Compradoras=("Organización compras", "nunique"),
        Proveedores_Actuales=("Proveedor/Centro suministrador", "nunique")
    ).reset_index().sort_values(by="Gasto_Total", ascending=False)
    
    total_gasto_lic = licit_df["Gasto_Total"].sum()
    licit_df["% Gasto Acum"] = (licit_df["Gasto_Total"] / total_gasto_lic * 100).cumsum()
    
    # Criterio de recomendación
    def recomendar_licitacion(row):
        if row["% Gasto Acum"] <= 80:
            return "🔥 Alta Prioridad (Clase A)"
        elif row["% Gasto Acum"] <= 95 and row["Organizaciones_Compradoras"] > 1:
            return "⚡ Prioridad Media (Corporativo Multi-Filial)"
        else:
            return "⚪ Baja Prioridad (Compra Spot / Local)"
            
    licit_df["Estrategia_Sourcing"] = licit_df.apply(recomendar_licitacion, axis=1)
    
    st.dataframe(
        licit_df.style.format({
            "Gasto_Total": "{:,.2f}",
            "Volumen_Total": "{:,.0f}",
            "% Gasto Acum": "{:.2f}%",
            "Organizaciones_Compradoras": "{:,.0f}",
            "Proveedores_Actuales": "{:,.0f}"
        }),
        use_container_width=True
    )

# ---------------------------------------------------------
# PESTAÑA 7: ANÁLISIS DE PRECIOS
# ---------------------------------------------------------
with tabs[5]:
    st.subheader("7. Comportamiento de Precios: Mínimo, Máximo y Promedio Ponderado")
    
    precios_df = df_filtrado[df_filtrado["Precio_Unitario_Real"] > 0].groupby(["Temporada", "Material", "Texto breve"]).agg(
        Precio_Min=("Precio_Unitario_Real", "min"),
        Precio_Max=("Precio_Unitario_Real", "max"),
        Precio_Prom_Aritmetico=("Precio_Unitario_Real", "mean"),
        Volumen_Total=("Cantidad de pedido", "sum"),
        Gasto_Total=("Valor neto de pedido", "sum")
    ).reset_index()
    
    precios_df["Precio_Prom_Ponderado"] = precios_df["Gasto_Total"] / precios_df["Volumen_Total"]
    precios_df["Dispersion_Precio_%"] = ((precios_df["Precio_Max"] - precios_df["Precio_Min"]) / precios_df["Precio_Min"]) * 100
    precios_df = precios_df.sort_values(by="Gasto_Total", ascending=False)
    
    st.dataframe(
        precios_df.style.format({
            "Precio_Min": "{:,.4f}",
            "Precio_Max": "{:,.4f}",
            "Precio_Prom_Aritmetico": "{:,.4f}",
            "Precio_Prom_Ponderado": "{:,.4f}",
            "Dispersion_Precio_%": "{:.1f}%",
            "Volumen_Total": "{:,.0f}",
            "Gasto_Total": "{:,.2f}"
        }),
        use_container_width=True
    )

# ---------------------------------------------------------
# PESTAÑA 8: RANKING DE PROVEEDORES
# ---------------------------------------------------------
with tabs[6]:
    st.subheader("8. Ranking de Proveedores por Gasto")
    
    ranking_prov = df_filtrado.groupby("Proveedor/Centro suministrador").agg(
        Gasto_Total=("Valor neto de pedido", "sum"),
        Volumen_Total=("Cantidad de pedido", "sum"),
        PO_Emitidas=("Documento compras", "nunique")
    ).reset_index().sort_values(by="Gasto_Total", ascending=False)
    
    total_gasto_prov = ranking_prov["Gasto_Total"].sum()
    ranking_prov["% Participacion"] = (ranking_prov["Gasto_Total"] / total_gasto_prov) * 100
    
    top_10_prov = ranking_prov.head(10)
    
    fig_prov, ax_prov = plt.subplots(figsize=(10, 5))
    sns.barplot(data=top_10_prov, y="Proveedor/Centro suministrador", x="Gasto_Total", palette="mako", ax=ax_prov)
    ax_prov.set_title("Top 10 Proveedores por Importe de Compra")
    ax_prov.set_xlabel("Gasto Total")
    ax_prov.set_ylabel("Proveedor")
    st.pyplot(fig_prov)
    
    st.dataframe(
        ranking_prov.style.format({
            "Gasto_Total": "{:,.2f}",
            "Volumen_Total": "{:,.0f}",
            "PO_Emitidas": "{:,.0f}",
            "% Participacion": "{:.2f}%"
        }),
        use_container_width=True
    )

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

def inferir_categoria(row):
    """
    Infiere la categoría analizando primero el prefijo del código Material (SAP)
    y luego el Texto breve como regla de respaldo.
    """
    material = str(row.get("Material", "")).strip().upper()
    texto = str(row.get("Texto breve", "")).strip().upper()

    prefijo = material.split("-")[0].strip() if "-" in material else material

    mapeo_prefijos = {
        # Packaging
        "PMPNT": "Punnets",
        "PMBOX": "Cajas",
        "PMLID": "Tapas",
        "PMCLA": "Clamshells",
        "PMCOR": "Esquineros",
        "PMPAL": "Pallets",
        "PMLAB": "Etiquetas",
        "PMBAG": "Bolsas",
        # Agrícola
        "FERT": "Fertilizantes",
        "PEST": "Pesticidas / Fitosanitarios",
        "ADDI": "Aditivos / Coadyuvantes",
        "STIM": "Bioestimulantes",
        "COND": "Acondicionadores de Suelo",
        "SANI": "Sanitizantes / Desinfección",
        "MUESTRA": "Muestras / Ensayos"
    }

    if prefijo in mapeo_prefijos:
        return mapeo_prefijos[prefijo]

    if any(k in texto for k in ["PUNNET", "TARRINA", "CESTA"]):
        return "Punnets"
    elif any(k in texto for k in ["CAJA", "BOX", "IFCO", "PLANCHA"]):
        return "Cajas"
    elif any(k in texto for k in ["TAPA", "LID"]):
        return "Tapas"
    elif any(k in texto for k in ["CLAMSHELL", "CLAM"]):
        return "Clamshells"
    elif any(k in texto for k in ["ESQUINERO", "CORNER"]):
        return "Esquineros"
    elif any(k in texto for k in ["NITRATO", "SULFATO", "FOSFATO", "POTASIO", "UREA", "ACIDO"]):
        return "Fertilizantes"
    elif any(k in texto for k in ["INSECTICIDA", "FUNGICIDA", "ACARICIDA", "HERBICIDA"]):
        return "Pesticidas / Fitosanitarios"
    
    return "Otros / Sin Categorizar"

@st.cache_data
def cargar_y_limpiar_datos(archivo_subido):
    xls = pd.ExcelFile(archivo_subido)
    hoja_correcta = None
    for sheet in xls.sheet_names:
        df_temp = pd.read_excel(xls, sheet_name=sheet, nrows=5)
        if "Material" in df_temp.columns or "Texto breve" in df_temp.columns:
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
    
    # 4. Homologación de Unidad de Medida
    if "Unidad medida pedido" in df.columns:
        df["Unidad medida pedido"] = df["Unidad medida pedido"].fillna("UN").astype(str).str.strip().str.upper()
    else:
        df["Unidad medida pedido"] = "UN"
        
    # 5. Inferencia de Categoría
    if "Categoría" not in df.columns or df["Categoría"].isna().all():
        df["Categoría"] = df.apply(inferir_categoria, axis=1)
    else:
        df["Categoría"] = df["Categoría"].fillna(df.apply(inferir_categoria, axis=1))
    
    # 6. Conversión a numéricos
    df["Cantidad de pedido"] = pd.to_numeric(df["Cantidad de pedido"], errors="coerce").fillna(0)
    df["Valor neto de pedido"] = pd.to_numeric(df["Valor neto de pedido"], errors="coerce").fillna(0)
    
    # 7. Precio Unitario Real considerando Cantidad Base de SAP
    if "Cantidad base" in df.columns:
        df["Cantidad base"] = pd.to_numeric(df["Cantidad base"], errors="coerce").fillna(1)
        df["Cantidad base"] = np.where(df["Cantidad base"] <= 0, 1, df["Cantidad base"])
        if "Precio neto" in df.columns:
            df["Precio_Unitario_Real"] = df["Precio neto"] / df["Cantidad base"]
        else:
            df["Precio_Unitario_Real"] = np.where(
                df["Cantidad de pedido"] > 0, 
                df["Valor neto de pedido"] / df["Cantidad de pedido"], 
                0
            )
    else:
        df["Precio_Unitario_Real"] = np.where(
            df["Cantidad de pedido"] > 0, 
            df["Valor neto de pedido"] / df["Cantidad de pedido"], 
            0
        )
        
    return df

# =========================================================
# INTERFAZ Y BARRA LATERAL (CARGA Y FILTROS)
# =========================================================

st.title("📦 Strategic Sourcing Analytics - Packaging & Agrícola")
st.markdown("Herramienta corporativa para análisis de demanda histórica, clasificación ABC, dispersión de precios y licitaciones.")

st.sidebar.header("📁 Carga de Archivo")
archivo = st.sidebar.file_uploader("Cargar reporte SAP (.xlsx)", type=["xlsx"])

if archivo is None:
    st.info("👈 Cargue un reporte Excel de compras de SAP en el menú lateral para iniciar.")
    st.stop()

df_clean = cargar_y_limpiar_datos(archivo)

st.sidebar.header("🔍 Filtros de Segmentación")

# Filtro: Organización de Compras
filiales = sorted(df_clean["Organización compras"].dropna().unique().tolist())
filial_sel = st.sidebar.multiselect("Organización de Compras:", filiales, default=filiales)

# Filtro: Categoría
categorias = sorted(df_clean["Categoría"].dropna().unique().tolist())
cat_sel = st.sidebar.multiselect("Categoría / Familia:", categorias, default=categorias)

# Filtro: Unidad de Medida
unidades = sorted(df_clean["Unidad medida pedido"].unique().tolist())
uom_sel = st.sidebar.multiselect("Unidad de Medida (UOM):", unidades, default=unidades)

# Filtro: Centro
centros = sorted(df_clean["Centro"].dropna().unique().tolist())
centro_sel = st.sidebar.multiselect("Centro / Almacén:", centros, default=centros)

# Filtro: Temporada
temporadas_disp = sorted(df_clean["Temporada"].unique().tolist())
temp_sel = st.sidebar.multiselect("Temporada:", temporadas_disp, default=temporadas_disp)

# Aplicación de filtros
df_filtrado = df_clean[
    (df_clean["Organización compras"].isin(filial_sel)) &
    (df_clean["Categoría"].isin(cat_sel)) &
    (df_clean["Unidad medida pedido"].isin(uom_sel)) &
    (df_clean["Centro"].isin(centro_sel)) &
    (df_clean["Temporada"].isin(temp_sel))
]

if df_filtrado.empty:
    st.warning("No se encontraron registros que cumplan con la combinación de filtros seleccionada.")
    st.stop()

# =========================================================
# INDICADORES GENERALES
# =========================================================

st.markdown("### Resumen Ejecutivo")
m1, m2, m3, m4 = st.columns(4)
m1.metric("Líneas de PO Válidas", f"{len(df_filtrado):,}")
m2.metric("Gasto Total Acumulado", f"{df_filtrado['Valor neto de pedido'].sum():,.2f}")
m3.metric("Materiales Distintos", f"{df_filtrado['Material'].nunique():,}")
m4.metric("Unidades de Medida", f"{', '.join(df_filtrado['Unidad medida pedido'].unique())}")

st.divider()

# =========================================================
# MÓDULOS DE ANÁLISIS DE NEGOCIO (PREGUNTAS 1 A 8)
# =========================================================

tabs = st.tabs([
    "1 & 2. Volumen y Gasto",
    "3. Clasificación ABC",
    "4. Recurrencia de Compra",
    "5. Variación Interanual (YoY)",
    "6. Priorización Licitación",
    "7. Análisis de Precios",
    "8. Ranking Proveedores"
])

# ---------------------------------------------------------
# 1 & 2. MATERIALES, VOLUMEN Y GASTO
# ---------------------------------------------------------
with tabs[0]:
    st.subheader("1 & 2. Materiales Comprados por Temporada, Volumen y Gasto")
    
    agrup_mat = df_filtrado.groupby(["Temporada", "Material", "Texto breve", "Unidad medida pedido"]).agg(
        Volumen_Total=("Cantidad de pedido", "sum"),
        Gasto_Total=("Valor neto de pedido", "sum"),
        Nro_Pedidos=("Documento compras", "nunique")
    ).reset_index().sort_values(by=["Temporada", "Gasto_Total"], ascending=[True, False])
    
    st.dataframe(
        agrup_mat.style.format({
            "Volumen_Total": "{:,.2f}",
            "Gasto_Total": "{:,.2f}",
            "Nro_Pedidos": "{:,.0f}"
        }),
        use_container_width=True
    )

# ---------------------------------------------------------
# 3. CLASIFICACIÓN ABC (PARETO DE GASTO)
# ---------------------------------------------------------
with tabs[1]:
    st.subheader("3. Clasificación ABC por Gasto Monetario")
    
    abc_df = df_filtrado.groupby(["Material", "Texto breve", "Unidad medida pedido"]).agg(
        Gasto_Total=("Valor neto de pedido", "sum"),
        Volumen_Total=("Cantidad de pedido", "sum")
    ).reset_index().sort_values(by="Gasto_Total", ascending=False)
    
    total_gasto = abc_df["Gasto_Total"].sum()
    abc_df["Participacion"] = (abc_df["Gasto_Total"] / total_gasto) * 100
    abc_df["Participacion_Acum"] = abc_df["Participacion"].cumsum()
    
    condiciones = [
        abc_df["Participacion_Acum"] <= 80,
        abc_df["Participacion_Acum"] <= 95
    ]
    abc_df["Clase_ABC"] = np.select(condiciones, ["A", "B"], default="C")
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Clase A (Gasto Crítico ~80%)", f"{(abc_df['Clase_ABC'] == 'A').sum()} ítems")
    c2.metric("Clase B (Gasto Medio ~15%)", f"{(abc_df['Clase_ABC'] == 'B').sum()} ítems")
    c3.metric("Clase C (Cola Larga ~5%)", f"{(abc_df['Clase_ABC'] == 'C').sum()} ítems")
    
    st.dataframe(
        abc_df.style.format({
            "Gasto_Total": "{:,.2f}",
            "Volumen_Total": "{:,.2f}",
            "Participacion": "{:.2f}%",
            "Participacion_Acum": "{:.2f}%"
        }),
        use_container_width=True
    )

# ---------------------------------------------------------
# 4. RECURRENCIA DE COMPRA
# ---------------------------------------------------------
with tabs[2]:
    st.subheader("4. Frecuencia y Recurrencia de Emisión de Órdenes")
    
    rec_df = df_filtrado.groupby(["Material", "Texto breve", "Unidad medida pedido"]).agg(
        Ordenes_Emitidas=("Documento compras", "nunique"),
        Lineas_PO=("Documento compras", "count"),
        Volumen_Total=("Cantidad de pedido", "sum")
    ).reset_index().sort_values(by="Ordenes_Emitidas", ascending=False)
    
    top_10 = rec_df.head(10)
    
    fig_rec, ax_rec = plt.subplots(figsize=(10, 4.5))
    sns.barplot(data=top_10, y="Texto breve", x="Ordenes_Emitidas", palette="Blues_r", ax=ax_rec)
    ax_rec.set_title("Top 10 Productos con Mayor Recurrencia de Pedidos (PO)")
    ax_rec.set_xlabel("Número de Órdenes Únicas")
    ax_rec.set_ylabel("Producto")
    st.pyplot(fig_rec)
    
    st.dataframe(
        rec_df.style.format({
            "Ordenes_Emitidas": "{:,.0f}",
            "Lineas_PO": "{:,.0f}",
            "Volumen_Total": "{:,.2f}"
        }),
        use_container_width=True
    )

# ---------------------------------------------------------
# 5. VARIACIÓN INTERANUAL (YoY)
# ---------------------------------------------------------
with tabs[3]:
    st.subheader("5. Variación Interanual (YoY)")
    
    temps = sorted(df_filtrado["Temporada"].unique().tolist())
    if len(temps) < 2:
        st.info(f"Se requieren al menos 2 temporadas para comparar variación interanual. Temporadas disponibles: {temps}")
    else:
        c_t1, c_t2 = st.columns(2)
        temp_base = c_t1.selectbox("Temporada Base (T1):", temps, index=0)
        temp_comp = c_t2.selectbox("Temporada Comparativa (T2):", temps, index=len(temps) - 1)
        
        pivot_vol = df_filtrado.pivot_table(
            index=["Material", "Texto breve", "Unidad medida pedido"],
            columns="Temporada",
            values="Cantidad de pedido",
            aggfunc="sum",
            fill_value=0
        )
        
        pivot_val = df_filtrado.pivot_table(
            index=["Material", "Texto breve", "Unidad medida pedido"],
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
                f"Vol_{temp_base}": "{:,.2f}",
                f"Vol_{temp_comp}": "{:,.2f}",
                "Var_Vol_%": "{:+.2f}%",
                f"Gasto_{temp_base}": "{:,.2f}",
                f"Gasto_{temp_comp}": "{:,.2f}",
                "Var_Gasto_%": "{:+.2f}%"
            }),
            use_container_width=True
        )

# ---------------------------------------------------------
# 6. PRIORIZACIÓN PARA LICITACIÓN CORPORATIVA
# ---------------------------------------------------------
with tabs[4]:
    st.subheader("6. Matriz de Recomendación de Licitación Corporativa")
    
    licit_df = df_filtrado.groupby(["Material", "Texto breve", "Unidad medida pedido"]).agg(
        Gasto_Total=("Valor neto de pedido", "sum"),
        Volumen_Total=("Cantidad de pedido", "sum"),
        Filiales_Compradoras=("Organización compras", "nunique"),
        Proveedores_Activos=("Proveedor/Centro suministrador", "nunique")
    ).reset_index().sort_values(by="Gasto_Total", ascending=False)
    
    total_gasto_lic = licit_df["Gasto_Total"].sum()
    licit_df["% Gasto Acum"] = (licit_df["Gasto_Total"] / total_gasto_lic * 100).cumsum()
    
    def catalogar_sourcing(row):
        if row["% Gasto Acum"] <= 80:
            return "🔥 Alta Prioridad (Clase A Corporativo)"
        elif row["% Gasto Acum"] <= 95 and row["Filiales_Compradoras"] > 1:
            return "⚡ Prioridad Media (Multi-Filial Consolidable)"
        else:
            return "⚪ Baja Prioridad (Compra Local / Spot)"
            
    licit_df["Recomendación_Estratégica"] = licit_df.apply(catalogar_sourcing, axis=1)
    
    st.dataframe(
        licit_df.style.format({
            "Gasto_Total": "{:,.2f}",
            "Volumen_Total": "{:,.2f}",
            "% Gasto Acum": "{:.2f}%",
            "Filiales_Compradoras": "{:,.0f}",
            "Proveedores_Activos": "{:,.0f}"
        }),
        use_container_width=True
    )

# ---------------------------------------------------------
# 7. ANÁLISIS DE PRECIOS POR UNIDAD DE MEDIDA
# ---------------------------------------------------------
with tabs[5]:
    st.subheader("7. Precios Históricos: Promedio Ponderado, Mínimo y Máximo")
    
    precios_df = df_filtrado[df_filtrado["Precio_Unitario_Real"] > 0].groupby(
        ["Temporada", "Material", "Texto breve", "Unidad medida pedido"]
    ).agg(
        Precio_Min=("Precio_Unitario_Real", "min"),
        Precio_Max=("Precio_Unitario_Real", "max"),
        Precio_Prom_Simple=("Precio_Unitario_Real", "mean"),
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
            "Precio_Prom_Simple": "{:,.4f}",
            "Precio_Prom_Ponderado": "{:,.4f}",
            "Dispersion_Precio_%": "{:.1f}%",
            "Volumen_Total": "{:,.2f}",
            "Gasto_Total": "{:,.2f}"
        }),
        use_container_width=True
    )

# ---------------------------------------------------------
# 8. RANKING DE PROVEEDORES
# ---------------------------------------------------------
with tabs[6]:
    st.subheader("8. Ranking de Proveedores por Gasto Total")
    
    ranking_prov = df_filtrado.groupby("Proveedor/Centro suministrador").agg(
        Gasto_Total=("Valor neto de pedido", "sum"),
        PO_Emitidas=("Documento compras", "nunique")
    ).reset_index().sort_values(by="Gasto_Total", ascending=False)
    
    total_gasto_p = ranking_prov["Gasto_Total"].sum()
    ranking_prov["% Participacion"] = (ranking_prov["Gasto_Total"] / total_gasto_p) * 100
    
    top_10_p = ranking_prov.head(10)
    
    fig_p, ax_p = plt.subplots(figsize=(10, 5))
    sns.barplot(data=top_10_p, y="Proveedor/Centro suministrador", x="Gasto_Total", palette="crest", ax=ax_p)
    ax_p.set_title("Top 10 Proveedores por Importe Adquirido")
    ax_p.set_xlabel("Gasto Total")
    ax_p.set_ylabel("Proveedor")
    st.pyplot(fig_p)
    
    st.dataframe(
        ranking_prov.style.format({
            "Gasto_Total": "{:,.2f}",
            "PO_Emitidas": "{:,.0f}",
            "% Participacion": "{:.2f}%"
        }),
        use_container_width=True
    )

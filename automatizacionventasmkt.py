import streamlit as st
import pandas as pd
import re
from datetime import datetime

st.set_page_config(page_title="Automatización Ventas MKT", page_icon="📊", layout="wide")

st.title("📊 Automatizador de Informe de Ventas - Marketing")
st.write("Sube los archivos .txt descargados de Amazon Seller Central para generar el informe unificado.")

# 1. Selectores de archivos para las dos cuentas
col1, col2 = st.columns(2)
with col1:
    archivo_jabiru = st.file_uploader("Subir TXT de JABIRU", type=["txt"])
with col2:
    archivo_turaco = st.file_uploader("Subir TXT de TURACO", type=["txt"])

def limpiar_sku(sku):
    if pd.isna(sku):
        return sku
    sku_str = str(sku).strip()
    # Eliminar prefijos comunes S, FR, IT, DE seguidos de espacio o guion, y puntos
    sku_str = re.sub(r'^(S|FR|IT|DE)[\s\-_]*', '', sku_str, flags=re.IGNORECASE)
    sku_str = sku_str.replace('.', '')
    return sku_str

def procesar_archivo(archivo, nombre_cuenta):
    if archivo is None:
        return None
    
    # Leer el archivo TXT separado por tabulaciones
    df = pd.read_csv(archivo, sep='\t')
    
    # Columnas requeridas por el negocio (mapeo tolerante a mayúsculas/minúsculas)
    columnas_interes = {
        'purchase-date': 'FECHA',
        'sku': 'REFERENCIA',
        'asin': 'ASIN',
        'quantity': 'CANTIDAD',
        'item-price': 'IMPORTE TOTAL',
        'ship-country': 'ship-country'
    }
    
    # Normalizar nombres de columnas del archivo para evitar fallos por minúsculas/mayúsculas
    df.columns = [col.lower().strip() for col in df.columns]
    
    # Verificar que existan las columnas
    for col_req in columnas_interes.keys():
        if col_req not in df.columns:
            st.error(f"No se encontró la columna '{col_req}' en el archivo de {nombre_cuenta}.")
            return None
            
    # Filtrar solo las columnas deseadas
    df = df[list(columnas_interes.keys())]
    df = df.rename(columns=columnas_interes)
    
    # Añadir columna de cuenta
    df['CUENTA'] = nombre_cuenta
    
    # --- LIMPIEZA DE DATOS ---
    
    # 1. Limpieza de Cantidad e Importe (eliminar vacíos y ceros)
    df = df.dropna(subset=['CANTIDAD', 'IMPORTE TOTAL'])
    df['CANTIDAD'] = pd.to_numeric(df['CANTIDAD'], errors='coerce')
    df['IMPORTE TOTAL'] = pd.to_numeric(df['IMPORTE TOTAL'], errors='coerce')
    df = df[(df['CANTIDAD'] > 0) & (df['IMPORTE TOTAL'] > 0)]
    
    # 2. Formatear Fecha y calcular Número de Semana (comenzando en Domingo -> week_start=7 o %U)
    # Amazon suele dar la fecha en formato ISO (ej. 2026-05-09T23:59:00)
    df['FECHA_DATETIME'] = pd.to_datetime(df['FECHA'], errors='coerce', utc=True)
    
    # Extraer número de semana estilo Excel NUM.DE.SEMANA(fecha; 1) -> Semana empieza en Domingo
    # %U calcula la semana del año empezando el domingo como primer día de la semana
    df['SEMANA'] = df['FECHA_DATETIME'].dt.strftime('%U').astype(float) + 1 
    df['FECHA'] = df['FECHA_DATETIME'].dt.strftime('%d/%m/%Y')
    
    # 3. Limpiar SKU (Referencia)
    df['REFERENCIA'] = df['REFERENCIA'].apply(limpiar_sku)
    
    # Reordenar columnas finales según la estructura solicitada
    columnas_finales = ['FECHA', 'SEMANA', 'REFERENCIA', 'ASIN', 'CANTIDAD', 'IMPORTE TOTAL', 'CUENTA', 'ship-country']
    df = df[columnas_finales]
    
    return df

# Procesar si se han subido los archivos
dataframes = []
if archivo_jabiru:
    df_j = procesar_archivo(archivo_jabiru, "JABIRU")
    if df_j is not None:
        dataframes.append(df_j)

if archivo_turaco:
    df_t = procesar_archivo(archivo_turaco, "TURACO")
    if df_t is not None:
        dataframes.append(df_t)

if len(dataframes) > 0:
    # Combinar ambos archivos de manera independiente y limpia
    resultado_final = pd.concat(dataframes, ignore_index=True)
    
    st.success("¡Archivos procesados y combinados con éxito!")
    
    # Mostrar vista previa
    st.subheader("Vista previa del informe unificado")
    st.dataframe(resultado_final.head(20))
    
    # Convertir a Excel para la descarga
    @st.cache_data
    def convertir_a_excel(df):
        # Guardar en buffer de memoria
        import io
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Ventas MKT')
        return buffer.getvalue()
    
    excel_data = convertir_a_excel(resultado_final)
    
    # Obtener el número de semana actual para el nombre del archivo
    num_semana_inf = int(resultado_final['SEMANA'].iloc[0]) if not resultado_final.empty else 0
    
    st.download_button(
        label="📥 Descargar Excel para Marketing",
        data=excel_data,
        file_name=f"Informe_Ventas_MKT_Semana_{num_semana_inf}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.info("📌 Recuerda enviar este archivo cada miércoles a: martacuesta@cecotec.es")
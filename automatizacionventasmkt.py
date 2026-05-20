import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Automatización Ventas MKT", page_icon="📊", layout="wide")

st.title("📊 Automatizador de Informe de Ventas - Marketing")
st.write("Sube los archivos .txt actualizados descargados de Amazon para generar el informe consolidado.")

# Selectores de archivos en la interfaz web (Carga volátil en memoria)
col1, col2 = st.columns(2)
with col1:
    archivo_jabiru = st.file_uploader("Subir TXT de JABIRU", type=["txt"])
with col2:
    archivo_turaco = st.file_uploader("Subir TXT de TURACO", type=["txt"])

def limpiar_sku(sku):
    if pd.isna(sku):
        return sku
    sku_str = str(sku).strip()
    # Quitar prefijos de país (S, FR, IT, DE) y puntos de forma segura [cite: 22]
    sku_str = re.sub(r'^(S|FR|IT|DE)[\s\-_]*', '', sku_str, flags=re.IGNORECASE)
    sku_str = sku_str.replace('.', '')
    return sku_str

def procesar_archivo(archivo, nombre_cuenta):
    if archivo is None:
        return None
    
    try:
        # El encoding 'utf-8-sig' es vital para ignorar caracteres ocultos de Amazon
        df = pd.read_csv(archivo, sep='\t', dtype=str, encoding='utf-8-sig')
        
        # Diccionario de equivalencias para las columnas requeridas [cite: 17, 18]
        mapeo = {
            'purchase-date': 'FECHA',
            'sku': 'REFERENCIA',
            'asin': 'ASIN',
            'quantity': 'CANTIDAD',
            'item-price': 'IMPORTE TOTAL',
            'ship-country': 'ship-country'
        }
        
        # Buscar correspondencias sin importar mayúsculas/minúsculas
        columnas_encontradas = {}
        for col_original in df.columns:
            col_normalizada = str(col_original).lower().strip()
            if col_normalizada in mapeo:
                columnas_encontradas[col_original] = mapeo[col_normalizada]
        
        # Validar que el archivo subido sea el correcto
        if len(columnas_encontradas) < 6:
            st.error(f"⚠️ Al archivo de {nombre_cuenta} le faltan columnas esenciales de Amazon.")
            return None
            
        # Filtrar y renombrar columnas
        df = df[list(columnas_encontradas.keys())]
        df = df.rename(columns=columnas_encontradas)
        df['CUENTA'] = nombre_cuenta
        
        # Limpieza de nulos y conversión numérica [cite: 23, 24]
        df = df.dropna(subset=['CANTIDAD', 'IMPORTE TOTAL'])
        df['CANTIDAD'] = pd.to_numeric(df['CANTIDAD'], errors='coerce')
        df['IMPORTE TOTAL'] = pd.to_numeric(df['IMPORTE TOTAL'], errors='coerce')
        df = df[(df['CANTIDAD'] > 0) & (df['IMPORTE TOTAL'] > 0)]
        
        # Extracción e interpretación instantánea de fechas (evita cuelgues)
        df['FECHA_CORTA'] = df['FECHA'].astype(str).str.slice(0, 10)
        df['FECHA_DATETIME'] = pd.to_datetime(df['FECHA_CORTA'], format='%Y-%m-%d', errors='coerce')
        
        # Calcular semana (Estilo Excel, Domingo a Sábado) y formatear texto [cite: 21]
        df['SEMANA'] = df['FECHA_DATETIME'].dt.strftime('%U').astype(float) + 1
        df['FECHA'] = df['FECHA_DATETIME'].dt.strftime('%d/%m/%Y')
        
        # Limpieza de referencias (SKU) [cite: 22]
        df['REFERENCIA'] = df['REFERENCIA'].apply(limpiar_sku)
        
        # Reordenación estructural final [cite: 25]
        columnas_finales = ['FECHA', 'SEMANA', 'REFERENCIA', 'ASIN', 'CANTIDAD', 'IMPORTE TOTAL', 'CUENTA', 'ship-country']
        df = df[columnas_finales].dropna(subset=['FECHA', 'SEMANA'])
        df['SEMANA'] = df['SEMANA'].astype(int)
        
        return df

    except Exception as e:
        st.error(f"Error procesando los datos de {nombre_cuenta}: {str(e)}")
        return None

# Unión y lógica de despliegue en pantalla
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
    resultado_final = pd.concat(dataframes, ignore_index=True)
    resultado_final = resultado_final.sort_values(by=['SEMANA', 'FECHA'], ascending=[True, True])
    
    st.success("🎯 ¡Archivos combinados y procesados con éxito!")
    st.subheader("Vista previa del Informe Final")
    st.dataframe(resultado_final)
    
    # Construcción del archivo Excel usando motores nativos estándar
    @st.cache_data
    def generar_excel_estandar(df):
        buffer = io.BytesIO()
        # openpyxl viene instalado por defecto en Streamlit Cloud, garantizando cero errores
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Ventas MKT')
        return buffer.getvalue()
    
    excel_final = generar_excel_estandar(resultado_final)
    num_semana = resultado_final['SEMANA'].iloc[0] if not resultado_final.empty else "Desconocida"
    
    st.download_button(
        label="📥 Descargar Excel Unificado para Marta Cuesta",
        data=excel_final,
        file_name=f"Ventas_MKT_Semana_{num_semana}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.info("📧 Enviar este reporte los miércoles a: martacuesta@cecotec.es [cite: 26, 27]")

import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Automatización Ventas MKT", page_icon="📊", layout="wide")

st.title("📊 Automatizador de Informe de Ventas - Marketing")
st.write("Sube los archivos .txt actualizados de Amazon para generar tu informe.")

# Selectores de archivos en la interfaz
col1, col2 = st.columns(2)
with col1:
    archivo_jabiru = st.file_uploader("Subir TXT de JABIRU", type=["txt"])
with col2:
    archivo_turaco = st.file_uploader("Subir TXT de TURACO", type=["txt"])

def limpiar_sku(sku):
    if pd.isna(sku):
        return sku
    sku_str = str(sku).strip()
    sku_str = re.sub(r'^(S|FR|IT|DE)[\s\-_]*', '', sku_str, flags=re.IGNORECASE)
    sku_str = sku_str.replace('.', '')
    return sku_str

def procesar_archivo(archivo, nombre_cuenta):
    if archivo is None:
        return None
    
    try:
        # 'utf-8-sig' elimina automáticamente caracteres ocultos/BOM que congelan la lectura
        df = pd.read_csv(archivo, sep='\t', dtype=str, encoding='utf-8-sig')
        
        # Diccionario de búsqueda directo sin bucles extraños
        mapeo = {
            'purchase-date': 'FECHA',
            'sku': 'REFERENCIA',
            'asin': 'ASIN',
            'quantity': 'CANTIDAD',
            'item-price': 'IMPORTE TOTAL',
            'ship-country': 'ship-country'
        }
        
        # Mapear las columnas existentes comparando de forma directa y limpia
        columnas_encontradas = {}
        for col_original in df.columns:
            col_normalizada = str(col_original).lower().strip()
            if col_normalizada in mapeo:
                columnas_encontradas[col_original] = mapeo[col_normalizada]
        
        # Verificar si tenemos las 6 columnas vitales
        if len(columnas_encontradas) < 6:
            st.error(f"⚠️ Al archivo de {nombre_cuenta} le faltan columnas. Asegúrate de que es el reporte correcto.")
            return None
            
        # Filtrar y renombrar usando el mapeo directo verificado
        df = df[list(columnas_encontradas.keys())]
        df = df.rename(columns=columnas_encontradas)
        
        # Añadir la cuenta
        df['CUENTA'] = nombre_cuenta
        
        # Limpieza rápida de filas vacías
        df = df.dropna(subset=['CANTIDAD', 'IMPORTE TOTAL'])
        
        # Conversión de tipos de datos de forma directa
        df['CANTIDAD'] = pd.to_numeric(df['CANTIDAD'], errors='coerce')
        df['IMPORTE TOTAL'] = pd.to_numeric(df['IMPORTE TOTAL'], errors='coerce')
        df = df[(df['CANTIDAD'] > 0) & (df['IMPORTE TOTAL'] > 0)]
        
        # Procesar fechas cortando texto de forma segura (sin interpretar zonas horarias complejas)
        df['FECHA_CORTA'] = df['FECHA'].astype(str).str.slice(0, 10)
        df['FECHA_DATETIME'] = pd.to_datetime(df['FECHA_CORTA'], format='%Y-%m-%d', errors='coerce')
        
        # Calcular semana y formatear visualmente
        df['SEMANA'] = df['FECHA_DATETIME'].dt.strftime('%U').astype(float) + 1
        df['FECHA'] = df['FECHA_DATETIME'].dt.strftime('%d/%m/%Y')
        
        # Limpiar SKU
        df['REFERENCIA'] = df['REFERENCIA'].apply(limpiar_sku)
        
        # Estructura final estricta
        columnas_finales = ['FECHA', 'SEMANA', 'REFERENCIA', 'ASIN', 'CANTIDAD', 'IMPORTE TOTAL', 'CUENTA', 'ship-country']
        df = df[columnas_finales]
        
        # Eliminar cualquier fila rota residual
        df = df.dropna(subset=['FECHA', 'SEMANA'])
        df['SEMANA'] = df['SEMANA'].astype(int)
        
        return df

    except Exception as e:
        st.error(f"Error procesando {nombre_cuenta}: {str(e)}")
        return None

# Ejecución de la lógica en Streamlit
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
    
    st.success("🎯 ¡Archivos procesados instantáneamente!")
    st.subheader("Vista previa del Informe Final")
    st.dataframe(resultado_final)
    
    # Renderizado a Excel físico
    @st.cache_data
    def generar_excel_formateado(df):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Ventas MKT')
            workbook  = writer.book
            worksheet = writer.sheets['Ventas MKT']
            
            formato_moneda = workbook.add_format({'num_format': '#,##0.00" €"'})
            formato_entero = workbook.add_format({'num_format': '#,##0', 'align': 'center'})
            formato_centrado = workbook.add_format({'align': 'center'})
            
            worksheet.set_column('A:A', 14, formato_centrado)
            worksheet.set_column('B:B', 10, formato_entero)
            worksheet.set_column('C:D', 18)
            worksheet.set_column('E:E', 12, formato_entero)
            worksheet.set_column('F:F', 16, formato_moneda)
            worksheet.set_column('G:H', 14, formato_centrado)
            
        return buffer.getvalue()
    
    excel_final = generar_excel_formateado(resultado_final)
    num_semana = resultado_final['SEMANA'].iloc[0] if not resultado_final.empty else "Desconocida"
    
    st.download_button(
        label="📥 Descargar Excel Unificado para Marta Cuesta",
        data=excel_final,
        file_name=f"Ventas_MKT_Semana_{num_semana}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.info(f"📧 Destinatario: martacuesta@cecotec.es")

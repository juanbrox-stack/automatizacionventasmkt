import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Automatización Ventas MKT", page_icon="📊", layout="wide")

st.title("📊 Automatizador de Informe de Ventas - Marketing")
st.write("Sube los archivos .txt reales descargados de Amazon Seller Central para generar el informe consolidado.")

# 1. Selectores de archivos en la interfaz [cite: 8, 10]
col1, col2 = st.columns(2)
with col1:
    archivo_jabiru = st.file_uploader("Subir TXT de JABIRU", type=["txt"])
with col2:
    archivo_turaco = st.file_uploader("Subir TXT de TURACO", type=["txt"])

def limpiar_sku(sku):
    if pd.isna(sku):
        return sku
    sku_str = str(sku).strip()
    # Limpiar prefijos de país (S, FR, IT, DE) seguidos de guion, espacio o guion bajo [cite: 22]
    sku_str = re.sub(r'^(S|FR|IT|DE)[\s\-_]*', '', sku_str, flags=re.IGNORECASE)
    # Quitar puntos del SKU [cite: 22]
    sku_str = sku_str.replace('.', '')
    return sku_str

def procesar_archivo(archivo, nombre_cuenta):
    if archivo is None:
        return None
    
    try:
        # Leer el archivo TXT nativo de Amazon (separado por tabulaciones) [cite: 15, 16]
        df = pd.read_csv(archivo, sep='\t', dtype=str)
        
        # Normalizar las cabeceras del archivo original (quitar espacios y pasar a minúsculas)
        df.columns = [col.lower().strip() for col in df.columns]
        
        # Mapeo de las columnas normalizadas a los nombres deseados por Marketing [cite: 17]
        mapeo_columnas = {
            'purchase-date': 'FECHA',
            'sku': 'REFERENCIA',
            'asin': 'ASIN',
            'quantity': 'CANTIDAD',
            'item-price': 'IMPORTE TOTAL',
            'ship-country': 'ship-country'
        }
        
        # Comprobar si todas las columnas necesarias existen en el DataFrame normalizado
        for col_req in mapeo_columnas.keys():
            if col_req not in df.columns:
                st.error(f"⚠️ El archivo de {nombre_cuenta} no contiene la columna necesaria: '{col_req}'")
                return None
        
        # Filtrar solo las columnas de interés usando los nombres normalizados
        df = df[list(mapeo_columnas.keys())]
        
        # Renombrar a las variables solicitadas por el negocio [cite: 17, 25]
        df = df.rename(columns=mapeo_columnas)
        
        # Añadir la columna de identificación de la cuenta [cite: 18, 25]
        df['CUENTA'] = nombre_cuenta
        
        # --- LIMPIEZA DE FILAS (CANTIDAD E IMPORTE) ---
        # Quitar nulos iniciales en las columnas críticas [cite: 23, 24]
        df = df.dropna(subset=['CANTIDAD', 'IMPORTE TOTAL'])
        
        # Convertir variables a numéricas de forma segura
        df['CANTIDAD'] = pd.to_numeric(df['CANTIDAD'], errors='coerce')
        df['IMPORTE TOTAL'] = pd.to_numeric(df['IMPORTE TOTAL'], errors='coerce')
        
        # Filtrar: Asegurar que no haya celdas vacías, ceros o valores negativos [cite: 23, 24]
        df = df[(df['CANTIDAD'] > 0) & (df['IMPORTE TOTAL'] > 0)]
        
        # --- PROCESAMIENTO DE FECHAS Y SEMANAS ---
        # Extraer los primeros 10 caracteres (YYYY-MM-DD) para evitar conflictos con la zona horaria de Amazon
        df['FECHA_CORTA'] = df['FECHA'].astype(str).str.slice(0, 10)
        df['FECHA_DATETIME'] = pd.to_datetime(df['FECHA_CORTA'], format='%Y-%m-%d', errors='coerce')
        
        # Calcular Número de Semana (Estilo Excel: Domingo a Sábado = modificador %U) [cite: 14, 21]
        df['SEMANA'] = df['FECHA_DATETIME'].dt.strftime('%U').astype(float) + 1
        
        # Formatear la fecha visual en formato dd/mm/aaaa [cite: 20]
        df['FECHA'] = df['FECHA_DATETIME'].dt.strftime('%d/%m/%Y')
        
        # Aplicar la limpieza de los SKU [cite: 22]
        df['REFERENCIA'] = df['REFERENCIA'].apply(limpiar_sku)
        
        # Reestructurar las columnas según el diseño final solicitado [cite: 25]
        columnas_finales = ['FECHA', 'SEMANA', 'REFERENCIA', 'ASIN', 'CANTIDAD', 'IMPORTE TOTAL', 'CUENTA', 'ship-country']
        df = df[columnas_finales]
        
        # Eliminar cualquier registro con conversión de fecha corrupta
        df = df.dropna(subset=['FECHA', 'SEMANA'])
        df['SEMANA'] = df['SEMANA'].astype(int)
        
        return df

    except Exception as e:
        st.error(f"Error interno al procesar {nombre_cuenta}: {e}")
        return None

# Flujo principal de la aplicación
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
    # Combinar los datos unificados [cite: 25]
    resultado_final = pd.concat(dataframes, ignore_index=True)
    
    # Ordenar por Semana y por Fecha de manera ascendente
    resultado_final = resultado_final.sort_values(by=['SEMANA', 'FECHA'], ascending=[True, True])
    
    st.success("🎯 ¡Archivos alineados y procesados con éxito!")
    
    st.subheader("Vista previa del Informe Final")
    st.dataframe(resultado_final)
    
    # Generador de Excel con formato físico de moneda e importes
    @st.cache_data
    def generar_excel_formateado(df):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Ventas MKT')
            
            workbook  = writer.book
            worksheet = writer.sheets['Ventas MKT']
            
            # Formatos de celda para emular el resultado del documento
            formato_moneda = workbook.add_format({'num_format': '#,##0.00" €"'})
            formato_entero = workbook.add_format({'num_format': '#,##0', 'align': 'center'})
            formato_centrado = workbook.add_format({'align': 'center'})
            
            # Anchos de columna configurados
            worksheet.set_column('A:A', 14, formato_centrado)  # Fecha
            worksheet.set_column('B:B', 10, formato_entero)    # Semana
            worksheet.set_column('C:D', 18)                    # Referencia y ASIN
            worksheet.set_column('E:E', 12, formato_entero)    # Cantidad
            worksheet.set_column('F:F', 16, formato_moneda)    # Importe Total
            worksheet.set_column('G:H', 14, formato_centrado)  # Cuenta y País
            
        return buffer.getvalue()
    
    excel_final = generar_excel_formateado(resultado_final)
    
    # Extraer la semana del informe para el nombre dinámico del archivo
    num_semana = resultado_final['SEMANA'].iloc[0] if not resultado_final.empty else "Desconocida"
    
    st.download_button(
        label="📥 Descargar Excel Unificado para Marta Cuesta",
        data=excel_final,
        file_name=f"Ventas_MKT_Semana_{num_semana}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.info(f"📧 Destinatario: **martacuesta@cecotec.es** (Proceso programado para los miércoles) [cite: 26, 27]")

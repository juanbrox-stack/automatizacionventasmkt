import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Automatización Ventas MKT", page_icon="📊", layout="wide")

st.title("📊 Automatizador de Informe de Ventas - Marketing")
st.write("Sube los archivos .txt descargados de Amazon Seller Central para generar el informe unificado.")

# Selectores de archivos para las dos cuentas
col1, col2 = st.columns(2)
with col1:
    archivo_jabiru = st.file_uploader("Subir TXT de JABIRU", type=["txt"])
with col2:
    archivo_turaco = st.file_uploader("Subir TXT de TURACO", type=["txt"])

def limpiar_sku(sku):
    if pd.isna(sku):
        return sku
    sku_str = str(sku).strip()
    # Eliminar prefijos S, FR, IT, DE (ignorando mayúsculas/minúsculas) seguidos de guion o espacio
    sku_str = re.sub(r'^(S|FR|IT|DE)[\s\-_]*', '', sku_str, flags=re.IGNORECASE)
    # Quitar los puntos intermedios o finales si existen
    sku_str = sku_str.replace('.', '')
    return sku_str

def procesar_archivo(archivo, nombre_cuenta):
    if archivo is None:
        return None
    
    # Leer el archivo TXT separado por tabulaciones (formato nativo de Amazon)
    df = pd.read_csv(archivo, sep='\t')
    
    # Mapeo de columnas requeridas por Marta Cuesta
    columnas_interes = {
        'purchase-date': 'FECHA',
        'sku': 'REFERENCIA',
        'asin': 'ASIN',
        'quantity': 'CANTIDAD',
        'item-price': 'IMPORTE TOTAL',
        'ship-country': 'ship-country'
    }
    
    # Normalizar columnas a minúsculas para evitar errores si Amazon cambia las cabeceras
    df.columns = [col.lower().strip() for col in df.columns]
    
    # Validar que todas las columnas necesarias existan
    for col_req in columnas_interes.keys():
        if col_req not in df.columns:
            st.error(f"Error: No se encontró la columna requerida '{col_req}' en el archivo de {nombre_cuenta}.")
            return None
            
    # Filtrar y renombrar
    df = df[list(columnas_interes.keys())]
    df = df.rename(columns=columnas_interes)
    
    # Añadir columna identificadora de la cuenta (JABIRU / TURACO)
    df['CUENTA'] = nombre_cuenta
    
    # --- LIMPIEZA Y PARSEO DE DATOS ---
    
    # 1. Convertir Cantidad e Importe a números y limpiar nulos/ceros
    df['CANTIDAD'] = pd.to_numeric(df['CANTIDAD'], errors='coerce')
    df['IMPORTE TOTAL'] = pd.to_numeric(df['IMPORTE TOTAL'], errors='coerce')
    
    # Eliminar filas con cantidades o importes vacíos, menores o iguales a cero
    df = df.dropna(subset=['CANTIDAD', 'IMPORTE TOTAL'])
    df = df[(df['CANTIDAD'] > 0) & (df['IMPORTE TOTAL'] > 0)]
    
    # 2. Procesamiento estricto de Fechas y Semanas (Semana de Amazon inicia en Domingo)
    df['FECHA_DATETIME'] = pd.to_datetime(df['purchase-date'], errors='coerce', utc=True)
    
    # Formato de semana '%U' inicia en Domingo (Equivalente a NUM.DE.SEMANA(fecha; 1) en Excel)
    df['SEMANA'] = df['FECHA_DATETIME'].dt.strftime('%U').astype(float) + 1
    df['FECHA'] = df['FECHA_DATETIME'].dt.strftime('%d/%m/%Y')
    
    # 3. Limpiar la columna REFERENCIA (SKU)
    df['REFERENCIA'] = df['REFERENCIA'].apply(limpiar_sku)
    
    # Reordenar las columnas para cumplir con la estructura final exacta
    columnas_finales = ['FECHA', 'SEMANA', 'REFERENCIA', 'ASIN', 'CANTIDAD', 'IMPORTE TOTAL', 'CUENTA', 'ship-country']
    df = df[columnas_finales]
    
    return df

# Procesar los archivos si se cargan en la interfaz
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
    # Concatenar de forma independiente ambos conjuntos de datos sin importar el desorden original
    resultado_final = pd.concat(dataframes, ignore_index=True)
    
    # Asegurar orden cronológico estricto en la visualización
    resultado_final = resultado_final.sort_values(by=['SEMANA', 'FECHA'], ascending=[True, True])
    
    st.success("🎯 ¡Archivos procesados y alineados perfectamente con el formato de ejemplo!")
    
    # Mostrar vista previa interactiva en la web
    st.subheader("Vista previa del resultado final")
    st.dataframe(resultado_final)
    
    # Compilar el archivo Excel aplicando formato numérico nativo
    @st.cache_data
    def generar_excel_formateado(df):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Ventas MKT')
            
            # Acceder a las herramientas de diseño de xlsxwriter
            workbook  = writer.book
            worksheet = writer.sheets['Ventas MKT']
            
            # Formato para la columna de dinero (€)
            formato_moneda = workbook.add_format({'num_format': '#,##0.00" €"'})
            # Formato para enteros en cantidad
            formato_entero = workbook.add_format({'num_format': '#,##0'})
            
            # Aplicar formatos a las columnas correspondientes (E = Cantidad, F = Importe)
            worksheet.set_column('E:E', 12, formato_entero)
            worksheet.set_column('F:F', 15, formato_moneda)
            worksheet.set_column('A:D', 15)
            worksheet.set_column('G:H', 12)
            
        return buffer.getvalue()
    
    excel_final = generar_excel_formateado(resultado_final)
    
    # Extraer dinámicamente el número de semana para nombrar el reporte
    num_semana = int(resultado_final['SEMANA'].iloc[0]) if not resultado_final.empty else 0
    
    st.download_button(
        label="📥 Descargar archivo Excel unificado",
        data=excel_final,
        file_name=f"Ventas_MKT_Semana_{num_semana}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.info("📨 Destinatario del informe: martacuesta@cecotec.es (Enviar todos los miércoles).")

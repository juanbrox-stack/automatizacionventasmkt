import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Automatización Ventas MKT", page_icon="📊", layout="wide")

st.title("📊 Automatizador de Informe de Ventas - Marketing")
st.write("Sube los archivos .txt reales descargados de Amazon Seller Central para generar el informe consolidado.")

# 1. Selectores de archivos en la interfaz
col1, col2 = st.columns(2)
with col1:
    archivo_jabiru = st.file_uploader("Subir TXT de JABIRU", type=["txt"])
with col2:
    archivo_turaco = st.file_uploader("Subir TXT de TURACO", type=["txt"])

def limpiar_sku(sku):
    if pd.isna(sku):
        return sku
    sku_str = str(sku).strip()
    # Eliminar prefijos S, FR, IT, DE seguidos de espacio, guion o guion bajo (case-insensitive) 
    sku_str = re.sub(r'^(S|FR|IT|DE)[\s\-_]*', '', sku_str, flags=re.IGNORECASE)
    # Eliminar puntos del SKU 
    sku_str = sku_str.replace('.', '')
    return sku_str

def procesar_archivo(archivo, nombre_cuenta):
    if archivo is None:
        return None
    
    try:
        # Leer el archivo de Amazon especificando codificación UTF-8 o ISO-8859-1 por si contiene caracteres especiales
        df = pd.read_csv(archivo, sep='\t', dtype=str)
        
        # Diccionario de mapeo de columnas requeridas [cite: 17, 18]
        columnas_interes = {
            'purchase-date': 'FECHA',
            'sku': 'REFERENCIA',
            'asin': 'ASIN',
            'quantity': 'CANTIDAD',
            'item-price': 'IMPORTE TOTAL',
            'ship-country': 'ship-country'
        }
        
        # Limpiar espacios y pasar a minúsculas las cabeceras originales del archivo para evitar desajustes [cite: 16]
        df.columns = [col.lower().strip() for col in df.columns]
        
        # Validar la existencia de las columnas requeridas [cite: 17]
        for col_req in columnas_interes.keys():
            if col_req not in df.columns:
                st.error(f"⚠️ El archivo de {nombre_cuenta} no contiene la columna '{col_req}'. Verifica que sea el reporte 'Todos los pedidos'.")
                return None
        
        # Filtrar solo lo que nos pide Marketing [cite: 17]
        df = df[list(columnas_interes.keys())]
        df = df.rename(columns=columnas_interes)
        
        # Asignar la cuenta correspondiente [cite: 18]
        df['CUENTA'] = nombre_cuenta
        
        # --- LIMPIEZA DE DATOS ---
        
        # Quitar filas donde Cantidad o Importe estén completamente vacíos [cite: 23, 24]
        df = df.dropna(subset=['CANTIDAD', 'IMPORTE TOTAL'])
        
        # Convertir a tipos numéricos de manera segura
        df['CANTIDAD'] = pd.to_numeric(df['CANTIDAD'], errors='coerce')
        df['IMPORTE TOTAL'] = pd.to_numeric(df['IMPORTE TOTAL'], errors='coerce')
        
        # Filtrar para no incluir celdas vacías, ceros o valores corruptos [cite: 23, 24]
        df = df.dropna(subset=['CANTIDAD', 'IMPORTE TOTAL'])
        df = df[(df['CANTIDAD'] > 0) & (df['IMPORTE TOTAL'] > 0)]
        
        # Procesar Fechas (Amazon usa formato ISO: 2026-05-10T11:47:15+02:00) 
        # Usamos t_str.split('T')[0] para asegurar la captura limpia de la fecha antes de la zona horaria
        df['FECHA_LIMPIA'] = df['FECHA'].apply(lambda x: str(x).split('T')[0] if pd.notna(x) else x)
        df['FECHA_DATETIME'] = pd.to_datetime(df['FECHA_LIMPIA'], format='%Y-%m-%d', errors='coerce')
        
        # Calcular Número de Semana de Amazon (Domingo a Sábado) -> Equivalente a NUM.DE.SEMANA(fecha; 1) 
        # %U calcula la semana iniciando en domingo
        df['SEMANA'] = df['FECHA_DATETIME'].dt.strftime('%U').astype(float) + 1
        
        # Forzar el formato visual solicitado para la fecha: dd/mm/aaaa 
        df['FECHA'] = df['FECHA_DATETIME'].dt.strftime('%d/%m/%Y')
        
        # Aplicar la limpieza avanzada a los SKU (Referencias) 
        df['REFERENCIA'] = df['REFERENCIA'].apply(limpiar_sku)
        
        # Reordenar las columnas según la estructura final estricta 
        columnas_finales = ['FECHA', 'SEMANA', 'REFERENCIA', 'ASIN', 'CANTIDAD', 'IMPORTE TOTAL', 'CUENTA', 'ship-country']
        df = df[columnas_finales]
        
        # Eliminar cualquier fila que haya quedado con fecha inválida
        df = df.dropna(subset=['FECHA', 'SEMANA'])
        df['SEMANA'] = df['SEMANA'].astype(int)
        
        return df

    except Exception as e:
        st.error(f"Error al procesar el archivo de {nombre_cuenta}: {e}")
        return None

# Ejecución del flujo de datos
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
    # Combinar los datos de ambas cuentas de manera limpia e independiente [cite: 16, 25]
    resultado_final = pd.concat(dataframes, ignore_index=True)
    
    # Ordenar cronológicamente (Por Semana y luego por Fecha)
    resultado_final = resultado_final.sort_values(by=['SEMANA', 'FECHA'], ascending=[True, True])
    
    st.success("🎯 ¡Tus archivos reales se han procesado, limpiado y unificado correctamente!")
    
    # Mostrar tabla interactiva en la pantalla
    st.subheader("Vista previa del Informe Final")
    st.dataframe(resultado_final)
    
    # Generador del archivo Excel físico formateado
    @st.cache_data
    def generar_excel_formateado(df):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Ventas MKT')
            
            workbook  = writer.book
            worksheet = writer.sheets['Ventas MKT']
            
            # Formatos de celda profesionales
            formato_moneda = workbook.add_format({'num_format': '#,##0.00" €"'})
            formato_entero = workbook.add_format({'num_format': '#,##0', 'align': 'center'})
            formato_centrado = workbook.add_format({'align': 'center'})
            
            # Aplicar anchos y formatos a las columnas 
            worksheet.set_column('A:A', 14, formato_centrado)  # Fecha
            worksheet.set_column('B:B', 10, formato_entero)    # Semana
            worksheet.set_column('C:D', 18)                    # Referencia y ASIN
            worksheet.set_column('E:E', 12, formato_entero)    # Cantidad
            worksheet.set_column('F:F', 16, formato_moneda)    # Importe Total
            worksheet.set_column('G:H', 14, formato_centrado)  # Cuenta y País
            
        return buffer.getvalue()
    
    excel_final = generar_excel_formateado(resultado_final)
    
    # Obtener el número de semana del reporte para el nombre automático del archivo [cite: 21]
    num_semana = resultado_final['SEMANA'].iloc[0] if not resultado_final.empty else "Desconocida"
    
    st.download_button(
        label="📥 Descargar Excel Unificado para Marta Cuesta",
        data=excel_final,
        file_name=f"Ventas_MKT_Semana_{num_semana}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.info(f"📧 Recuerda enviar este archivo adjunto a: **martacuesta@cecotec.es** indicando que son las ventas de la semana {num_semana}[cite: 26].")

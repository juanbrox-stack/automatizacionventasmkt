import streamlit as st
import pandas as pd
import re
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

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
    sku_str = re.sub(r'^(S|FR|IT|DE)[\s\-_]*', '', sku_str, flags=re.IGNORECASE)
    sku_str = sku_str.replace('.', '')
    return sku_str

def procesar_archivo(archivo, nombre_cuenta):
    if archivo is None:
        return None
    
    try:
        df = pd.read_csv(archivo, sep='\t', dtype=str, encoding='utf-8-sig')
        
        mapeo = {
            'purchase-date': 'FECHA',
            'sku': 'REFERENCIA',
            'asin': 'ASIN',
            'quantity': 'CANTIDAD',
            'item-price': 'IMPORTE TOTAL',
            'ship-country': 'ship-country'
        }
        
        columnas_encontradas = {}
        for col_original in df.columns:
            col_normalizada = str(col_original).lower().strip()
            if col_normalizada in mapeo:
                columnas_encontradas[col_original] = mapeo[col_normalizada]
        
        if len(columnas_encontradas) < 6:
            st.error(f"⚠️ Al archivo de {nombre_cuenta} le faltan columnas esenciales de Amazon.")
            return None
            
        df = df[list(columnas_encontradas.keys())]
        df = df.rename(columns=columnas_encontradas)
        df['CUENTA'] = nombre_cuenta
        
        df = df.dropna(subset=['CANTIDAD', 'IMPORTE TOTAL'])
        df['CANTIDAD'] = pd.to_numeric(df['CANTIDAD'], errors='coerce')
        df['IMPORTE TOTAL'] = pd.to_numeric(df['IMPORTE TOTAL'], errors='coerce')
        df = df[(df['CANTIDAD'] > 0) & (df['IMPORTE TOTAL'] > 0)]
        
        df['FECHA_CORTA'] = df['FECHA'].astype(str).str.slice(0, 10)
        df['FECHA_DATETIME'] = pd.to_datetime(df['FECHA_CORTA'], format='%Y-%m-%d', errors='coerce')
        
        df['SEMANA'] = df['FECHA_DATETIME'].dt.strftime('%U').astype(float) + 1
        df['FECHA'] = df['FECHA_DATETIME'].dt.strftime('%d/%m/%Y')
        
        df['REFERENCIA'] = df['REFERENCIA'].apply(limpiar_sku)
        
        columnas_finales = ['FECHA', 'SEMANA', 'REFERENCIA', 'ASIN', 'CANTIDAD', 'IMPORTE TOTAL', 'CUENTA', 'ship-country', 'FECHA_DATETIME']
        df = df[columnas_finales].dropna(subset=['FECHA', 'SEMANA'])
        df['SEMANA'] = df['SEMANA'].astype(int)
        
        return df

    except Exception as e:
        st.error(f"Error procesando los datos de {nombre_cuenta}: {str(e)}")
        return None

def enviar_correo_marta(excel_bytes, nombre_archivo, rango_fechas):
    try:
        # Obtener credenciales seguras desde los Secrets del servidor
        remitente = st.secrets["correo"]["usuario"]
        password = st.secrets["correo"]["password"]
        smtp_server = st.secrets["correo"]["servidor_smtp"]
        puerto = int(st.secrets["correo"]["puerto"])
        
        destinatario = "martacuesta@cecotec.es"
        
        # Configurar la estructura del correo electrónico
        msg = MIMEMultipart()
        msg['From'] = remitente
        msg['To'] = destinatario
        msg['Subject'] = f"Informe de Ventas Amazon (Periodo: {rango_fechas})"
        
        cuerpo = f"""Hola Marta,

Adjunto el informe resumen de ventas generadas en Amazon (cuentas Jabiru y Turaco) correspondientes al periodo {rango_fechas}.

Un saludo,

"""
        msg.attach(MIMEText(cuerpo, 'plain'))
        
        # Adjuntar el binario del archivo de Excel
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(excel_bytes)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={nombre_archivo}')
        msg.attach(part)
        
        # Conexión y envío seguro a través de SMTP de Gmail
        server = smtplib.SMTP(smtp_server, puerto)
        server.starttls()
        server.login(remitente, password)
        server.sendmail(remitente, destinatario, msg.as_string())
        server.quit()
        
        return True
    except Exception as e:
        st.error(f"Error al enviar el correo: {e}")
        return False

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
    resultado_final = resultado_final.sort_values(by=['SEMANA', 'FECHA_DATETIME'], ascending=[True, True])
    
    # --- CALCULAR EL RANGO DE FECHAS DINÁMICO ---
    fecha_minima = resultado_final['FECHA_DATETIME'].min().strftime('%d/%m/%Y')
    fecha_maxima = resultado_final['FECHA_DATETIME'].max().strftime('%d/%m/%Y')
    rango_dias_texto = f"del {fecha_minima} al {fecha_maxima}"
    
    # Eliminar la columna auxiliar datetime antes de mostrar la tabla para que quede limpia
    num_semana = resultado_final['SEMANA'].iloc[0] if not resultado_final.empty else "Desconocida"
    df_mostrar = resultado_final.drop(columns=['FECHA_DATETIME'])
    
    st.success("🎯 ¡Archivos combinados y procesados con éxito!")
    st.info(f"📅 Rango de días detectado automáticamente: **{rango_dias_texto}**")
    
    st.subheader("Vista previa del Informe Final")
    st.dataframe(df_mostrar)
    
    # Construcción del archivo Excel en memoria
    @st.cache_data
    def generar_excel_estandar(df):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Ventas MKT')
        return buffer.getvalue()
    
    excel_final = generar_excel_estandar(df_mostrar)
    nombre_del_excel = f"Ventas_MKT_Semana_{num_semana}.xlsx"
    
    # --- SECCIÓN DE ACCIONES ---
    st.write("---")
    st.subheader("🚀 Acciones Disponibles")
    
    col_descarga, col_correo = st.columns(2)
    
    with col_descarga:
        st.download_button(
            label="📥 Descargar copia en mi ordenador (Excel)",
            data=excel_final,
            file_name=nombre_del_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    with col_correo:
        if st.button("📧 Enviar informe directo a Marta Cuesta", use_container_width=True):
            with st.spinner("Enviando correo con el archivo adjunto..."):
                exito = enviar_correo_marta(excel_final, nombre_del_excel, rango_dias_texto)
                if exito:
                    st.success(f"📩 ¡Correo enviado con éxito a martacuesta@cecotec.es (Periodo: {rango_dias_texto})!")

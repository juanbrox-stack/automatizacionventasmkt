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

# Selectores de archivos en la interfaz web
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
        
        # Procesamiento rápido de fechas
        df['FECHA_CORTA'] = df['FECHA'].astype(str).str.slice(0, 10)
        df['FECHA_DATETIME'] = pd.to_datetime(df['FECHA_CORTA'], format='%Y-%m-%d', errors='coerce')
        
        # Calcular semana (Estilo Excel: Domingo a Sábado = %U)
        df['SEMANA'] = df['FECHA_DATETIME'].dt.strftime('%U').astype(float).fillna(0).astype(int) + 1
        
        # Formatear la fecha visual como dd/mm/aaaa
        df['FECHA'] = df['FECHA_DATETIME'].dt.strftime('%d/%m/%Y')
        
        df['REFERENCIA'] = df['REFERENCIA'].apply(limpiar_sku)
        
        return df

    except Exception as e:
        st.error(f"Error procesando los datos de {nombre_cuenta}: {str(e)}")
        return None

def enviar_correo_marta(excel_bytes, nombre_archivo, num_semana):
    try:
        remitente = st.secrets["correo"]["usuario"]
        password = st.secrets["correo"]["password"]
        smtp_server = st.secrets["correo"]["servidor_smtp"]
        puerto = int(st.secrets["correo"]["puerto"])
        
        destinatario = "martacuesta@cecotec.es"
        
        msg = MIMEMultipart()
        msg['From'] = remitente
        msg['To'] = destinatario
        msg['Subject'] = f"Informe de Ventas Amazon - Semana {num_semana}"
        
        cuerpo = f"""Hola Marta,

Adjunto el informe resumen de las ventas generadas en Amazon (cuentas Jabiru y Turaco) correspondientes a la semana anterior (Semana {num_semana}).

Un saludo.

"""
        msg.attach(MIMEText(cuerpo, 'plain'))
        
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(excel_bytes)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={nombre_archivo}')
        msg.attach(part)
        
        server = smtplib.SMTP(smtp_server, puerto)
        server.starttls()
        server.login(remitente, password)
        server.sendmail(remitente, destinatario, msg.as_string())
        server.quit()
        
        return True
    except Exception as e:
        st.error(f"Error al enviar el correo: {e}")
        return False

# Lógica de unificación y filtrado estricto por número de semana
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
    
    # --- FILTRADO DE SEMANA ANTERIOR DIRECTO ---
    # 1. Encontrar el número de semana más alto presente en el informe unificado
    semana_maxima = int(resultado_final['SEMANA'].max())
    
    # 2. Eliminar cualquier fila que pertenezca a una semana menor a la máxima
    resultado_final = resultado_final[resultado_final['SEMANA'] == semana_maxima]
    
    # Estructura de columnas requerida
    columnas_finales = ['FECHA', 'SEMANA', 'REFERENCIA', 'ASIN', 'CANTIDAD', 'IMPORTE TOTAL', 'CUENTA', 'ship-country']
    resultado_final = resultado_final[columnas_finales].dropna(subset=['FECHA'])
    
    # Ordenar cronológicamente por fecha
    resultado_final = resultado_final.sort_values(by=['FECHA'], ascending=[True])
    
    st.success(f"🎯 ¡Archivos procesados con éxito! Se ha limpiado el reporte dejando únicamente la **Semana {semana_maxima}**.")
    
    st.subheader(f"Vista previa del Informe Final (Semana {semana_maxima})")
    st.dataframe(resultado_final)
    
    # Generar Excel en memoria
    @st.cache_data
    def generar_excel_estandar(df):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Ventas MKT')
        return buffer.getvalue()
    
    excel_final = generar_excel_estandar(resultado_final)
    nombre_del_excel = f"Ventas_MKT_Semana_{semana_maxima}.xlsx"
    
    # Secciones de descargas y envíos
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
                exito = enviar_correo_marta(excel_final, nombre_del_excel, semana_maxima)
                if exito:
                    st.success(f"📩 ¡Correo enviado con éxito a Marketing!")

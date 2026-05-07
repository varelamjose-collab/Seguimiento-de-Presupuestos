import streamlit as st
import pandas as pd
from datetime import datetime
import io
import base64
import smtplib
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
from PIL import Image

st.set_page_config(page_title="Seguimiento Presupuestos", page_icon="💰", layout="centered")

st.title("💰 Seguimiento de Presupuestos")

# ==================== CONFIGURACIÓN ====================
CORREO_DESTINO = "fmo@fundacionmasaveu.com"  # 👈 CAMBIAR POR EL CORREO DESEADO
# =======================================================

# Barra lateral
st.sidebar.header("👷 Datos del trabajador")
nombre = st.sidebar.text_input("Tu nombre")
fecha = st.sidebar.date_input("Fecha", datetime.now())

# Inicializar datos en memoria
if 'datos_presupuesto' not in st.session_state:
    st.session_state.datos_presupuesto = pd.DataFrame(columns=[
        "Numero_Albaran",
        "Fecha",
        "Trabajador",
        "Gasto_Euros",
        "Comentarios",
        "Foto_Nombre"
    ])

# Inicializar fotos
if 'fotos_data' not in st.session_state:
    st.session_state.fotos_data = {}

# ==================== FORMULARIO ====================
st.subheader("📝 Registrar nuevo gasto")

with st.form("form_presupuesto"):
    num_albaran = st.text_input("📄 Número de albarán *", placeholder="Ej: ALB-2024-001")
    
    gasto = st.number_input("💶 Importe del gasto (€)", min_value=0.01, step=0.01, format="%.2f")
    
    comentarios = st.text_area("📝 Comentarios", placeholder="Descripción del gasto...")
    
    # Extra: subir foto
    st.markdown("---")
    st.markdown("📸 **Foto del albarán (opcional)**")
    foto_subida = st.file_uploader("Selecciona una imagen", type=["jpg", "jpeg", "png"])
    
    foto_bytes = None
    if foto_subida is not None:
        foto_bytes = foto_subida.getvalue()
        st.image(foto_subida, caption="Vista previa", width=150)
    
    enviar = st.form_submit_button("✅ Registrar gasto")
    
    if enviar:
        if not nombre:
            st.error("❌ Escribe tu nombre en el menú lateral")
        elif not num_albaran:
            st.error("❌ El número de albarán es obligatorio")
        elif gasto <= 0:
            st.error("❌ El importe debe ser mayor que 0")
        else:
            nombre_foto = ""
            if foto_bytes is not None:
                nombre_foto = f"albaran_{num_albaran}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
                st.session_state.fotos_data[nombre_foto] = foto_bytes
            
            nuevo_registro = pd.DataFrame([{
                "Numero_Albaran": num_albaran,
                "Fecha": fecha.strftime("%Y-%m-%d"),
                "Trabajador": nombre,
                "Gasto_Euros": gasto,
                "Comentarios": comentarios,
                "Foto_Nombre": nombre_foto
            }])
            
            st.session_state.datos_presupuesto = pd.concat([st.session_state.datos_presupuesto, nuevo_registro], ignore_index=True)
            
            st.success(f"✅ Gasto registrado - Albarán: {num_albaran}")
            st.balloons()

# ==================== MOSTRAR REGISTROS ====================
st.markdown("---")
st.subheader("📋 Registros de gastos")

if len(st.session_state.datos_presupuesto) > 0:
    # Mostrar tabla
    df_mostrar = st.session_state.datos_presupuesto.drop(columns=['Foto_Nombre'], errors='ignore')
    st.dataframe(df_mostrar, use_container_width=True)
    
    # Total gastado
    total_gastado = st.session_state.datos_presupuesto['Gasto_Euros'].sum()
    st.metric("💰 Total gastado", f"{total_gastado:,.2f} €")
    
    # Mostrar fotos si hay
    if len(st.session_state.fotos_data) > 0:
        st.subheader("📸 Fotos de albaranes")
        for foto_nombre, foto_bytes in st.session_state.fotos_data.items():
            albaran_asociado = st.session_state.datos_presupuesto[
                st.session_state.datos_presupuesto['Foto_Nombre'] == foto_nombre
            ]['Numero_Albaran'].values
            albaran_texto = albaran_asociado[0] if len(albaran_asociado) > 0 else "Desconocido"
            
            with st.expander(f"📄 Albarán: {albaran_texto}"):
                st.image(foto_bytes, use_container_width=True)
    
    # ==================== FUNCIÓN CREAR EXCEL ====================
    def crear_excel():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export = st.session_state.datos_presupuesto.copy()
            df_export = df_export.drop(columns=['Foto_Nombre'], errors='ignore')
            df_export.to_excel(writer, index=False, sheet_name='Gastos')
        return output.getvalue()
    
    # ==================== BOTONES ====================
    col1, col2 = st.columns(2)
    
    with col1:
        excel_data = crear_excel()
        b64 = base64.b64encode(excel_data).decode()
        fecha_str = datetime.now().strftime("%Y%m%d")
        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="presupuesto_{fecha_str}.xlsx">📥 Descargar Excel</a>'
        st.markdown(href, unsafe_allow_html=True)
    
    with col2:
        if st.button("📧 Enviar por correo", type="primary"):
            try:
                EMAIL_REMITENTE = "fmo@fundacionmasaveu.com"  # 👈 CAMBIAR
                # Usar secrets o poner contraseña directamente
                EMAIL_PASSWORD = st.secrets["gmail"]["password"] if "secrets" in dir(st) else "TU_CONTRASEÑA_APP"
                
                msg = MIMEMultipart()
                msg['From'] = EMAIL_REMITENTE
                msg['To'] = CORREO_DESTINO
                msg['Subject'] = f"Informe Presupuestos - {datetime.now().strftime('%d/%m/%Y')}"
                
                cuerpo = f"""
INFORME DE SEGUIMIENTO DE PRESUPUESTOS

Fecha de envio: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
Registros totales: {len(st.session_state.datos_presupuesto)}
Total gastado: {total_gastado:,.2f} €
Trabajador: {nombre}

Adjunto encontraras el Excel con todos los gastos.
"""
                msg.attach(MIMEText(cuerpo.encode('utf-8'), 'plain', 'utf-8'))
                
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(excel_data)
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="presupuesto_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx"')
                msg.attach(part)
                
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.starttls()
                server.login(EMAIL_REMITENTE, EMAIL_PASSWORD)
                server.send_message(msg)
                server.quit()
                
                st.success(f"✅ Correo enviado a {CORREO_DESTINO}")
                st.balloons()
                
            except Exception as e:
                st.error(f"Error: {str(e)}")
    
    # Botón limpiar
    if st.button("🗑️ Borrar todos los registros"):
        st.session_state.datos_presupuesto = pd.DataFrame(columns=[
            "Numero_Albaran", "Fecha", "Trabajador", "Gasto_Euros", "Comentarios", "Foto_Nombre"
        ])
        st.session_state.fotos_data = {}
        st.success("Registros borrados")
        st.rerun()

else:
    st.info("ℹ️ Aún no hay registros. Completa el formulario para comenzar.")

# ==================== PIE ====================
st.markdown("---")
st.caption("💰 App de Seguimiento de Presupuestos | PEDRO MASAVEU, 18 | OVIEDO | 33007")

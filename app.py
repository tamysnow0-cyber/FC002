import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import io

st.set_page_config(page_title="Generador de Oficios", page_icon="🚗")
st.title("🚗 Generador Automático de Oficios de Comisión")
st.write("Llena los siguientes datos para descargar tu oficio.")

@st.cache_data
def cargar_datos():
    return pd.read_excel('Información del empleado_FINALd.xlsx')

try:
    df = cargar_datos()
except Exception as e:
    st.error("❌ No se pudo cargar la base de datos. Verifica los archivos.")
    st.stop()

with st.form("formulario_oficio"):
    num_empleado = st.number_input("1. Número de Empleado:", min_value=1, step=1, format="%d")
    placa_input = st.text_input("2. Placas de la unidad (Ej. HM4036G):").strip().upper()
    lugar_input = st.text_input("3. ¿Lugar y fecha? (Ej. MUNICIPIO,REGIÓN.LOCALIDAD,ETC):").strip()
    
    # --- CAJÓN DE OPCIONES PARA EL MOTIVO ---
    opciones_motivo = [
        "Realizar trámites, levantamiento de información para dictamen técnico",
        "Elaborar constancias de operación",
        "Reunión con concesionarios",
        "Asistencia a ruta de Transformación",
        "Cursos para operadores",
        "Asistencia a Mesas de acercamiento a la paz",
        "Asistencia a Reunión",
        "Dejar correspondencia",
        "Otro (escribir manualmente)"
    ]
    motivo_seleccion = st.selectbox("4. Selecciona la finalidad de la comisión:", opciones_motivo)
    
    if motivo_seleccion == "Otro (escribir manualmente)":
        motivo_input = st.text_input("Escribe la finalidad de la comisión:").strip()
    else:
        motivo_input = motivo_seleccion
    # ----------------------------------------
    
    hora_salida = st.time_input("5. Hora de salida:")
    
    # --- NUEVO: SELECTOR DE FIRMANTE ---
    opciones_firmante = [
        "Ing. Sandra Saraí Hernández López",
        "Dr. José Antonio Pérez Sánchez"
    ]
    firmante_seleccion = st.selectbox("6. Selecciona quién autoriza (Firmante):", opciones_firmante)
    # -----------------------------------
    
    generar = st.form_submit_button("Generar Oficio")

if generar:
    if not num_empleado or not placa_input or not lugar_input or not motivo_input:
        st.warning("⚠️ Por favor, llena todos los campos.")
    else:
        empleado_data = df[df['No. empleado'] == num_empleado]
        vehiculo_data = df[df['placa'].astype(str).str.upper() == placa_input]
        
        if empleado_data.empty:
            st.error(f"⚠️ No se encontró el empleado {num_empleado}.")
        elif vehiculo_data.empty:
            st.error(f"⚠️ No se encontró la placa '{placa_input}'.")
        else:
            datos_emp = empleado_data.iloc[0]
            datos_veh = vehiculo_data.iloc[0]
            
            # --- DETECTOR DE GÉNERO ---
            #genero_excel = str(datos_emp.get('Género', 'M')).strip().upper()
            #if genero_excel in ['F', 'MUJER', 'FEMENINO']:
                #palabra_genero = "comisionada"
            #else:
                #palabra_genero = "comisionado"
            
            meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
            ahora = datetime.now()
            fecha_larga = f"{ahora.day} de {meses[ahora.month - 1]} de {ahora.year}"
            hora_formateada = hora_salida.strftime("%H:%M")
            
            modelo = str(datos_veh['Modelo'])
            if modelo.endswith('.0'):
                modelo = modelo[:-2]
                
            # --- CONFIGURACIÓN DINÁMICA DEL FIRMANTE ---
            if firmante_seleccion == "Ing. Sandra Saraí Hernández López":
                nombre_firmante = "Ing. Sandra Saraí Hernández López"
                cargo_firmante = "Directora de Movilidad e Ingeniería del\nSistema de Transporte Convencional de Hidalgo"
            else:
                nombre_firmante = "Dr. José Antonio Pérez Sánchez"
                cargo_firmante = "Director General del Sistema de Transporte\nConvencional de Hidalgo"
            # -------------------------------------------

            contexto = {
                'Fecha': fecha_larga, 
                'Nombre': str(datos_emp['Nombre']),
                'Adscripcion': str(datos_emp['Adscripción']),
                'Puesto': str(datos_emp['Puesto']),
                'Nombramiento': str(datos_emp['Nombramiento']),
                'RFC': str(datos_emp['RFC']),
                'Unidad': str(datos_veh['Unidad']),
                'Modelo': modelo,
                'Placa': str(datos_veh['placa']), 
                'Lugar_Fecha': lugar_input,
                'Motivo': motivo_input, 
                'Hora_Salida': hora_formateada,
                'Comisionado':str(datos_emp['Comisionado']),
                'Nombre_Firmante': nombre_firmante,
                'Cargo_Firmante': cargo_firmante
            }
            
            try:
                doc = DocxTemplate('OFICIO COMISIÓN 2026_finald.docx')
                doc.render(contexto)
                
                bio = io.BytesIO()
                doc.save(bio)
                nombre_archivo_salida = f"Oficio_{datos_emp['Nombre'].replace(' ', '_')}_{placa_input}.docx"
                
                st.success("✅ ¡Oficio generado exitosamente!")
                st.download_button(
                    label="📥 Descargar Documento Word",
                    data=bio.getvalue(),
                    file_name=nombre_archivo_salida,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"❌ Ocurrió un error: {e}")

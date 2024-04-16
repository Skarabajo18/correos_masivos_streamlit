import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import random
import time  # Importa el módulo time

# Inicializar el estado de la sesión si aún no existe
if 'saludos' not in st.session_state:
    st.session_state['saludos'] = ['']
if 'despedidas' not in st.session_state:
    st.session_state['despedidas'] = ['']

# Función para añadir un nuevo saludo
def add_saludo():
    st.session_state['saludos'].append('')

# Función para añadir una nueva despedida
def add_despedida():
    st.session_state['despedidas'].append('')

# Función para enviar correos
def enviar_correos(email_address, email_password, df, saludos, despedidas, intervalo):
    smtp_server = 'smtp.office365.com'
    smtp_port = 587

    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(email_address, email_password)

    for index, row in df.iterrows():
        destinatario = str(row['CorreoDestinatario'])
        cc = str(row['CC'])
        cc_list = [c.strip() for c in cc.split(';')] if cc.lower() != 'nan' else []

        if destinatario.lower() != 'nan' and '@' in destinatario:
            asunto = row['Asunto']
            descripcion = str(row['Descripción']).replace('_x000D_', '')

            saludo = random.choice(saludos)
            despedida = random.choice(despedidas)
            cuerpo_correo = f"{saludo}\n\n{descripcion}\n\n{despedida}"

            mensaje = MIMEMultipart()
            mensaje['From'] = email_address
            mensaje['To'] = destinatario
            mensaje['CC'] = ', '.join(cc_list)
            mensaje['Subject'] = asunto

            mensaje.attach(MIMEText(cuerpo_correo, 'plain'))

            all_recipients = [destinatario] + cc_list

            server.sendmail(email_address, all_recipients, mensaje.as_string())

            # Espera por el intervalo de tiempo especificado antes de enviar el próximo correo
            time.sleep(intervalo)

    server.quit()
    st.success('Correos enviados exitosamente!')

# Streamlit UI
st.title('Envío de correos masivos')

email_address = st.text_input('Dirección de correo electrónico', '')
email_password = st.text_input('Contraseña', type='password')

st.write('Saludos:')
for i in range(len(st.session_state['saludos'])):
    st.session_state['saludos'][i] = st.text_area(f'Saludo {i+1}', value=st.session_state['saludos'][i], key=f'saludo{i}', height=150)

st.button('Añadir otro saludo', on_click=add_saludo)

st.write('Despedidas:')
for i in range(len(st.session_state['despedidas'])):
    st.session_state['despedidas'][i] = st.text_area(f'Despedida {i+1}', value=st.session_state['despedidas'][i], key=f'despedida{i}', height=150)

st.button('Añadir otra despedida', on_click=add_despedida)

# Selector de intervalo de tiempo
intervalo = st.number_input('Intervalo de tiempo entre correos (en segundos)', min_value=1, value=30)

uploaded_file = st.file_uploader("Elige un archivo Excel", type=['xlsx'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write('Vista previa de los datos:')
    st.dataframe(df.head())

if st.button('Enviar correos'):
    if not email_address or not email_password:
        st.error('Por favor, ingresa una dirección de correo electrónico y una contraseña válidas.')
    elif not st.session_state['saludos'][0] or not st.session_state['despedidas'][0]:
        st.error('Por favor, introduce al menos un saludo y una despedida.')
    else:
        # Incluye el intervalo seleccionado al llamar a la función enviar_correos
        enviar_correos(email_address, email_password, df, st.session_state['saludos'], st.session_state['despedidas'], intervalo)

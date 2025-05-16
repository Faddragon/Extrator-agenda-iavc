import streamlit as st
import pandas as pd
import os
import json
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from io import BytesIO

# Configurações
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
CALENDAR_ID = 'jh8dpotn9etu9o231tlldvn6ms@group.calendar.google.com'

# Autenticação com conta de serviço
def autenticar_google():
    os.makedirs(".streamlit", exist_ok=True)
    with open(".streamlit/credentials.json", "w") as f:
        json.dump(st.secrets["google_credentials"].to_dict(), f)

    credentials = service_account.Credentials.from_service_account_file(
        ".streamlit/credentials.json",
        scopes=SCOPES
    )
    return build("calendar", "v3", credentials=credentials)

# Buscar eventos
def buscar_eventos(service, data_inicio_iso, data_fim_iso):
    eventos_formatados = []
    page_token = None

    while True:
        eventos = service.events().list(
            calendarId=CALENDAR_ID,
            timeMin=data_inicio_iso + 'T00:00:00Z',
            timeMax=data_fim_iso + 'T23:59:59Z',
            maxResults=2500,
            singleEvents=True,
            orderBy='startTime',
            pageToken=page_token
        ).execute()

        for evento in eventos.get('items', []):
            titulo = evento.get('summary', '').strip()
            if titulo in ['Modelo agendamento', 'Dados do hospital']:
                continue
            eventos_formatados.append({
                'Resumo': titulo,
                'Início': evento.get('start').get('dateTime', evento.get('start').get('date')),
                'Término': evento.get('end').get('dateTime', evento.get('end').get('date')),
                'Local': evento.get('location'),
                'Descrição': evento.get('description')
            })

        page_token = eventos.get('nextPageToken')
        if not page_token:
            break

    return eventos_formatados

# Interface
st.set_page_config(page_title="📆 Extrator de Agenda IAVC", layout="wide")
st.title("📅 Extrator de Eventos do Google Agenda")

col1, col2 = st.columns(2)
with col1:
    data_inicio = st.date_input("📆 Data inicial", value=datetime.today())
with col2:
    data_fim = st.date_input("📆 Data final", value=datetime.today())

# Entrada da senha
senha = st.text_input("🔐 Digite a senha para liberar o download:", type="password")

if st.button("🔍 Buscar e preparar arquivo"):
    with st.spinner("🔐 Autenticando e processando..."):
        try:
            service = autenticar_google()
            eventos = buscar_eventos(
                service,
                data_inicio.strftime('%Y-%m-%d'),
                data_fim.strftime('%Y-%m-%d')
            )

            if not eventos:
                st.warning("⚠️ Nenhum evento encontrado.")
            else:
                df = pd.DataFrame(eventos)
                for col in ['Início', 'Término']:
                    df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d/%m/%Y %H:%M')

                # Armazenar em buffer
                nome_arquivo = f"agenda_{data_inicio.strftime('%d-%m-%Y')}_a_{data_fim.strftime('%d-%m-%Y')}.xlsx"
                buffer = BytesIO()
                df.to_excel(buffer, index=False, engine='openpyxl')
                buffer.seek(0)

                if senha == "ccpiavc2025":
                    st.success("✅ Senha correta. Clique para baixar:")
                    st.download_button(
                        label="📥 Baixar como Excel",
                        data=buffer,
                        file_name=nome_arquivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ Senha incorreta. Digite a senha correta para liberar o download.")
        except Exception as e:
            st.error(f"❌ Erro: {e}")

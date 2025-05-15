import streamlit as st
import pandas as pd
import os
import pickle
import json
from datetime import datetime
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

# Configurações iniciais
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
CALENDAR_ID = 'jh8dpotn9etu9o231tlldvn6ms@group.calendar.google.com'

# 🔐 Função para autenticar com Google via st.secrets
def autenticar_google():
    creds = None

    if os.path.exists("token.pkl"):
        with open("token.pkl", "rb") as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Salva as credenciais do secret como arquivo temporário
            os.makedirs(".streamlit", exist_ok=True)
            with open(".streamlit/credentials.json", "w") as f:
                json.dump(st.secrets["google_credentials"], f)

            flow = InstalledAppFlow.from_client_secrets_file(".streamlit/credentials.json", SCOPES)
            creds = flow.run_console()

        with open("token.pkl", "wb") as token:
            pickle.dump(creds, token)

    return build("calendar", "v3", credentials=creds)

# 📅 Função para buscar eventos
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

# 🎨 Interface Streamlit
st.set_page_config(page_title="📆 Extrator de Agenda IAVC", layout="wide")

# Exibir logotipo se houver
if os.path.exists("logo_iavc.png"):
    st.image("logo_iavc.png", width=200)

st.title("📅 Extrator de Eventos do Google Agenda")

# Seletor de datas
col1, col2 = st.columns(2)
with col1:
    data_inicio = st.date_input("📆 Data inicial", value=datetime.today())
with col2:
    data_fim = st.date_input("📆 Data final", value=datetime.today())

# Botão principal
if st.button("🔍 Buscar eventos"):
    with st.spinner("Autenticando e coletando eventos da agenda..."):
        try:
            service = autenticar_google()
            eventos = buscar_eventos(
                service,
                data_inicio.strftime('%Y-%m-%d'),
                data_fim.strftime('%Y-%m-%d')
            )

            if eventos:
                df = pd.DataFrame(eventos)

                # Formatar datas no estilo brasileiro
                for col in ['Início', 'Término']:
                    df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d/%m/%Y %H:%M')

                st.success(f"✅ {len(df)} eventos encontrados.")
                st.dataframe(df, use_container_width=True)

                buffer = df.to_excel(index=False, engine='openpyxl')
                nome_arquivo = f"agenda_{data_inicio.strftime('%d-%m-%Y')}_a_{data_fim.strftime('%d-%m-%Y')}.xlsx"
                st.download_button("📥 Baixar como Excel", data=buffer, file_name=nome_arquivo)
            else:
                st.warning("⚠️ Nenhum evento encontrado no período.")
        except Exception as e:
            st.error(f"Erro: {e}")

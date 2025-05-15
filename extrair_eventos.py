import pandas as pd
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from datetime import datetime

# Escopos necessários para ler a agenda
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

def autenticar_google():
    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
    creds = flow.run_local_server(port=0)
    return creds

def buscar_eventos(service, data_inicio_iso, data_fim_iso):
    eventos_formatados = []
    page_token = None

    while True:
        eventos = service.events().list(
            calendarId='jh8dpotn9etu9o231tlldvn6ms@group.calendar.google.com',
            timeMin=data_inicio_iso + 'T00:00:00Z',
            timeMax=data_fim_iso + 'T23:59:59Z',
            maxResults=2500,  # Aumentado o limite
            singleEvents=True,
            orderBy='startTime',
            pageToken=page_token
        ).execute()

        for evento in eventos.get('items', []):
            titulo = evento.get('summary', '').strip()

            # Ignorar eventos com esses títulos
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

def main():
    print("📆 Extração de eventos do Google Agenda")
    data_inicio_str = input("🔹 Digite a data inicial (DD-MM-AAAA): ").strip()
    data_fim_str = input("🔹 Digite a data final   (DD-MM-AAAA): ").strip()
    
    try:
        data_inicio = datetime.strptime(data_inicio_str, "%d-%m-%Y")
        data_fim = datetime.strptime(data_fim_str, "%d-%m-%Y")
    except ValueError:
        print("❌ Datas inválidas. Use o formato DD-MM-AAAA.")
        return

    # Converte para o formato ISO que a API espera
    data_inicio_iso = data_inicio.strftime("%Y-%m-%d")
    data_fim_iso = data_fim.strftime("%Y-%m-%d")

    print("🔐 Autenticando...")
    creds = autenticar_google()
    service = build('calendar', 'v3', credentials=creds)

    print("📥 Buscando eventos...")
    eventos = buscar_eventos(service, data_inicio_iso, data_fim_iso)

    if not eventos:
        print("⚠️ Nenhum evento encontrado no período.")
    else:
        df = pd.DataFrame(eventos)
        nome_arquivo = f"agenda_{data_inicio_str}_a_{data_fim_str}.xlsx"
        df.to_excel(nome_arquivo, index=False)
        print(f"✅ Planilha gerada com sucesso: {nome_arquivo}")

if __name__ == '__main__':
    main()

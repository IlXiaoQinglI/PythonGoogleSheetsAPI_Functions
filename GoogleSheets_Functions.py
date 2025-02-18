from __future__ import print_function
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# Retorna o valor de determinada/as célula/as
def lersheet(link,page,cel):
    try:
        texto = str(sheet.values().get(spreadsheetId=link,
                                        range=f'{page}!{cel}').execute().get('values', ))
        return texto
    except HttpError as err:
        print(err)


# Option se refere às opções de preenchimento, 1 -> USER_ENTERED, 2 -> RAW,  3 -> INPUT_VALUE_OPTION_UNSPECIFIED
def mudarsheet(link,page,cel,text, option):
    while True:
        try:
            if option == 1:
                result = sheet.values().update(spreadsheetId=link,
                                               range=f'{page}!{cel}',
                                               valueInputOption='USER_ENTERED',
                                               body={'values': text}).execute()
            elif option == 2:
                result = sheet.values().update(spreadsheetId=link,
                                               range=f'{page}!{cel}',
                                               valueInputOption='RAW',
                                               body={'values': text}).execute()
            elif option == 3:
                result = sheet.values().update(spreadsheetId=link,
                                               range=f'{page}!{cel}',
                                               valueInputOption='INPUT_VALUE_OPTION_UNSPECIFIED',
                                               body={'values': text}).execute()
            values = str(result.get('values', []))
            break
        except HttpError as err:
            print(err)


# Realiza a conexão com a API
def sheetson():
    global creds
    global sheet
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Na váriavel LocalCreds, você deve inserir o caminho para a pasta onde fica localizado os arquivos token.json e credentials.json
    LocalCreds = r"C:\Caminho\Para\Pasta\Credenciaisgoogle"

    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(fr'{LocalCreds}\token.json'):
        creds = Credentials.from_authorized_user_file(
            fr'{LocalCreds}\token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                fr'{LocalCreds}\credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(fr'{LocalCreds}\token.json', 'w') as token:
            token.write(creds.to_json())
    try:
        service = build('sheets', 'v4', credentials=creds)
    except:
        DISCOVERY_SERVICE_URL = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
        service = build('sheets', 'v4', credentials=creds, discoveryServiceUrl=DISCOVERY_SERVICE_URL)
    sheet = service.spreadsheets()

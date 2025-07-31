import os.path
import pickle
import re
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from helpers import carregar_config

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def get_spreadsheet_id_from_url(url):
    if not url: return None
    match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', url)
    return match.group(1) if match else None

def get_service(app_interface=None):
    creds = None; config = carregar_config()
    credentials_path = config.get("credentials_path")
    if not credentials_path or not os.path.exists(credentials_path):
        if app_interface: app_interface.show_error_message("Ficheiro de Credenciais Não Definido", "Por favor, selecione o 'credentials.json' nas Opções.")
        else: print("ERRO: Ficheiro de Credenciais Não Definido. Configure no app.")
        return None
    
    token_filename = 'token_principal.pickle'
    token_path = os.path.join(os.path.dirname(credentials_path), token_filename)
    
    if os.path.exists(token_path):
        with open(token_path, 'rb') as token: creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try: 
                creds.refresh(Request())
            except Exception as e:
                log_msg = f"Erro ao atualizar token: {e}. Re-autorize."
                if app_interface: app_interface.log(log_msg)
                else: print(log_msg)
                creds = None
        if not creds:
            try:
                flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
                log_msg = "Uma janela do navegador será aberta para autorização."
                if app_interface: app_interface.show_info_message("Autorização Necessária", log_msg)
                else: print(log_msg)
                creds = flow.run_local_server(port=0)
            except Exception as e:
                log_msg = f"Falha na Autorização: {e}"
                if app_interface: app_interface.show_error_message("Falha na Autorização", log_msg)
                else: print(log_msg)
                return None
        with open(token_path, 'wb') as token: pickle.dump(creds, token)
    
    try: 
        return build('sheets', 'v4', credentials=creds)
    except Exception as e:
        log_msg = f"Falha ao construir o serviço da API: {e}"
        if app_interface: app_interface.show_error_message("Erro de Serviço", log_msg)
        else: print(log_msg)
        return None

def get_service_for_url(url, app_interface=None):
    spreadsheet_id = get_spreadsheet_id_from_url(url)
    if not spreadsheet_id:
        if app_interface: app_interface.show_error_message("URL Inválido", f"O URL fornecido não parece ser de uma planilha válida:\n{url}")
        return None, None
    service = get_service(app_interface)
    return service, spreadsheet_id

def get_sheet_names(service, spreadsheet_id, app_interface=None):
    if not service or not spreadsheet_id: return []
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        return [sheet['properties']['title'] for sheet in spreadsheet.get('sheets', [])]
    except HttpError as e:
        if app_interface: app_interface.show_error_message("Erro de API", f"Não foi possível obter os nomes das abas. Verifique o URL e permissões.\n\nErro: {e}")
        return []

def get_sheet_id(service, spreadsheet_id, sheet_name, app_interface=None):
    """Obtém o ID de uma aba pelo seu nome."""
    if not service or not spreadsheet_id: return None
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        for sheet in spreadsheet.get('sheets', []):
            if sheet['properties']['title'] == sheet_name:
                return sheet['properties']['sheetId']
        return None
    except HttpError as e:
        if app_interface: app_interface.log(f"Erro ao obter ID da aba '{sheet_name}': {e}")
        return None

def read_rows(service, spreadsheet_id, range_name, app_interface=None):
    if not service or not spreadsheet_id: return []
    try:
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        return result.get('values', [])
    except HttpError as err:
        if err.resp.status == 400 and 'Unable to parse range' in str(err):
            log_msg = f"Aviso: A aba '{range_name}' não foi encontrada. Retornando lista vazia."
            if app_interface: app_interface.log(log_msg)
            else: print(log_msg)
        else:
            log_msg = f"ERRO ao ler as linhas da aba '{range_name}': {err}"
            if app_interface: app_interface.log(log_msg)
            else: print(log_msg)
        return []
    except Exception as e:
        log_msg = f"ERRO ao ler as linhas da aba '{range_name}': {e}"
        if app_interface: app_interface.log(log_msg)
        else: print(log_msg)
        return []

def read_values(service, spreadsheet_id, range_name, app_interface=None):
    rows = read_rows(service, spreadsheet_id, f"'{range_name}'!A:A", app_interface)
    return [item for sublist in rows for item in sublist if item]

def append_values(service, spreadsheet_id, range_name, values, app_interface=None):
    if not service or not spreadsheet_id or not values: return False
    try:
        body = {'values': values}
        service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id, 
            range=range_name, 
            valueInputOption='USER_ENTERED', 
            insertDataOption='INSERT_ROWS', 
            body=body
        ).execute()
        return True
    except Exception as e:
        log_msg = f"ERRO ao adicionar dados na aba '{range_name}': {e}"
        if app_interface: app_interface.log(log_msg)
        else: print(log_msg)
        return False

def clear_sheet(service, spreadsheet_id, range_name, app_interface=None):
    """Limpa todos os dados de uma aba."""
    if not service or not spreadsheet_id: return False
    try:
        service.spreadsheets().values().clear(spreadsheetId=spreadsheet_id, range=range_name, body={}).execute()
        return True
    except Exception as e:
        log_msg = f"ERRO ao limpar a aba '{range_name}': {e}"
        if app_interface: app_interface.log(log_msg)
        else: print(log_msg)
        return False

def overwrite_values(service, spreadsheet_id, range_name, values, app_interface=None):
    if not service or not spreadsheet_id: return False
    try:
        log_msg_clear = f"--- Limpando a aba '{range_name}' para atualização..."
        if app_interface: app_interface.log(log_msg_clear)
        else: print(log_msg_clear)
        
        clear_sheet(service, spreadsheet_id, range_name, app_interface)
        
        if values:
            log_msg_write = f"--- Escrevendo {len(values)} registros atualizados em '{range_name}'..."
            if app_interface: app_interface.log(log_msg_write)
            else: print(log_msg_write)
            
            body = {'values': values}
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id, 
                range=f"'{range_name}'!A1", 
                valueInputOption='USER_ENTERED', 
                body=body
            ).execute()
        return True
    except Exception as e:
        log_msg = f"ERRO ao sobrescrever dados na aba '{range_name}': {e}"
        if app_interface: app_interface.log(log_msg)
        else: print(log_msg)
        return False

def prepare_update_cell_request(sheet_name, row_index, col_index, new_value):
    col_letter = chr(64 + col_index)
    return {
        'range': f"'{sheet_name}'!{col_letter}{row_index}",
        'values': [[new_value]]
    }

def batch_update_values(service, spreadsheet_id, update_requests, app_interface=None):
    if not service or not spreadsheet_id or not update_requests: return False
    try:
        body = {'valueInputOption': 'USER_ENTERED', 'data': update_requests}
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
        return True
    except Exception as e:
        if app_interface: app_interface.log(f"ERRO ao executar atualização em lote: {e}")
        return False

def delete_rows_by_value(service, spreadsheet_id, sheet_name, col_to_check_index, values_to_delete, app_interface=None):
    """Deleta linhas de uma aba se o valor em uma coluna específica corresponde."""
    if not service or not spreadsheet_id or not values_to_delete: return False
    
    sheet_id = get_sheet_id(service, spreadsheet_id, sheet_name, app_interface)
    if sheet_id is None:
        if app_interface: app_interface.log(f"ERRO: Não foi possível encontrar o ID da aba '{sheet_name}' para deletar linhas.")
        return False

    if app_interface: app_interface.log(f"Lendo a aba '{sheet_name}' para encontrar linhas a serem deletadas...")
    all_rows = read_rows(service, spreadsheet_id, sheet_name, app_interface)
    if not all_rows: return True # Nada a fazer

    requests = []
    # Itera de trás para frente para que os índices não mudem após uma exclusão
    for i in range(len(all_rows) - 1, -1, -1):
        row = all_rows[i]
        if row and len(row) > col_to_check_index and row[col_to_check_index] in values_to_delete:
            requests.append({
                "deleteDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": i,
                        "endIndex": i + 1
                    }
                }
            })
    
    if not requests:
        if app_interface: app_interface.log("Nenhuma linha correspondente encontrada para deletar.")
        return True

    try:
        if app_interface: app_interface.log(f"Enviando pedido para deletar {len(requests)} linhas...")
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': requests}
        ).execute()
        if app_interface: app_interface.log("Linhas deletadas com sucesso.")
        return True
    except Exception as e:
        if app_interface: app_interface.log(f"ERRO ao deletar linhas da aba '{sheet_name}': {e}")
        return False

def ensure_sheets_exist(service, spreadsheet_id, sheet_names, app_interface=None):
    if not service or not spreadsheet_id: return
    try:
        existing_sheets = get_sheet_names(service, spreadsheet_id, app_interface)
        requests = [{'addSheet': {'properties': {'title': name}}} for name in sheet_names if name not in existing_sheets]
        if requests:
            titles = [req['addSheet']['properties']['title'] for req in requests]
            log_msg = f"Criando abas necessárias: {titles}"
            if app_interface: app_interface.log(log_msg)
            else: print(log_msg)
            service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={'requests': requests}).execute()
    except Exception as e:
        log_msg = f"ERRO ao verificar/criar abas: {e}"
        if app_interface: app_interface.log(log_msg)
        else: print(log_msg)

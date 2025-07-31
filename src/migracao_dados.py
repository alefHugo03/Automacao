import sys
import time

sys.path.append('src')

import sheets_integration
from helpers import carregar_config

# --- CONFIGURAÇÃO ---
ABA_CONTROLE_GERAL = "ControleGeral"
ABAS_ANTIGAS_STATUS = {"liberados": "LB", "bloqueados": "BQ"}
ABA_OUTROS = "outros_status"
ABA_MAILING = "enviados_maling"

# ALTERADO: Mapeamento de colunas atualizado para a nova estrutura
COL_TELEFONE = 0
COL_STATUS = 1
COL_DATA_VERIFICACAO = 2
COL_STATUS_SECUNDARIO = 3 # NOVO
COL_DATA_MAILING = 4       # Índice alterado
COL_OUTROS = 5             # Índice alterado

TAMANHO_LOTE = 2000 

def main():
    print("--- INICIANDO SCRIPT DE MIGRAÇÃO DE DADOS (V3 - Com Status Secundário) ---")
    print("AVISO: Este script irá limpar e reescrever a aba 'ControleGeral' com a nova estrutura de colunas.")
    print("Certifique-se de ter um backup da sua planilha.")
    
    confirmacao = input("Deseja continuar? (s/n): ")
    if confirmacao.lower() != 's':
        print("Migração cancelada pelo usuário.")
        return

    try:
        print("\n[PASSO 1/4] Conectando à API do Google Sheets...")
        service = sheets_integration.get_service()
        if not service: return
        config = carregar_config()
        spreadsheet_id = sheets_integration.get_spreadsheet_id_from_url(config.get("spreadsheet_url"))
        if not spreadsheet_id: return
        print("Conexão estabelecida.")

        print("\n[PASSO 2/4] Lendo dados das abas antigas...")
        dados_consolidados = {}
        
        # O processo de leitura é o mesmo
        for nome_aba, status_cod in ABAS_ANTIGAS_STATUS.items():
            print(f"Lendo '{nome_aba}'...")
            rows = sheets_integration.read_rows(service, spreadsheet_id, nome_aba)
            if not rows: continue
            for row in rows:
                if not row or len(row) < 2: continue
                telefone, data_verificacao = row[0], row[1]
                if telefone not in dados_consolidados: dados_consolidados[telefone] = {}
                dados_consolidados[telefone]['status'] = status_cod
                dados_consolidados[telefone]['data_verificacao'] = data_verificacao

        print(f"Lendo '{ABA_OUTROS}'...")
        rows = sheets_integration.read_rows(service, spreadsheet_id, ABA_OUTROS)
        if rows:
            for row in rows:
                if not row or len(row) < 1: continue
                telefone_completo = row[0]
                telefone = telefone_completo.split(" - ")[0]
                outro_status = telefone_completo.split("Status: ")[1] if "Status: " in telefone_completo else telefone_completo
                if telefone not in dados_consolidados: dados_consolidados[telefone] = {}
                dados_consolidados[telefone]['outros_status'] = outro_status

        print(f"Lendo '{ABA_MAILING}'...")
        rows = sheets_integration.read_rows(service, spreadsheet_id, ABA_MAILING)
        if rows:
            for row in rows:
                if not row or len(row) < 2: continue
                telefone, data_mailing = row[0], row[1]
                if telefone not in dados_consolidados: dados_consolidados[telefone] = {}
                dados_consolidados[telefone]['data_mailing'] = data_mailing

        print(f"Total de {len(dados_consolidados)} registros únicos encontrados.")

        print("\n[PASSO 3/4] Preparando dados para gravação na nova estrutura...")
        # ALTERADO: Novo cabeçalho
        header = ['Telefone', 'Status', 'Data Verificacao', 'Status Secundario', 'Data Mailing', 'Outros Status']
        dados_para_escrever = []

        for telefone, dados in dados_consolidados.items():
            nova_linha = [''] * len(header)
            nova_linha[COL_TELEFONE] = telefone
            nova_linha[COL_STATUS] = dados.get('status', '')
            nova_linha[COL_DATA_VERIFICACAO] = dados.get('data_verificacao', '')
            nova_linha[COL_DATA_MAILING] = dados.get('data_mailing', '')
            nova_linha[COL_OUTROS] = dados.get('outros_status', '')

            # NOVO: Lógica para preencher o Status Secundário
            if dados.get('status') == 'BQ':
                nova_linha[COL_STATUS_SECUNDARIO] = 'bloqueado'
            elif dados.get('data_mailing'): # Se existe data de mailing, foi enviado
                nova_linha[COL_STATUS_SECUNDARIO] = 'enviado'

            dados_para_escrever.append(nova_linha)
        
        print(f"{len(dados_para_escrever)} linhas de dados prontas.")

        print(f"\n[PASSO 4/4] Escrevendo dados na aba '{ABA_CONTROLE_GERAL}' em lotes...")
        sheets_integration.ensure_sheets_exist(service, spreadsheet_id, [ABA_CONTROLE_GERAL])
        print(f"Limpando a aba '{ABA_CONTROLE_GERAL}'...")
        sheets_integration.clear_sheet(service, spreadsheet_id, ABA_CONTROLE_GERAL)
        print("Escrevendo o cabeçalho...")
        sheets_integration.append_values(service, spreadsheet_id, ABA_CONTROLE_GERAL, [header])

        total_escrito = 0
        for i in range(0, len(dados_para_escrever), TAMANHO_LOTE):
            lote = dados_para_escrever[i:i + TAMANHO_LOTE]
            num_lote = (i // TAMANHO_LOTE) + 1
            print(f"Escrevendo lote {num_lote}...")
            sucesso_lote = sheets_integration.append_values(service, spreadsheet_id, ABA_CONTROLE_GERAL, lote)
            if sucesso_lote:
                total_escrito += len(lote)
                print(f"Lote {num_lote} OK. Total: {total_escrito}.")
                time.sleep(1.5)
            else:
                print(f"ERRO: Falha ao escrever o lote {num_lote}.")
                return

        print("\n--- MIGRAÇÃO CONCLUÍDA COM SUCESSO! ---")

    except Exception as e:
        print(f"\nOcorreu um erro inesperado: {e}")

if __name__ == "__main__":
    main()
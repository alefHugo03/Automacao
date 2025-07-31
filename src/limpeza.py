# src/limpeza.py
import sheets_integration
from helpers import carregar_config

# Constantes para manter a consistência
ABA_CONTROLE_GERAL = "ControleGeral"
COL_TELEFONE = 0

def limpar_duplicatas_controle_geral(app):
    """Limpa duplicatas da aba principal 'ControleGeral'."""
    app.log(f">>> [LIMPEZA AUTOMÁTICA] Verificando duplicatas em '{ABA_CONTROLE_GERAL}'...")
    service, spreadsheet_id = app.sheet_service, app.spreadsheet_id
    if not service or not spreadsheet_id:
        app.log("!!! [LIMPEZA] ERRO: Planilha principal não configurada.")
        return
    try:
        dados_completos = sheets_integration.read_rows(service, spreadsheet_id, ABA_CONTROLE_GERAL, app)
        if not dados_completos or len(dados_completos) < 2:
            app.log(f"    -> Nenhuma duplicata para limpar em '{ABA_CONTROLE_GERAL}'.")
            return
        
        header = dados_completos[0]
        dados_unicos = {}
        # Itera sobre os dados para manter apenas a última ocorrência de cada telefone
        for linha in dados_completos[1:]:
            if linha and len(linha) > COL_TELEFONE and linha[COL_TELEFONE]:
                dados_unicos[linha[COL_TELEFONE]] = linha
        
        dados_limpos = [header] + list(dados_unicos.values())
        removidos = len(dados_completos) - len(dados_limpos)
        
        if removidos > 0:
            app.log(f"    -> {removidos} duplicatas encontradas. Reescrevendo '{ABA_CONTROLE_GERAL}'...")
            sheets_integration.overwrite_values(service, spreadsheet_id, ABA_CONTROLE_GERAL, dados_limpos, app)
            app.log(f"    -> Limpeza de '{ABA_CONTROLE_GERAL}' concluída.")
        else:
            app.log(f"    -> Nenhuma duplicata encontrada em '{ABA_CONTROLE_GERAL}'.")

    except Exception as e:
        app.log(f"!!! [LIMPEZA] ERRO CRÍTICO durante a limpeza de '{ABA_CONTROLE_GERAL}': {e}")

def limpar_duplicatas_mailing(app):
    """Limpa duplicatas da planilha de mailing."""
    app.log(">>> [LIMPEZA AUTOMÁTICA] Verificando duplicatas na planilha de Mailing...")
    try:
        config = carregar_config()
        maling_url = config.get("maling_spreadsheet_url")
        if not maling_url:
            app.log("!!! [LIMPEZA] ERRO: URL da planilha de mailing não definida.")
            return

        service, spreadsheet_id = sheets_integration.get_service_for_url(maling_url, app)
        if not service or not spreadsheet_id:
            app.log("!!! [LIMPEZA] ERRO: Não foi possível conectar à planilha de mailing.")
            return

        nome_aba = "Sheet1"
        dados_completos = sheets_integration.read_rows(service, spreadsheet_id, nome_aba, app)
        
        if not dados_completos:
            app.log(f"    -> A aba '{nome_aba}' do mailing está vazia. Nenhuma duplicata para limpar.")
            return

        dados_unicos = {}
        for linha in dados_completos:
            if linha and len(linha) > 0 and linha[0]:
                dados_unicos[linha[0]] = linha
        
        dados_limpos = list(dados_unicos.values())
        removidos = len(dados_completos) - len(dados_limpos)

        if removidos > 0:
            app.log(f"    -> {removidos} duplicatas encontradas. Reescrevendo mailing...")
            sheets_integration.overwrite_values(service, spreadsheet_id, nome_aba, dados_limpos, app)
            app.log("    -> Limpeza da planilha de mailing concluída.")
        else:
            app.log("    -> Nenhuma duplicata encontrada na planilha de mailing.")

    except Exception as e:
        app.log(f"!!! [LIMPEZA] ERRO CRÍTICO durante a limpeza do mailing: {e}")

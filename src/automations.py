import os
import random
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from typing import TYPE_CHECKING, Dict, Any
from datetime import datetime

if TYPE_CHECKING:
    from main import App

import sheets_integration
import limpeza # NOVO: Importa o módulo de limpeza
from helpers import carregar_config

# --- CONSTANTES ---
ABA_CONTROLE_GERAL = "ControleGeral"
ABA_TELEFONES = "telefones"
PROCESSO_CLARO = "claro"
PROCESSO_ENVIAR_MALING = "maling"
UI_UPDATE_INTERVAL = 5

# Índices de coluna para a aba ControleGeral
COL_TELEFONE = 0
COL_STATUS = 1
COL_DATA_VERIFICACAO = 2
COL_STATUS_SECUNDARIO = 3
COL_DATA_MAILING = 4
COL_OUTROS = 5

ABAS_GERENCIADAS = [ABA_CONTROLE_GERAL, ABA_TELEFONES]

NOMES_FUNCIONARIOS = ["ALEF", "LARISSA", "HEVI", "TATHI"]
MESES_PT = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
    7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
}

def iniciar_navegador(headless=True, use_session=False, driver_path=None):
    """
    Função para iniciar o navegador Chrome com as configurações desejadas.
    """
    opcoes = webdriver.ChromeOptions()
    if headless:
        opcoes.add_argument("--headless=new")
    if use_session:
        user_data_path = os.path.join(os.path.expanduser("~"), "verificador_app_session")
        os.makedirs(user_data_path, exist_ok=True)
        opcoes.add_argument(f"user-data-dir={user_data_path}")
    opcoes.add_argument("--no-sandbox")
    opcoes.add_argument("--disable-dev-shm-usage")
    opcoes.add_argument("--disable-gpu")
    opcoes.add_argument("--window-size=1920,1080")
    opcoes.add_argument("--log-level=3")
    opcoes.add_argument("--disable-blink-features=AutomationControlled")
    opcoes.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
    opcoes.add_experimental_option('useAutomationExtension', False)
    
    try:
        service = Service() if not driver_path else Service(driver_path)
        return webdriver.Chrome(service=service, options=opcoes)
    except WebDriverException as e:
        if "failed to download" in str(e) or "unable to obtain" in str(e):
             raise Exception(f"Falha ao baixar/localizar o ChromeDriver automaticamente. Verifique sua conexão ou firewall.\nUse o botão 'Definir Caminho do ChromeDriver' nas configurações para apontar para o executável manualmente.\nErro: {e}")
        raise Exception(f"Falha ao iniciar o ChromeDriver. Verifique se o Google Chrome está instalado. Erro: {e}")

def carregar_dados_controle_geral(app: 'App') -> Dict[str, Dict[str, Any]]:
    """
    Lê a aba ControleGeral e a transforma em um dicionário para acesso rápido.
    """
    app.log(f">>> Lendo a base de dados principal: '{ABA_CONTROLE_GERAL}'...")
    dados_map = {}
    rows = sheets_integration.read_rows(app.sheet_service, app.spreadsheet_id, ABA_CONTROLE_GERAL, app)
    if not rows or len(rows) < 1:
        app.log(f">>> A aba '{ABA_CONTROLE_GERAL}' está vazia ou não foi encontrada.")
        return {}
    
    for i, row in enumerate(rows[1:]): # Pula o cabeçalho
        if not row or not row[COL_TELEFONE]:
            continue
        telefone = row[COL_TELEFONE]
        dados_map[telefone] = {
            "row_index": i + 2,
            "status": row[COL_STATUS] if len(row) > COL_STATUS else "",
            "data_verificacao": row[COL_DATA_VERIFICACAO] if len(row) > COL_DATA_VERIFICACAO else "",
            "data_mailing": row[COL_DATA_MAILING] if len(row) > COL_DATA_MAILING else "",
        }
    app.log(f">>> {len(dados_map)} registros carregados da base de dados.")
    return dados_map

def executar_verificacao(app: 'App', usuario, senha, headless, driver_path, stop_event):
    """
    Executa a automação de verificação de 'Não Perturbe' com reanálise e sincronização de status.
    """
    stats = {"liberados": 0, "bloqueados": 0, "outros": 0, "erros": 0, "pulados": 0, "processados": 0, "total": 0, "novos": 0, "reanalisados": 0}
    navegador = None
    try:
        # NOVO: Executa a limpeza antes de iniciar o processo
        limpeza.limpar_duplicatas_controle_geral(app)

        data_hoje_str = datetime.now().strftime('%d/%m/%Y')
        app.log(f">>> [{PROCESSO_CLARO.upper()}] Automação iniciada.")
        
        dados_controle = carregar_dados_controle_geral(app)
        
        app.log(f">>> Lendo a lista de entrada '{ABA_TELEFONES}' para adicionar números novos...")
        telefones_novos_raw = sheets_integration.read_values(app.sheet_service, app.spreadsheet_id, ABA_TELEFONES, app)
        telefones_novos = set([t for t in telefones_novos_raw if t and len(t) == 11])
        
        telefones_existentes = set(dados_controle.keys())
        telefones_para_verificar = list(telefones_existentes.union(telefones_novos))
        
        app.log(f">>> Total de {len(telefones_para_verificar)} números únicos a serem considerados para verificação.")

        stats["total"] = len(telefones_para_verificar)
        if stats["total"] == 0:
            app.log(">>> Nenhuma base de dados ou número novo para verificar.")
            app.show_info_message("Concluído", "Nenhum número para verificar.")
            return

        config = carregar_config()
        maling_url = config.get("maling_spreadsheet_url")
        service_destino, id_destino = None, None
        if maling_url:
            app.log(">>> Conectando à planilha de mailing para sincronização...")
            service_destino, id_destino = sheets_integration.get_service_for_url(maling_url, app)

        app.log(">>> Iniciando navegador e fazendo login...")
        navegador = iniciar_navegador(headless, False, driver_path)
        navegador.get("https://app.conexaoclarobrasil.com.br/controle/telefonia/consulta-lei-nao-perturbe")
        WebDriverWait(navegador, 20).until(EC.visibility_of_element_located(("id", "email"))).send_keys(usuario)
        navegador.find_element("id", "password").send_keys(senha)
        WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[type="submit"]'))).click()
        WebDriverWait(navegador, 20).until(EC.presence_of_element_located(("id", "input-telefones")))
        app.log(">>> Login realizado.")

        novas_linhas_para_adicionar = []
        pedidos_update_principal = []
        pedidos_delete_mailing = []
        
        for i, telefone in enumerate(telefones_para_verificar):
            if stop_event.is_set(): break
            stats["processados"] += 1
            
            reanalisar = False
            row_index_existente = None
            status_antigo = None

            if telefone in dados_controle:
                dados_existentes = dados_controle[telefone]
                status_antigo = dados_existentes.get("status")
                data_verificacao_str = dados_existentes.get("data_verificacao", "")
                
                if not data_verificacao_str: reanalisar = True
                else:
                    try:
                        if (datetime.now() - datetime.strptime(data_verificacao_str, '%d/%m/%Y')).days >= 30: reanalisar = True
                    except ValueError: reanalisar = True

                if not reanalisar:
                    stats["pulados"] += 1
                    if i % 100 == 0: app.log(f"--- PULANDO: {telefone} (verificado nos últimos 30 dias)")
                    if i % UI_UPDATE_INTERVAL == 0: app.atualizar_estatisticas(stats, PROCESSO_CLARO)
                    continue
                else:
                    app.log(f"--- REANALISANDO: {telefone} (verificação expirou)")
                    row_index_existente = dados_existentes['row_index']
            else:
                app.log(f"--- PROCESSANDO NOVO: {telefone}")

            try:
                campo = WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.ID, "input-telefones")))
                campo.clear(); campo.send_keys(telefone); time.sleep(0.3); campo.send_keys(Keys.ENTER)
                botao = WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '//button/span[text()="Consultar"]')))
                navegador.execute_script("arguments[0].click();", botao)
                time.sleep(0.8)
                linha_resultado = WebDriverWait(navegador, 30).until(EC.presence_of_element_located((By.XPATH, f"//tbody/tr[td[1][contains(text(), '{telefone}')]]")))
                situacao = linha_resultado.find_elements(By.TAG_NAME, "td")[2].text.strip().upper()
                app.log(f"    -> RESULTADO: {situacao}")

                status_final = ""
                status_secundario_final = ""
                outros_final = ""

                if "LIBERADO" in situacao or "LB" in situacao:
                    stats["liberados"] += 1
                    status_final = "LB"
                elif "BLOQUEADO" in situacao or "BQ" in situacao:
                    stats["bloqueados"] += 1
                    status_final = "BQ"
                    status_secundario_final = "bloqueado"
                else:
                    stats["outros"] += 1
                    outros_final = situacao
                
                if status_antigo == "LB" and status_final == "BQ" and service_destino:
                    app.log(f"    -> MUDANÇA DE STATUS: {telefone} de LIBERADO para BLOQUEADO. Removendo do mailing...")
                    pedidos_delete_mailing.append(telefone)

                if reanalisar and row_index_existente:
                    stats["reanalisados"] += 1
                    pedidos_update_principal.append(sheets_integration.prepare_update_cell_request(ABA_CONTROLE_GERAL, row_index_existente, COL_STATUS + 1, status_final))
                    pedidos_update_principal.append(sheets_integration.prepare_update_cell_request(ABA_CONTROLE_GERAL, row_index_existente, COL_DATA_VERIFICACAO + 1, data_hoje_str))
                    pedidos_update_principal.append(sheets_integration.prepare_update_cell_request(ABA_CONTROLE_GERAL, row_index_existente, COL_STATUS_SECUNDARIO + 1, status_secundario_final))
                    pedidos_update_principal.append(sheets_integration.prepare_update_cell_request(ABA_CONTROLE_GERAL, row_index_existente, COL_OUTROS + 1, outros_final))
                else:
                    stats["novos"] += 1
                    nova_linha = [""] * 6 
                    nova_linha[COL_TELEFONE] = telefone
                    nova_linha[COL_STATUS] = status_final
                    nova_linha[COL_DATA_VERIFICACAO] = data_hoje_str
                    nova_linha[COL_STATUS_SECUNDARIO] = status_secundario_final
                    nova_linha[COL_OUTROS] = outros_final
                    novas_linhas_para_adicionar.append(nova_linha)
            
            except Exception as e:
                stats["erros"] += 1
                app.log(f"    -> ERRO ao processar {telefone}: {str(e).splitlines()[0]}")
                navegador.refresh()
                WebDriverWait(navegador, 20).until(EC.presence_of_element_located(("id", "input-telefones")))

            if i % UI_UPDATE_INTERVAL == 0: app.atualizar_estatisticas(stats, PROCESSO_CLARO)
            time.sleep(random.uniform(0.5, 1.2))

        if novas_linhas_para_adicionar:
            app.log(f"\n>>> Enviando {len(novas_linhas_para_adicionar)} novos registros para '{ABA_CONTROLE_GERAL}'...")
            sheets_integration.append_values(app.sheet_service, app.spreadsheet_id, ABA_CONTROLE_GERAL, novas_linhas_para_adicionar, app)
        
        if pedidos_update_principal:
            app.log(f"\n>>> Atualizando {stats['reanalisados']} registros existentes em '{ABA_CONTROLE_GERAL}'...")
            sheets_integration.batch_update_values(app.sheet_service, app.spreadsheet_id, pedidos_update_principal, app)

        if pedidos_delete_mailing and service_destino:
            app.log(f"\n>>> Removendo {len(pedidos_delete_mailing)} números da planilha de mailing devido à mudança de status...")
            sheets_integration.delete_rows_by_value(service_destino, id_destino, "Sheet1", 0, pedidos_delete_mailing, app)

        if not stop_event.is_set():
            app.log(f">>> Limpando a aba de entrada '{ABA_TELEFONES}'...")
            sheets_integration.clear_sheet(app.sheet_service, app.spreadsheet_id, ABA_TELEFONES, app)
            app.show_info_message("Processo Concluído", f"Verificação finalizada!\n\nNovos: {stats['novos']}\nReanalisados: {stats['reanalisados']}\nPulados: {stats['pulados']}\nErros: {stats['erros']}")
            
    except Exception as e:
        app.log(f"!!! [{PROCESSO_CLARO.upper()}] ERRO CRÍTICO: {e}")
        app.show_error_message("Erro Crítico", f"Ocorreu um erro fatal:\n\n{e}")
    finally:
        if navegador: navegador.quit()
        app.atualizar_estatisticas(stats, PROCESSO_CLARO)
        app.log(">>> Automação finalizada.")
        app.resetar_controles_processo(PROCESSO_CLARO)

def executar_envio_maling(app: 'App', maling_url, stop_event):
    """
    Executa o envio de números liberados para a planilha de mailing.
    """
    stats = {"enviados": 0, "retidos": 0, "erros": 0, "processados": 0, "total": 0}
    try:
        # NOVO: Executa a limpeza antes de iniciar o processo
        limpeza.limpar_duplicatas_mailing(app)

        data_hoje = datetime.now()
        data_hoje_str = data_hoje.strftime('%d/%m/%Y')
        app.log(f">>> [{PROCESSO_ENVIAR_MALING.upper()}] Processo iniciado.")

        app.log(">>> Conectando à planilha de destino (mailing)...")
        service_destino, id_destino = sheets_integration.get_service_for_url(maling_url, app)
        if not service_destino or not id_destino: return

        app.log(">>> Lendo dados da planilha de mailing (destino)...")
        dados_mailing_destino = sheets_integration.read_rows(service_destino, id_destino, "Sheet1!A:B", app)
        mapa_destino = {row[0]: {'row_index': i + 1, 'funcionario': row[1].strip().upper()} for i, row in enumerate(dados_mailing_destino) if row and len(row) > 1 and row[1]}
        app.log(f">>> {len(mapa_destino)} leads com atendentes encontrados na planilha de mailing.")

        dados_controle = carregar_dados_controle_geral(app)
        if not dados_controle: 
            app.log(">>> Base de dados principal está vazia.")
            return

        numeros_para_processar = []
        for telefone, dados in dados_controle.items():
            if dados.get("status") == "LB":
                data_mailing_str = dados.get("data_mailing", "")
                if not data_mailing_str:
                    numeros_para_processar.append(telefone)
                else:
                    try:
                        if (data_hoje - datetime.strptime(data_mailing_str, '%d/%m/%Y')).days >= 30:
                            numeros_para_processar.append(telefone)
                        else:
                            stats["retidos"] += 1
                    except ValueError:
                        numeros_para_processar.append(telefone)
        
        stats["total"] = len(numeros_para_processar)
        if stats["total"] == 0:
            app.log(">>> Nenhum número liberado para envio no momento.")
            app.show_info_message("Concluído", "Nenhum número novo ou reenvio para fazer no Mailing.")
            return

        app.log(f">>> {stats['total']} números para processar. {stats['retidos']} retidos (no prazo de 30 dias).")
        
        pedidos_update_destino = []
        pedidos_update_principal = []
        numeros_para_adicionar_final = []

        for telefone in numeros_para_processar:
            if stop_event.is_set(): break
            stats["processados"] += 1

            if telefone in mapa_destino and mapa_destino[telefone]['funcionario'] in NOMES_FUNCIONARIOS:
                app.log(f"--- Reatribuindo lead: {telefone}")
                dados_lead = mapa_destino[telefone]
                linha_alvo = dados_lead['row_index']
                funcionario = dados_lead['funcionario'].capitalize()
                
                data_envio_anterior_str = dados_controle[telefone].get('data_mailing', '')
                mes_anterior = ""
                if data_envio_anterior_str:
                    mes_anterior = MESES_PT.get(datetime.strptime(data_envio_anterior_str, '%d/%m/%Y').month, "")

                texto_historico = f"enviado em {mes_anterior} por {funcionario}" if mes_anterior else f"enviado por {funcionario}"
                
                pedidos_update_destino.append(sheets_integration.prepare_update_cell_request("Sheet1", linha_alvo, 3, texto_historico))
                pedidos_update_destino.append(sheets_integration.prepare_update_cell_request("Sheet1", linha_alvo, 2, ""))
            
            else:
                numeros_para_adicionar_final.append([telefone])

            row_index_principal = dados_controle[telefone]['row_index']
            pedidos_update_principal.append(sheets_integration.prepare_update_cell_request(
                ABA_CONTROLE_GERAL, row_index_principal, COL_DATA_MAILING + 1, data_hoje_str))
            pedidos_update_principal.append(sheets_integration.prepare_update_cell_request(
                ABA_CONTROLE_GERAL, row_index_principal, COL_STATUS_SECUNDARIO + 1, "enviado"))
            stats["enviados"] += 1

        if pedidos_update_destino:
            app.log(f">>> Executando {len(pedidos_update_destino)} atualizações de histórico na planilha de mailing...")
            sheets_integration.batch_update_values(service_destino, id_destino, pedidos_update_destino, app)
        
        if numeros_para_adicionar_final:
            app.log(f">>> Adicionando {len(numeros_para_adicionar_final)} novos números à planilha de mailing...")
            sheets_integration.append_values(service_destino, id_destino, "Sheet1", numeros_para_adicionar_final, app)

        if pedidos_update_principal:
            app.log(f">>> Atualizando {stats['enviados']} registros na planilha principal '{ABA_CONTROLE_GERAL}'...")
            sheets_integration.batch_update_values(app.sheet_service, app.spreadsheet_id, pedidos_update_principal, app)
        
        app.show_info_message("Processo Concluído", f"{stats['enviados']} leads foram processados e enviados/liberados na planilha de mailing.")

    except Exception as e:
        app.log(f"!!! [{PROCESSO_ENVIAR_MALING.upper()}] ERRO CRÍTICO: {e}")
        app.show_error_message("Erro Crítico", f"Ocorreu um erro fatal: {e}")
    finally:
        app.log(">>> Processo de mailing finalizado.")
        app.resetar_controles_processo(PROCESSO_ENVIAR_MALING)

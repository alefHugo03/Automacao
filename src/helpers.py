# src/helpers.py (VERSÃO FINAL COMPLETA)
import os
import sys
import json
import shutil

def get_base_path():
    """Retorna o caminho base correto, seja executando como script ou .exe."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

def get_caminho_recurso(caminho_relativo):
    """Retorna o caminho absoluto para um recurso, funcionando em modo dev e compilado."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__)))
    return os.path.join(base_path, caminho_relativo)

CAMINHO_CONFIG_USUARIO = os.path.join(get_base_path(), "config.json")

def carregar_config():
    """Carrega as configurações do utilizador. Se não existir, copia o padrão."""
    if not os.path.exists(CAMINHO_CONFIG_USUARIO):
        caminho_config_padrao = get_caminho_recurso('config.json')
        if os.path.exists(caminho_config_padrao):
            try:
                shutil.copy2(caminho_config_padrao, CAMINHO_CONFIG_USUARIO)
            except Exception as e:
                print(f"Erro ao copiar config padrão: {e}")
                return {}
        else:
            salvar_config({})
            return {}
    try:
        with open(CAMINHO_CONFIG_USUARIO, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, FileNotFoundError):
        return {}

def salvar_config(config_dict):
    """Salva as configurações no ficheiro do utilizador."""
    try:
        with open(CAMINHO_CONFIG_USUARIO, "w", encoding="utf-8") as f:
            json.dump(config_dict, f, indent=4)
    except Exception as e:
        print(f"Erro ao salvar configuração: {e}")

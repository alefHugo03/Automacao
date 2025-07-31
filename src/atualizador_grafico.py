# src/atualizador_grafico.py (VERSÃO FINAL COMPLETA)
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import threading
import requests
import zipfile
import os
import sys
import io
import time

# --- CONFIGURAÇÕES DA ATUALIZAÇÃO ---
GITHUB_REPO_OWNER = "alefHugo03"
GITHUB_REPO_NAME = "Automacao"
ASSET_NAME = "VerificadorApp.zip"
APP_VERSION_ATUAL = "3.5.0" # Lembre-se de atualizar esta versão

class UpdaterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Atualizador de Aplicação")
        self.geometry("500x250")
        self.resizable(False, False)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(self.main_frame, text="Pronto para verificar atualizações.", font=ctk.CTkFont(size=14))
        self.status_label.grid(row=0, column=0, padx=10, pady=10)

        self.progress_bar = ctk.CTkProgressBar(self.main_frame)
        self.progress_bar.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        self.progress_bar.set(0)
        
        self.progress_label = ctk.CTkLabel(self.main_frame, text="")
        self.progress_label.grid(row=2, column=0, padx=10, pady=5)

        self.action_button = ctk.CTkButton(self.main_frame, text="Verificar Atualizações", height=40, command=self.iniciar_verificacao)
        self.action_button.grid(row=3, column=0, padx=20, pady=15, sticky="ew")
        
        # Inicia a verificação automaticamente ao abrir
        self.after(500, self.iniciar_verificacao)

    def log(self, message):
        self.status_label.configure(text=message)
        self.update_idletasks()

    def iniciar_verificacao(self):
        self.action_button.configure(state="disabled", text="Verificando...")
        self.progress_bar.set(0)
        self.progress_label.configure(text="")
        threading.Thread(target=self.run_update_process, daemon=True).start()

    def run_update_process(self):
        try:
            self.log("Procurando a versão mais recente...")
            api_url = f"https://api.github.com/repos/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/releases/latest"
            response = requests.get(api_url, timeout=10)
            response.raise_for_status()
            latest_release = response.json()
            latest_version = latest_release["tag_name"].lstrip('v')
            self.log(f"Versão mais recente encontrada: {latest_version}")

            if latest_version <= APP_VERSION_ATUAL:
                self.log(f"Você já está na versão mais recente ({APP_VERSION_ATUAL}).")
                messagebox.showinfo("Tudo Certo!", "Você já está usando a versão mais recente do programa.")
                self.action_button.configure(state="normal", text="Verificar Novamente")
                self.after(1000, self.destroy)
                return

            if not messagebox.askyesno("Nova Versão Disponível!", f"Uma nova versão ({latest_version}) está disponível. A sua é {APP_VERSION_ATUAL}.\n\nDeseja baixar e instalar agora?"):
                self.action_button.configure(state="normal", text="Verificar Atualizações")
                self.destroy()
                return

            asset_url = ""
            for asset in latest_release['assets']:
                if asset['name'] == ASSET_NAME:
                    asset_url = asset['browser_download_url']
                    break
            
            if not asset_url: raise Exception(f"Não foi possível encontrar o arquivo '{ASSET_NAME}' na release.")

            self.log(f"Baixando {ASSET_NAME}..."); self.action_button.configure(text="Baixando...")
            
            download_response = requests.get(asset_url, stream=True, timeout=300)
            download_response.raise_for_status()
            
            total_size = int(download_response.headers.get('content-length', 0))
            bytes_downloaded = 0
            
            zip_in_memory = io.BytesIO()
            for chunk in download_response.iter_content(chunk_size=8192):
                zip_in_memory.write(chunk)
                bytes_downloaded += len(chunk)
                progress = bytes_downloaded / total_size if total_size > 0 else 0
                self.progress_bar.set(progress)
                self.progress_label.configure(text=f"{int(progress * 100)}%")
                self.update_idletasks()

            self.log("Download concluído. Instalando..."); self.action_button.configure(text="Instalando...")
            
            app_path = os.path.dirname(sys.executable)
            zip_in_memory.seek(0)
            with zipfile.ZipFile(zip_in_memory, 'r') as zip_ref:
                zip_ref.extractall(app_path)

            self.log("Atualização instalada com sucesso!")
            messagebox.showinfo("Sucesso!", "O programa foi atualizado. Por favor, inicie-o novamente.")
            self.destroy()

        except requests.exceptions.RequestException as e:
            self.log("Erro de conexão."); messagebox.showerror("Erro de Rede", f"Não foi possível conectar ao GitHub.\n\nVerifique sua conexão com a internet.\nDetalhes: {e}")
            self.destroy()
        except Exception as e:
            self.log("Ocorreu um erro durante a atualização."); messagebox.showerror("Erro", f"Não foi possível concluir a atualização.\n\nDetalhes: {e}")
            self.action_button.configure(state="normal", text="Tentar Novamente")

if __name__ == "__main__":
    try:
        if os.path.basename(sys.executable).lower() != 'atualizador.exe':
            app = UpdaterApp(); app.mainloop()
        else:
            time.sleep(1.5); app = UpdaterApp(); app.mainloop()
    except Exception as e:
        messagebox.showerror("Erro Crítico", f"O atualizador não pôde ser iniciado.\n\n{e}")

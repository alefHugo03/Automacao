# src/main.py (VERSÃO SEM AGENDADOR E MODO EFICIÊNCIA)
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog
import threading
import os
import sys
import subprocess
from typing import Literal
import webbrowser

from helpers import carregar_config, salvar_config
import automations
import ui_tabs
import sheets_integration

class LoginWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Login Não Perturbe")
        self.geometry("350x200")
        self.resizable(False, False)
        self.credentials = None
        self.grid_columnconfigure(0, weight=1)
        self.transient(parent)
        self.grab_set()
        ctk.CTkLabel(self, text="Utilizador:").grid(row=0, column=0, padx=20, pady=(20, 5), sticky="w")
        self.user_entry = ctk.CTkEntry(self, width=200)
        self.user_entry.grid(row=1, column=0, padx=20, sticky="ew")
        ctk.CTkLabel(self, text="Senha:").grid(row=2, column=0, padx=20, pady=(5, 5), sticky="w")
        self.pass_entry = ctk.CTkEntry(self, show="*", width=200)
        self.pass_entry.grid(row=3, column=0, padx=20, sticky="ew")
        login_button = ctk.CTkButton(self, text="Entrar", command=self._on_login, height=40)
        login_button.grid(row=4, column=0, padx=20, pady=20, sticky="ew")
        self.user_entry.focus()
        self.pass_entry.bind("<Return>", self._on_login)
        self.user_entry.bind("<Return>", lambda e: self.pass_entry.focus())
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _on_login(self, event=None):
        user = self.user_entry.get()
        password = self.pass_entry.get()
        if not user or not password:
            messagebox.showwarning("Campos Vazios", "Por favor, preencha o utilizador e a senha.", parent=self)
            return
        self.credentials = (user, password)
        self.destroy()

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        try:
            # --- Eventos de controle para threads ---
            self.stop_event_claro = threading.Event()
            self.stop_event_maling = threading.Event()
            self.stop_event_limpeza = threading.Event()
            self.stop_event_limpeza_maling = threading.Event()

            self.sheet_service = None
            self.spreadsheet_id = None
            self.chromedriver_path_var = tk.StringVar()

            self.title("Verificador de Números")
            self.minsize(850, 700)
            ctk.set_appearance_mode("dark")
            ctk.set_default_color_theme("blue")

            self.grid_rowconfigure(0, weight=1)
            self.grid_columnconfigure(0, weight=1)
            
            self.settings_panel_visible = False

            # --- Estrutura da UI ---
            self.main_container = ctk.CTkFrame(self, fg_color="transparent")
            self.main_container.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
            self.main_container.grid_columnconfigure(0, weight=1)
            self.main_container.grid_rowconfigure(0, weight=0)
            self.main_container.grid_rowconfigure(1, weight=5)
            self.main_container.grid_rowconfigure(2, weight=2)

            header_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
            header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
            header_frame.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(header_frame, text="Automação do Não Perturbe", font=ctk.CTkFont(size=24, weight="bold")).grid(row=0, column=0, sticky="w")
            
            settings_button = ctk.CTkButton(header_frame, text="☰", font=ctk.CTkFont(size=22), width=40, height=40, fg_color="transparent", command=self.toggle_settings_panel)
            settings_button.grid(row=0, column=1, sticky="e")
            
            self.tab_view = ctk.CTkTabview(self.main_container, anchor="w")
            self.tab_view.grid(row=1, column=0, sticky="nsew")
            self.tab_adicionar = self.tab_view.add("1. Adicionar Números")
            self.tab_executar_claro = self.tab_view.add("2. Não Perturbe")
            self.tab_enviar_maling = self.tab_view.add("3. Enviar para Mailing")

            self.log_textbox = ctk.CTkTextbox(self.main_container, state="disabled", font=("Courier New", 12))
            self.log_textbox.grid(row=2, column=0, pady=(10,0), sticky="nsew")

            # --- Criação dos componentes da UI ---
            self.settings_panel = ui_tabs.criar_painel_configuracoes(self)
            ui_tabs.criar_aba_adicionar(self)
            ui_tabs.criar_aba_executar_claro(self)
            ui_tabs.criar_aba_enviar_maling(self)

            self.log(">>> Bem-vindo! Programa pronto para uso.")
            
            self.carregar_configuracoes_iniciais()
            self.protocol("WM_DELETE_WINDOW", self.on_closing)

        except Exception as e:
            messagebox.showerror("Erro na Inicialização", f"Ocorreu um erro crítico ao iniciar a aplicação:\n\n{e}")
            self.destroy()

    def carregar_configuracoes_iniciais(self):
        config = carregar_config()
        self.spreadsheet_url_entry.insert(0, config.get("spreadsheet_url", ""))
        self.maling_url_entry.insert(0, config.get("maling_spreadsheet_url", ""))
        
        credentials_path = config.get("credentials_path")
        if credentials_path and os.path.exists(credentials_path):
            self.label_credentials_path.configure(text=f"Ficheiro definido: ...{os.path.basename(credentials_path)}", text_color="lightgreen")
            
        driver_path = config.get("driver_path", "")
        if driver_path and os.path.exists(driver_path):
            self.chromedriver_path_var.set(driver_path)
            self.label_chromedriver_path.configure(text=f"Caminho definido: ...{os.path.basename(driver_path)}")

    def on_closing(self):
        if messagebox.askokcancel("Sair", "Tem a certeza que quer sair?"):
            self.destroy()

    # --- Funções de automação manual (restante do código) ---
    def iniciar_limpeza_duplicatas(self):
        if not self.sheet_service or not self.spreadsheet_id:
            self.show_error_message("Não Conectado", "Conecte-se a uma Planilha Google primeiro.")
            return
        if messagebox.askyesno("Confirmar Limpeza", "Isto irá verificar as abas da PLANILHA PRINCIPAL e remover duplicatas.\n\nEsta ação não pode ser desfeita.\nDeseja continuar?"):
            self.log(">>> Iniciando limpeza da planilha principal...")
            self.iniciar_processo_thread(automations.remover_duplicatas_de_todas_abas, self.stop_event_limpeza)

    def iniciar_limpeza_maling(self):
        url = self.maling_url_entry.get()
        if not url:
            self.show_error_message("URL Ausente", "Insira o URL da planilha de Mailing para fazer a limpeza.")
            return
        if messagebox.askyesno("Confirmar Limpeza de Mailing", "Isto irá verificar a 'Sheet1' da PLANILHA DE MAILING e remover duplicatas.\n\nEsta ação não pode ser desfeita.\nDeseja continuar?"):
            self.log(">>> Iniciando limpeza da planilha de mailing...")
            self.iniciar_processo_thread(automations.remover_duplicatas_do_mailing, self.stop_event_limpeza_maling)

    def toggle_settings_panel(self):
        if self.settings_panel_visible:
            self.settings_panel.place_forget()
        else:
            self.settings_panel.place(relx=1.0, rely=0, anchor="ne", relwidth=0.4, relheight=1.0)
        self.settings_panel_visible = not self.settings_panel_visible

    def selecionar_chromedriver_path(self):
        filetypes = [("ChromeDriver", "chromedriver.exe")] if sys.platform == "win32" else [("ChromeDriver", "chromedriver"), ("Todos", "*.*")]
        caminho = filedialog.askopenfilename(title="Selecione o executável do ChromeDriver", filetypes=filetypes)
        if caminho:
            self.chromedriver_path_var.set(caminho)
            self.log(f">>> Caminho do ChromeDriver definido: {caminho}")
            self.label_chromedriver_path.configure(text=f"Caminho definido: ...{os.path.basename(caminho)}")
            config = carregar_config()
            config["driver_path"] = caminho
            salvar_config(config)

    def iniciar_atualizacao(self):
        try:
            base_path = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else None
            if not base_path:
                self.show_info_message("Ação de Desenvolvimento", "Esta ação é para o usuário final.")
                return
            updater_filename = "Atualizador.exe" if sys.platform == "win32" else "Atualizador"
            updater_path = os.path.join(base_path, updater_filename)
            if not os.path.exists(updater_path):
                self.show_error_message("Atualizador Não Encontrado", f"O arquivo '{updater_filename}' não foi encontrado.")
                return
            if messagebox.askyesno("Confirmar Atualização", "Isto irá FECHAR a aplicação atual e abrir o atualizador.\nDeseja continuar?"):
                if sys.platform != "win32":
                    try: os.chmod(updater_path, 0o755)
                    except Exception as e: self.log(f"Aviso: Não foi possível definir permissão de execução: {e}")
                subprocess.Popen([updater_path])
                self.after(200, self.on_closing)
        except Exception as e:
            self.show_error_message("Erro ao Iniciar Atualizador", f"Não foi possível iniciar o atualizador.\n\nErro: {e}")

    def selecionar_credentials_path(self):
        caminho = filedialog.askopenfilename(title="Selecione o seu ficheiro credentials.json", filetypes=[("JSON", "*.json"), ("Todos", "*.*")])
        if caminho:
            config = carregar_config()
            config["credentials_path"] = caminho
            salvar_config(config)
            self.log(f">>> Ficheiro de credenciais definido: {caminho}")
            self.label_credentials_path.configure(text=f"Ficheiro definido: ...{os.path.basename(caminho)}", text_color="lightgreen")
            self.show_info_message("Sucesso", "Ficheiro guardado. Insira o URL da planilha e conecte.")

    def salvar_e_conectar_sheets(self):
        url = self.spreadsheet_url_entry.get()
        if not url:
            self.show_warning_message("URL Ausente", "Por favor, insira o URL da sua Planilha Google.")
            return
        config = carregar_config()
        config["spreadsheet_url"] = url
        salvar_config(config)
        self.status_conexao_label.configure(text="Status: A autenticar...", text_color="orange")
        self.log(">>> A tentar conectar com a API do Google Sheets...")
        threading.Thread(target=self._conectar_worker, daemon=True).start()

    def _conectar_worker(self):
        service = sheets_integration.get_service(self)
        if service:
            self.sheet_service = service
            spreadsheet_id = sheets_integration.get_spreadsheet_id_from_url(carregar_config().get("spreadsheet_url"))
            if spreadsheet_id:
                self.spreadsheet_id = spreadsheet_id
                self.after(0, lambda: self.status_conexao_label.configure(text="Status: Conectado!", text_color="lightgreen"))
                self.log(f">>> Conexão estabelecida! ID: {spreadsheet_id}")
                sheets_integration.ensure_sheets_exist(self.sheet_service, self.spreadsheet_id, automations.ABAS_GERENCIADAS, self)
            else:
                self.after(0, lambda: self.status_conexao_label.configure(text="Status: Erro - URL inválido", text_color="red"))
                self.log("!!! ERRO: O URL fornecido não parece ser válido.")
        else:
            self.after(0, lambda: self.status_conexao_label.configure(text="Status: Falha na conexão", text_color="red"))
            self.log("!!! ERRO: Falha ao conectar com o Google.")

    def iniciar_automacao_claro(self):
        if not self.sheet_service or not self.spreadsheet_id:
            self.show_error_message("Não Conectado", "Conecte-se a uma Planilha Google nas opções.")
            return
        login_dialog = LoginWindow(self)
        self.wait_window(login_dialog)
        if login_dialog.credentials:
            usuario, senha = login_dialog.credentials
            self.start_button_claro.grid_remove()
            self.stop_button_claro.grid()
            self.resetar_estatisticas(automations.PROCESSO_CLARO)
            args = (automations.executar_verificacao, self.stop_event_claro, usuario, senha, self.headless_claro_var.get(), self.chromedriver_path_var.get())
            self.iniciar_processo_thread(*args)
        else:
            self.log(">>> [CLARO] Login cancelado.")

    def iniciar_automacao_maling(self):
        if not self.sheet_service or not self.spreadsheet_id:
            self.show_error_message("Não Conectado", "Conecte-se a uma Planilha Google nas opções.")
            return
        url = self.maling_url_entry.get()
        if not url:
            self.show_error_message("URL Ausente", "Insira o URL da planilha de destino.")
            return
        config = carregar_config()
        config["maling_spreadsheet_url"] = url
        salvar_config(config)
        self.log(">>> URL de mailing guardada.")
        self.start_button_maling.grid_remove()
        self.stop_button_maling.grid()
        self.resetar_estatisticas(automations.PROCESSO_ENVIAR_MALING)
        args = (automations.executar_envio_maling, self.stop_event_maling, url)
        self.iniciar_processo_thread(*args)

    def iniciar_processo_thread(self, target_func, stop_event, *args):
        stop_event.clear()
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", tk.END)
        self.log_textbox.configure(state="disabled")
        final_args = (self,) + args + (stop_event,)
        thread = threading.Thread(target=target_func, args=final_args, daemon=True)
        thread.start()

    def solicitar_parada(self, processo: Literal["claro", "maling"]):
        event_map = { "claro": self.stop_event_claro, "maling": self.stop_event_maling }
        button_map = { "claro": self.stop_button_claro, "maling": self.stop_button_maling }
        event_map[processo].set()
        button_map[processo].configure(state="disabled", text="A parar...")
        self.log(f">>> [{processo.upper()}] PEDIDO DE PARAGEM RECEBIDO...")

    def show_error_message(self, title, message): self.after(0, lambda: messagebox.showerror(title, message, parent=self))
    def show_info_message(self, title, message): self.after(0, lambda: messagebox.showinfo(title, message, parent=self))
    def show_warning_message(self, title, message): self.after(0, lambda: messagebox.showwarning(title, message, parent=self))
    def abrir_planilha_navegador(self):
        url = carregar_config().get("spreadsheet_url")
        if not url: self.show_error_message("Erro", "Nenhum URL de planilha definido."); return
        self.log(f">>> ABRINDO: '{url}'"); webbrowser.open(url)

    def _on_key_release(self, event=None):
        if event and event.keysym in ("Up", "Down", "Left", "Right", "Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R"): return
        self._validar_e_formatar_texto()

    def _formatar_numero(self, num_str):
        if len(num_str) > 2: num_str = f"({num_str[:2]}) {num_str[2:]}"
        if len(num_str) > 9: num_str = f"{num_str[:9]}-{num_str[9:]}"
        return num_str

    def _validar_e_formatar_texto(self):
        try:
            original = self.entry_numeros.get("1.0", "end-1c"); cursor_pos = self.entry_numeros.index(tk.INSERT)
            line, char = map(int, cursor_pos.split('.')); linha_antes = self.entry_numeros.get(f"{line}.0", f"{line}.{char}")
            digitos_antes = sum(c.isdigit() for c in linha_antes); linhas = original.split('\n')
            formatadas = [self._formatar_numero("".join(filter(str.isdigit, l))[:11]) for l in linhas]
            texto_formatado = "\n".join(formatadas)
            if texto_formatado != original:
                nova_pos = 0; digitos_contados = 0; linha_alvo = formatadas[line - 1]
                for c in linha_alvo:
                    if digitos_contados == digitos_antes: break
                    if c.isdigit(): digitos_contados += 1
                    nova_pos += 1
                if digitos_contados < digitos_antes: nova_pos = len(linha_alvo)
                self.entry_numeros.delete("1.0", tk.END); self.entry_numeros.insert("1.0", texto_formatado)
                self.entry_numeros.mark_set(tk.INSERT, f"{line}.{nova_pos}")
        except Exception as e: self.log(f"Erro ao formatar números: {e}")

    def adicionar_numeros(self, lista):
        if not self.sheet_service or not self.spreadsheet_id: self.show_error_message("Não Conectado", "Conecte-se a uma planilha."); return
        if not lista: self.show_warning_message("Aviso", "Nenhum número válido para adicionar."); return
        self.log("Lendo números existentes..."); existentes = set(sheets_integration.read_values(self.sheet_service, self.spreadsheet_id, automations.FICHEIRO_TELEFONES, self))
        novos = [num for num in lista if num not in existentes]
        if not novos: self.show_info_message("Informação", "Todos os números fornecidos já constam na planilha."); return
        if sheets_integration.append_values(self.sheet_service, self.spreadsheet_id, automations.FICHEIRO_TELEFONES, [[n] for n in novos], self):
            msg = f"SUCESSO: {len(novos)} novos números adicionados.\n{len(lista) - len(novos)} duplicados ignorados."
            self.show_info_message("Sucesso", msg); self.log(msg); self.entry_numeros.delete("1.0", tk.END)
        else: self.show_error_message("Erro", "Falha ao adicionar números.")

    def extrair_e_validar_numeros(self, texto): return ["".join(filter(str.isdigit, l)) for l in texto.splitlines() if len("".join(filter(str.isdigit, l))) == 11]
    def adicionar_numeros_da_caixa(self): self.adicionar_numeros(list(set(self.extrair_e_validar_numeros(self.entry_numeros.get("1.0", "end-1c")))))
    def carregar_numeros_de_arquivo(self):
        caminho = filedialog.askopenfilename(title="Selecione um ficheiro de texto", filetypes=[("Texto", "*.txt"), ("Todos", "*.*")])
        if not caminho: self.log(">>> Carregamento cancelado."); return
        try:
            with open(caminho, "r", encoding="utf-8") as f: conteudo = f.read()
            numeros = list(set(self.extrair_e_validar_numeros(conteudo)))
            if not numeros: self.show_warning_message("Ficheiro Vazio", "Nenhum número válido encontrado."); return
            self.log(f">>> {len(numeros)} números válidos lidos. Adicionando..."); self.adicionar_numeros(numeros)
        except Exception as e: self.show_error_message("Erro ao Ler Ficheiro", f"Ocorreu um erro: {e}"); self.log(f"!!! ERRO ao ler o ficheiro {caminho}: {e}")

    def resetar_estatisticas(self, processo: Literal["claro", "maling"]):
        if processo == automations.PROCESSO_CLARO:
            self.progress_bar_claro.set(0); self.label_progresso_claro.configure(text="0%")
            self.label_liberados.configure(text="Liberados: 0"); self.label_bloqueados.configure(text="Bloqueados: 0")
            self.label_pulados_claro.configure(text="Pulados: 0"); self.label_erros_claro.configure(text="Erros: 0")
        elif processo == automations.PROCESSO_ENVIAR_MALING:
            self.progress_bar_maling.set(0); self.label_progresso_maling.configure(text="0%")
            self.label_enviados_maling.configure(text="Enviados: 0"); self.label_retidos_maling.configure(text="Retidos: 0")
            self.label_expirados_maling.configure(text="Expirados: 0"); self.label_erros_maling.configure(text="Erros: 0")

    def _atualizar_estatisticas_thread_safe(self, stats, processo):
        progresso = stats["processados"] / stats["total"] if stats["total"] > 0 else 0; porcentagem_txt = f"{int(progresso * 100)}%"
        if processo == automations.PROCESSO_CLARO:
            self.label_progresso_claro.configure(text=porcentagem_txt); self.progress_bar_claro.set(progresso)
            self.label_liberados.configure(text=f"Liberados: {stats['liberados']}"); self.label_bloqueados.configure(text=f"Bloqueados: {stats['bloqueados']}")
            self.label_pulados_claro.configure(text=f"Pulados: {stats['pulados']}"); self.label_erros_claro.configure(text=f"Erros: {stats['erros']}")
        elif processo == automations.PROCESSO_ENVIAR_MALING:
            self.label_progresso_maling.configure(text=porcentagem_txt); self.progress_bar_maling.set(progresso)
            self.label_enviados_maling.configure(text=f"Enviados: {stats['enviados']}"); self.label_retidos_maling.configure(text=f"Retidos: {stats['retidos']}")
            self.label_expirados_maling.configure(text=f"Expirados: {stats['expirados']}"); self.label_erros_maling.configure(text=f"Erros: {stats['erros']}")

    def atualizar_estatisticas(self, stats, processo): self.after(0, lambda: self._atualizar_estatisticas_thread_safe(stats, processo))
    def log(self, message): self.after(0, lambda: self._log_thread_safe(message))
    def _log_thread_safe(self, message):
        self.log_textbox.configure(state="normal"); self.log_textbox.insert(tk.END, message + "\n"); self.log_textbox.configure(state="disabled"); self.log_textbox.see(tk.END)

    def _resetar_controles_processo_thread_safe(self, processo):
        mapa = {
            automations.PROCESSO_CLARO: (self.start_button_claro, self.stop_button_claro, "2. Não Perturbe"),
            automations.PROCESSO_ENVIAR_MALING: (self.start_button_maling, self.stop_button_maling, "3. Enviar para Mailing")
        }
        start_btn, stop_btn, tab_name = mapa[processo]
        start_btn.grid(); stop_btn.grid_remove()
        stop_btn.configure(state="normal", text=f"Parar {tab_name.split('.')[1].strip()}")
        self.tab_view.set(tab_name)
        
    def resetar_controles_processo(self, processo): self.after(0, lambda: self._resetar_controles_processo_thread_safe(processo))

if __name__ == "__main__":
    main_app = App()
    main_app.mainloop()
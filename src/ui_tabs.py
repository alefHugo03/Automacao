# src/ui_tabs.py
import customtkinter as ctk
import winreg # Usado para detetar a versão do Chrome no Windows
import subprocess # Usado para detetar a versão em outros sistemas
import os
import sys

def get_chrome_version():
    """Deteta a versão do Google Chrome instalada."""
    try:
        if sys.platform == "win32":
            # Caminho para a chave de registo do Chrome
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon")
            version, _ = winreg.QueryValueEx(key, "version")
            return version
        elif sys.platform == "linux" or sys.platform == "linux2":
            # Tenta executar o comando no Linux
            process = subprocess.Popen(['google-chrome', '--version'], stdout=subprocess.PIPE)
            version = process.communicate()[0].decode('UTF-8').strip().split()[-1]
            return version
        elif sys.platform == "darwin":
            # Tenta executar o comando no macOS
            process = subprocess.Popen(['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '--version'], stdout=subprocess.pipe)
            version = process.communicate()[0].decode('UTF-8').strip().split()[-1]
            return version
    except Exception:
        return "Não detetada"
    return "Não detetada"


def criar_aba_adicionar(self):
    self.tab_adicionar.grid_columnconfigure(0, weight=1)
    self.tab_adicionar.grid_rowconfigure(0, weight=1)
    
    content_frame = ctk.CTkFrame(self.tab_adicionar, fg_color="transparent")
    content_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
    content_frame.grid_columnconfigure(0, weight=1)
    content_frame.grid_rowconfigure(2, weight=1)

    ctk.CTkLabel(content_frame, text="Adicionar Novos Números", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, sticky="w")
    ctk.CTkLabel(content_frame, text="Cole os números abaixo ou carregue de um ficheiro .txt.", anchor="w").grid(row=1, column=0, pady=(5, 10), sticky="ew")
    
    self.entry_numeros = ctk.CTkTextbox(content_frame, font=("Courier New", 13))
    self.entry_numeros.grid(row=2, column=0, sticky="nsew")
    self.entry_numeros.bind("<KeyRelease>", self._on_key_release)
    
    buttons_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
    buttons_frame.grid(row=3, column=0, pady=(15, 0), sticky="ew")
    buttons_frame.grid_columnconfigure((0, 1), weight=1)

    ctk.CTkButton(buttons_frame, text="Adicionar Números da Caixa", command=self.adicionar_numeros_da_caixa, height=40).grid(row=0, column=0, padx=(0, 5), sticky="ew")
    ctk.CTkButton(buttons_frame, text="Carregar de Ficheiro .txt", command=self.carregar_numeros_de_arquivo, height=40).grid(row=0, column=1, padx=(5, 0), sticky="ew")

def criar_aba_executar_claro(self):
    self.tab_executar_claro.grid_columnconfigure(0, weight=1)
    self.tab_executar_claro.grid_rowconfigure(0, weight=1)

    content_frame = ctk.CTkFrame(self.tab_executar_claro, fg_color="transparent")
    content_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
    content_frame.grid_columnconfigure(0, weight=1)
    
    ctk.CTkLabel(content_frame, text="Verificação 'Não Perturbe'", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, sticky="w", pady=(0,10))
    
    self.headless_claro_var = ctk.BooleanVar(value=True)
    ctk.CTkCheckBox(content_frame, text="Executar em modo invisível (recomendado)", variable=self.headless_claro_var).grid(row=1, column=0, pady=5, sticky="w")
    
    button_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
    button_frame.grid(row=2, column=0, sticky="ew", pady=10)
    button_frame.grid_columnconfigure(0, weight=1)
    
    self.start_button_claro = ctk.CTkButton(button_frame, text="Iniciar Verificação", command=self.iniciar_automacao_claro, height=50, font=ctk.CTkFont(size=16, weight="bold"))
    self.start_button_claro.grid(row=0, column=0, sticky="ew")
    
    self.stop_button_claro = ctk.CTkButton(button_frame, text="Parar Verificação", command=lambda: self.solicitar_parada("claro"), height=50, font=ctk.CTkFont(size=16, weight="bold"), fg_color="#d9534f", hover_color="#c9302c")
    self.stop_button_claro.grid(row=0, column=0, sticky="ew")
    self.stop_button_claro.grid_remove()
    
    stats_frame = ctk.CTkFrame(content_frame)
    stats_frame.grid(row=3, column=0, pady=(10, 0), sticky="ew")
    stats_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
    
    progress_frame = ctk.CTkFrame(stats_frame, fg_color="transparent")
    progress_frame.grid(row=0, column=0, columnspan=4, padx=10, pady=(10, 5), sticky="ew")
    progress_frame.grid_columnconfigure(0, weight=1)
    self.progress_bar_claro = ctk.CTkProgressBar(progress_frame)
    self.progress_bar_claro.set(0)
    self.progress_bar_claro.grid(row=0, column=0, sticky="ew")
    self.label_progresso_claro = ctk.CTkLabel(progress_frame, text="0%", font=ctk.CTkFont(size=12))
    self.label_progresso_claro.grid(row=0, column=1, padx=(10,0))
    
    self.label_liberados = ctk.CTkLabel(stats_frame, text="Liberados: 0", font=("Arial", 12, "bold"), text_color="#5cb85c")
    self.label_liberados.grid(row=1, column=0, padx=5, pady=10)
    self.label_bloqueados = ctk.CTkLabel(stats_frame, text="Bloqueados: 0", font=("Arial", 12, "bold"), text_color="#d9534f")
    self.label_bloqueados.grid(row=1, column=1, padx=5, pady=10)
    self.label_pulados_claro = ctk.CTkLabel(stats_frame, text="Pulados: 0", font=("Arial", 12, "bold"), text_color="#f0ad4e")
    self.label_pulados_claro.grid(row=1, column=2, padx=5, pady=10)
    self.label_erros_claro = ctk.CTkLabel(stats_frame, text="Erros: 0", font=("Arial", 12, "bold"), text_color="#ffffff")
    self.label_erros_claro.grid(row=1, column=3, padx=5, pady=10)

def criar_aba_enviar_maling(self):
    self.tab_enviar_maling.grid_columnconfigure(0, weight=1)
    self.tab_enviar_maling.grid_rowconfigure(0, weight=1)

    content_frame = ctk.CTkFrame(self.tab_enviar_maling, fg_color="transparent")
    content_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
    content_frame.grid_columnconfigure(0, weight=1)

    ctk.CTkLabel(content_frame, text="Enviar para Planilha de Mailing", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, sticky="w", pady=(0,5))
    ctk.CTkLabel(content_frame, text="Envia números da aba 'liberados' para a planilha de destino,\nrespeitando um intervalo de 30 dias para reenvio.", anchor="w").grid(row=1, column=0, pady=5, sticky="ew")

    ctk.CTkLabel(content_frame, text="URL da Planilha de Destino (Mailing):", anchor="w").grid(row=2, column=0, sticky="w", pady=(10,0))
    self.maling_url_entry = ctk.CTkEntry(content_frame, placeholder_text="Cole o URL da planilha de destino aqui")
    self.maling_url_entry.grid(row=3, column=0, pady=(5, 10), sticky="ew")

    button_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
    button_frame.grid(row=4, column=0, sticky="ew", pady=10)
    button_frame.grid_columnconfigure(0, weight=1)

    self.start_button_maling = ctk.CTkButton(button_frame, text="Iniciar Envio para Mailing", command=self.iniciar_automacao_maling, height=50, font=ctk.CTkFont(size=16, weight="bold"))
    self.start_button_maling.grid(row=0, column=0, sticky="ew")

    self.stop_button_maling = ctk.CTkButton(button_frame, text="Parar Envio", command=lambda: self.solicitar_parada("maling"), height=50, font=ctk.CTkFont(size=16, weight="bold"), fg_color="#d9534f", hover_color="#c9302c")
    self.stop_button_maling.grid(row=0, column=0, sticky="ew")
    self.stop_button_maling.grid_remove()
    
    stats_frame = ctk.CTkFrame(content_frame)
    stats_frame.grid(row=5, column=0, pady=(15, 0), sticky="ew")
    stats_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

    progress_frame = ctk.CTkFrame(stats_frame, fg_color="transparent")
    progress_frame.grid(row=0, column=0, columnspan=4, padx=10, pady=(10, 5), sticky="ew")
    progress_frame.grid_columnconfigure(0, weight=1)
    self.progress_bar_maling = ctk.CTkProgressBar(progress_frame)
    self.progress_bar_maling.set(0)
    self.progress_bar_maling.grid(row=0, column=0, sticky="ew")
    self.label_progresso_maling = ctk.CTkLabel(progress_frame, text="0%", font=ctk.CTkFont(size=12))
    self.label_progresso_maling.grid(row=0, column=1, padx=(10, 0))

    self.label_enviados_maling = ctk.CTkLabel(stats_frame, text="Enviados: 0", font=("Arial", 12, "bold"), text_color="#5cb85c")
    self.label_enviados_maling.grid(row=1, column=0, padx=5, pady=10)
    self.label_retidos_maling = ctk.CTkLabel(stats_frame, text="Retidos: 0", font=("Arial", 12, "bold"), text_color="#3399ff")
    self.label_retidos_maling.grid(row=1, column=1, padx=5, pady=10)
    self.label_expirados_maling = ctk.CTkLabel(stats_frame, text="Expirados: 0", font=("Arial", 12, "bold"), text_color="#f0ad4e")
    self.label_expirados_maling.grid(row=1, column=2, padx=5, pady=10)
    self.label_erros_maling = ctk.CTkLabel(stats_frame, text="Erros: 0", font=("Arial", 12, "bold"), text_color="#ffffff")
    self.label_erros_maling.grid(row=1, column=3, padx=5, pady=10)

def criar_painel_configuracoes(self):
    settings_panel = ctk.CTkFrame(self, border_width=2)
    settings_panel.grid_columnconfigure(0, weight=1)

    header_frame = ctk.CTkFrame(settings_panel, fg_color="transparent")
    header_frame.grid(row=0, column=0, padx=20, pady=(20, 15), sticky="ew")
    header_frame.grid_columnconfigure(0, weight=1)

    ctk.CTkLabel(header_frame, text="Configurações", font=ctk.CTkFont(size=20, weight="bold")).grid(row=0, column=0, sticky="w")
    back_button = ctk.CTkButton(header_frame, text="< Voltar", width=80, command=self.toggle_settings_panel)
    back_button.grid(row=0, column=1, sticky="e")
    
    scroll_frame = ctk.CTkScrollableFrame(settings_panel, fg_color="transparent")
    scroll_frame.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
    settings_panel.grid_rowconfigure(1, weight=1)
    scroll_frame.grid_columnconfigure(0, weight=1)

    # --- Seção do Google Sheets ---
    g_sheets_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
    g_sheets_frame.grid(row=0, column=0, padx=10, pady=0, sticky="ew")
    g_sheets_frame.grid_columnconfigure(0, weight=1)
    
    ctk.CTkLabel(g_sheets_frame, text="Planilha Principal", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, sticky="w")
    ctk.CTkButton(g_sheets_frame, text="1. Selecionar 'credentials.json'", command=self.selecionar_credentials_path).grid(row=1, column=0, pady=(10,0), sticky="ew")
    self.label_credentials_path = ctk.CTkLabel(g_sheets_frame, text="Ficheiro de credenciais: Nenhum", font=("Arial", 10), text_color="gray")
    self.label_credentials_path.grid(row=2, column=0, pady=(2, 10), sticky="w")
    ctk.CTkLabel(g_sheets_frame, text="URL da Planilha Principal:", anchor="w").grid(row=3, column=0, pady=(5,0), sticky="w")
    self.spreadsheet_url_entry = ctk.CTkEntry(g_sheets_frame, placeholder_text="Cole o URL da sua Planilha Google")
    self.spreadsheet_url_entry.grid(row=4, column=0, pady=(0, 10), sticky="ew")
    self.save_url_button = ctk.CTkButton(g_sheets_frame, text="2. Salvar e Conectar", command=self.salvar_e_conectar_sheets, font=ctk.CTkFont(weight="bold"))
    self.save_url_button.grid(row=5, column=0, sticky="ew")
    self.status_conexao_label = ctk.CTkLabel(g_sheets_frame, text="Status: Desconectado", text_color="gray")
    self.status_conexao_label.grid(row=6, column=0, pady=(2, 10), sticky="w")
    ctk.CTkButton(g_sheets_frame, text="Abrir Planilha no Navegador", command=self.abrir_planilha_navegador, fg_color="transparent", border_width=1, text_color=("gray10", "gray90")).grid(row=7, column=0, pady=(10,0), sticky="ew")

    separator2 = ctk.CTkFrame(scroll_frame, height=2, fg_color="gray25")
    separator2.grid(row=1, column=0, padx=10, pady=20, sticky="ew")

    # --- Seção Avançada ---
    update_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
    update_frame.grid(row=2, column=0, padx=10, pady=0, sticky="ew")
    update_frame.grid_columnconfigure(0, weight=1)
    
    ctk.CTkLabel(update_frame, text="Avançado", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, sticky="w")
    
    # NOVO: Label para mostrar a versão do Chrome
    chrome_version_frame = ctk.CTkFrame(update_frame, fg_color="transparent")
    chrome_version_frame.grid(row=1, column=0, pady=(10,0), sticky="ew")
    ctk.CTkLabel(chrome_version_frame, text="Versão do Chrome detetada:", font=("Arial", 12, "bold")).pack(side="left")
    self.label_chrome_version = ctk.CTkLabel(chrome_version_frame, text=get_chrome_version(), font=("Arial", 12))
    self.label_chrome_version.pack(side="left", padx=5)

    ctk.CTkButton(update_frame, text="Definir Caminho do ChromeDriver", command=self.selecionar_chromedriver_path).grid(row=2, column=0, pady=(10,0), sticky="ew")
    self.label_chromedriver_path = ctk.CTkLabel(update_frame, text="Caminho do driver: Automático", font=("Arial", 10), text_color="gray")
    self.label_chromedriver_path.grid(row=3, column=0, pady=(2,10), sticky="w")
    ctk.CTkButton(update_frame, text="Procurar Atualizações", command=self.iniciar_atualizacao).grid(row=4, column=0, pady=5, sticky="ew")
    
    return settings_panel

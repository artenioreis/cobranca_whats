import tkinter as tk
from tkinter import ttk, messagebox
import pyodbc
import json
import os
import base64
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import locale
from datetime import datetime
from urllib.parse import quote
import re

# Configuração de locale
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_ALL, '')
    except:
        pass

class AppIntegrada:
    def __init__(self, root):
        """
        Inicializa a janela principal e configurações.
        """
        self.root = root
        self.root.title("Sistema Integrado de Cobrança")
        self.root.geometry("550x720")

        # Variáveis de Configuração
        self.servidor = tk.StringVar()
        self.banco = tk.StringVar()
        self.usuario = tk.StringVar()
        self.senha = tk.StringVar()
        self.lembrar_senha = tk.BooleanVar(value=True)

        # Variáveis de Controle
        self.navegador = None
        self.conn_sql = None

        self.criar_interface()
        self.carregar_configuracoes()

    def criar_interface(self):
        self.frame = ttk.Frame(self.root, padding="15")
        self.frame.pack(fill=tk.BOTH, expand=True)

        # Título
        ttk.Label(self.frame, text="Configuração do SQL Server", font=('Arial', 12, 'bold')).grid(
            row=0, column=0, columnspan=2, pady=10)

        # Campos de Entrada
        campos = [
            ("Servidor:", self.servidor, 1),
            ("Banco de Dados:", self.banco, 2),
            ("Usuário:", self.usuario, 3),
            ("Senha:", self.senha, 4)
        ]

        for texto, var, linha in campos:
            ttk.Label(self.frame, text=texto).grid(row=linha, column=0, sticky=tk.W, pady=5)
            show_char = "*" if texto == "Senha:" else ""
            ttk.Entry(self.frame, textvariable=var, width=35, show=show_char).grid(
                row=linha, column=1, pady=5, sticky=tk.W)

        # Checkbox
        ttk.Checkbutton(
            self.frame,
            text="Lembrar senha",
            variable=self.lembrar_senha
        ).grid(row=5, column=0, columnspan=2, pady=5, sticky=tk.W)

        # Botão Principal
        style = ttk.Style()
        style.configure('Accent.TButton', font=('Arial', 10, 'bold'), foreground='black', background="#3CA592")
        
        ttk.Button(
            self.frame,
            text="Conectar, Buscar Dados e Enviar WhatsApp",
            command=self.iniciar_processo_completo,
            style='Accent.TButton'
        ).grid(row=6, column=0, columnspan=2, pady=15, sticky=tk.EW)

        # Área de Status
        self.status_text = tk.Text(self.frame, height=18, width=60, state=tk.DISABLED, wrap=tk.WORD)
        self.status_text.grid(row=7, column=0, columnspan=2, pady=5)

        # Botão Fechar
        ttk.Button(
            self.frame,
            text="Fechar Aplicação",
            command=self.on_closing
        ).grid(row=8, column=0, columnspan=2, pady=(5,10), sticky=tk.EW)

    def simple_encrypt(self, text):
        return base64.b64encode(text.encode('utf-8')).decode('utf-8')

    def simple_decrypt(self, text):
        try:
            return base64.b64decode(text.encode('utf-8')).decode('utf-8')
        except:
            return ""

    def carregar_configuracoes(self):
        config_file = 'config_db.json'
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.servidor.set(config.get('servidor', ''))
                    self.banco.set(config.get('banco', ''))
                    self.usuario.set(config.get('usuario', ''))
                    
                    senha_cripto = config.get('senha', '')
                    if senha_cripto:
                        self.senha.set(self.simple_decrypt(senha_cripto))
                    
                    self.lembrar_senha.set(config.get('lembrar_senha', True))
                self.atualizar_status("Configurações carregadas.")
            except Exception as e:
                self.atualizar_status(f"Erro ao carregar config: {e}")

    def salvar_configuracoes(self):
        config = {
            'servidor': self.servidor.get(),
            'banco': self.banco.get(),
            'usuario': self.usuario.get(),
            'lembrar_senha': self.lembrar_senha.get()
        }
        if self.lembrar_senha.get():
            config['senha'] = self.simple_encrypt(self.senha.get())
        else:
            config['senha'] = ""

        try:
            with open('config_db.json', 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4)
            return True
        except Exception as e:
            self.atualizar_status(f"Erro ao salvar config: {e}")
            return False

    def conectar_sql_server(self):
        server = self.servidor.get()
        database = self.banco.get()
        username = self.usuario.get()
        password = self.senha.get()

        if not all([server, database, username, password]):
            messagebox.showerror("Erro", "Preencha todos os campos do SQL.")
            return False

        try:
            conn_str = (
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={server};DATABASE={database};"
                f"UID={username};PWD={password};"
                f"TrustServerCertificate=yes;"
            )
            self.atualizar_status(f"🔄 Conectando ao SQL: {server}...")
            self.conn_sql = pyodbc.connect(conn_str)
            
            cursor = self.conn_sql.cursor()
            cursor.execute("SELECT @@VERSION")
            self.atualizar_status("✅ Conexão SQL estabelecida com sucesso!")
            return True
        except pyodbc.Error as e:
            self.atualizar_status(f"❌ Erro SQL: {e}")
            messagebox.showerror("Erro Conexão", str(e))
            return False

    def buscar_dados_cobranca(self):
        if not self.conn_sql:
            return None

        query = """
        SELECT TOP 6
            Razao_social, Num_Bloqueto, Vlr_Documento,
            Dat_Vencimento, Cod_Barra, fone1
        FROM V_CTREC7DIAS
        """
        try:
            self.atualizar_status("🔄 Buscando dados na view V_CTREC3DIAS...")
            cursor = self.conn_sql.cursor()
            cursor.execute(query)
            colunas = [col[0] for col in cursor.description]
            dados = [dict(zip(colunas, row)) for row in cursor.fetchall()]

            if not dados:
                self.atualizar_status("⚠️ Nenhum registro encontrado na view.")
                return None
            
            self.atualizar_status(f"✅ {len(dados)} registros encontrados.")
            return dados
        except pyodbc.Error as e:
            self.atualizar_status(f"❌ Erro na consulta: {e}")
            return None

    def limpar_numero_telefone(self, telefone, codigo_pais="55"):
        if not telefone: return None
        num_limpo = re.sub(r'\D', '', str(telefone))
        
        # Lógica simples para tratar DDD e 9º dígito
        if len(num_limpo) < 10: 
            return None # Número muito curto
            
        if not num_limpo.startswith(codigo_pais):
            num_limpo = codigo_pais + num_limpo
            
        return num_limpo

    def formatar_mensagem(self, cliente):
        try:
            val = cliente['Vlr_Documento']
            val_fmt = locale.currency(float(val), grouping=True) if val else "R$ 0,00"
            
            dt = cliente['Dat_Vencimento']
            dt_fmt = dt.strftime('%d/%m/%Y') if isinstance(dt, datetime) else str(dt)

            msg = (
                f"Prezado(a) {cliente.get('Razao_social', 'Cliente')},\n\n"
                f"Lembrete: O boleto referente à fatura {cliente.get('Num_Bloqueto', 'N/A')} "
                f"no valor de {val_fmt} vencerá em breve ({dt_fmt}).\n"
                f"Código de barras: {cliente.get('Cod_Barra', 'N/A')}\n\n"
                f"Pagamento via internet banking ou lotéricas."
            )
            return msg
        except Exception as e:
            self.atualizar_status(f"Erro ao formatar mensagem: {e}")
            return None

    def abrir_whatsapp_web(self):
        try:
            self.atualizar_status("\n🔄 Abrindo WhatsApp Web...")
            self.root.update()
            
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            # Salva sessão do usuário
            user_data = os.path.join(os.getcwd(), "chrome_profile_wpp")
            if not os.path.exists(user_data): os.makedirs(user_data)
            options.add_argument(f"user-data-dir={user_data}")

            self.navegador = webdriver.Chrome(options=options)
            self.navegador.get("https://web.whatsapp.com")
            
            self.atualizar_status("Aguardando login (escaneie o QR Code)...")
            WebDriverWait(self.navegador, 120).until(
                EC.presence_of_element_located((By.ID, "side"))
            )
            self.atualizar_status("✅ WhatsApp Conectado!")
            return True
        except Exception as e:
            self.atualizar_status(f"❌ Erro ao abrir Navegador: {e}")
            return False

    def enviar_mensagem(self, telefone, texto):
        if not self.navegador: return False
        
        num_fmt = self.limpar_numero_telefone(telefone)
        if not num_fmt:
            self.atualizar_status(f"⚠️ Telefone inválido: {telefone}")
            return False

        url = f"https://web.whatsapp.com/send?phone={num_fmt}&text={quote(texto)}"
        try:
            self.navegador.get(url)
            time.sleep(5) # Aguarda carregamento inicial

            # Botão Enviar
            btn_xpath = '//span[@data-icon="send"]'
            btn = WebDriverWait(self.navegador, 20).until(
                EC.element_to_be_clickable((By.XPATH, btn_xpath))
            )
            btn.click()
            time.sleep(3) # Aguarda envio
            self.atualizar_status(f"✔️ Enviado para {num_fmt}")
            return True
        except TimeoutException:
            self.atualizar_status(f"❌ Falha/Timeout ao enviar para {num_fmt}")
            return False
        except Exception as e:
            self.atualizar_status(f"❌ Erro envio: {e}")
            return False

    def iniciar_processo_completo(self):
        self.status_text.config(state=tk.NORMAL)
        self.status_text.delete(1.0, tk.END)
        self.status_text.config(state=tk.DISABLED)

        if not self.salvar_configuracoes(): return
        if not self.conectar_sql_server(): return

        dados = self.buscar_dados_cobranca()
        if not dados:
            if self.conn_sql: self.conn_sql.close()
            return

        if not self.abrir_whatsapp_web():
            if self.conn_sql: self.conn_sql.close()
            return

        sucesso = 0
        self.atualizar_status(f"\n🚀 Enviando {len(dados)} mensagens...")
        
        for cli in dados:
            msg = self.formatar_mensagem(cli)
            fone = cli.get('fone1')
            if msg and fone:
                if self.enviar_mensagem(fone, msg):
                    sucesso += 1
            else:
                self.atualizar_status(f"⚠️ Dados incompletos para {cli.get('Razao_social')}")
            
        self.atualizar_status(f"\n🏁 Finalizado! Sucesso: {sucesso}/{len(dados)}")
        
        if self.conn_sql:
            self.conn_sql.close()
        
        messagebox.showinfo("Concluído", f"Processo finalizado.\nEnviados: {sucesso}")

    def atualizar_status(self, msg):
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, msg + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def on_closing(self):
        if self.navegador:
            try: self.navegador.quit()
            except: pass
        if self.conn_sql:
            try: self.conn_sql.close()
            except: pass
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = AppIntegrada(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()
import pyodbc
import json
import os
import base64

def simple_decrypt(text):
    try:
        return base64.b64decode(text.encode('utf-8')).decode('utf-8')
    except:
        return ""

def atualizar_telefones_teste():
    """
    Executa updates no banco de dados baseados no arquivo SQL fornecido.
    Útil para configurar números de teste antes de rodar o envio.
    """
    # 1. Tenta ler a configuração existente para conectar
    config_file = 'config_db.json'
    if not os.path.exists(config_file):
        print("❌ Arquivo 'config_db.json' não encontrado. Rode o app principal primeiro e salve a configuração.")
        return

    with open(config_file, 'r', encoding='utf-8') as f:
        config = json.load(f)

    server = config.get('servidor')
    database = config.get('banco')
    username = config.get('usuario')
    password = simple_decrypt(config.get('senha', ''))

    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={server};DATABASE={database};"
        f"UID={username};PWD={password};"
        f"TrustServerCertificate=yes;"
    )

    try:
        print(f"🔄 Conectando ao banco {database} em {server}...")
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()

        # Lista de updates conforme seu arquivo SQL
        updates = [
            ("CLINICA ONSMBV LTDA", "8581829648"),  # Rafael
            ("CV DISTRIBUIDORA DE PRODUTOS HOSPITALARES LTDA", "88888888"), # Não vai enviar
            ("IVONEIDE DE SOUZA SOARES", "8597518182"), # Daniel
            ("SINGULAR ODONTOLOGIA LJ LTDA", "8588917343"), # Ariel
            ("L B & FILHOS EMPREENDIMENTOS LTDA", "") # Vazio
        ]

        print("🚀 Iniciando atualizações de telefone...")
        for razao, fone in updates:
            # Atenção: Usar parâmetros (?) previne SQL Injection e erros com aspas
            sql = "UPDATE CLIEN SET Fone1 = ? WHERE Razao_Social = ?"
            cursor.execute(sql, (fone, razao))
            print(f"   ✔️ Atualizado: {razao} -> {fone}")

        conn.commit()
        print("\n✅ Todas as atualizações foram aplicadas com sucesso!")
        conn.close()

    except pyodbc.Error as e:
        print(f"\n❌ Erro ao conectar ou executar SQL: {e}")

if __name__ == "__main__":
    atualizar_telefones_teste()
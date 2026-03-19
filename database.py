import sqlite3, os
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(__file__), "data", "agente_email.db")

def get_conn():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_conn()
    c = conn.cursor()

    # Tabela de contatos (importada das agendas)
    c.execute("""
    CREATE TABLE IF NOT EXISTS contatos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        orgao TEXT,
        email TEXT,
        email_alternativo TEXT,
        telefone TEXT,
        endereco TEXT,
        categoria TEXT,
        observacoes TEXT,
        fonte TEXT,
        criado_em TEXT DEFAULT (datetime('now','localtime')),
        atualizado_em TEXT DEFAULT (datetime('now','localtime'))
    )""")

    # Tabela de processos/intimações do PJe
    c.execute("""
    CREATE TABLE IF NOT EXISTS processos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        numero_processo TEXT NOT NULL,
        tipo TEXT,
        destinatario TEXT,
        endereco_destinatario TEXT,
        email_destinatario TEXT,
        email_encontrado_em TEXT,
        assunto_email TEXT,
        arquivo_mandado TEXT,
        arquivo_anexo TEXT,
        status TEXT DEFAULT 'pendente',
        pje_url TEXT,
        expedido_em TEXT,
        distribuido_em TEXT,
        criado_em TEXT DEFAULT (datetime('now','localtime')),
        processado_em TEXT
    )""")

    # Histórico de e-mails enviados
    c.execute("""
    CREATE TABLE IF NOT EXISTS emails_enviados (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        processo_id INTEGER,
        numero_processo TEXT,
        destinatario TEXT,
        email TEXT,
        assunto TEXT,
        arquivo_mandado TEXT,
        arquivo_anexo TEXT,
        enviado_em TEXT DEFAULT (datetime('now','localtime')),
        revisado_por TEXT DEFAULT 'aglanio.carvalho@trf1.jus.br',
        FOREIGN KEY(processo_id) REFERENCES processos(id)
    )""")

    # Templates de e-mail
    c.execute("""
    CREATE TABLE IF NOT EXISTS templates (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT UNIQUE NOT NULL,
        assunto TEXT NOT NULL,
        corpo TEXT NOT NULL,
        tipo_intimacao TEXT,
        ativo INTEGER DEFAULT 1,
        criado_em TEXT DEFAULT (datetime('now','localtime'))
    )""")

    # Inserir template padrão se não existir
    c.execute("SELECT COUNT(*) FROM templates WHERE nome='padrao_intimacao'")
    if c.fetchone()[0] == 0:
        c.execute("""
        INSERT INTO templates (nome, assunto, corpo, tipo_intimacao) VALUES (
            'padrao_intimacao',
            'Mandado {numero_processo} - {tipo} - SJPI',
            'Prezado(a) Senhor(a),\n\nEncaminhamos em anexo o mandado judicial referente ao processo {numero_processo}, expedido pela Seção Judiciária do Piauí - SJPI.\n\nAtenciosamente,\n\nOficial de Justiça\nSeção Judiciária do Piauí - TRF 1ª Região\naglanio.carvalho@trf1.jus.br',
            'Intimação'
        )""")

    # Adiciona coluna tratamento se ainda não existir
    try:
        c.execute("ALTER TABLE contatos ADD COLUMN tratamento TEXT")
        conn.commit()
    except Exception:
        pass  # já existe

    # Adiciona coluna grau (1=1ºgrau, 2=2ºgrau) se ainda não existir
    try:
        c.execute("ALTER TABLE processos ADD COLUMN grau INTEGER DEFAULT 1")
        conn.commit()
    except Exception:
        pass  # já existe

    # ── Escritório Virtual de IA ──────────────────────────────────────────────
    c.execute("""
    CREATE TABLE IF NOT EXISTS escritorio_agentes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        cargo TEXT NOT NULL,
        avatar TEXT DEFAULT '🤖',
        cor TEXT DEFAULT '#7c3aed',
        system_prompt TEXT,
        status TEXT DEFAULT 'online',
        posicao INTEGER DEFAULT 0,
        criado_em TEXT DEFAULT (datetime('now','localtime'))
    )""")

    c.execute("""
    CREATE TABLE IF NOT EXISTS escritorio_conversas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        agente_id INTEGER,
        tipo TEXT DEFAULT 'individual',
        participantes TEXT,
        criado_em TEXT DEFAULT (datetime('now','localtime'))
    )""")

    c.execute("""
    CREATE TABLE IF NOT EXISTS escritorio_mensagens (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        conversa_id INTEGER,
        remetente TEXT NOT NULL,
        conteudo TEXT NOT NULL,
        criado_em TEXT DEFAULT (datetime('now','localtime'))
    )""")

    c.execute("""
    CREATE TABLE IF NOT EXISTS escritorio_tarefas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        agente_id INTEGER,
        titulo TEXT NOT NULL,
        descricao TEXT,
        status TEXT DEFAULT 'pendente',
        prioridade TEXT DEFAULT 'normal',
        criado_em TEXT DEFAULT (datetime('now','localtime')),
        concluido_em TEXT
    )""")

    c.execute("""
    CREATE TABLE IF NOT EXISTS escritorio_docs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        agente_id INTEGER,
        titulo TEXT NOT NULL,
        conteudo TEXT,
        criado_em TEXT DEFAULT (datetime('now','localtime'))
    )""")

    # Inserir agentes padrão se a tabela estiver vazia
    c.execute("SELECT COUNT(*) FROM escritorio_agentes")
    if c.fetchone()[0] == 0:
        agentes_padrao = [
            ("Ricardo", "CEO / Coordenador", "👨‍💼", "#7c3aed",
             "Você é Ricardo, CEO e coordenador da equipe. Coordena os outros agentes, define estratégias e resolve questões jurídicas complexas. É experiente, estratégico e objetivo."),
            ("Marina", "Redatora Jurídica", "⚖️", "#0ea5e9",
             "Você é Marina, especialista em redação jurídica. Redige petições, ofícios, mandados e documentos legais com precisão técnica. Conhece profundamente o direito processual."),
            ("Carlos", "Pesquisador", "🔍", "#10b981",
             "Você é Carlos, pesquisador jurídico. Busca jurisprudência, legislação e doutrina relevantes para cada caso. Analisa precedentes e sugere fundamentos legais."),
            ("Beatriz", "Agenda & Prazos", "📅", "#f59e0b",
             "Você é Beatriz, responsável por agenda e controle de prazos processuais. Organiza compromissos, alerta sobre vencimentos e mantém o escritório pontual."),
        ]
        for nome, cargo, avatar, cor, prompt in agentes_padrao:
            c.execute(
                "INSERT INTO escritorio_agentes (nome, cargo, avatar, cor, system_prompt) VALUES (?,?,?,?,?)",
                (nome, cargo, avatar, cor, prompt)
            )

    conn.commit()
    conn.close()
    print(f"[DB] Banco inicializado: {DB_PATH}")

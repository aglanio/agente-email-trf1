"""
app.py — Backend FastAPI do Agente E-mail Pessoal.
Porta: 8090
"""
import os
import sqlite3

# Carrega .env se existir
_env_path = os.path.join(os.path.dirname(__file__), ".env")
if os.path.exists(_env_path):
    with open(_env_path) as _f:
        for _line in _f:
            _line = _line.strip()
            if _line and not _line.startswith('#') and '=' in _line:
                _k, _, _v = _line.partition('=')
                os.environ.setdefault(_k.strip(), _v.strip())
import shutil
import tempfile
from pathlib import Path
from typing import Optional
from datetime import datetime

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from database import DB_PATH, init_db
from importer import importar_todas
from pdf_extractor import extrair_info_pdf, montar_assunto, montar_corpo
from email_matcher import buscar_email_completo, salvar_email_manual, buscar_sugestoes

# ── SMTP Config (funciona em Linux/Docker) ───────────────────────────────────
SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.office365.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER", "aglanio.carvalho@trf1.jus.br")
SMTP_PASS = os.environ.get("SMTP_PASS", "")
REMETENTE = SMTP_USER


def enviar_email_smtp(
    destinatario_email: str,
    assunto: str,
    corpo: str,
    arquivo_mandado: Optional[str] = None,
    arquivo_anexo: Optional[str] = None,
    numero_processo: Optional[str] = None,
) -> dict:
    """Envia email direto via SMTP (Office 365)."""
    if not SMTP_PASS:
        return {"ok": False, "erro": "SMTP_PASS não configurado no .env"}
    if not destinatario_email:
        return {"ok": False, "erro": "E-mail do destinatário não encontrado"}

    try:
        msg = MIMEMultipart()
        msg["From"] = REMETENTE
        msg["To"] = destinatario_email
        msg["Subject"] = assunto or f"Processo {numero_processo or ''}"
        msg.attach(MIMEText(corpo or _corpo_padrao_smtp(numero_processo), "plain", "utf-8"))

        # Anexar arquivos
        for filepath in [arquivo_mandado, arquivo_anexo]:
            if filepath and Path(filepath).exists():
                part = MIMEBase("application", "octet-stream")
                with open(filepath, "rb") as f:
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename={Path(filepath).name}")
                msg.attach(part)

        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
        server.starttls()
        server.login(SMTP_USER, SMTP_PASS)
        server.sendmail(REMETENTE, destinatario_email, msg.as_string())
        server.quit()

        return {
            "ok": True,
            "msg": f"E-mail enviado para {destinatario_email}",
            "assunto": assunto,
            "mode": "smtp",
        }
    except Exception as e:
        return {"ok": False, "erro": f"Erro SMTP: {str(e)}"}


def _corpo_padrao_smtp(numero_processo: Optional[str] = None) -> str:
    from datetime import date
    proc = numero_processo or "[número do processo]"
    meses = ["janeiro","fevereiro","março","abril","maio","junho",
             "julho","agosto","setembro","outubro","novembro","dezembro"]
    hoje = date.today()
    data_str = f"Teresina, {hoje.day} de {meses[hoje.month-1]} de {hoje.year}"
    return f"""Prezado(a) Senhor(a),

Encaminhamos em anexo o mandado judicial referente ao processo {proc}, expedido pela Seção Judiciária do Piauí - SJPI.

Atenciosamente,

{data_str}

Aglanio Frota Moura Carvalho
Oficial de Justiça Avaliador Federal     PI100327
Seção Judiciária do Piauí - TRF 1ª Região
aglanio.carvalho@trf1.jus.br
"""


def verificar_smtp() -> dict:
    """Verifica se SMTP está configurado e funcional."""
    if not SMTP_PASS:
        return {"ok": False, "erro": "SMTP_PASS não configurado", "available": False}
    try:
        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=10)
        server.starttls()
        server.login(SMTP_USER, SMTP_PASS)
        server.quit()
        return {"ok": True, "available": True, "mode": "smtp", "host": SMTP_HOST, "user": SMTP_USER}
    except Exception as e:
        return {"ok": False, "erro": str(e), "available": False}


# Outlook integration (Windows-only via pywin32)
OUTLOOK_AVAILABLE = False
try:
    from outlook_agent import (
        criar_rascunho,
        criar_rascunhos_em_lote,
        verificar_outlook_disponivel,
        exportar_contatos_outlook,
        exportar_enviados_com_tratamento,
    )
    OUTLOOK_AVAILABLE = True
except ImportError:
    # Linux/Docker: use SMTP fallback
    def criar_rascunho(**kwargs):
        return enviar_email_smtp(**kwargs)
    def criar_rascunhos_em_lote(processos, **kwargs):
        resultados = []
        for p in processos:
            email = p.get("email_destinatario", "")
            if not email:
                resultados.append({"numero_processo": p.get("numero_processo", ""), "email": email, "status": "erro", "msg": "E-mail do destinatário não encontrado"})
                continue
            from pdf_extractor import montar_corpo
            corpo = montar_corpo(p)
            res = enviar_email_smtp(
                destinatario_email=email,
                assunto=p.get("assunto_email", ""),
                corpo=corpo,
                arquivo_mandado=p.get("arquivo_mandado"),
                arquivo_anexo=p.get("arquivo_anexo"),
                numero_processo=p.get("numero_processo"),
            )
            if res["ok"]:
                resultados.append({"numero_processo": p.get("numero_processo", ""), "email": email, "status": "rascunho_aberto", "msg": res["msg"]})
                _atualizar_status_processo_db(p.get("id"), "enviado")
            else:
                resultados.append({"numero_processo": p.get("numero_processo", ""), "email": email, "status": "erro", "msg": res.get("erro", "Erro")})
        return resultados
    def verificar_outlook_disponivel():
        return verificar_smtp()
    def exportar_contatos_outlook(**kwargs):
        return 0
    def exportar_enviados_com_tratamento(**kwargs):
        return {"ok": False, "erro": "Outlook não disponível em Linux. Contatos devem ser importados via agenda DOCX."}


def _atualizar_status_processo_db(processo_id, status):
    if not processo_id:
        return
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.execute("UPDATE processos SET status = ? WHERE id = ?", (status, processo_id))
        conn.commit()
        conn.close()
    except Exception:
        pass

# ── Downloads temporários ────────────────────────────────────────────────────
DOWNLOADS_DIR = Path(__file__).parent / "downloads"
DOWNLOADS_DIR.mkdir(exist_ok=True)

# ── App ───────────────────────────────────────────────────────────────────────
app = FastAPI(title="Agente E-mail TRF1", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["X-Processo", "X-Email", "X-Acao", "X-Id"],
)

@app.on_event("startup")
async def startup():
    init_db()


# ── Modelos Pydantic ──────────────────────────────────────────────────────────

class ProcessoUpdate(BaseModel):
    destinatario: Optional[str] = None
    email_destinatario: Optional[str] = None
    assunto_email: Optional[str] = None
    status: Optional[str] = None
    grau: Optional[int] = None

class EmailManual(BaseModel):
    nome_orgao: str
    email: str

class RascunhoRequest(BaseModel):
    processo_ids: list[int]

class ProcessoCreate(BaseModel):
    numero_processo: str
    tipo: Optional[str] = None
    destinatario: Optional[str] = None
    endereco_destinatario: Optional[str] = None
    email_destinatario: Optional[str] = None
    assunto_email: Optional[str] = None


# ── Status / health ───────────────────────────────────────────────────────────

@app.get("/api/status")
def status():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM contatos WHERE email != ''")
    total_contatos = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM processos")
    total_processos = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM processos WHERE status = 'pendente'")
    pendentes = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM processos WHERE status = 'rascunho'")
    rascunhos = cur.fetchone()[0]
    conn.close()

    outlook_info = verificar_outlook_disponivel()

    return {
        "ok": True,
        "contatos": total_contatos,
        "processos": total_processos,
        "pendentes": pendentes,
        "rascunhos": rascunhos,
        "outlook": outlook_info,
    }


# ── Contatos ──────────────────────────────────────────────────────────────────

@app.get("/api/contatos")
def listar_contatos(q: Optional[str] = None, limit: int = 50):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    if q:
        cur.execute(
            """SELECT * FROM contatos WHERE nome LIKE ? OR orgao LIKE ? OR email LIKE ?
               ORDER BY nome LIMIT ?""",
            (f"%{q}%", f"%{q}%", f"%{q}%", limit),
        )
    else:
        cur.execute("SELECT * FROM contatos ORDER BY nome LIMIT ?", (limit,))
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return {"contatos": rows, "total": len(rows)}


@app.post("/api/contatos/importar")
def importar_contatos(background_tasks: BackgroundTasks):
    """Importa contatos das agendas DOCX + Outlook."""
    def _importar():
        n_docx = importar_todas()
        n_outlook = exportar_contatos_outlook(destino_db=True)
        return n_docx + n_outlook

    background_tasks.add_task(_importar)
    return {"ok": True, "msg": "Importação iniciada em segundo plano"}


@app.post("/api/contatos/importar/sync")
def importar_contatos_sync():
    """Importa contatos de forma síncrona (pode demorar)."""
    n_docx = importar_todas()
    n_outlook = exportar_contatos_outlook(destino_db=True)
    return {"ok": True, "importados_docx": n_docx, "importados_outlook": n_outlook}


@app.post("/api/contatos/salvar-email")
def salvar_email_contato(data: EmailManual):
    ok = salvar_email_manual(data.nome_orgao, data.email)
    return {"ok": ok}


@app.get("/api/contatos/buscar")
def buscar_email(nome: str):
    """Busca email de um destinatário pelo nome/órgão."""
    resultado = buscar_email_completo(nome)
    sugestoes = buscar_sugestoes(nome)
    resultado["sugestoes"] = sugestoes
    return resultado


@app.post("/api/outlook/importar-enviados")
def importar_enviados_outlook(background_tasks: BackgroundTasks):
    """
    Varre TODOS os e-mails enviados do Outlook, extrai destinatários únicos
    com formas de tratamento (Ilmo. Sr., Exmo. Sr., etc.) e salva na base de contatos.
    """
    resultado = exportar_enviados_com_tratamento()
    return resultado


@app.get("/api/contatos/{contato_id}/tratamento")
def obter_tratamento(contato_id: int):
    """Retorna a forma de tratamento de um contato."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("SELECT nome, email, tratamento FROM contatos WHERE id = ?", (contato_id,))
    row = cur.fetchone()
    conn.close()
    if not row:
        raise HTTPException(404, "Contato não encontrado")
    return dict(row)


# ── Processos ─────────────────────────────────────────────────────────────────

@app.get("/api/processos")
def listar_processos(status: Optional[str] = None, limit: int = 100):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    if status:
        cur.execute(
            "SELECT * FROM processos WHERE status = ? ORDER BY id DESC LIMIT ?",
            (status, limit),
        )
    else:
        cur.execute("SELECT * FROM processos ORDER BY id DESC LIMIT ?", (limit,))
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return {"processos": rows, "total": len(rows)}


@app.get("/api/processos/{processo_id}")
def obter_processo(processo_id: int):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("SELECT * FROM processos WHERE id = ?", (processo_id,))
    row = cur.fetchone()
    conn.close()
    if not row:
        raise HTTPException(404, "Processo não encontrado")
    return dict(row)


@app.post("/api/processos")
def criar_processo(data: ProcessoCreate):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        """INSERT INTO processos
           (numero_processo, tipo, destinatario, endereco_destinatario,
            email_destinatario, assunto_email, status)
           VALUES (?, ?, ?, ?, ?, ?, 'pendente')""",
        (
            data.numero_processo, data.tipo, data.destinatario,
            data.endereco_destinatario, data.email_destinatario,
            data.assunto_email,
        ),
    )
    processo_id = cur.lastrowid
    conn.commit()
    conn.close()
    return {"ok": True, "id": processo_id}


@app.patch("/api/processos/{processo_id}")
def atualizar_processo(processo_id: int, data: ProcessoUpdate):
    conn = sqlite3.connect(DB_PATH)
    updates = {}
    if data.destinatario is not None:
        updates["destinatario"] = data.destinatario
    if data.email_destinatario is not None:
        updates["email_destinatario"] = data.email_destinatario
    if data.assunto_email is not None:
        updates["assunto_email"] = data.assunto_email
    if data.status is not None:
        updates["status"] = data.status
    if data.grau is not None:
        updates["grau"] = data.grau

    if updates:
        sets = ", ".join(f"{k} = ?" for k in updates)
        vals = list(updates.values()) + [processo_id]
        conn.execute(f"UPDATE processos SET {sets} WHERE id = ?", vals)
        conn.commit()
    conn.close()
    return {"ok": True}


@app.delete("/api/processos/{processo_id}")
def deletar_processo(processo_id: int):
    conn = sqlite3.connect(DB_PATH)
    conn.execute("DELETE FROM processos WHERE id = ?", (processo_id,))
    conn.commit()
    conn.close()
    return {"ok": True}


# ── Upload de PDFs ────────────────────────────────────────────────────────────

@app.post("/api/processos/upload-pdf")
async def upload_pdf(
    mandado: UploadFile = File(...),
    anexo: Optional[UploadFile] = File(None),
):
    """
    Recebe os 2 PDFs do processo, extrai informações e cria registro.
    Retorna processo com email já buscado.
    """
    # Salvar mandado
    mandado_path = DOWNLOADS_DIR / f"mandado_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{mandado.filename}"
    with open(mandado_path, "wb") as f:
        f.write(await mandado.read())

    # Salvar anexo (opcional)
    anexo_path = None
    if anexo and anexo.filename:
        anexo_path = DOWNLOADS_DIR / f"anexo_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{anexo.filename}"
        with open(anexo_path, "wb") as f:
            f.write(await anexo.read())

    # Extrair informações do mandado
    info = extrair_info_pdf(str(mandado_path))
    assunto = montar_assunto(info)
    corpo = montar_corpo(info)

    # Buscar email
    destinatario = info.get("destinatario") or ""
    email_result = buscar_email_completo(destinatario) if destinatario else {}

    # Criar registro no banco
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        """INSERT OR REPLACE INTO processos
           (numero_processo, tipo, destinatario, endereco_destinatario,
            email_destinatario, email_encontrado_em, assunto_email,
            arquivo_mandado, arquivo_anexo, status)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pendente')""",
        (
            info.get("numero_processo", ""),
            info.get("tipo", "intimacao"),
            destinatario,
            info.get("endereco", ""),
            email_result.get("email", ""),
            email_result.get("fonte", ""),
            assunto,
            str(mandado_path),
            str(anexo_path) if anexo_path else None,
        ),
    )
    processo_id = cur.lastrowid
    conn.commit()
    conn.close()

    return {
        "ok": True,
        "id": processo_id,
        "numero_processo": info.get("numero_processo"),
        "tipo": info.get("tipo"),
        "destinatario": destinatario,
        "email": email_result.get("email", ""),
        "email_confianca": email_result.get("confianca", ""),
        "email_fonte": email_result.get("fonte", ""),
        "sugestoes": email_result.get("sugestoes", []),
        "assunto": assunto,
        "arquivo_mandado": str(mandado_path),
        "arquivo_anexo": str(anexo_path) if anexo_path else None,
    }


# ── Upload de arquivos para processo existente ────────────────────────────────

@app.post("/api/processos/{processo_id}/arquivos")
async def upload_arquivos_processo(
    processo_id: int,
    mandado: Optional[UploadFile] = File(None),
    anexo: Optional[UploadFile] = File(None),
):
    """Anexa PDFs a um processo já criado."""
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT id FROM processos WHERE id = ?", (processo_id,))
    if not cur.fetchone():
        conn.close()
        raise HTTPException(404, "Processo não encontrado")

    updates = {}
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')

    if mandado and mandado.filename:
        path = DOWNLOADS_DIR / f"mandado_{processo_id}_{ts}_{mandado.filename}"
        with open(path, "wb") as f:
            f.write(await mandado.read())
        updates["arquivo_mandado"] = str(path)

    if anexo and anexo.filename:
        path = DOWNLOADS_DIR / f"anexo_{processo_id}_{ts}_{anexo.filename}"
        with open(path, "wb") as f:
            f.write(await anexo.read())
        updates["arquivo_anexo"] = str(path)

    if updates:
        sets = ", ".join(f"{k} = ?" for k in updates)
        vals = list(updates.values()) + [processo_id]
        conn.execute(f"UPDATE processos SET {sets} WHERE id = ?", vals)
        conn.commit()
    conn.close()
    return {"ok": True, "atualizados": list(updates.keys())}


# ── Rascunhos Outlook ─────────────────────────────────────────────────────────

@app.post("/api/rascunhos/criar")
def criar_rascunho_endpoint(data: RascunhoRequest):
    """Cria rascunhos no Outlook para os processos selecionados."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    processos_data = []
    for pid in data.processo_ids:
        cur.execute("SELECT * FROM processos WHERE id = ?", (pid,))
        row = cur.fetchone()
        if row:
            p = dict(row)
            # Monta corpo se não tiver
            if not p.get("corpo_email"):
                from pdf_extractor import montar_corpo
                p["corpo_email"] = montar_corpo(p)
            processos_data.append(p)

    conn.close()

    if not processos_data:
        return {"ok": False, "erro": "Nenhum processo encontrado"}

    resultados = criar_rascunhos_em_lote(processos_data)
    sucessos = sum(1 for r in resultados if r["status"] == "rascunho_aberto")
    erros = [r for r in resultados if r["status"] == "erro"]

    return {
        "ok": True,
        "total": len(resultados),
        "sucessos": sucessos,
        "erros": erros,
        "resultados": resultados,
    }


@app.post("/api/rascunhos/criar-um/{processo_id}")
def criar_rascunho_unico(processo_id: int):
    """Cria um único rascunho no Outlook."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("SELECT * FROM processos WHERE id = ?", (processo_id,))
    row = cur.fetchone()
    conn.close()

    if not row:
        raise HTTPException(404, "Processo não encontrado")

    p = dict(row)
    if not p.get("email_destinatario"):
        raise HTTPException(400, "E-mail do destinatário não encontrado")

    from pdf_extractor import montar_corpo
    corpo = montar_corpo(p)

    res = criar_rascunho(
        destinatario_email=p["email_destinatario"],
        assunto=p.get("assunto_email", ""),
        corpo=corpo,
        arquivo_mandado=p.get("arquivo_mandado"),
        arquivo_anexo=p.get("arquivo_anexo"),
        numero_processo=p.get("numero_processo"),
    )

    if res["ok"]:
        conn = sqlite3.connect(DB_PATH)
        status_novo = "enviado" if res.get("mode") == "smtp" else "rascunho"
        conn.execute(f"UPDATE processos SET status = '{status_novo}' WHERE id = ?", (processo_id,))
        conn.commit()
        conn.close()

    return res


# ── Gerar .eml para abrir no Outlook ─────────────────────────────────────────

@app.get("/api/rascunhos/eml/{processo_id}")
def gerar_eml(processo_id: int):
    """Gera arquivo .eml que o usuário baixa e abre no Outlook para revisar e enviar."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("SELECT * FROM processos WHERE id = ?", (processo_id,))
    row = cur.fetchone()
    conn.close()

    if not row:
        raise HTTPException(404, "Processo não encontrado")

    p = dict(row)
    if not p.get("email_destinatario"):
        raise HTTPException(400, "E-mail do destinatário não encontrado")

    from pdf_extractor import montar_corpo
    corpo = montar_corpo(p)
    assunto = p.get("assunto_email") or f"Processo {p.get('numero_processo', '')}"

    # Monta mensagem MIME
    msg = MIMEMultipart()
    msg["From"] = REMETENTE
    msg["To"] = p["email_destinatario"]
    msg["Subject"] = assunto
    msg["X-Unsent"] = "1"  # Marca como NÃO enviado — Outlook abre em modo compose
    msg.attach(MIMEText(corpo, "plain", "utf-8"))

    # Anexar arquivos
    for filepath in [p.get("arquivo_mandado"), p.get("arquivo_anexo")]:
        if filepath and Path(filepath).exists():
            part = MIMEBase("application", "octet-stream")
            with open(filepath, "rb") as f:
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={Path(filepath).name}")
            msg.attach(part)

    # Atualiza status
    conn = sqlite3.connect(DB_PATH)
    conn.execute("UPDATE processos SET status = 'rascunho' WHERE id = ?", (processo_id,))
    conn.commit()
    conn.close()

    # Salva .eml temporário e retorna
    numero = p.get("numero_processo", str(processo_id)).replace("/", "-")
    filename = f"rascunho_{numero}.eml"
    eml_path = DOWNLOADS_DIR / filename
    with open(eml_path, "w", encoding="utf-8") as f:
        f.write(msg.as_string())

    return FileResponse(
        path=str(eml_path),
        filename=filename,
        media_type="application/octet-stream",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/api/rascunhos/eml-lote")
def gerar_eml_lote(data: RascunhoRequest):
    """Gera .eml para múltiplos processos. Retorna lista de IDs com download disponível."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    resultados = []
    for pid in data.processo_ids:
        cur.execute("SELECT * FROM processos WHERE id = ?", (pid,))
        row = cur.fetchone()
        if not row:
            resultados.append({"id": pid, "ok": False, "erro": "Não encontrado"})
            continue
        p = dict(row)
        if not p.get("email_destinatario"):
            resultados.append({"id": pid, "ok": False, "erro": "Sem e-mail", "numero_processo": p.get("numero_processo", "")})
            continue

        from pdf_extractor import montar_corpo
        corpo = montar_corpo(p)
        assunto = p.get("assunto_email") or f"Processo {p.get('numero_processo', '')}"

        msg = MIMEMultipart()
        msg["From"] = REMETENTE
        msg["To"] = p["email_destinatario"]
        msg["Subject"] = assunto
        msg["X-Unsent"] = "1"
        msg.attach(MIMEText(corpo, "plain", "utf-8"))

        for filepath in [p.get("arquivo_mandado"), p.get("arquivo_anexo")]:
            if filepath and Path(filepath).exists():
                part = MIMEBase("application", "octet-stream")
                with open(filepath, "rb") as f:
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename={Path(filepath).name}")
                msg.attach(part)

        numero = p.get("numero_processo", str(pid)).replace("/", "-")
        filename = f"rascunho_{numero}.eml"
        eml_path = DOWNLOADS_DIR / filename
        with open(eml_path, "w", encoding="utf-8") as f:
            f.write(msg.as_string())

        conn.execute("UPDATE processos SET status = 'rascunho' WHERE id = ?", (pid,))
        resultados.append({"id": pid, "ok": True, "numero_processo": p.get("numero_processo", ""), "filename": filename})

    conn.commit()
    conn.close()

    sucessos = sum(1 for r in resultados if r.get("ok"))
    return {"ok": True, "total": len(resultados), "sucessos": sucessos, "resultados": resultados}


# ── Envio direto via SMTP ────────────────────────────────────────────────────

class EnvioRequest(BaseModel):
    processo_ids: list[int]

@app.post("/api/email/enviar/{processo_id}")
def enviar_email_endpoint(processo_id: int):
    """Envia e-mail direto via SMTP para um processo."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("SELECT * FROM processos WHERE id = ?", (processo_id,))
    row = cur.fetchone()
    conn.close()

    if not row:
        raise HTTPException(404, "Processo não encontrado")

    p = dict(row)
    if not p.get("email_destinatario"):
        raise HTTPException(400, "E-mail do destinatário não encontrado")

    from pdf_extractor import montar_corpo
    corpo = montar_corpo(p)

    res = enviar_email_smtp(
        destinatario_email=p["email_destinatario"],
        assunto=p.get("assunto_email", ""),
        corpo=corpo,
        arquivo_mandado=p.get("arquivo_mandado"),
        arquivo_anexo=p.get("arquivo_anexo"),
        numero_processo=p.get("numero_processo"),
    )

    if res["ok"]:
        # Registra envio e atualiza status
        conn = sqlite3.connect(DB_PATH)
        conn.execute("UPDATE processos SET status = 'enviado' WHERE id = ?", (processo_id,))
        conn.execute(
            """INSERT INTO emails_enviados
               (processo_id, numero_processo, destinatario, email, assunto, arquivo_mandado, arquivo_anexo, enviado_em)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (processo_id, p.get("numero_processo", ""), p.get("destinatario", ""),
             p["email_destinatario"], p.get("assunto_email", ""),
             p.get("arquivo_mandado", ""), p.get("arquivo_anexo", ""),
             datetime.now().isoformat()),
        )
        conn.commit()
        conn.close()

    return res


@app.post("/api/email/enviar-lote")
def enviar_email_lote(data: EnvioRequest):
    """Envia e-mails em lote via SMTP."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    resultados = []
    for pid in data.processo_ids:
        cur.execute("SELECT * FROM processos WHERE id = ?", (pid,))
        row = cur.fetchone()
        if not row:
            continue
        p = dict(row)
        if not p.get("email_destinatario"):
            resultados.append({"numero_processo": p.get("numero_processo", ""), "email": "", "status": "erro", "msg": "Sem e-mail"})
            continue

        from pdf_extractor import montar_corpo
        corpo = montar_corpo(p)
        res = enviar_email_smtp(
            destinatario_email=p["email_destinatario"],
            assunto=p.get("assunto_email", ""),
            corpo=corpo,
            arquivo_mandado=p.get("arquivo_mandado"),
            arquivo_anexo=p.get("arquivo_anexo"),
            numero_processo=p.get("numero_processo"),
        )
        if res["ok"]:
            conn.execute("UPDATE processos SET status = 'enviado' WHERE id = ?", (pid,))
            conn.execute(
                """INSERT INTO emails_enviados
                   (processo_id, numero_processo, destinatario, email, assunto, arquivo_mandado, arquivo_anexo, enviado_em)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                (pid, p.get("numero_processo", ""), p.get("destinatario", ""),
                 p["email_destinatario"], p.get("assunto_email", ""),
                 p.get("arquivo_mandado", ""), p.get("arquivo_anexo", ""),
                 datetime.now().isoformat()),
            )
            resultados.append({"numero_processo": p.get("numero_processo", ""), "email": p["email_destinatario"], "status": "enviado", "msg": res["msg"]})
        else:
            resultados.append({"numero_processo": p.get("numero_processo", ""), "email": p.get("email_destinatario", ""), "status": "erro", "msg": res.get("erro", "Erro")})

    conn.commit()
    conn.close()

    sucessos = sum(1 for r in resultados if r["status"] == "enviado")
    return {"ok": True, "total": len(resultados), "sucessos": sucessos, "erros": [r for r in resultados if r["status"] == "erro"], "resultados": resultados}


@app.get("/api/email/preview/{processo_id}")
def preview_email(processo_id: int):
    """Retorna preview do e-mail sem enviar."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("SELECT * FROM processos WHERE id = ?", (processo_id,))
    row = cur.fetchone()
    conn.close()

    if not row:
        raise HTTPException(404, "Processo não encontrado")

    p = dict(row)
    from pdf_extractor import montar_corpo
    corpo = montar_corpo(p)
    anexos = []
    if p.get("arquivo_mandado") and Path(p["arquivo_mandado"]).exists():
        anexos.append(Path(p["arquivo_mandado"]).name)
    if p.get("arquivo_anexo") and Path(p["arquivo_anexo"]).exists():
        anexos.append(Path(p["arquivo_anexo"]).name)

    return {
        "ok": True,
        "preview": {
            "to": p.get("email_destinatario", ""),
            "from": REMETENTE,
            "subject": p.get("assunto_email", f"Processo {p.get('numero_processo', '')}"),
            "body": corpo,
            "attachments": anexos,
        }
    }


# ── Histórico de enviados ─────────────────────────────────────────────────────

@app.get("/api/historico")
def listar_historico(limit: int = 50):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("SELECT * FROM emails_enviados ORDER BY enviado_em DESC LIMIT ?", (limit,))
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return {"historico": rows, "total": len(rows)}


@app.post("/api/historico/registrar/{processo_id}")
def registrar_envio(processo_id: int):
    """Registra que o usuário enviou o e-mail de um processo."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("SELECT * FROM processos WHERE id = ?", (processo_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        raise HTTPException(404, "Processo não encontrado")

    p = dict(row)
    cur.execute(
        """INSERT INTO emails_enviados
           (processo_id, numero_processo, destinatario, email, assunto, arquivo_mandado, arquivo_anexo, enviado_em)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
        (
            processo_id,
            p.get("numero_processo", ""),
            p.get("destinatario", ""),
            p.get("email_destinatario", ""),
            p.get("assunto_email", ""),
            p.get("arquivo_mandado", ""),
            p.get("arquivo_anexo", ""),
            datetime.now().isoformat(),
        ),
    )
    conn.execute("UPDATE processos SET status = 'enviado' WHERE id = ?", (processo_id,))
    conn.commit()
    conn.close()
    return {"ok": True}


# ── Templates ─────────────────────────────────────────────────────────────────

@app.get("/api/templates")
def listar_templates():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("SELECT * FROM templates ORDER BY id")
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return {"templates": rows}


# ── Upload em Lote de PDFs ────────────────────────────────────────────────────

@app.post("/api/processos/upload-lote")
async def upload_lote(arquivos: list[UploadFile] = File(...)):
    """
    Recebe múltiplos PDFs de uma vez.
    Extrai o número CNJ de cada um e cria/atualiza processos automaticamente.
    Tenta parear mandado+anexo pelo número do processo.
    """
    resultados = []
    # Primeiro, salvar todos e extrair info
    infos = []
    for arq in arquivos:
        ts = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
        path = DOWNLOADS_DIR / f"lote_{ts}_{arq.filename}"
        with open(path, "wb") as f:
            f.write(await arq.read())
        info = extrair_info_pdf(str(path))
        info["_path"] = str(path)
        info["_filename"] = arq.filename or ""
        infos.append(info)

    # Agrupar por número de processo
    por_processo = {}
    sem_numero = []
    for info in infos:
        num = info.get("numero_processo", "")
        if num:
            por_processo.setdefault(num, []).append(info)
        else:
            sem_numero.append(info)

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    for num, docs in por_processo.items():
        # Decidir qual é mandado e qual é anexo
        mandado_path = None
        anexo_path = None
        info_principal = docs[0]

        for doc in docs:
            fname = doc["_filename"].lower()
            if "anexo" in fname:
                anexo_path = doc["_path"]
            elif "mandado" in fname or not mandado_path:
                mandado_path = doc["_path"]
                info_principal = doc

        # Se tem 2+ docs e nenhum foi identificado como anexo, o segundo vira anexo
        if len(docs) >= 2 and not anexo_path:
            for doc in docs:
                if doc["_path"] != mandado_path:
                    anexo_path = doc["_path"]
                    break

        # Buscar email
        destinatario = info_principal.get("destinatario") or ""
        email_result = buscar_email_completo(destinatario) if destinatario else {}
        assunto = montar_assunto(info_principal)

        # Verificar se processo já existe
        cur.execute("SELECT id FROM processos WHERE numero_processo = ?", (num,))
        existing = cur.fetchone()

        if existing:
            # Atualizar arquivos
            updates = {}
            if mandado_path:
                updates["arquivo_mandado"] = mandado_path
            if anexo_path:
                updates["arquivo_anexo"] = anexo_path
            if updates:
                sets = ", ".join(f"{k} = ?" for k in updates)
                vals = list(updates.values()) + [existing[0]]
                cur.execute(f"UPDATE processos SET {sets} WHERE id = ?", vals)
            resultados.append({
                "numero_processo": num, "acao": "atualizado",
                "id": existing[0], "ok": True,
            })
        else:
            cur.execute(
                """INSERT INTO processos
                   (numero_processo, tipo, destinatario, endereco_destinatario,
                    email_destinatario, email_encontrado_em, assunto_email,
                    arquivo_mandado, arquivo_anexo, status)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pendente')""",
                (
                    num,
                    info_principal.get("tipo", "intimacao"),
                    destinatario,
                    info_principal.get("endereco", ""),
                    email_result.get("email", ""),
                    email_result.get("fonte", ""),
                    assunto,
                    mandado_path,
                    anexo_path,
                ),
            )
            resultados.append({
                "numero_processo": num, "acao": "criado",
                "id": cur.lastrowid, "ok": True,
                "destinatario": destinatario,
                "email": email_result.get("email", ""),
            })

    # PDFs sem número CNJ
    for info in sem_numero:
        resultados.append({
            "numero_processo": "", "acao": "ignorado",
            "ok": False, "erro": "Número CNJ não encontrado",
            "arquivo": info["_filename"],
        })

    conn.commit()
    conn.close()

    return {
        "ok": True,
        "total": len(arquivos),
        "processados": len(por_processo),
        "ignorados": len(sem_numero),
        "resultados": resultados,
    }


# ── Captura HTML do PJe (Bookmarklet) ────────────────────────────────────────

class CapturaPainelRequest(BaseModel):
    processos: list[dict]  # [{numero, destinatario, endereco, tipo}, ...]

@app.post("/api/pje/captura-painel")
async def captura_painel_pje(data: CapturaPainelRequest):
    """
    Recebe múltiplos processos extraídos do painel PJe via bookmarklet.
    Cria todos de uma vez.
    """
    import re as _re
    from pdf_extractor import _extrair_email_pje

    resultados = []
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    for proc in data.processos:
        numero = proc.get("numero", "").strip()
        if not numero:
            continue

        destinatario = proc.get("destinatario", "").strip()
        endereco = proc.get("endereco", "").strip()
        tipo = proc.get("tipo", "intimacao").strip()

        # Extrair email do endereço (padrão PJe sem @)
        email_direto = _extrair_email_pje(endereco) if endereco else ""
        email_result = {}
        if email_direto:
            email_result = {"email": email_direto, "fonte": "pje_endereco"}
        elif destinatario:
            email_result = buscar_email_completo(destinatario)

        email = email_result.get("email", "")
        fonte = email_result.get("fonte", "")

        # Verificar se já existe
        cur.execute("SELECT id, email_destinatario FROM processos WHERE numero_processo = ?", (numero,))
        existing = cur.fetchone()

        if existing:
            updates = {}
            if destinatario:
                updates["destinatario"] = destinatario
            if email and not existing[1]:
                updates["email_destinatario"] = email
                updates["email_encontrado_em"] = fonte
            if updates:
                sets = ", ".join(f"{k} = ?" for k in updates)
                vals = list(updates.values()) + [existing[0]]
                cur.execute(f"UPDATE processos SET {sets} WHERE id = ?", vals)
            resultados.append({"numero": numero, "id": existing[0], "acao": "atualizado",
                               "email": email or (existing[1] or ""), "ok": True})
        else:
            assunto = f"Intimação - Processo {numero} - SJPI"
            cur.execute(
                """INSERT INTO processos
                   (numero_processo, tipo, destinatario, endereco_destinatario,
                    email_destinatario, email_encontrado_em, assunto_email, status)
                   VALUES (?, ?, ?, ?, ?, ?, ?, 'pendente')""",
                (numero, tipo or "intimacao", destinatario, endereco, email, fonte, assunto),
            )
            resultados.append({"numero": numero, "id": cur.lastrowid, "acao": "criado",
                               "email": email, "ok": True})

    conn.commit()
    conn.close()

    return {
        "ok": True,
        "total": len(resultados),
        "criados": sum(1 for r in resultados if r["acao"] == "criado"),
        "atualizados": sum(1 for r in resultados if r["acao"] == "atualizado"),
        "com_email": sum(1 for r in resultados if r.get("email")),
        "resultados": resultados,
    }


class CapturaHtmlRequest(BaseModel):
    html: str
    url: Optional[str] = ""
    titulo: Optional[str] = ""

@app.post("/api/pje/captura-pdf")
async def captura_pdf_pje(data: CapturaHtmlRequest):
    """
    Recebe HTML do mandado e devolve o PDF para download.
    Também cria/atualiza o processo no banco.
    """
    import re as _re
    import html as html_mod
    from pdf_extractor import _extrair_email_pje

    html_content = data.html
    if not html_content or len(html_content) < 50:
        raise HTTPException(400, "HTML vazio")

    # Extrair número CNJ
    num_match = _re.search(r'\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}', html_content)
    numero_processo = num_match.group(0) if num_match else ""
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    num_safe = numero_processo.replace("-", "_").replace(".", "_") if numero_processo else ts

    # Salvar HTML com estilos inline para imprimir bonito
    html_completo = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>body{{font-family:serif;margin:20mm;font-size:12pt}}
table{{border-collapse:collapse;width:100%}}
td,th{{border:1px solid #ccc;padding:4px 8px;font-size:11pt}}</style>
</head><body>{html_content}</body></html>"""

    pdf_path = DOWNLOADS_DIR / f"mandado_{num_safe}.pdf"
    html_path = DOWNLOADS_DIR / f"mandado_{num_safe}.html"

    # Salvar HTML sempre (funciona como fallback e para anexar no email)
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_completo)

    # Tentar converter pra PDF
    pdf_ok = False
    try:
        from weasyprint import HTML as WeasyprintHTML
        WeasyprintHTML(string=html_completo).write_pdf(str(pdf_path))
        pdf_ok = True
    except Exception:
        pass

    arquivo_final = str(pdf_path) if pdf_ok else str(html_path)

    # Extrair info do texto
    texto_limpo = _re.sub(r'<[^>]+>', ' ', html_content)
    texto_limpo = html_mod.unescape(texto_limpo)

    # Destinatário
    destinatario = ""
    dest_match = _re.search(
        r'INTIMA[ÇC][ÃA]O\s+DE[:\s]+(.+?)(?:\n|Endere)',
        texto_limpo, _re.IGNORECASE | _re.DOTALL
    )
    if dest_match:
        destinatario = dest_match.group(1).strip()[:200]
    if not destinatario:
        dest2 = _re.search(r'Destinat[aá]rio\(?s?\)?\s*(.+?)(?:\s{3,}|Rua |Endere|CEP|Expedi)',
                           texto_limpo, _re.IGNORECASE)
        if dest2:
            destinatario = dest2.group(1).strip()[:200]

    # Email
    email_direto = _extrair_email_pje(texto_limpo)
    email_result = {}
    if email_direto:
        email_result = {"email": email_direto, "fonte": "pje_endereco"}
    elif destinatario:
        email_result = buscar_email_completo(destinatario)
    email = email_result.get("email", "")

    # Criar/atualizar processo
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    acao = "criado"

    if numero_processo:
        cur.execute("SELECT id, email_destinatario FROM processos WHERE numero_processo = ?", (numero_processo,))
        existing = cur.fetchone()
        if existing:
            updates = {"arquivo_mandado": arquivo_final}
            if destinatario:
                updates["destinatario"] = destinatario
            if email and not existing[1]:
                updates["email_destinatario"] = email
                updates["email_encontrado_em"] = email_result.get("fonte", "")
            sets = ", ".join(f"{k} = ?" for k in updates)
            vals = list(updates.values()) + [existing[0]]
            cur.execute(f"UPDATE processos SET {sets} WHERE id = ?", vals)
            conn.commit()
            conn.close()
            # Retornar arquivo para download
            return FileResponse(
                arquivo_final,
                filename=f"mandado_{num_safe}.{'pdf' if pdf_ok else 'html'}",
                media_type="application/pdf" if pdf_ok else "text/html",
                headers={
                    "X-Processo": numero_processo,
                    "X-Email": email or "",
                    "X-Acao": "atualizado",
                    "X-Id": str(existing[0]),
                },
            )

    assunto = f"Intimação - Processo {numero_processo} - SJPI" if numero_processo else "Intimação - SJPI"
    cur.execute(
        """INSERT INTO processos
           (numero_processo, tipo, destinatario, email_destinatario,
            email_encontrado_em, assunto_email, arquivo_mandado, status)
           VALUES (?, 'intimacao', ?, ?, ?, ?, ?, 'pendente')""",
        (numero_processo, destinatario, email, email_result.get("fonte", ""), assunto, arquivo_final),
    )
    pid = cur.lastrowid
    conn.commit()
    conn.close()

    return FileResponse(
        arquivo_final,
        filename=f"mandado_{num_safe}.{'pdf' if pdf_ok else 'html'}",
        media_type="application/pdf" if pdf_ok else "text/html",
        headers={
            "X-Processo": numero_processo,
            "X-Email": email or "",
            "X-Acao": "criado",
            "X-Id": str(pid),
        },
    )


@app.post("/api/pje/captura-html")
async def captura_html_pje(data: CapturaHtmlRequest):
    """
    Recebe HTML capturado do PJe via bookmarklet.
    Converte para PDF e cria processo.
    """
    import re as _re

    html = data.html
    if not html or len(html) < 50:
        raise HTTPException(400, "HTML vazio ou muito curto")

    # Extrair número CNJ do HTML
    num_match = _re.search(r'\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}', html)
    numero_processo = num_match.group(0) if num_match else ""

    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    num_safe = numero_processo.replace("-", "_").replace(".", "_") if numero_processo else ts

    # Tentar converter HTML para PDF
    pdf_path = DOWNLOADS_DIR / f"pje_captura_{num_safe}.pdf"
    pdf_ok = False

    # Método 1: weasyprint
    try:
        from weasyprint import HTML as WeasyprintHTML
        WeasyprintHTML(string=html).write_pdf(str(pdf_path))
        pdf_ok = True
    except Exception:
        pass

    # Método 2: salvar como HTML (fallback)
    if not pdf_ok:
        html_path = DOWNLOADS_DIR / f"pje_captura_{num_safe}.html"
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html)
        # Usar o HTML como "mandado"
        pdf_path = html_path

    # Extrair info do texto do HTML
    import html as html_mod
    texto_limpo = _re.sub(r'<[^>]+>', ' ', html)
    texto_limpo = html_mod.unescape(texto_limpo)

    # Buscar destinatário — múltiplos padrões
    destinatario = ""
    # Padrão 1: "INTIMAÇÃO DE: ..."
    dest_match = _re.search(
        r'INTIMA[ÇC][ÃA]O\s+DE[:\s]+(.+?)(?:\n|Endere)',
        texto_limpo, _re.IGNORECASE | _re.DOTALL
    )
    if dest_match:
        destinatario = dest_match.group(1).strip()[:200]
    # Padrão 2: "Destinatário(s) ..." (painel PJe)
    if not destinatario:
        dest2 = _re.search(r'Destinat[aá]rio\(?s?\)?\s*(.+?)(?:\s{3,}|Rua |Endere|CEP|Expedi)',
                           texto_limpo, _re.IGNORECASE)
        if dest2:
            destinatario = dest2.group(1).strip()[:200]

    # Extrair email direto do HTML (padrão PJe: email sem @ no endereço)
    from pdf_extractor import _extrair_email_pje
    email_direto = _extrair_email_pje(texto_limpo)

    email_result = {}
    if email_direto:
        email_result = {"email": email_direto, "fonte": "pje_endereco"}
    elif destinatario:
        email_result = buscar_email_completo(destinatario)

    # Criar processo
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    # Verificar se já existe
    if numero_processo:
        cur.execute("SELECT id, email_destinatario FROM processos WHERE numero_processo = ?", (numero_processo,))
        existing = cur.fetchone()
        if existing:
            updates = {"arquivo_mandado": str(pdf_path)}
            # Atualizar destinatário e email se não tinha antes
            if destinatario:
                updates["destinatario"] = destinatario
            email_final = email_result.get("email", "")
            if email_final and not existing[1]:
                updates["email_destinatario"] = email_final
                if email_result.get("fonte"):
                    updates["email_encontrado_em"] = email_result["fonte"]
            sets = ", ".join(f"{k} = ?" for k in updates)
            vals = list(updates.values()) + [existing[0]]
            cur.execute(f"UPDATE processos SET {sets} WHERE id = ?", vals)
            conn.commit()
            conn.close()
            return {"ok": True, "id": existing[0], "acao": "atualizado",
                    "numero_processo": numero_processo,
                    "destinatario": destinatario,
                    "email": email_final or (existing[1] or ""),
                    "pdf": pdf_ok}

    assunto = f"Intimação - Processo {numero_processo} - SJPI" if numero_processo else "Intimação - SJPI"
    cur.execute(
        """INSERT INTO processos
           (numero_processo, tipo, destinatario, email_destinatario,
            email_encontrado_em, assunto_email, arquivo_mandado, status)
           VALUES (?, 'intimacao', ?, ?, ?, ?, ?, 'pendente')""",
        (
            numero_processo, destinatario,
            email_result.get("email", ""),
            email_result.get("fonte", ""),
            assunto, str(pdf_path),
        ),
    )
    pid = cur.lastrowid
    conn.commit()
    conn.close()

    return {
        "ok": True, "id": pid, "acao": "criado",
        "numero_processo": numero_processo,
        "destinatario": destinatario,
        "email": email_result.get("email", ""),
        "pdf": pdf_ok,
    }


# ── Monitorar pasta Downloads ─────────────────────────────────────────────────

_watcher_running = False

@app.post("/api/watcher/iniciar")
def iniciar_watcher():
    """Inicia monitoramento da pasta Downloads do usuário."""
    global _watcher_running
    import threading

    downloads_path = Path(os.environ.get("USERPROFILE", r"C:\Users\aglan")) / "Downloads"
    if not downloads_path.exists():
        return {"ok": False, "erro": "Pasta Downloads não encontrada"}

    if _watcher_running:
        return {"ok": True, "msg": "Watcher já está rodando", "pasta": str(downloads_path)}

    def watch_loop():
        global _watcher_running
        _watcher_running = True
        import time
        seen = set()
        # Marcar arquivos existentes como já vistos
        for f in downloads_path.glob("*.pdf"):
            seen.add(f.name)
        for f in downloads_path.glob("*.html"):
            seen.add(f.name)

        while _watcher_running:
            try:
                for ext in ["*.pdf", "*.html"]:
                    for f in downloads_path.glob(ext):
                        if f.name in seen:
                            continue
                        # Esperar arquivo terminar de ser escrito
                        time.sleep(1)
                        try:
                            size1 = f.stat().st_size
                            time.sleep(0.5)
                            size2 = f.stat().st_size
                            if size1 != size2:
                                continue  # Ainda escrevendo
                        except Exception:
                            continue

                        seen.add(f.name)

                        # Só processar se parece ser do PJe (tem mandado/intimação no nome ou conteúdo)
                        fname_lower = f.name.lower()
                        is_pje = any(kw in fname_lower for kw in
                                     ["mandado", "intimacao", "intimação", "citacao", "citação", "pje"])

                        if not is_pje and f.suffix == ".pdf":
                            # Verificar conteúdo
                            try:
                                info = extrair_info_pdf(str(f))
                                if info.get("numero_processo"):
                                    is_pje = True
                            except Exception:
                                pass

                        if not is_pje:
                            continue

                        # Importar o arquivo
                        try:
                            _importar_arquivo_watcher(f)
                        except Exception as e:
                            print(f"[watcher] Erro ao importar {f.name}: {e}")

                time.sleep(2)
            except Exception as e:
                print(f"[watcher] Erro: {e}")
                time.sleep(5)

    t = threading.Thread(target=watch_loop, daemon=True)
    t.start()
    return {"ok": True, "msg": "Watcher iniciado", "pasta": str(downloads_path)}


@app.post("/api/watcher/parar")
def parar_watcher():
    global _watcher_running
    _watcher_running = False
    return {"ok": True, "msg": "Watcher parado"}


@app.get("/api/watcher/status")
def status_watcher():
    return {"rodando": _watcher_running}


def _importar_arquivo_watcher(filepath: Path):
    """Importa um PDF/HTML da pasta Downloads para o sistema."""
    # Copiar para downloads do app
    dest = DOWNLOADS_DIR / f"watcher_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{filepath.name}"
    shutil.copy2(str(filepath), str(dest))

    # Extrair info
    if filepath.suffix == ".pdf":
        info = extrair_info_pdf(str(dest))
    else:
        # HTML - extrair texto
        import re as _re
        import html as html_mod
        with open(dest, "r", encoding="utf-8", errors="ignore") as f:
            html_content = f.read()
        texto = _re.sub(r'<[^>]+>', ' ', html_content)
        texto = html_mod.unescape(texto)
        num_match = _re.search(r'\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}', texto)
        info = {
            "numero_processo": num_match.group(0) if num_match else "",
            "destinatario": "",
            "tipo": "intimacao",
            "endereco": "",
            "texto_completo": texto,
        }

    numero = info.get("numero_processo", "")
    if not numero:
        return

    destinatario = info.get("destinatario", "")
    from pdf_extractor import _extrair_email_pje
    email_direto = _extrair_email_pje(info.get("texto_completo", "") or info.get("endereco", ""))
    email_result = {}
    if email_direto:
        email_result = {"email": email_direto, "fonte": "pje_endereco"}
    elif destinatario:
        email_result = buscar_email_completo(destinatario)

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT id FROM processos WHERE numero_processo = ?", (numero,))
    existing = cur.fetchone()

    if existing:
        cur.execute("UPDATE processos SET arquivo_mandado = ? WHERE id = ?",
                    (str(dest), existing[0]))
    else:
        assunto = montar_assunto(info)
        cur.execute(
            """INSERT INTO processos
               (numero_processo, tipo, destinatario, endereco_destinatario,
                email_destinatario, email_encontrado_em, assunto_email,
                arquivo_mandado, status)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'pendente')""",
            (numero, info.get("tipo", "intimacao"), destinatario,
             info.get("endereco", ""), email_result.get("email", ""),
             email_result.get("fonte", ""), assunto, str(dest)),
        )

    conn.commit()
    conn.close()
    print(f"[watcher] Importado: {numero} de {filepath.name}")


# ── PJe Download Automático ───────────────────────────────────────────────────

class PjePuxarRequest(BaseModel):
    grau: int = 1

@app.post("/api/pje/puxar-tudo")
def pje_puxar_tudo(data: PjePuxarRequest):
    """
    Conecta ao Chrome via CDP e baixa TODOS os mandados do painel PJe.
    Usa puxar_todos_pdfs.py (script que funciona perfeitamente).
    Chrome precisa estar aberto com --remote-debugging-port=9222.
    """
    import asyncio
    from puxar_todos_pdfs import puxar_todos
    from pdf_extractor import extrair_info_pdf, _extrair_email_pje

    # Rodar a função async em um event loop novo
    loop = asyncio.new_event_loop()
    try:
        download_results = loop.run_until_complete(puxar_todos(grau=data.grau))
    except Exception as e:
        return {"ok": False, "erro": f"Erro ao puxar PDFs: {e}", "resultados": []}
    finally:
        loop.close()

    if not download_results:
        return {"ok": True, "total": 0, "baixados": 0,
                "erro": "Nenhum processo encontrado no painel", "resultados": []}

    total = len(download_results)
    baixados = sum(1 for r in download_results if r.get("ok"))

    # Processar cada PDF baixado: extrair info e salvar no banco
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    resultados_final = []

    for r in download_results:
        numero = r.get("numero", "")
        pdf_path = r.get("path", "")
        ok = r.get("ok", False)

        resultado = {
            "numero_processo": numero,
            "ok": ok,
            "erro": r.get("msg") if not ok else None,
            "acao": None,
            "email": "",
        }

        if not ok or not numero:
            resultados_final.append(resultado)
            continue

        # Extrair info do PDF via pdf_extractor
        info = {}
        if pdf_path and os.path.exists(pdf_path):
            info = extrair_info_pdf(pdf_path)

        destinatario = info.get("destinatario", "")
        endereco = info.get("endereco", "")
        tipo = info.get("tipo", "intimacao") or "intimacao"

        # Buscar email: primeiro do PDF, depois do endereço, depois da agenda
        email = ""
        email_fonte = ""
        pdf_email = info.get("email", "")
        if pdf_email:
            email = pdf_email
            email_fonte = "pdf_extraido"
        else:
            email_direto = _extrair_email_pje(endereco) if endereco else ""
            if email_direto:
                email = email_direto
                email_fonte = "pje_endereco"
            elif destinatario:
                email_result = buscar_email_completo(destinatario)
                email = email_result.get("email", "")
                email_fonte = email_result.get("fonte", "")

        assunto = montar_assunto(info)

        # Verificar se já existe no banco
        cur.execute("SELECT id FROM processos WHERE numero_processo = ?", (numero,))
        existing = cur.fetchone()

        if existing:
            updates = {}
            if pdf_path:
                updates["arquivo_mandado"] = pdf_path
            if destinatario:
                updates["destinatario"] = destinatario
            if endereco:
                updates["endereco_destinatario"] = endereco
            if email:
                updates["email_destinatario"] = email
                updates["email_encontrado_em"] = email_fonte
            if tipo:
                updates["tipo"] = tipo
            if updates:
                sets = ", ".join(f"{k} = ?" for k in updates)
                vals = list(updates.values()) + [existing[0]]
                cur.execute(f"UPDATE processos SET {sets} WHERE id = ?", vals)
            resultado["id"] = existing[0]
            resultado["acao"] = "atualizado"
            resultado["email"] = email
        else:
            cur.execute(
                """INSERT INTO processos
                   (numero_processo, tipo, destinatario, endereco_destinatario,
                    email_destinatario, email_encontrado_em, assunto_email,
                    arquivo_mandado, status, grau)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'pendente', ?)""",
                (numero, tipo, destinatario, endereco, email,
                 email_fonte, assunto, pdf_path or "", data.grau),
            )
            resultado["id"] = cur.lastrowid
            resultado["acao"] = "criado"
            resultado["email"] = email

        resultados_final.append(resultado)

    conn.commit()
    conn.close()

    return {
        "ok": True,
        "total": total,
        "baixados": baixados,
        "resultados": resultados_final,
    }


class PjeRequest(BaseModel):
    processo_ids: list[int]

@app.post("/api/pje/baixar")
def pje_baixar(data: PjeRequest, background_tasks: BackgroundTasks):
    """
    Inicia o download dos mandados do PJe para os processos informados.
    Abre o Chrome do usuário (com sessão ativa) e baixa os PDFs.
    Usa o campo 'grau' de cada processo para saber se é PJe 1G ou 2G.
    """
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    processos_data = []
    for pid in data.processo_ids:
        cur.execute("SELECT id, numero_processo, grau FROM processos WHERE id = ?", (pid,))
        row = cur.fetchone()
        if row:
            processos_data.append({
                "id": row["id"],
                "numero_processo": row["numero_processo"],
                "grau": row["grau"] or 1,
            })
    conn.close()

    if not processos_data:
        return {"ok": False, "erro": "Nenhum processo encontrado"}

    try:
        from pje_downloader import baixar_documentos_pje_sync
    except ImportError as e:
        return {"ok": False, "erro": str(e)}

    resultados = baixar_documentos_pje_sync(processos_data)

    # Salvar caminhos dos arquivos baixados no banco
    conn2 = sqlite3.connect(DB_PATH)
    for r in resultados:
        if r.get("ok"):
            pid = r.get("id")
            if pid:
                updates = {}
                if r.get("mandado"):
                    updates["arquivo_mandado"] = r["mandado"]
                if r.get("anexo"):
                    updates["arquivo_anexo"] = r["anexo"]
                if updates:
                    sets = ", ".join(f"{k} = ?" for k in updates)
                    vals = list(updates.values()) + [pid]
                    conn2.execute(f"UPDATE processos SET {sets} WHERE id = ?", vals)
    conn2.commit()
    conn2.close()

    sucessos = sum(1 for r in resultados if r.get("ok"))
    return {
        "ok": True,
        "total": len(resultados),
        "sucessos": sucessos,
        "resultados": resultados,
    }


# ── Escritório Virtual de IA ──────────────────────────────────────────────────

class NovoAgente(BaseModel):
    nome: str
    cargo: str
    avatar: Optional[str] = "🤖"
    cor: Optional[str] = "#7c3aed"
    system_prompt: Optional[str] = None

class EscritorioMsg(BaseModel):
    agente_id: Optional[int] = None
    conversa_id: Optional[int] = None
    mensagem: str
    tipo: Optional[str] = "individual"  # individual | reuniao
    participantes: Optional[list[int]] = None  # para reunião

class NovaTarefa(BaseModel):
    agente_id: int
    titulo: str
    descricao: Optional[str] = None
    prioridade: Optional[str] = "normal"

class NovoDoc(BaseModel):
    agente_id: int
    titulo: str
    conteudo: Optional[str] = None


@app.get("/api/escritorio/agentes")
def listar_agentes():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    rows = [dict(r) for r in conn.execute("SELECT * FROM escritorio_agentes ORDER BY posicao, id")]
    conn.close()
    return {"agentes": rows}


@app.post("/api/escritorio/agentes")
def criar_agente(body: NovoAgente):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        "INSERT INTO escritorio_agentes (nome, cargo, avatar, cor, system_prompt) VALUES (?,?,?,?,?)",
        (body.nome, body.cargo, body.avatar, body.cor, body.system_prompt)
    )
    agente_id = cur.lastrowid
    conn.commit()
    conn.close()
    return {"ok": True, "id": agente_id}


@app.put("/api/escritorio/agentes/{agente_id}")
def atualizar_agente(agente_id: int, body: NovoAgente):
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        "UPDATE escritorio_agentes SET nome=?,cargo=?,avatar=?,cor=?,system_prompt=? WHERE id=?",
        (body.nome, body.cargo, body.avatar, body.cor, body.system_prompt, agente_id)
    )
    conn.commit()
    conn.close()
    return {"ok": True}


@app.delete("/api/escritorio/agentes/{agente_id}")
def deletar_agente(agente_id: int):
    conn = sqlite3.connect(DB_PATH)
    conn.execute("DELETE FROM escritorio_agentes WHERE id=?", (agente_id,))
    conn.commit()
    conn.close()
    return {"ok": True}


@app.get("/api/escritorio/conversas/{agente_id}")
def listar_conversas(agente_id: int):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    rows = [dict(r) for r in conn.execute(
        "SELECT * FROM escritorio_conversas WHERE agente_id=? ORDER BY criado_em DESC LIMIT 10",
        (agente_id,)
    )]
    conn.close()
    return {"conversas": rows}


@app.get("/api/escritorio/mensagens/{conversa_id}")
def listar_mensagens(conversa_id: int):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    rows = [dict(r) for r in conn.execute(
        "SELECT * FROM escritorio_mensagens WHERE conversa_id=? ORDER BY id",
        (conversa_id,)
    )]
    conn.close()
    return {"mensagens": rows}


@app.post("/api/escritorio/chat")
async def chat_agente(body: EscritorioMsg):
    """Chat com um agente ou sala de reunião com múltiplos agentes."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row

    if body.tipo == "reuniao" and body.participantes:
        agentes = [dict(r) for r in conn.execute(
            f"SELECT * FROM escritorio_agentes WHERE id IN ({','.join('?'*len(body.participantes))})",
            body.participantes
        )]
    elif body.agente_id:
        row = conn.execute("SELECT * FROM escritorio_agentes WHERE id=?", (body.agente_id,)).fetchone()
        agentes = [dict(row)] if row else []
    else:
        conn.close()
        raise HTTPException(400, "Informe agente_id ou participantes")

    if not agentes:
        conn.close()
        raise HTTPException(404, "Agente(s) não encontrados")

    # Criar/reutilizar conversa
    conversa_id = body.conversa_id
    if not conversa_id:
        participantes_str = ",".join(str(a["id"]) for a in agentes)
        cur = conn.execute(
            "INSERT INTO escritorio_conversas (agente_id, tipo, participantes) VALUES (?,?,?)",
            (body.agente_id, body.tipo, participantes_str)
        )
        conversa_id = cur.lastrowid

    # Salvar mensagem do usuário
    conn.execute(
        "INSERT INTO escritorio_mensagens (conversa_id, remetente, conteudo) VALUES (?,?,?)",
        (conversa_id, "você", body.mensagem)
    )
    conn.commit()

    # Buscar histórico para contexto
    historico = [dict(r) for r in conn.execute(
        "SELECT * FROM escritorio_mensagens WHERE conversa_id=? ORDER BY id DESC LIMIT 20",
        (conversa_id,)
    )]
    historico.reverse()

    # Chamar LLM para cada agente
    try:
        import openai as _openai
        client = _openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY", ""))
    except Exception:
        conn.close()
        return {"ok": False, "erro": "OpenAI não configurado. Defina OPENAI_API_KEY no .env"}

    respostas = []
    for agente in agentes:
        sys_prompt = agente.get("system_prompt") or f"Você é {agente['nome']}, {agente['cargo']}."
        if body.tipo == "reuniao":
            outros = [a["nome"] for a in agentes if a["id"] != agente["id"]]
            if outros:
                sys_prompt += f" Você está em reunião com: {', '.join(outros)}."

        msgs_llm = [{"role": "system", "content": sys_prompt}]
        for m in historico[:-1]:
            role = "user" if m["remetente"] == "você" else "assistant"
            msgs_llm.append({"role": role, "content": m["conteudo"]})
        msgs_llm.append({"role": "user", "content": body.mensagem})

        try:
            resp = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=msgs_llm,
                max_tokens=800,
            )
            texto = resp.choices[0].message.content or ""
        except Exception as e:
            texto = f"[Erro: {e}]"

        # Salvar resposta
        conn.execute(
            "INSERT INTO escritorio_mensagens (conversa_id, remetente, conteudo) VALUES (?,?,?)",
            (conversa_id, agente["nome"], texto)
        )
        respostas.append({"agente": agente["nome"], "avatar": agente["avatar"], "resposta": texto})

    conn.commit()
    conn.close()
    return {"ok": True, "conversa_id": conversa_id, "respostas": respostas}


@app.get("/api/escritorio/tarefas/{agente_id}")
def listar_tarefas(agente_id: int):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    rows = [dict(r) for r in conn.execute(
        "SELECT * FROM escritorio_tarefas WHERE agente_id=? ORDER BY id DESC",
        (agente_id,)
    )]
    conn.close()
    return {"tarefas": rows}


@app.post("/api/escritorio/tarefas")
def criar_tarefa(body: NovaTarefa):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        "INSERT INTO escritorio_tarefas (agente_id, titulo, descricao, prioridade) VALUES (?,?,?,?)",
        (body.agente_id, body.titulo, body.descricao, body.prioridade)
    )
    conn.commit()
    tid = cur.lastrowid
    conn.close()
    return {"ok": True, "id": tid}


@app.patch("/api/escritorio/tarefas/{tarefa_id}")
def atualizar_tarefa(tarefa_id: int, status: str):
    conn = sqlite3.connect(DB_PATH)
    ts = datetime.now().isoformat() if status == "concluida" else None
    conn.execute(
        "UPDATE escritorio_tarefas SET status=?, concluido_em=? WHERE id=?",
        (status, ts, tarefa_id)
    )
    conn.commit()
    conn.close()
    return {"ok": True}


@app.get("/api/escritorio/docs/{agente_id}")
def listar_docs(agente_id: int):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    rows = [dict(r) for r in conn.execute(
        "SELECT * FROM escritorio_docs WHERE agente_id=? ORDER BY id DESC",
        (agente_id,)
    )]
    conn.close()
    return {"docs": rows}


@app.post("/api/escritorio/docs")
def criar_doc(body: NovoDoc):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        "INSERT INTO escritorio_docs (agente_id, titulo, conteudo) VALUES (?,?,?)",
        (body.agente_id, body.titulo, body.conteudo)
    )
    conn.commit()
    did = cur.lastrowid
    conn.close()
    return {"ok": True, "id": did}


# ── Frontend (SPA) ────────────────────────────────────────────────────────────

frontend_dir = Path(__file__).parent / "frontend"
if frontend_dir.exists():
    app.mount("/static", StaticFiles(directory=str(frontend_dir)), name="static")

    @app.get("/")
    def serve_frontend():
        return FileResponse(str(frontend_dir / "index.html"))


# ── Entrada ───────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import uvicorn
    print("=" * 50)
    print("  Agente E-mail TRF1 — http://localhost:8090")
    print("=" * 50)
    uvicorn.run("app:app", host="0.0.0.0", port=8090, reload=True)

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
    return f"""{data_str}

Prezado(a) Senhor(a),

Encaminhamos em anexo o mandado judicial referente ao processo {proc}, expedido pela Seção Judiciária do Piauí - SJPI.

Atenciosamente,

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


# ── PJe Download ──────────────────────────────────────────────────────────────

class PjeRequest(BaseModel):
    processo_ids: list[int]

@app.post("/api/pje/baixar")
def pje_baixar(data: PjeRequest, background_tasks: BackgroundTasks):
    """
    Inicia o download dos mandados do PJe para os processos informados.
    Abre o Chrome do usuário (com sessão ativa) e baixa os PDFs.
    """
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    numeros = []
    processo_map = {}  # numero → id
    for pid in data.processo_ids:
        cur.execute("SELECT id, numero_processo FROM processos WHERE id = ?", (pid,))
        row = cur.fetchone()
        if row:
            numeros.append(row["numero_processo"])
            processo_map[row["numero_processo"]] = row["id"]
    conn.close()

    if not numeros:
        return {"ok": False, "erro": "Nenhum processo encontrado"}

    try:
        from pje_downloader import baixar_documentos_pje_sync
    except ImportError as e:
        return {"ok": False, "erro": str(e)}

    resultados = baixar_documentos_pje_sync(numeros)

    # Salvar caminhos dos arquivos baixados no banco
    conn2 = sqlite3.connect(DB_PATH)
    for r in resultados:
        if r["ok"] and r.get("arquivo"):
            pid = processo_map.get(r["numero"])
            if pid:
                conn2.execute(
                    "UPDATE processos SET arquivo_mandado = ? WHERE id = ?",
                    (r["arquivo"], pid)
                )
    conn2.commit()
    conn2.close()

    sucessos = sum(1 for r in resultados if r["ok"])
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

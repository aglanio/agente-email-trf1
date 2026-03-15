"""
outlook_agent.py — Automação do Outlook via win32com para criar rascunhos de e-mail.
Abre até 3 rascunhos por vez para não sobrecarregar o PC.
"""
import os
import re
import time
from pathlib import Path
from typing import Optional
import sqlite3
from database import DB_PATH

# Máximo de rascunhos abertos de uma vez
MAX_RASCUNHOS = 3

REMETENTE = "aglanio.carvalho@trf1.jus.br"


def criar_rascunho(
    destinatario_email: str,
    assunto: str,
    corpo: str,
    arquivo_mandado: Optional[str] = None,
    arquivo_anexo: Optional[str] = None,
    numero_processo: Optional[str] = None,
) -> dict:
    """
    Cria um rascunho no Outlook com os dados do processo.
    Não envia — deixa aberto para revisão do usuário.
    Retorna {"ok": True, "msg": "..."} ou {"ok": False, "erro": "..."}
    """
    try:
        import win32com.client
    except ImportError:
        return {"ok": False, "erro": "pywin32 não instalado. Execute: pip install pywin32"}

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # olMailItem

        # Configurar remetente (conta TRF1)
        try:
            contas = outlook.Session.Accounts
            for conta in contas:
                if REMETENTE.lower() in conta.SmtpAddress.lower():
                    mail._oleobj_.Invoke(
                        0xF045,  # PR_SENT_REPRESENTING_EMAIL_ADDRESS
                        0, 8, 1, conta
                    )
                    break
        except Exception:
            pass  # Usa conta padrão se não encontrar

        # Destinatário, assunto e corpo
        mail.To = destinatario_email
        mail.Subject = assunto or f"Processo {numero_processo or ''}"
        mail.Body = corpo or _corpo_padrao(numero_processo)

        # Anexar mandado
        if arquivo_mandado and Path(arquivo_mandado).exists():
            mail.Attachments.Add(str(Path(arquivo_mandado).resolve()))

        # Anexar anexo do processo
        if arquivo_anexo and Path(arquivo_anexo).exists():
            mail.Attachments.Add(str(Path(arquivo_anexo).resolve()))

        # Salvar como rascunho (não envia)
        mail.Save()

        # Abrir janela para o usuário revisar
        mail.Display(False)  # False = não modal, não bloqueia

        return {
            "ok": True,
            "msg": f"Rascunho criado para {destinatario_email}",
            "assunto": assunto,
        }

    except Exception as e:
        return {"ok": False, "erro": str(e)}


def criar_rascunhos_em_lote(processos: list[dict], delay_segundos: float = 1.5) -> list[dict]:
    """
    Cria rascunhos em lote, MAX_RASCUNHOS por vez.
    processos: lista de dicts com campos do processo + email_destinatario resolvido.
    """
    resultados = []
    total = len(processos)

    for i, proc in enumerate(processos):
        resultado = {
            "numero_processo": proc.get("numero_processo", ""),
            "email": proc.get("email_destinatario", ""),
            "status": "erro",
            "msg": "",
        }

        res = criar_rascunho(
            destinatario_email=proc.get("email_destinatario") or "",
            assunto=proc.get("assunto_email", ""),
            corpo=proc.get("corpo_email", ""),
            arquivo_mandado=proc.get("arquivo_mandado"),
            arquivo_anexo=proc.get("arquivo_anexo"),
            numero_processo=proc.get("numero_processo"),
        )

        if res["ok"]:
            resultado["status"] = "rascunho_aberto"
            resultado["msg"] = res["msg"]
            # Atualizar status no banco
            _atualizar_status_processo(proc.get("id"), "rascunho")
        else:
            resultado["status"] = "erro"
            resultado["msg"] = res.get("erro", "Erro desconhecido")

        resultados.append(resultado)

        # Pausa entre rascunhos para não travar o Outlook
        if (i + 1) % MAX_RASCUNHOS == 0 and (i + 1) < total:
            time.sleep(delay_segundos * 2)  # pausa maior entre lotes
        elif i + 1 < total:
            time.sleep(delay_segundos)

    return resultados


def verificar_outlook_disponivel() -> dict:
    """Verifica se o Outlook está instalado e acessível."""
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        session = outlook.Session
        contas = session.Accounts
        lista_contas = []
        for conta in contas:
            try:
                lista_contas.append({
                    "email": conta.SmtpAddress,
                    "nome": conta.DisplayName,
                })
            except Exception:
                pass
        return {
            "ok": True,
            "contas": lista_contas,
            "conta_trf1": any(REMETENTE.lower() in c["email"].lower() for c in lista_contas),
        }
    except ImportError:
        return {"ok": False, "erro": "pywin32 não instalado"}
    except Exception as e:
        return {"ok": False, "erro": str(e)}


def buscar_outlook_contatos(nome_orgao: str, limit: int = 5) -> list[dict]:
    """
    Busca no catálogo de endereços do Outlook.
    Complementar à busca no histórico de enviados/recebidos.
    """
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        # Catálogo de endereços global
        try:
            gal = namespace.AddressLists.Item("Catálogo de Endereços Global")
            entries = gal.AddressEntries
            resultados = []
            for entry in entries:
                try:
                    if nome_orgao.lower() in (entry.Name or "").lower():
                        email = entry.Address
                        if email:
                            resultados.append({
                                "email": email,
                                "nome": entry.Name,
                                "fonte": "gal",
                            })
                            if len(resultados) >= limit:
                                break
                except Exception:
                    continue
            return resultados
        except Exception:
            return []
    except Exception:
        return []


def exportar_contatos_outlook(destino_db: bool = True) -> int:
    """
    Exporta contatos do Outlook para a base SQLite.
    Retorna número de contatos importados.
    """
    try:
        import win32com.client
    except ImportError:
        return 0

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        contacts_folder = namespace.GetDefaultFolder(10)  # olFolderContacts
        items = contacts_folder.Items

        contatos = []
        for item in items:
            try:
                email = getattr(item, "Email1Address", "") or ""
                nome = getattr(item, "FullName", "") or getattr(item, "CompanyName", "") or ""
                orgao = getattr(item, "CompanyName", "") or ""
                if email and nome:
                    contatos.append({
                        "nome": nome,
                        "orgao": orgao,
                        "email": email,
                        "fonte": "outlook_contatos",
                    })
            except Exception:
                continue

        if destino_db and contatos:
            conn = sqlite3.connect(DB_PATH)
            cur = conn.cursor()
            inseridos = 0
            for c in contatos:
                cur.execute(
                    "SELECT id FROM contatos WHERE email = ?", (c["email"],)
                )
                if not cur.fetchone():
                    cur.execute(
                        """INSERT INTO contatos (nome, orgao, email, fonte)
                           VALUES (?, ?, ?, ?)""",
                        (c["nome"], c["orgao"], c["email"], c["fonte"]),
                    )
                    inseridos += 1
            conn.commit()
            conn.close()
            return inseridos

        return len(contatos)

    except Exception as e:
        print(f"[exportar_outlook] erro: {e}")
        return 0


# ── helpers internos ──────────────────────────────────────────────────────────

def _corpo_padrao(numero_processo: Optional[str] = None) -> str:
    from datetime import date
    import locale
    proc = numero_processo or "[número do processo]"
    meses = ["janeiro","fevereiro","março","abril","maio","junho",
             "julho","agosto","setembro","outubro","novembro","dezembro"]
    hoje = date.today()
    data_str = f"Teresina, {hoje.day} de {meses[hoje.month-1]} de {hoje.year}"
    return f"""{data_str}

Prezado(a) Senhor(a),

Encaminho em anexo o mandado/intimação referente ao processo {proc}, para as providências cabíveis.

Atenciosamente,

Aglanio Frota Moura Carvalho
Oficial de Justiça Avaliador Federal     PI100327
Seção Judiciária do Piauí - TRF 1ª Região
aglanio.carvalho@trf1.jus.br
"""


def _atualizar_status_processo(processo_id: Optional[int], novo_status: str):
    if not processo_id:
        return
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.execute(
            "UPDATE processos SET status = ? WHERE id = ?",
            (novo_status, processo_id),
        )
        conn.commit()
        conn.close()
    except Exception:
        pass


def exportar_enviados_com_tratamento(limite: int = 5000) -> dict:
    """
    Varre TODOS os e-mails enviados do Outlook e importa destinatários únicos
    com suas formas de tratamento (Ilmo. Sr., Exmo. Sr., Excelência, etc.)
    para a tabela contatos.
    Retorna {"ok": True, "inseridos": N, "atualizados": N, "total_varredura": N, "unicos": N}
    """
    try:
        import win32com.client
    except ImportError:
        return {"ok": False, "erro": "pywin32 não instalado"}

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        sent_folder = namespace.GetDefaultFolder(5)  # olFolderSentMail
        items = sent_folder.Items
        items.Sort("[SentOn]", True)  # mais recentes primeiro

        # email → {nome, tratamento, orgao}  (primeiro encontrado = mais recente)
        encontrados: dict[str, dict] = {}
        total_varredura = 0

        for item in items:
            try:
                total_varredura += 1
                if total_varredura > limite:
                    break

                corpo = getattr(item, "Body", "") or ""
                tratamento = _extrair_tratamento(corpo)

                recipients = item.Recipients
                for r in recipients:
                    email = _extrair_email_recipient_outlook(r)
                    if not email or "@" not in email:
                        continue
                    if "trf1.jus.br" in email.lower():
                        continue
                    if email not in encontrados:
                        nome = getattr(r, "Name", "") or ""
                        encontrados[email] = {
                            "email": email,
                            "nome": nome,
                            "tratamento": tratamento,
                        }
                    elif tratamento and not encontrados[email].get("tratamento"):
                        # se encontrou tratamento em e-mail mais antigo e ainda não tem, atualiza
                        encontrados[email]["tratamento"] = tratamento
            except Exception:
                continue

        # Salvar no banco
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        inseridos = 0
        atualizados = 0

        for email, dados in encontrados.items():
            cur.execute("SELECT id, tratamento FROM contatos WHERE email = ?", (email,))
            row = cur.fetchone()
            if row:
                # Atualiza tratamento se estava vazio e agora temos
                if not row["tratamento"] and dados["tratamento"]:
                    cur.execute(
                        "UPDATE contatos SET tratamento = ?, atualizado_em = datetime('now','localtime') WHERE id = ?",
                        (dados["tratamento"], row["id"]),
                    )
                    atualizados += 1
            else:
                cur.execute(
                    """INSERT INTO contatos (nome, orgao, email, tratamento, fonte)
                       VALUES (?, ?, ?, ?, 'outlook_enviados')""",
                    (dados["nome"], "", email, dados["tratamento"] or ""),
                )
                inseridos += 1

        conn.commit()
        conn.close()

        return {
            "ok": True,
            "total_varredura": total_varredura,
            "unicos": len(encontrados),
            "inseridos": inseridos,
            "atualizados": atualizados,
        }

    except Exception as e:
        return {"ok": False, "erro": str(e)}


def _extrair_tratamento(corpo: str) -> str:
    """
    Extrai a forma de tratamento usada na saudação do e-mail.
    Verifica as primeiras 8 linhas do corpo.
    """
    if not corpo:
        return ""

    # Primeiras 8 linhas não vazias
    linhas = [l.strip() for l in corpo.split("\n") if l.strip()][:8]
    texto = " ".join(linhas)

    # Padrões em ordem de especificidade (mais específico primeiro)
    padroes = [
        (r"Exmo\.?\s*Sr\.?\s*Des(?:embargador)?", "Exmo. Sr. Desembargador"),
        (r"Exmo\.?\s*Sr\.?\s*Juiz", "Exmo. Sr. Juiz"),
        (r"Exmo\.?\s*Sr\.?\s*Procurador", "Exmo. Sr. Procurador"),
        (r"Exmo\.?\s*Sr\.?\s*Delegado", "Exmo. Sr. Delegado"),
        (r"Excelent[ií]ssim[ao]\s+Senhor\b", "Excelentíssimo Senhor"),
        (r"Excelent[ií]ssim[ao]\s+Senhora\b", "Excelentíssima Senhora"),
        (r"Exmo\.?\s*Sr\.?", "Exmo. Sr."),
        (r"Exma\.?\s*Sra\.?", "Exma. Sra."),
        (r"Ilmo\.?\s*Sr\.?", "Ilmo. Sr."),
        (r"Ilma\.?\s*Sra\.?", "Ilma. Sra."),
        (r"Meritíssimo\b", "Meritíssimo"),
        (r"Mmo\.?\s*(?:Sr\.?\s*)?Juiz", "Mmo. Juiz"),
        (r"V\.?\s*Ex[aª]\.?\b", "V. Exa."),
        (r"Vossa\s+Excel[êe]ncia\b", "Vossa Excelência"),
        (r"Prezad[oa]\(a\)\s+Senhor\(a\)", "Prezado(a) Senhor(a)"),
        (r"Prezad[oa]\s+Senhor\b", "Prezado Senhor"),
        (r"Prezada\s+Senhora\b", "Prezada Senhora"),
        (r"Prezad[oa]\b", "Prezado"),
        (r"Dra?\.\s+\w", "Dr."),
        (r"Senhor\b", "Senhor"),
    ]

    for pattern, label in padroes:
        if re.search(pattern, texto, re.IGNORECASE):
            return label

    return ""


def _extrair_email_recipient_outlook(recipient) -> Optional[str]:
    """Extrai email SMTP de um objeto Recipient do Outlook."""
    try:
        addr = recipient.AddressEntry
        if addr.AddressEntryUserType == 0:  # Exchange user
            ex_user = addr.GetExchangeUser()
            if ex_user:
                return ex_user.PrimarySmtpAddress
        return addr.Address
    except Exception:
        try:
            return recipient.Address
        except Exception:
            return None

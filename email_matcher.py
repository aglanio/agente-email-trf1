"""
email_matcher.py — Busca email do destinatário na base de contatos ou Outlook.
"""
import sqlite3
import re
from pathlib import Path
from typing import Optional
from database import DB_PATH


def buscar_email_por_nome(nome_orgao: str) -> Optional[dict]:
    """
    Busca email na base de contatos pelo nome ou órgão.
    Retorna dict com {email, nome, orgao, fonte} ou None.
    """
    if not nome_orgao or len(nome_orgao.strip()) < 3:
        return None

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    # Normaliza para busca
    termos = _normalizar_termos(nome_orgao)

    resultados = []
    for termo in termos:
        if len(termo) < 3:
            continue
        cur.execute(
            """
            SELECT nome, orgao, email, email_alternativo, categoria, fonte
            FROM contatos
            WHERE (nome LIKE ? OR orgao LIKE ?) AND email != ''
            ORDER BY
                CASE WHEN nome LIKE ? THEN 0 ELSE 1 END,
                LENGTH(nome) ASC
            LIMIT 10
            """,
            (f"%{termo}%", f"%{termo}%", f"%{termo}%"),
        )
        rows = cur.fetchall()
        for row in rows:
            resultados.append(dict(row))

    conn.close()

    if not resultados:
        return None

    # Prioriza match mais exato
    melhor = _melhor_match(nome_orgao, resultados)
    if melhor:
        return {
            "email": melhor["email"],
            "email_alternativo": melhor.get("email_alternativo", ""),
            "nome": melhor["nome"],
            "orgao": melhor["orgao"],
            "categoria": melhor["categoria"],
            "fonte": melhor["fonte"] or "agenda",
        }
    return None


def buscar_email_outlook(nome_orgao: str) -> Optional[dict]:
    """
    Busca email nos itens enviados/recebidos do Outlook.
    Retorna dict {email, nome, fonte:'outlook'} ou None.
    """
    try:
        import win32com.client
    except ImportError:
        return None

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        termos = _normalizar_termos(nome_orgao)
        encontrados = {}

        # Busca na caixa de saída
        try:
            sent_folder = namespace.GetDefaultFolder(5)  # olFolderSentMail
            items = sent_folder.Items
            items.Sort("[SentOn]", True)
            count = 0
            for item in items:
                if count > 500:
                    break
                try:
                    count += 1
                    for termo in termos:
                        if len(termo) < 4:
                            continue
                        to_name = getattr(item, "To", "") or ""
                        subject = getattr(item, "Subject", "") or ""
                        if termo.lower() in to_name.lower() or termo.lower() in subject.lower():
                            # Extrai email dos destinatários
                            recipients = item.Recipients
                            for r in recipients:
                                email = _extrair_email_recipient(r)
                                if email and "trf1.jus.br" not in email:
                                    if email not in encontrados:
                                        encontrados[email] = {
                                            "email": email,
                                            "nome": to_name,
                                            "orgao": nome_orgao,
                                            "fonte": "outlook_enviados",
                                            "score": 0,
                                        }
                                    encontrados[email]["score"] += 1
                except Exception:
                    continue
        except Exception:
            pass

        # Busca na caixa de entrada
        try:
            inbox = namespace.GetDefaultFolder(6)  # olFolderInbox
            items = inbox.Items
            items.Sort("[ReceivedTime]", True)
            count = 0
            for item in items:
                if count > 300:
                    break
                try:
                    count += 1
                    for termo in termos:
                        if len(termo) < 4:
                            continue
                        sender_name = getattr(item, "SenderName", "") or ""
                        sender_email = getattr(item, "SenderEmailAddress", "") or ""
                        if termo.lower() in sender_name.lower():
                            if sender_email and "trf1.jus.br" not in sender_email:
                                if sender_email not in encontrados:
                                    encontrados[sender_email] = {
                                        "email": sender_email,
                                        "nome": sender_name,
                                        "orgao": nome_orgao,
                                        "fonte": "outlook_recebidos",
                                        "score": 0,
                                    }
                                encontrados[sender_email]["score"] += 1
                except Exception:
                    continue
        except Exception:
            pass

        if encontrados:
            # Retorna o que aparece mais vezes
            melhor = max(encontrados.values(), key=lambda x: x["score"])
            return melhor

    except Exception as e:
        print(f"[outlook] erro: {e}")

    return None


def buscar_email_completo(nome_orgao: str) -> dict:
    """
    Busca email em todas as fontes disponíveis.
    Retorna dict com resultado e metadados de onde foi encontrado.
    """
    resultado = {
        "email": "",
        "email_alternativo": "",
        "nome": "",
        "orgao": nome_orgao,
        "fonte": "nao_encontrado",
        "confianca": "baixa",
        "sugestoes": [],
    }

    # 1. Processos anteriores com email já usado (maior prioridade — uso real)
    match_proc = buscar_email_em_processos(nome_orgao)
    if match_proc:
        resultado.update(match_proc)
        resultado["confianca"] = "alta"
        # ainda busca sugestões para o usuário poder trocar
        resultado["sugestoes"] = buscar_sugestoes(nome_orgao)
        return resultado

    # 2. Base de contatos (agenda DOCX)
    match_agenda = buscar_email_por_nome(nome_orgao)
    if match_agenda:
        resultado.update(match_agenda)
        resultado["confianca"] = "alta"
        resultado["sugestoes"] = buscar_sugestoes(nome_orgao)
        return resultado

    # 3. Outlook histórico
    match_outlook = buscar_email_outlook(nome_orgao)
    if match_outlook:
        resultado.update(match_outlook)
        resultado["confianca"] = "media"
        resultado["sugestoes"] = buscar_sugestoes(nome_orgao)
        return resultado

    # 4. Sugestões parciais
    sugestoes = buscar_sugestoes(nome_orgao)
    resultado["sugestoes"] = sugestoes
    if sugestoes:
        resultado["confianca"] = "sugestao"

    return resultado


def buscar_email_em_processos(nome_orgao: str) -> Optional[dict]:
    """
    Busca email em processos anteriores com destinatário/órgão similar.
    Retorna o email mais recente encontrado.
    """
    if not nome_orgao or len(nome_orgao.strip()) < 3:
        return None

    termos = _normalizar_termos(nome_orgao)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    candidatos: dict[str, int] = {}  # email -> contagem

    for termo in termos:
        if len(termo) < 3:
            continue
        cur.execute(
            """
            SELECT email_destinatario, destinatario
            FROM processos
            WHERE email_destinatario != '' AND email_destinatario IS NOT NULL
              AND (destinatario LIKE ? OR email_destinatario LIKE ?)
            ORDER BY id DESC
            LIMIT 20
            """,
            (f"%{termo}%", f"%{termo}%"),
        )
        for row in cur.fetchall():
            email = row["email_destinatario"] or ""
            if email and "@" in email:
                candidatos[email] = candidatos.get(email, 0) + 1

    conn.close()

    if not candidatos:
        return None

    # Pega o email mais frequente
    melhor_email = max(candidatos, key=lambda e: candidatos[e])
    return {
        "email": melhor_email,
        "nome": nome_orgao,
        "orgao": nome_orgao,
        "fonte": "processos_anteriores",
    }


def buscar_sugestoes(nome_orgao: str, limit: int = 5) -> list:
    """Retorna lista de sugestões parciais da base."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    termos = nome_orgao.split()[:3]  # até 3 palavras
    sugestoes = []
    seen = set()

    for termo in termos:
        if len(termo) < 4:
            continue
        cur.execute(
            """
            SELECT nome, orgao, email, categoria
            FROM contatos
            WHERE (nome LIKE ? OR orgao LIKE ?) AND email != ''
            LIMIT ?
            """,
            (f"%{termo}%", f"%{termo}%", limit),
        )
        for row in cur.fetchall():
            key = row["email"]
            if key not in seen:
                seen.add(key)
                sugestoes.append(dict(row))

    conn.close()
    return sugestoes[:limit]


def salvar_email_manual(nome_orgao: str, email: str, fonte: str = "manual") -> bool:
    """Salva email manualmente na base para aprendizado futuro."""
    conn = sqlite3.connect(DB_PATH)
    try:
        cur = conn.cursor()
        # Verifica se já existe
        cur.execute("SELECT id FROM contatos WHERE orgao = ? OR nome = ?", (nome_orgao, nome_orgao))
        row = cur.fetchone()
        if row:
            cur.execute(
                "UPDATE contatos SET email = ?, fonte = ? WHERE id = ?",
                (email, fonte, row[0]),
            )
        else:
            cur.execute(
                "INSERT INTO contatos (nome, orgao, email, fonte) VALUES (?, ?, ?, ?)",
                (nome_orgao, nome_orgao, email, fonte),
            )
        conn.commit()
        return True
    except Exception as e:
        print(f"[salvar_email] erro: {e}")
        return False
    finally:
        conn.close()


# ── helpers internos ──────────────────────────────────────────────────────────

def _normalizar_termos(texto: str) -> list[str]:
    """Extrai termos relevantes de um nome/órgão para busca."""
    # Remove stopwords comuns
    stopwords = {
        "de", "da", "do", "das", "dos", "e", "em", "a", "o", "as", "os",
        "por", "para", "com", "no", "na", "nos", "nas", "ao", "aos",
        "se", "um", "uma", "uns", "umas",
    }
    palavras = re.findall(r'\b\w{3,}\b', texto.lower())
    termos = [p for p in palavras if p not in stopwords]

    # Adiciona o nome completo e partes significativas
    resultado = [texto]  # busca com nome completo
    resultado.extend(termos)

    # Siglas (maiúsculas no original)
    siglas = re.findall(r'\b[A-Z]{2,}\b', texto)
    resultado.extend([s.lower() for s in siglas])

    return list(dict.fromkeys(resultado))  # deduplica mantendo ordem


def _melhor_match(query: str, resultados: list) -> Optional[dict]:
    """Retorna o resultado com melhor score de similaridade."""
    if not resultados:
        return None

    query_low = query.lower()
    palavras_query = set(re.findall(r'\b\w{3,}\b', query_low))

    def score(r):
        nome_low = (r.get("nome") or "").lower()
        orgao_low = (r.get("orgao") or "").lower()
        combined = nome_low + " " + orgao_low

        palavras_r = set(re.findall(r'\b\w{3,}\b', combined))
        intersecao = palavras_query & palavras_r
        s = len(intersecao) * 10

        # Bonus por match exato de substring
        if query_low in combined:
            s += 50
        elif any(p in combined for p in query_low.split()[:2]):
            s += 20

        return s

    melhor = max(resultados, key=score)
    if score(melhor) >= 10:
        return melhor
    return None


def _extrair_email_recipient(recipient) -> Optional[str]:
    """Extrai email de um objeto Recipient do Outlook."""
    try:
        addr = recipient.AddressEntry
        if addr.AddressEntryUserType == 0:  # olExchangeUserAddressEntry
            ex_user = addr.GetExchangeUser()
            if ex_user:
                return ex_user.PrimarySmtpAddress
        return addr.Address
    except Exception:
        try:
            return recipient.Address
        except Exception:
            return None

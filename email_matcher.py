"""
email_matcher.py — Busca email do destinatário na base de contatos ou Outlook.
v2 — Matching melhorado com scoring mais rigoroso.
"""
import sqlite3
import re
from pathlib import Path
from typing import Optional
from database import DB_PATH


# Emails que NUNCA devem ser retornados (internos, genéricos, etc.)
EMAILS_BLACKLIST = {
    "pje1g@trf1.jus.br",
    "pje2g@trf1.jus.br",
    "pje@trf1.jus.br",
    "aglanio@hotmail.com",  # email pessoal do usuário
}

# Stopwords para normalização
STOPWORDS = {
    "de", "da", "do", "das", "dos", "e", "em", "a", "o", "as", "os",
    "por", "para", "com", "no", "na", "nos", "nas", "ao", "aos",
    "se", "um", "uma", "uns", "umas", "que", "ou",
}


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

    if not nome_orgao or len(nome_orgao.strip()) < 3:
        return resultado

    # 1. Busca por domínio institucional (mais confiável — RECEITA→rfb.gov.br)
    match_inst = buscar_email_por_instituicao(nome_orgao)
    if match_inst and _email_valido(match_inst.get("email", "")):
        resultado.update(match_inst)
        resultado["confianca"] = "alta"
        resultado["sugestoes"] = buscar_sugestoes(nome_orgao)
        return resultado

    # 2. Base de contatos — match exato ou quase exato do nome
    match_agenda = buscar_email_por_nome(nome_orgao)
    if match_agenda and _email_valido(match_agenda.get("email", "")):
        resultado.update(match_agenda)
        resultado["confianca"] = "alta"
        resultado["sugestoes"] = buscar_sugestoes(nome_orgao)
        return resultado

    # 3. Processos anteriores — só se match forte (>= 70% palavras)
    match_proc = buscar_email_em_processos(nome_orgao)
    if match_proc and _email_valido(match_proc.get("email", "")):
        resultado.update(match_proc)
        resultado["confianca"] = "alta"
        resultado["sugestoes"] = buscar_sugestoes(nome_orgao)
        return resultado

    # 4. Outlook histórico
    match_outlook = buscar_email_outlook(nome_orgao)
    if match_outlook and _email_valido(match_outlook.get("email", "")):
        resultado.update(match_outlook)
        resultado["confianca"] = "media"
        resultado["sugestoes"] = buscar_sugestoes(nome_orgao)
        return resultado

    # 5. Sugestões parciais
    sugestoes = buscar_sugestoes(nome_orgao)
    resultado["sugestoes"] = sugestoes
    if sugestoes:
        resultado["confianca"] = "sugestao"

    return resultado


def _email_valido(email: str) -> bool:
    """Verifica se o email é válido e não está na blacklist."""
    if not email or "@" not in email:
        return False
    return email.lower() not in EMAILS_BLACKLIST


# ── Mapeamento instituição → domínio ──────────────────────────────────────────

INSTITUICAO_DOMINIOS = {
    "receita federal": "rfb.gov.br",
    "delegado da receita": "rfb.gov.br",
    "delegacia da receita": "rfb.gov.br",
    "secretaria da receita": "rfb.gov.br",
    "inss": "inss.gov.br",
    "superintendencia do inss": "inss.gov.br",
    "superintendência do inss": "inss.gov.br",
    "gerente do inss": "inss.gov.br",
    "gerência do inss": "inss.gov.br",
    "superintendente do inss": "inss.gov.br",
    "agencia do inss": "inss.gov.br",
    "agência do inss": "inss.gov.br",
    "ceab/dj": "inss.gov.br",
    "superintendência regional": "inss.gov.br",
    "caixa economica": "caixa.gov.br",
    "caixa econômica": "caixa.gov.br",
    "caixa econ": "caixa.gov.br",
    "gerente da agência": "caixa.gov.br",
    "gerente da agencia": "caixa.gov.br",
    "banco do brasil": "bb.com.br",
    "policia federal": "pf.gov.br",
    "polícia federal": "pf.gov.br",
    "delegado da policia federal": "pf.gov.br",
    "ibama": "ibama.gov.br",
    "incra": "incra.gov.br",
    "funai": "funai.gov.br",
    "dnit": "dnit.gov.br",
    "ufpi": "ufpi.edu.br",
    "universidade federal do piaui": "ufpi.edu.br",
    "universidade federal do piauí": "ufpi.edu.br",
    "ifpi": "ifpi.edu.br",
    "correios": "correios.com.br",
    "anatel": "anatel.gov.br",
    "anvisa": "anvisa.gov.br",
    "antt": "antt.gov.br",
    "codevasf": "codevasf.gov.br",
    "funasa": "funasa.gov.br",
    "icmbio": "icmbio.gov.br",
    "exercito": "eb.mil.br",
    "exército": "eb.mil.br",
    "marinha": "marinha.mil.br",
    "aeronautica": "fab.mil.br",
    "aeronáutica": "fab.mil.br",
    "procuradoria": "agu.gov.br",
    "advocacia geral da uniao": "agu.gov.br",
    "advocacia geral da união": "agu.gov.br",
    "policia civil": "pc.pi.gov.br",
    "polícia civil": "pc.pi.gov.br",
    "bradesco": "bradesco.com.br",
    "itau": "itau-unibanco.com.br",
    "itaú": "itau-unibanco.com.br",
    "santander": "santander.com.br",
    "detran": "detran.pi.gov.br",
    "municipio de teresina": "teresina.pi.gov.br",
    "município de teresina": "teresina.pi.gov.br",
    "prefeitura de teresina": "teresina.pi.gov.br",
    "estado do piaui": "pi.gov.br",
    "estado do piauí": "pi.gov.br",
    "governo do piaui": "pi.gov.br",
    "governo do piauí": "pi.gov.br",
}


def buscar_email_por_instituicao(nome_orgao: str) -> Optional[dict]:
    """
    Busca email pela instituição do destinatário.
    Mapeia palavras-chave → domínio → busca na base de contatos.
    """
    if not nome_orgao:
        return None

    nome_low = _limpar_texto(nome_orgao)

    # Encontrar domínio pela instituição (match mais longo primeiro)
    dominio = None
    melhor_match_len = 0
    for chave, dom in INSTITUICAO_DOMINIOS.items():
        if chave in nome_low and len(chave) > melhor_match_len:
            dominio = dom
            melhor_match_len = len(chave)

    if not dominio:
        return None

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute(
        """SELECT nome, orgao, email, email_alternativo, categoria, fonte
           FROM contatos
           WHERE email LIKE ? AND email != ''
           ORDER BY id DESC""",
        (f"%@{dominio}",),
    )
    rows = cur.fetchall()

    if not rows:
        cur.execute(
            """SELECT nome, orgao, email, email_alternativo, categoria, fonte
               FROM contatos
               WHERE email LIKE ? AND email != ''
               ORDER BY id DESC""",
            (f"%{dominio}%",),
        )
        rows = cur.fetchall()

    conn.close()

    if not rows:
        return None

    emails_unicos = list({r["email"]: dict(r) for r in rows}.values())

    if len(emails_unicos) == 1:
        r = emails_unicos[0]
        return {
            "email": r["email"],
            "email_alternativo": r.get("email_alternativo", ""),
            "nome": r["nome"],
            "orgao": r.get("orgao", ""),
            "fonte": "instituicao_dominio",
        }

    # Vários emails — scoring detalhado
    palavras_query = _extrair_palavras(nome_low)

    def score_inst(r):
        s = 0
        email = (r.get("email") or "").lower()
        nome_r = _limpar_texto(r.get("nome") or "")
        orgao_r = _limpar_texto(r.get("orgao") or "")
        combined = nome_r + " " + orgao_r

        palavras_r = _extrair_palavras(combined)
        intersecao = palavras_query & palavras_r

        # Palavras em comum (cada uma vale pontos)
        s += len(intersecao) * 15

        # Email com "judicial" é preferido (mandados são judiciais)
        if "judicial" in email:
            s += 50

        # Nome do contato contém nome completo do destinatário
        if nome_low in combined:
            s += 100

        # Cargo similar (delegado, gerente, superintendente)
        cargos_query = _extrair_cargos(nome_low)
        cargos_r = _extrair_cargos(combined)
        if cargos_query & cargos_r:
            s += 30

        # Cidade no nome (teresina, piaui)
        for cidade in ["teresina", "piaui", "piauí", "parnaiba", "parnaíba",
                        "floriano", "picos", "campo maior"]:
            if cidade in nome_low and cidade in combined:
                s += 25

        return s

    melhor = max(emails_unicos, key=score_inst)
    if score_inst(melhor) >= 15:
        return {
            "email": melhor["email"],
            "email_alternativo": melhor.get("email_alternativo", ""),
            "nome": melhor["nome"],
            "orgao": melhor.get("orgao", ""),
            "fonte": "instituicao_dominio",
        }
    return None


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

    nome_low = _limpar_texto(nome_orgao)
    termos = _normalizar_termos(nome_orgao)
    palavras_query = _extrair_palavras(nome_low)

    resultados = []
    seen_emails = set()

    # Primeiro: busca exata pelo nome completo
    cur.execute(
        """SELECT nome, orgao, email, email_alternativo, categoria, fonte
           FROM contatos
           WHERE LOWER(nome) = ? AND email != ''
           LIMIT 5""",
        (nome_low,),
    )
    for row in cur.fetchall():
        r = dict(row)
        if r["email"] not in seen_emails and _email_valido(r["email"]):
            seen_emails.add(r["email"])
            r["_score"] = 200  # match exato
            resultados.append(r)

    # Segundo: busca por termos significativos (palavras 4+ chars)
    for termo in termos:
        if len(termo) < 4:
            continue
        cur.execute(
            """SELECT nome, orgao, email, email_alternativo, categoria, fonte
               FROM contatos
               WHERE (LOWER(nome) LIKE ? OR LOWER(orgao) LIKE ?) AND email != ''
               LIMIT 15""",
            (f"%{termo.lower()}%", f"%{termo.lower()}%"),
        )
        for row in cur.fetchall():
            r = dict(row)
            if r["email"] in seen_emails or not _email_valido(r["email"]):
                continue
            seen_emails.add(r["email"])

            # Calcular score
            nome_r = _limpar_texto(r.get("nome") or "")
            orgao_r = _limpar_texto(r.get("orgao") or "")
            combined = nome_r + " " + orgao_r
            palavras_r = _extrair_palavras(combined)

            intersecao = palavras_query & palavras_r
            score = len(intersecao) * 15

            # Bonus match exato de substring
            if nome_low in combined:
                score += 100
            elif combined in nome_low:
                score += 80

            # Proporção de palavras em comum
            if palavras_query:
                ratio = len(intersecao) / len(palavras_query)
                if ratio >= 0.7:
                    score += 50
                elif ratio >= 0.5:
                    score += 25

            r["_score"] = score
            resultados.append(r)

    conn.close()

    if not resultados:
        return None

    # Pegar o melhor score
    melhor = max(resultados, key=lambda r: r["_score"])
    if melhor["_score"] >= 30:
        return {
            "email": melhor["email"],
            "email_alternativo": melhor.get("email_alternativo", ""),
            "nome": melhor["nome"],
            "orgao": melhor.get("orgao", ""),
            "categoria": melhor.get("categoria", ""),
            "fonte": melhor.get("fonte") or "agenda",
        }
    return None


def buscar_email_em_processos(nome_orgao: str) -> Optional[dict]:
    """
    Busca email em processos anteriores com destinatário similar.
    Só retorna se match for forte (>= 60% das palavras em comum).
    """
    if not nome_orgao or len(nome_orgao.strip()) < 3:
        return None

    nome_low = _limpar_texto(nome_orgao)
    palavras_query = _extrair_palavras(nome_low)

    if not palavras_query:
        return None

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    # Buscar processos com email preenchido
    cur.execute(
        """SELECT DISTINCT email_destinatario, destinatario
           FROM processos
           WHERE email_destinatario IS NOT NULL
             AND email_destinatario != ''
           ORDER BY id DESC
           LIMIT 200"""
    )
    rows = cur.fetchall()
    conn.close()

    candidatos = {}  # email -> {score, count, dest}

    for row in rows:
        email = row["email_destinatario"] or ""
        dest = row["destinatario"] or ""

        if not _email_valido(email):
            continue

        dest_low = _limpar_texto(dest)
        palavras_dest = _extrair_palavras(dest_low)

        # Match exato
        if dest_low == nome_low:
            ratio = 1.0
        elif palavras_dest:
            intersecao = palavras_query & palavras_dest
            ratio = len(intersecao) / max(len(palavras_query), 1)
        else:
            ratio = 0

        # Só aceitar se >= 60% das palavras em comum
        if ratio >= 0.6:
            if email not in candidatos:
                candidatos[email] = {"score": 0, "count": 0, "dest": dest}
            candidatos[email]["score"] += ratio * 10
            candidatos[email]["count"] += 1

    if not candidatos:
        return None

    # Pegar email com melhor score
    melhor_email = max(candidatos, key=lambda e: candidatos[e]["score"])
    info = candidatos[melhor_email]

    if info["score"] >= 6:  # pelo menos 60% match
        return {
            "email": melhor_email,
            "nome": info["dest"],
            "orgao": nome_orgao,
            "fonte": "processos_anteriores",
        }
    return None


def buscar_email_outlook(nome_orgao: str) -> Optional[dict]:
    """
    Busca email nos itens enviados/recebidos do Outlook.
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
            sent_folder = namespace.GetDefaultFolder(5)
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
                            recipients = item.Recipients
                            for r in recipients:
                                email = _extrair_email_recipient(r)
                                if email and _email_valido(email):
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
            inbox = namespace.GetDefaultFolder(6)
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
                            if sender_email and _email_valido(sender_email):
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
            melhor = max(encontrados.values(), key=lambda x: x["score"])
            return melhor

    except Exception as e:
        print(f"[outlook] erro: {e}")

    return None


def buscar_sugestoes(nome_orgao: str, limit: int = 5) -> list:
    """Retorna lista de sugestões parciais da base."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    termos = _normalizar_termos(nome_orgao)
    sugestoes = []
    seen = set()

    for termo in termos:
        if len(termo) < 4:
            continue
        cur.execute(
            """SELECT nome, orgao, email, categoria
               FROM contatos
               WHERE (LOWER(nome) LIKE ? OR LOWER(orgao) LIKE ?) AND email != ''
               LIMIT ?""",
            (f"%{termo.lower()}%", f"%{termo.lower()}%", limit),
        )
        for row in cur.fetchall():
            email = row["email"]
            if email not in seen and _email_valido(email):
                seen.add(email)
                sugestoes.append(dict(row))

    conn.close()
    return sugestoes[:limit]


def salvar_email_manual(nome_orgao: str, email: str, fonte: str = "manual") -> bool:
    """Salva email manualmente na base para aprendizado futuro."""
    conn = sqlite3.connect(DB_PATH)
    try:
        cur = conn.cursor()
        cur.execute("SELECT id FROM contatos WHERE orgao = ? OR nome = ?",
                     (nome_orgao, nome_orgao))
        row = cur.fetchone()
        if row:
            cur.execute("UPDATE contatos SET email = ?, fonte = ? WHERE id = ?",
                        (email, fonte, row[0]))
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

def _limpar_texto(texto: str) -> str:
    """Normaliza texto: lowercase, remove acentos parciais, trim."""
    return (texto or "").lower().strip()


def _extrair_palavras(texto: str) -> set:
    """Extrai palavras significativas (3+ chars, sem stopwords)."""
    palavras = set(re.findall(r'\b\w{3,}\b', texto.lower()))
    return palavras - STOPWORDS


def _extrair_cargos(texto: str) -> set:
    """Extrai cargos/funções do texto."""
    cargos = {"delegado", "delegada", "gerente", "superintendente",
              "diretor", "diretora", "presidente", "chefe", "coordenador",
              "coordenadora", "procurador", "procuradora", "juiz", "juíza",
              "prefeito", "prefeita", "secretario", "secretária", "reitor",
              "reitora", "oficial"}
    palavras = set(re.findall(r'\b\w{3,}\b', texto.lower()))
    return palavras & cargos


def _normalizar_termos(texto: str) -> list[str]:
    """Extrai termos relevantes de um nome/órgão para busca."""
    palavras = re.findall(r'\b\w{3,}\b', texto.lower())
    termos = [p for p in palavras if p not in STOPWORDS]

    resultado = []
    resultado.extend(termos)

    # Siglas (maiúsculas no original)
    siglas = re.findall(r'\b[A-Z]{2,}\b', texto)
    resultado.extend([s.lower() for s in siglas])

    return list(dict.fromkeys(resultado))


def _extrair_email_recipient(recipient) -> Optional[str]:
    """Extrai email de um objeto Recipient do Outlook."""
    try:
        addr = recipient.AddressEntry
        if addr.AddressEntryUserType == 0:
            ex_user = addr.GetExchangeUser()
            if ex_user:
                return ex_user.PrimarySmtpAddress
        return addr.Address
    except Exception:
        try:
            return recipient.Address
        except Exception:
            return None

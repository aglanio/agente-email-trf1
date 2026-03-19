"""
Extrai informações dos PDFs do PJe (mandado e anexo).
"""
import re, os

def extrair_info_pdf(pdf_path: str) -> dict:
    """Extrai número do processo, destinatário e e-mail de um PDF do PJe."""
    resultado = {
        "numero_processo": "",
        "tipo": "",
        "destinatario": "",
        "email": "",
        "endereco": "",
        "texto_completo": "",
    }
    if not os.path.exists(pdf_path):
        return resultado
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            texto = "\n".join(p.extract_text() or "" for p in pdf.pages)
        resultado["texto_completo"] = texto
    except ImportError:
        # Fallback sem pdfplumber
        try:
            import subprocess
            r = subprocess.run(["python", "-c",
                f"import sys; sys.stdout.reconfigure(encoding='utf-8')"],
                capture_output=True, text=True)
        except Exception:
            pass
        return resultado
    except Exception as e:
        print(f"Erro ao ler PDF {pdf_path}: {e}")
        return resultado

    # Número do processo (formato TRF: 0000000-00.0000.0.00.0000)
    proc = re.search(
        r'\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}',
        texto
    )
    if proc:
        resultado["numero_processo"] = proc.group(0)

    # Tipo de ato (Intimação, Citação, etc.)
    tipo = re.search(
        r'\b(Intima[çc][ãa]o|Cita[çc][ãa]o|Notifica[çc][ãa]o|Mandado)\b',
        texto, re.IGNORECASE
    )
    if tipo:
        resultado["tipo"] = tipo.group(0)

    # Destinatário: busca padrão "INTIMAÇÃO DE:" ou "DESTINATÁRIO:"
    dest = re.search(
        r'INTIMA[ÇC][ÃA]O\s+DE[:\s]+(.+?)(?:\n|Endere)',
        texto, re.IGNORECASE | re.DOTALL
    )
    if dest:
        resultado["destinatario"] = dest.group(1).strip()[:200]
    else:
        # Fallback: linha após "Destinatário(s)"
        dest2 = re.search(r'Destinat[aá]rio\(?s?\)?\s*[:\n]+(.+)', texto, re.IGNORECASE)
        if dest2:
            resultado["destinatario"] = dest2.group(1).strip()[:200]

    # E-mail embutido no endereço (padrão PJe)
    email = _extrair_email_pje(texto)
    if email:
        resultado["email"] = email

    # Endereço
    end = re.search(r'Endere[çc]o[:\s]+(.+?)(?:\n|CEP|$)', texto, re.IGNORECASE)
    if end:
        resultado["endereco"] = end.group(1).strip()[:300]

    return resultado

def _extrair_email_pje(texto: str) -> str:
    """
    O PJe insere o e-mail dentro do campo endereço sem o '@'.
    Ex: 'Rua X, 100, aps16001120inss.gov.br, Centro, TERESINA'
    Converte para 'aps16001120@inss.gov.br'
    """
    # 1. E-mails normais
    emails = re.findall(r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}', texto)
    if emails:
        # Filtra e-mails que não são do remetente
        for e in emails:
            if 'aglanio' not in e and 'trf1' not in e.lower():
                return e.lower()

    # 2. Padrão PJe sem @: prefixo + domínio colados
    dominios_conhecidos = [
        r'inss\.gov\.br', r'caixa\.gov\.br', r'bb\.com\.br',
        r'trf1\.jus\.br', r'pgfn\.fazenda\.gov\.br', r'mpf\.mp\.br',
        r'mppi\.mp\.br', r'pf\.gov\.br', r'fazenda\.gov\.br',
        r'receita\.fazenda\.gov\.br', r'economia\.gov\.br',
        r'trabalho\.gov\.br', r'saude\.gov\.br', r'mec\.gov\.br',
        r'anatel\.gov\.br', r'ibama\.gov\.br', r'dnit\.gov\.br',
        r'gmail\.com', r'hotmail\.com', r'outlook\.com',
    ]
    for dom in dominios_conhecidos:
        pat = r'([a-zA-Z0-9._%-]+)(' + dom + r')'
        m = re.search(pat, texto)
        if m:
            prefixo = m.group(1).rstrip('.')
            dominio = m.group(2)
            # Verifica se tem @ (email normal já tratado acima)
            if '@' not in prefixo:
                return f"{prefixo}@{dominio}".lower()

    return ""

def montar_assunto(info: dict, template_assunto: str = None) -> str:
    """Gera o assunto do e-mail baseado nas informações do processo."""
    num = info.get("numero_processo", "")
    tipo = info.get("tipo") or "Mandado"
    if template_assunto:
        return template_assunto.format(
            numero_processo=num,
            tipo=tipo,
        )
    return f"{tipo} - Processo {num} - SJPI"

def montar_corpo(info: dict, template_corpo: str = None) -> str:
    """Gera o corpo do e-mail."""
    num = info.get("numero_processo", "")
    tipo = info.get("tipo") or "Mandado"
    dest = info.get("destinatario") or "Senhor(a)"

    if template_corpo:
        return template_corpo.format(
            numero_processo=num,
            tipo=tipo,
            destinatario=dest,
        )

    from datetime import date as _date
    meses = ["janeiro","fevereiro","março","abril","maio","junho",
             "julho","agosto","setembro","outubro","novembro","dezembro"]
    hoje = _date.today()
    data_str = f"Teresina, {hoje.day} de {meses[hoje.month-1]} de {hoje.year}"
    tipo_str = (tipo or "documento").lower()
    return (
        f"Prezado(a) Senhor(a),\n\n"
        f"Encaminhamos em anexo o {tipo_str} judicial referente ao "
        f"processo {num}, expedido pela Seção Judiciária do Piauí - SJPI.\n\n"
        f"Atenciosamente,\n\n"
        f"{data_str}\n\n"
        f"Aglanio Frota Moura Carvalho\n"
        f"Oficial de Justiça Avaliador Federal     PI100327\n"
        f"Seção Judiciária do Piauí - TRF 1ª Região\n"
        f"aglanio.carvalho@trf1.jus.br"
    )

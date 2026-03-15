"""
Importa contatos das agendas DOCX para o banco SQLite.
"""
import re, os, zipfile
from xml.etree import ElementTree as ET
from database import get_conn, init_db

PASTA_AGENDAS = r"C:\Users\aglan\OneDrive\Documentos\claude projetos\teams"

CATEGORIAS = {
    "AGENDA 01": "Autarquias/Órgãos Federais",
    "AGENDA 02": "Advogados/Peritos",
    "AGENDA 03": "INSS",
    "AGENDA 04": "Universidades/IF",
}

def ler_docx(path):
    try:
        with zipfile.ZipFile(path) as z:
            with z.open('word/document.xml') as f:
                tree = ET.parse(f)
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        linhas = []
        for p in tree.findall('.//w:p', ns):
            linha = ''.join(t.text or '' for t in p.findall('.//w:t', ns)).strip()
            if linha:
                linhas.append(linha)
        return linhas
    except Exception as e:
        print(f"Erro ao ler {path}: {e}")
        return []

def extrair_emails(texto):
    """Extrai e-mails de um texto, corrigindo padrões sem @."""
    # E-mails normais
    emails = re.findall(r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}', texto)
    # Padrão PJe: "aps16001120inss.gov.br" → "aps16001120@inss.gov.br"
    pje_pattern = re.findall(
        r'\b([a-zA-Z0-9._%-]+)(inss\.gov\.br|caixa\.gov\.br|bb\.com\.br|'
        r'trf1\.jus\.br|jfpi\.jus\.br|mppi\.mp\.br|pgfn\.fazenda\.gov\.br)\b',
        texto
    )
    for prefixo, dominio in pje_pattern:
        email_fixed = f"{prefixo}@{dominio}"
        if '@' not in prefixo and email_fixed not in emails:
            emails.append(email_fixed)
    return list(set(e.lower().strip() for e in emails if e.strip()))

def parse_bloco_contato(linhas, idx, categoria, fonte):
    """Tenta extrair um contato de um bloco de linhas."""
    bloco = '\n'.join(linhas[max(0,idx-1):idx+5])
    emails = extrair_emails(bloco)
    if not emails:
        return None

    nome = linhas[idx].strip() if idx < len(linhas) else ""
    orgao = ""
    telefone = re.search(r'[\(\d]{2,}[\d\s\-\.\/\(\)]{5,}', bloco)
    tel_str = telefone.group(0).strip() if telefone else ""

    return {
        "nome": nome[:200],
        "orgao": orgao[:200],
        "email": emails[0],
        "email_alternativo": emails[1] if len(emails) > 1 else "",
        "telefone": tel_str[:100],
        "endereco": "",
        "categoria": categoria,
        "observacoes": bloco[:500],
        "fonte": fonte,
    }

def importar_agenda(path, categoria, fonte):
    linhas = ler_docx(path)
    contatos = []
    texto_completo = '\n'.join(linhas)

    # Extração inteligente: busca padrões "nome + email" por proximidade
    blocos = []
    bloco_atual = []
    for linha in linhas:
        if any(sep in linha for sep in ['01)', '02)', '03)', '---', '====', 'ATENÇÃO']):
            if bloco_atual:
                blocos.append(bloco_atual)
            bloco_atual = []
        bloco_atual.append(linha)
    if bloco_atual:
        blocos.append(bloco_atual)

    conn = get_conn()
    c = conn.cursor()
    inseridos = 0

    # Processar linha a linha buscando e-mails
    for i, linha in enumerate(linhas):
        emails = extrair_emails(linha)
        if not emails:
            # Verifica se linha adjacente tem e-mail
            prox = linhas[i+1] if i+1 < len(linhas) else ""
            emails = extrair_emails(prox)
            if not emails:
                continue

        # Determinar nome/orgão: linha atual ou anterior
        nome = linha.strip()
        orgao = linhas[i-1].strip() if i > 0 else ""

        # Se nome parece telefone, usar anterior
        if re.match(r'^[\d\s\(\)\-\.\/]+$', nome):
            nome = orgao
            orgao = linhas[i-2].strip() if i > 1 else ""

        # Bloco de contexto para observações
        ctx = '\n'.join(linhas[max(0,i-2):i+4])
        tel = re.search(r'(?:Tel|Fone|Fax)?[:\s]*(?:\(86\)|\(85\)|86|85)[\s\d\.\-]{7,}', ctx)

        for email in emails:
            # Verifica duplicata
            c.execute("SELECT id FROM contatos WHERE email=?", (email,))
            if c.fetchone():
                continue

            c.execute("""
                INSERT INTO contatos (nome, orgao, email, telefone, categoria, observacoes, fonte)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                nome[:200],
                orgao[:200],
                email,
                tel.group(0).strip()[:80] if tel else "",
                categoria,
                ctx[:500],
                fonte,
            ))
            inseridos += 1

    conn.commit()
    conn.close()
    return inseridos

def importar_todas():
    init_db()
    total = 0
    for arq in os.listdir(PASTA_AGENDAS):
        if not arq.endswith('.docx'):
            continue
        path = os.path.join(PASTA_AGENDAS, arq)
        cat_key = arq[:8].strip()
        categoria = CATEGORIAS.get(cat_key, "Outros")
        n = importar_agenda(path, categoria, arq)
        print(f"  {arq}: {n} contatos importados")
        total += n
    print(f"\nTotal importado: {total} contatos")
    return total

if __name__ == "__main__":
    importar_todas()

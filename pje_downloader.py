"""
pje_downloader.py — Baixa mandados do PJe TRF1 via Playwright.
Usa o perfil do Chrome do usuário (sessão e certificados já autenticados).
"""
import asyncio
import os
import time
from pathlib import Path
from typing import Optional

DOWNLOADS_DIR = Path(__file__).parent / "downloads"
DOWNLOADS_DIR.mkdir(exist_ok=True)

# URL base do PJe TRF1
PJE_BASE = "https://pje2g.trf1.jus.br/pje"
PJE_PAINEL = PJE_BASE + "/Painel/painel_usuario/Paniel_Usuario_Oficial_Justica/listView.seam"

# Perfil do Chrome do usuário (tem certificados e sessão)
CHROME_PROFILE = Path(os.environ.get("CHROME_USER_DATA",
    r"C:\Users\aglan\AppData\Local\Google\Chrome\User Data"))


async def baixar_documentos_pje(
    numeros_processo: list[str],
    on_progress=None,
) -> list[dict]:
    """
    Abre o PJe no Chrome do usuário e baixa os PDFs dos mandados.
    Retorna lista de {"numero": str, "ok": bool, "arquivo": str|None, "erro": str|None}
    """
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        return [{"numero": n, "ok": False, "erro": "playwright não instalado. Execute: pip install playwright && playwright install chromium"} for n in numeros_processo]

    resultados = []

    async with async_playwright() as p:
        # Abrir Chrome com o perfil do usuário (tem sessão PJe ativa)
        try:
            context = await p.chromium.launch_persistent_context(
                user_data_dir=str(CHROME_PROFILE),
                channel="chrome",  # usa Chrome instalado (não Chromium baixado)
                headless=False,    # visível para o usuário acompanhar / autenticar
                downloads_path=str(DOWNLOADS_DIR),
                args=["--no-sandbox", "--disable-blink-features=AutomationControlled"],
            )
        except Exception as e:
            # Fallback: Chromium sem perfil
            try:
                context = await p.chromium.launch_persistent_context(
                    user_data_dir=str(CHROME_PROFILE),
                    headless=False,
                    downloads_path=str(DOWNLOADS_DIR),
                )
            except Exception as e2:
                return [{"numero": n, "ok": False, "erro": f"Erro ao abrir browser: {e2}"} for n in numeros_processo]

        page = await context.new_page()

        # Navegar para o painel do Oficial de Justiça
        if on_progress:
            on_progress({"etapa": "abrindo_pje", "msg": "Abrindo painel PJe..."})

        try:
            await page.goto(PJE_PAINEL, wait_until="networkidle", timeout=30000)
            # Aguardar um pouco para ver se precisa de login
            await page.wait_for_timeout(3000)

            # Verificar se está logado (procurar por elementos que indicam autenticação)
            url_atual = page.url
            if "login" in url_atual.lower() or "certificado" in url_atual.lower():
                if on_progress:
                    on_progress({"etapa": "aguardando_login", "msg": "Faça login no PJe... aguardando 60 segundos."})
                # Aguardar até 60 segundos para o usuário autenticar
                for _ in range(60):
                    await page.wait_for_timeout(1000)
                    if "login" not in page.url.lower():
                        break

        except Exception as e:
            if on_progress:
                on_progress({"etapa": "erro", "msg": f"Erro ao abrir PJe: {e}"})

        # Para cada número de processo, tentar encontrar e baixar o PDF
        for numero in numeros_processo:
            resultado = {"numero": numero, "ok": False, "arquivo": None, "erro": None}

            if on_progress:
                on_progress({"etapa": "buscando", "msg": f"Buscando {numero}...", "numero": numero})

            try:
                arquivo = await _baixar_mandado(page, numero, context)
                if arquivo:
                    resultado["ok"] = True
                    resultado["arquivo"] = str(arquivo)
                else:
                    # Salvar screenshot para debug dos seletores
                    debug_img = DOWNLOADS_DIR / f"debug_{numero[:20]}.png"
                    try:
                        await page.screenshot(path=str(debug_img))
                        resultado["erro"] = f"Mandado não encontrado. Screenshot salvo: {debug_img.name}"
                    except Exception:
                        resultado["erro"] = "Mandado não encontrado na página"
            except Exception as e:
                resultado["erro"] = str(e)

            resultados.append(resultado)
            await page.wait_for_timeout(1500)  # pausa entre processos

        await context.close()

    return resultados


async def _baixar_mandado(page, numero_processo: str, context) -> Optional[Path]:
    """
    No painel PJe do Oficial de Justiça, localiza a linha do processo e baixa o mandado.

    O painel exibe uma tabela com colunas:
    - Esquerda: botões de ação (impressora, editar, abrir)
    - Centro: dados do ato (número, destinatário, datas)
    - Direita: botão de anexo/download do documento

    Estratégia:
    1. Acha a linha da tabela que contém o número do processo
    2. Tenta o botão DIREITO (download direto do anexo)
    3. Tenta o botão ESQUERDO (impressora → nova aba → salvar como PDF)
    """
    try:
        await page.goto(PJE_PAINEL, wait_until="networkidle", timeout=20000)
        await page.wait_for_timeout(2000)

        # Parte do número para localizar a linha (primeiros 15 chars são únicos)
        num_curto = numero_processo[:15]

        nome_arquivo = f"mandado_{numero_processo.replace('-', '_').replace('.', '_')}.pdf"
        destino = DOWNLOADS_DIR / nome_arquivo

        # ── Tentar encontrar a LINHA da tabela que contém o processo ─────────
        # O PJe usa <tr> com células contendo o número do processo
        linha = None
        row_selectors = [
            f"tr:has-text('{numero_processo}')",
            f"tr:has-text('{num_curto}')",
        ]
        for sel in row_selectors:
            try:
                linha = await page.wait_for_selector(sel, timeout=4000)
                if linha:
                    break
            except Exception:
                continue

        if linha:
            # ── Botão DIREITO: ícone de anexo/documento (último td) ──────────
            # No PJe, o último td da linha de ato costuma ter o ícone de arquivo
            botao_direito_sels = [
                "td:last-child span",
                "td:last-child a",
                "td:last-child button",
                "td:last-child input[type='image']",
            ]
            for sel in botao_direito_sels:
                try:
                    btn = await linha.query_selector(sel)
                    if not btn:
                        continue
                    # Tenta download direto
                    try:
                        async with page.expect_download(timeout=15000) as dl_info:
                            await btn.click()
                        dl = await dl_info.value
                        await dl.save_as(str(destino))
                        return destino
                    except Exception:
                        # Pode ter aberto nova aba em vez de download
                        pass
                except Exception:
                    continue

            # ── Botão ESQUERDO: impressora (primeiro td) ─────────────────────
            # Clica no ícone de impressora → nova aba com documento → page.pdf()
            botao_esq_sels = [
                "td:first-child button:first-child",
                "td:first-child a:first-child",
                "td:first-child span:first-child",
                "td:first-child input[type='image']:first-child",
            ]
            for sel in botao_esq_sels:
                try:
                    btn = await linha.query_selector(sel)
                    if not btn:
                        continue
                    try:
                        async with context.expect_page() as new_page_info:
                            await btn.click()
                        doc_page = await new_page_info.value
                        await doc_page.wait_for_load_state("networkidle", timeout=20000)
                        await doc_page.wait_for_timeout(1500)
                        await doc_page.pdf(path=str(destino), format="A4", print_background=True)
                        await doc_page.close()
                        if destino.exists() and destino.stat().st_size > 2048:
                            return destino
                    except Exception:
                        pass
                except Exception:
                    continue

        # ── Fallback: tentar download em toda a página ────────────────────────
        return await _encontrar_e_baixar_pdf(page, numero_processo, context)

    except Exception as e:
        print(f"[pje_downloader] erro em {numero_processo}: {e}")
        return None


async def _encontrar_e_baixar_pdf(page, numero_processo: str, context) -> Optional[Path]:
    """
    Na página do processo (painel PJe), tenta baixar o PDF do mandado.

    Estratégia baseada na UI do PJe TRF1:
    1. Botão DIREITO (ícone de anexo/documento) — download direto do arquivo
    2. Botão ESQUERDO (ícone de impressora) — abre nova aba com o documento → salvar como PDF
    """
    nome_arquivo = f"mandado_{numero_processo.replace('-', '_').replace('.', '_')}.pdf"
    destino = DOWNLOADS_DIR / nome_arquivo

    # ── Estratégia 1: Botão de anexo (ícone direito) ─────────────────────────
    # O PJe tem um ícone de documento/clipe no lado direito da linha do mandado.
    # Seletores comuns para esse botão no PJe TRF1:
    anexo_selectors = [
        # ícone de arquivo no final da linha (último td da tabela de atos)
        "td:last-child span[class*='ui-icon-document']",
        "td:last-child a[class*='download']",
        "td:last-child button[title*='Baixar']",
        "td:last-child button[title*='Download']",
        "td:last-child a[href*='download']",
        "td:last-child a[href*='documento']",
        "td:last-child a[title*='Baixar']",
        "td:last-child span[title*='Baixar']",
        # ícone de PDF genérico
        "a[href*='.pdf']",
        "a[title*='Baixar documento']",
        "span[class*='ui-icon-arrowthick'] + a",
    ]

    for sel in anexo_selectors:
        try:
            el = await page.wait_for_selector(sel, timeout=1500)
            if not el:
                continue
            async with page.expect_download(timeout=25000) as dl_info:
                await el.click()
            download = await dl_info.value
            await download.save_as(str(destino))
            return destino
        except Exception:
            continue

    # ── Estratégia 2: Botão de impressora (ícone esquerdo) ───────────────────
    # Clica no ícone de impressora → abre nova aba com o HTML/PDF do documento.
    # Usamos page.pdf() para salvar sem abrir o dialog do SO.
    impressora_selectors = [
        "button[title*='Imprimir']",
        "button[title*='imprimir']",
        "a[title*='Imprimir']",
        "span[class*='ui-icon-print']",
        "span[class*='print']",
        # primeiro botão da linha (costuma ser impressora no PJe)
        "td:first-child button:first-child",
        "td:first-child a:first-child",
    ]

    for sel in impressora_selectors:
        try:
            el = await page.wait_for_selector(sel, timeout=1500)
            if not el:
                continue

            # Tenta capturar nova aba que abre ao clicar no botão de impressora
            try:
                async with context.expect_page() as new_page_info:
                    await el.click()
                doc_page = await new_page_info.value
                await doc_page.wait_for_load_state("networkidle", timeout=20000)
                # Salvar como PDF usando o Playwright (sem dialog do SO)
                await doc_page.pdf(path=str(destino), format="A4")
                await doc_page.close()
                if destino.exists() and destino.stat().st_size > 1024:
                    return destino
            except Exception:
                # Se não abriu nova aba, tenta salvar a página atual como PDF
                await page.pdf(path=str(destino), format="A4")
                if destino.exists() and destino.stat().st_size > 1024:
                    return destino

        except Exception:
            continue

    # ── Estratégia 3: Qualquer link/botão que dispare download ───────────────
    fallback_selectors = [
        "a:has-text('Mandado')",
        "a:has-text('mandado')",
        "a:has-text('Baixar')",
        "a:has-text('Imprimir')",
        "td:has-text('Mandado') a",
    ]
    for sel in fallback_selectors:
        try:
            el = await page.wait_for_selector(sel, timeout=1500)
            if not el:
                continue
            async with page.expect_download(timeout=20000) as dl_info:
                await el.click()
            download = await dl_info.value
            await download.save_as(str(destino))
            return destino
        except Exception:
            continue

    return None


# ── Versão síncrona para uso no FastAPI ──────────────────────────────────────

def baixar_documentos_pje_sync(
    numeros_processo: list[str],
    on_progress=None,
) -> list[dict]:
    """Wrapper síncrono para uso em endpoints FastAPI."""
    try:
        loop = asyncio.new_event_loop()
        return loop.run_until_complete(
            baixar_documentos_pje(numeros_processo, on_progress)
        )
    except Exception as e:
        return [{"numero": n, "ok": False, "erro": str(e)} for n in numeros_processo]
    finally:
        loop.close()

"""
pje_downloader.py — Baixa mandados do PJe TRF1 via Playwright.
Usa o Chrome do usuário (sessão PJe já autenticada com certificado).

Fluxo:
1. Abre o painel do Oficial de Justiça (1G ou 2G)
2. Localiza o processo pelo número
3. Clica no ícone de Imprimir (🖨️) → salva mandado como PDF
4. Clica em "Gerar PDF" nos Anexos → baixa o anexo
"""
import asyncio
import os
import re
import sqlite3
from pathlib import Path
from typing import Optional

DOWNLOADS_DIR = Path(__file__).parent / "downloads"
DOWNLOADS_DIR.mkdir(exist_ok=True)

PJE_URLS = {
    1: "https://pje1g.trf1.jus.br/pje/Painel/painel_usuario/Paniel_Usuario_Oficial_Justica/listView.seam",
    2: "https://pje2g.trf1.jus.br/pje/Painel/painel_usuario/Paniel_Usuario_Oficial_Justica/listView.seam",
}

# Chrome profile do usuário (tem certificado digital e sessão PJe)
CHROME_PROFILE = os.environ.get(
    "CHROME_USER_DATA",
    r"C:\Users\aglan\AppData\Local\Google\Chrome\User Data"
)


async def baixar_documentos_pje(
    processos: list[dict],
    on_progress=None,
) -> list[dict]:
    """
    Baixa mandados e anexos do PJe para cada processo.
    processos: lista de dicts com 'id', 'numero_processo', 'grau' (1 ou 2)
    Retorna lista de resultados com status de cada download.
    """
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        return [{"numero_processo": p.get("numero_processo", ""), "ok": False,
                 "erro": "playwright não instalado"} for p in processos]

    resultados = []

    async with async_playwright() as pw:
        # Conecta ao Chrome que JÁ ESTÁ ABERTO via CDP (Chrome DevTools Protocol)
        # O Chrome precisa ter sido iniciado com --remote-debugging-port=9222
        # Caso contrário, abre uma nova instância com o perfil do usuário
        browser = None
        context = None
        page = None

        # Tenta conectar ao Chrome aberto via CDP
        try:
            browser = await pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
            context = browser.contexts[0] if browser.contexts else await browser.new_context()
            page = await context.new_page()
            if on_progress:
                on_progress({"etapa": "conectado", "msg": "Conectado ao Chrome aberto!"})
        except Exception:
            # CDP não disponível — abre Chrome com perfil do usuário
            try:
                context = await pw.chromium.launch_persistent_context(
                    user_data_dir=CHROME_PROFILE,
                    channel="chrome",
                    headless=False,
                    accept_downloads=True,
                    args=["--no-sandbox", "--disable-blink-features=AutomationControlled"],
                )
                page = context.pages[0] if context.pages else await context.new_page()
            except Exception as e:
                try:
                    browser = await pw.chromium.launch(channel="chrome", headless=False)
                    context = await browser.new_context(accept_downloads=True)
                    page = await context.new_page()
                except Exception as e2:
                    return [{"numero_processo": p.get("numero_processo", ""),
                             "ok": False, "erro": f"Feche o Chrome e tente novamente, ou inicie com: chrome --remote-debugging-port=9222"} for p in processos]

        # Agrupar por grau para não ficar alternando
        por_grau = {}
        for p in processos:
            g = p.get("grau", 1)
            por_grau.setdefault(g, []).append(p)

        for grau, procs in por_grau.items():
            url_painel = PJE_URLS.get(grau, PJE_URLS[1])

            if on_progress:
                on_progress({"etapa": "abrindo_pje", "grau": grau,
                             "msg": f"Abrindo PJe {grau}º grau..."})

            try:
                await page.goto(url_painel, wait_until="networkidle", timeout=30000)
                await page.wait_for_timeout(3000)

                # Verificar se precisa login
                if "login" in page.url.lower() or "certificado" in page.url.lower():
                    if on_progress:
                        on_progress({"etapa": "aguardando_login",
                                     "msg": "Faça login no PJe com certificado... aguardando 90s"})
                    for _ in range(90):
                        await page.wait_for_timeout(1000)
                        if "listView" in page.url or "painel" in page.url.lower():
                            break

            except Exception as e:
                for p in procs:
                    resultados.append({"numero_processo": p.get("numero_processo", ""),
                                       "ok": False, "erro": f"Erro PJe: {e}"})
                continue

            # Para cada processo neste grau
            for proc in procs:
                numero = proc.get("numero_processo", "")
                proc_id = proc.get("id")
                resultado = {"numero_processo": numero, "id": proc_id,
                             "ok": False, "mandado": None, "anexo": None, "erro": None}

                if on_progress:
                    on_progress({"etapa": "buscando", "numero": numero,
                                 "msg": f"Buscando {numero}..."})

                try:
                    mandado, anexo = await _baixar_processo(page, context, numero, grau)
                    if mandado:
                        resultado["mandado"] = str(mandado)
                        resultado["ok"] = True
                    if anexo:
                        resultado["anexo"] = str(anexo)
                        resultado["ok"] = True
                    if not mandado and not anexo:
                        resultado["erro"] = "Processo não encontrado no painel"
                except Exception as e:
                    resultado["erro"] = str(e)

                resultados.append(resultado)
                await page.wait_for_timeout(1000)

        await context.close()
        await browser.close()

    return resultados


async def _baixar_processo(page, context, numero_processo: str, grau: int):
    """
    No painel do Oficial, localiza o processo e baixa:
    1. Mandado (botão Imprimir 🖨️ → nova aba → salvar como PDF)
    2. Anexo (botão "Gerar PDF" na coluna Anexos)
    """
    mandado_path = None
    anexo_path = None
    num_safe = numero_processo.replace("-", "_").replace(".", "_")

    # Garantir que está no painel
    url_painel = PJE_URLS.get(grau, PJE_URLS[1])
    if "listView" not in page.url:
        await page.goto(url_painel, wait_until="networkidle", timeout=20000)
        await page.wait_for_timeout(2000)

    # Localizar a linha do processo na tabela
    # O PJe mostra o número em negrito: "Usucap 1029324-46.2021.4.01.4000 - Intimação"
    num_curto = numero_processo[:20]  # suficiente para ser único

    # Procurar texto que contém o número do processo
    try:
        linha = await page.wait_for_selector(
            f"tr:has-text('{num_curto}')", timeout=5000
        )
    except Exception:
        # Pode estar em outra página, tentar scroll
        try:
            # Tentar buscar pelo campo de pesquisa se existir
            campos_busca = await page.query_selector_all("input[type='text']")
            for campo in campos_busca:
                placeholder = await campo.get_attribute("placeholder") or ""
                name = await campo.get_attribute("name") or ""
                if "processo" in placeholder.lower() or "processo" in name.lower():
                    await campo.fill(numero_processo)
                    # Clicar botão pesquisar
                    pesquisar = await page.query_selector("button:has-text('PESQUISAR')")
                    if pesquisar:
                        await pesquisar.click()
                        await page.wait_for_timeout(3000)
                    break

            linha = await page.wait_for_selector(
                f"tr:has-text('{num_curto}')", timeout=5000
            )
        except Exception:
            return None, None

    if not linha:
        return None, None

    # ═══ 1. MANDADO — Botão Imprimir (🖨️) ═══
    # É o PRIMEIRO ícone/botão na linha (lado esquerdo)
    # Pelas imagens: ícones são 🖨️ ✏️ 🔗 na primeira célula
    try:
        # Seletores para o botão de imprimir (primeiro botão da linha)
        print_sels = [
            "input[type='image'][title*='mprimir']",
            "button[title*='mprimir']",
            "a[title*='mprimir']",
            "span[title*='mprimir']",
            "input[type='image']:first-of-type",
            "button:first-of-type",
        ]
        for sel in print_sels:
            btn = await linha.query_selector(sel)
            if btn:
                # Clicar abre nova aba com o documento HTML
                try:
                    async with context.expect_page(timeout=15000) as new_page_info:
                        await btn.click()
                    doc_page = await new_page_info.value
                    await doc_page.wait_for_load_state("networkidle", timeout=20000)
                    await doc_page.wait_for_timeout(2000)

                    # Salvar como PDF
                    mandado_file = DOWNLOADS_DIR / f"mandado_{num_safe}.pdf"
                    await doc_page.pdf(
                        path=str(mandado_file),
                        format="A4",
                        print_background=True,
                        margin={"top": "10mm", "bottom": "10mm", "left": "10mm", "right": "10mm"},
                    )
                    await doc_page.close()

                    if mandado_file.exists() and mandado_file.stat().st_size > 1024:
                        mandado_path = mandado_file
                        break
                except Exception as e:
                    print(f"[pje] Erro imprimir {numero_processo}: {e}")
                    continue
    except Exception as e:
        print(f"[pje] Erro buscando botão imprimir: {e}")

    # ═══ 2. ANEXO — Botão "Gerar PDF" (coluna Anexos, lado direito) ═══
    try:
        anexo_sels = [
            "input[type='image'][title*='erar PDF']",
            "button[title*='erar PDF']",
            "a[title*='erar PDF']",
            "td:last-child input[type='image']",
            "td:last-child button",
            "td:last-child a[href*='documento']",
        ]
        for sel in anexo_sels:
            btn = await linha.query_selector(sel)
            if btn:
                try:
                    # Pode ser download direto ou nova aba
                    try:
                        async with page.expect_download(timeout=15000) as dl_info:
                            await btn.click()
                        download = await dl_info.value
                        anexo_file = DOWNLOADS_DIR / f"anexo_{num_safe}.pdf"
                        await download.save_as(str(anexo_file))
                        if anexo_file.exists() and anexo_file.stat().st_size > 512:
                            anexo_path = anexo_file
                            break
                    except Exception:
                        # Pode abrir nova aba
                        try:
                            async with context.expect_page(timeout=10000) as new_page_info:
                                await btn.click()
                            doc_page = await new_page_info.value
                            await doc_page.wait_for_load_state("networkidle", timeout=15000)
                            await doc_page.wait_for_timeout(1500)
                            anexo_file = DOWNLOADS_DIR / f"anexo_{num_safe}.pdf"
                            await doc_page.pdf(path=str(anexo_file), format="A4", print_background=True)
                            await doc_page.close()
                            if anexo_file.exists() and anexo_file.stat().st_size > 512:
                                anexo_path = anexo_file
                                break
                        except Exception:
                            pass
                except Exception:
                    continue
    except Exception as e:
        print(f"[pje] Erro buscando anexo: {e}")

    return mandado_path, anexo_path


# ── Versão síncrona para FastAPI ──────────────────────────────────────────────

def baixar_documentos_pje_sync(
    processos: list[dict],
    on_progress=None,
) -> list[dict]:
    """Wrapper síncrono."""
    try:
        loop = asyncio.new_event_loop()
        return loop.run_until_complete(
            baixar_documentos_pje(processos, on_progress)
        )
    except Exception as e:
        return [{"numero_processo": p.get("numero_processo", ""),
                 "ok": False, "erro": str(e)} for p in processos]
    finally:
        loop.close()

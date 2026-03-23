"""
pje_downloader.py — Baixa mandados do PJe TRF1 via Playwright CDP.
Conecta ao Chrome do usuário (já autenticado com certificado).
Usa as abas JÁ ABERTAS do PJe — não cria páginas novas.
"""
import asyncio
import base64
import os
import re
from pathlib import Path

DOWNLOADS_DIR = Path(__file__).parent / "downloads"
DOWNLOADS_DIR.mkdir(exist_ok=True)

CDP_URL = "http://127.0.0.1:9222"


async def puxar_todos_pje(grau: int = 1, on_progress=None) -> dict:
    """
    Conecta ao Chrome via CDP.
    Encontra a aba do PJe já aberta e logada.
    Lê todos os processos da tabela.
    Clica no ícone de imprimir de cada → salva como PDF.
    """
    from playwright.async_api import async_playwright

    resultados = []
    pje_host = f"pje{grau}g.trf1.jus.br"

    async with async_playwright() as pw:
        try:
            browser = await pw.chromium.connect_over_cdp(CDP_URL)
        except Exception as e:
            return {
                "ok": False,
                "erro": f"Chrome não encontrado na porta 9222. Execute iniciar_chrome_debug.bat primeiro. ({e})",
                "resultados": [],
            }

        context = browser.contexts[0]

        # Encontrar aba do PJe já aberta
        pje_page = None
        for p in context.pages:
            if pje_host in p.url and "listView" in p.url:
                pje_page = p
                break

        if not pje_page:
            # Tentar qualquer aba do PJe
            for p in context.pages:
                if pje_host in p.url:
                    pje_page = p
                    break

        if not pje_page:
            return {
                "ok": False,
                "erro": f"Nenhuma aba do PJe {grau}G encontrada. Abra o PJe e faça login primeiro.",
                "resultados": [],
            }

        if on_progress:
            on_progress({"etapa": "conectado", "msg": f"Conectado à aba PJe {grau}G!"})

        # Se não está no painel, navegar
        if "listView" not in pje_page.url:
            url_painel = f"https://{pje_host}/pje/Painel/painel_usuario/Paniel_Usuario_Oficial_Justica/listView.seam"
            await pje_page.goto(url_painel, wait_until="domcontentloaded", timeout=30000)
            await pje_page.wait_for_timeout(3000)

        # Trazer aba pro foco
        await pje_page.bring_to_front()
        await pje_page.wait_for_timeout(2000)

        if on_progress:
            on_progress({"etapa": "lendo", "msg": "Lendo processos do painel..."})

        # Extrair todos os processos da tabela
        processos_info = await pje_page.evaluate("""() => {
            const rows = document.querySelectorAll('tr');
            const re = /\\d{7}-\\d{2}\\.\\d{4}\\.\\d\\.\\d{2}\\.\\d{4}/;
            const procs = [];
            const seen = new Set();
            rows.forEach((tr) => {
                const txt = tr.innerText || '';
                const m = txt.match(re);
                if (!m || seen.has(m[0])) return;
                seen.add(m[0]);

                // Destinatário
                let dest = '';
                const dm = txt.match(/Destinat[aá]rio\\(?s?\\)?\\s*(.+?)(?=\\s{2,}|Rua |Endere|CEP|Expedi|$)/i);
                if (dm) dest = dm[1].trim();

                // Endereço (coluna com Rua/CEP)
                let end = '';
                const cells = tr.querySelectorAll('td');
                for (let i = 0; i < cells.length; i++) {
                    const ct = cells[i].innerText;
                    if (ct.match(/Rua |Av |CEP|\\d{5}-\\d{3}/i)) {
                        end = ct.trim();
                        break;
                    }
                }

                procs.push({
                    numero: m[0],
                    destinatario: dest.substring(0, 150),
                    endereco: end.substring(0, 300),
                });
            });
            return procs;
        }""")

        total = len(processos_info)
        if on_progress:
            on_progress({"etapa": "encontrados", "msg": f"{total} processos encontrados"})

        if total == 0:
            return {"ok": True, "total": 0, "baixados": 0,
                    "erro": "Nenhum processo encontrado no painel", "resultados": []}

        # Para cada processo, clicar no botão de imprimir
        for i, proc in enumerate(processos_info):
            numero = proc["numero"]
            num_safe = numero.replace("-", "_").replace(".", "_")

            if on_progress:
                on_progress({"etapa": "baixando",
                             "msg": f"[{i+1}/{total}] {numero}"})

            resultado = {
                "numero_processo": numero,
                "destinatario": proc.get("destinatario", ""),
                "endereco": proc.get("endereco", ""),
                "ok": False, "mandado": None, "erro": None,
            }

            try:
                # Contar páginas antes do clique
                pages_before = len(context.pages)

                # Clicar no primeiro ícone da linha do processo (imprimir)
                clicked = await pje_page.evaluate("""(num) => {
                    const rows = document.querySelectorAll('tr');
                    for (const tr of rows) {
                        if (!(tr.innerText || '').includes(num)) continue;
                        // Primeiro: input[type=image] (ícones do PJe são inputs de imagem)
                        const imgs = tr.querySelectorAll('input[type="image"]');
                        if (imgs.length > 0) { imgs[0].click(); return 'input_image'; }
                        // Segundo: qualquer link/botão
                        const links = tr.querySelectorAll('a, button');
                        if (links.length > 0) { links[0].click(); return 'link'; }
                        return 'no_button';
                    }
                    return 'not_found';
                }""", numero)

                if clicked in ("not_found", "no_button"):
                    resultado["erro"] = f"Botão não encontrado ({clicked})"
                    resultados.append(resultado)
                    continue

                # Esperar nova aba abrir
                await pje_page.wait_for_timeout(3000)

                # Verificar se abriu nova aba
                doc_page = None
                if len(context.pages) > pages_before:
                    # Pegar a última página aberta
                    doc_page = context.pages[-1]
                    await doc_page.wait_for_load_state("domcontentloaded", timeout=15000)
                    await doc_page.wait_for_timeout(2000)

                if doc_page and doc_page != pje_page:
                    mandado_file = DOWNLOADS_DIR / f"mandado_{num_safe}.pdf"

                    # Tentar salvar como PDF via CDP
                    try:
                        cdp = await doc_page.context.new_cdp_session(doc_page)
                        pdf_data = await cdp.send("Page.printToPDF", {
                            "printBackground": True,
                            "marginTop": 0.4,
                            "marginBottom": 0.4,
                            "marginLeft": 0.4,
                            "marginRight": 0.4,
                        })
                        await cdp.detach()

                        pdf_bytes = base64.b64decode(pdf_data["data"])
                        with open(mandado_file, "wb") as f:
                            f.write(pdf_bytes)

                        if mandado_file.exists() and mandado_file.stat().st_size > 500:
                            resultado["mandado"] = str(mandado_file)
                            resultado["ok"] = True

                    except Exception as pdf_err:
                        # Fallback: salvar HTML
                        html_file = DOWNLOADS_DIR / f"mandado_{num_safe}.html"
                        try:
                            html_content = await doc_page.content()
                            with open(html_file, "w", encoding="utf-8") as f:
                                f.write(html_content)
                            resultado["mandado"] = str(html_file)
                            resultado["ok"] = True
                            resultado["erro"] = f"Salvo como HTML ({pdf_err})"
                        except Exception as html_err:
                            resultado["erro"] = f"PDF: {pdf_err} | HTML: {html_err}"

                    # Fechar a aba do documento
                    try:
                        await doc_page.close()
                    except Exception:
                        pass

                    # Voltar foco pro painel
                    await pje_page.bring_to_front()
                    await pje_page.wait_for_timeout(1000)

                else:
                    resultado["erro"] = "Nova aba não abriu após clique"

            except Exception as e:
                resultado["erro"] = str(e)

            resultados.append(resultado)

        # NÃO fechar o browser - é o Chrome do usuário!

    return {
        "ok": True,
        "total": total,
        "baixados": sum(1 for r in resultados if r["ok"]),
        "resultados": resultados,
    }


def puxar_todos_pje_sync(grau: int = 1, on_progress=None) -> dict:
    """Wrapper síncrono."""
    loop = asyncio.new_event_loop()
    try:
        return loop.run_until_complete(puxar_todos_pje(grau, on_progress))
    except Exception as e:
        return {"ok": False, "erro": str(e), "resultados": []}
    finally:
        loop.close()


# Compatibilidade
def baixar_documentos_pje_sync(processos, on_progress=None):
    graus = set(p.get("grau", 1) for p in processos)
    all_results = []
    for g in graus:
        result = puxar_todos_pje_sync(g, on_progress)
        all_results.extend(result.get("resultados", []))
    return all_results

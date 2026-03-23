"""
puxar_todos_pdfs.py — Baixa TODOS os mandados do PJe como PDF.
Método: clica no botão "Gerar PDF" de cada mandado via Playwright CDP.
Chrome precisa estar aberto com --remote-debugging-port=9222.
"""
import asyncio
import os
import sys
import time
from pathlib import Path

os.environ.setdefault("PYTHONIOENCODING", "utf-8")

DOWNLOADS_DIR = Path(__file__).parent / "downloads"
DOWNLOADS_DIR.mkdir(exist_ok=True)
CDP_URL = "http://127.0.0.1:9222"


async def puxar_todos(grau: int = 2):
    from playwright.async_api import async_playwright

    pje_host = f"pje{grau}g.trf1.jus.br"
    resultados = []

    async with async_playwright() as pw:
        try:
            browser = await pw.chromium.connect_over_cdp(CDP_URL)
        except Exception as e:
            print(f"[ERRO] Chrome não encontrado: {e}")
            return resultados

        ctx = browser.contexts[0]

        # Encontrar aba do painel PJe
        painel = None
        for p in ctx.pages:
            if pje_host in p.url and "listView" in p.url:
                painel = p
                break

        if not painel:
            print(f"[ERRO] Aba do PJe {grau}G não encontrada!")
            return resultados

        print(f"[OK] Conectado ao PJe {grau}G: {painel.url[:60]}")
        await painel.bring_to_front()
        await painel.wait_for_timeout(2000)

        # Pegar todos os botões "Gerar PDF" e o número do processo de cada linha
        dados = await painel.evaluate(r"""() => {
            const rows = document.querySelectorAll('tr');
            const re = /\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}/;
            const procs = [];
            const seen = new Set();

            rows.forEach(tr => {
                const txt = tr.innerText || '';
                const m = txt.match(re);
                if (!m || seen.has(m[0])) return;
                seen.add(m[0]);

                const pdfBtn = tr.querySelector('a[title="Gerar PDF"]');
                if (!pdfBtn) return;

                // Endereço
                let endereco = '';
                const cells = tr.querySelectorAll('td');
                for (let c of cells) {
                    if ((c.innerText || '').match(/Rua |Av |CEP|\d{5}-\d{3}/i)) {
                        endereco = c.innerText.trim();
                        break;
                    }
                }

                // Destinatário
                let dest = '';
                const dm = txt.match(/Destinat[aá]rio\(?s?\)?\s*(.+?)(?=Expedi|$)/i);
                if (dm) dest = dm[1].trim().substring(0, 150);

                procs.push({
                    numero: m[0],
                    btnId: pdfBtn.id,
                    endereco: endereco.substring(0, 300),
                    destinatario: dest,
                });
            });
            return procs;
        }""")

        total = len(dados)
        print(f"[INFO] {total} processos encontrados\n")

        for i, proc in enumerate(dados):
            numero = proc["numero"]
            btn_id = proc["btnId"]
            pdf_path = DOWNLOADS_DIR / f"{numero}.pdf"

            print(f"[{i+1}/{total}] {numero}...", end=" ", flush=True)

            # Se já existe com tamanho razoável, pular
            if pdf_path.exists() and pdf_path.stat().st_size > 1000:
                print(f"[SKIP] já existe ({pdf_path.stat().st_size // 1024}KB)")
                resultados.append({
                    "numero": numero, "ok": True,
                    "path": str(pdf_path), "msg": "já existe",
                    "destinatario": proc.get("destinatario", ""),
                    "endereco": proc.get("endereco", ""),
                })
                continue

            try:
                # Clicar em "Gerar PDF" via JavaScript e capturar download
                async with painel.expect_download(timeout=30000) as dl_info:
                    await painel.evaluate(
                        f'document.getElementById("{btn_id}").click()'
                    )

                download = await dl_info.value
                fname = download.suggested_filename or f"{numero}.pdf"
                save_path = str(pdf_path)
                await download.save_as(save_path)

                size = os.path.getsize(save_path)
                size_kb = size // 1024
                print(f"[OK] {size_kb}KB")

                resultados.append({
                    "numero": numero, "ok": True,
                    "path": save_path, "size": size,
                    "destinatario": proc.get("destinatario", ""),
                    "endereco": proc.get("endereco", ""),
                })

                await painel.wait_for_timeout(2000)

            except Exception as e:
                print(f"[ERRO] {e}")
                resultados.append({
                    "numero": numero, "ok": False, "msg": str(e),
                    "destinatario": proc.get("destinatario", ""),
                    "endereco": proc.get("endereco", ""),
                })
                await painel.wait_for_timeout(3000)

    # Resumo
    ok = sum(1 for r in resultados if r.get("ok"))
    print(f"\n{'='*50}")
    print(f"[OK] {ok}/{total} PDFs baixados com sucesso")
    print(f"[DIR] Pasta: {DOWNLOADS_DIR}")
    return resultados


if __name__ == "__main__":
    grau = int(sys.argv[1]) if len(sys.argv) > 1 else 2
    asyncio.run(puxar_todos(grau))

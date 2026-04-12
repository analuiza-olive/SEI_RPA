import asyncio
import os
import openpyxl
from playwright.async_api import async_playwright

# ── CONFIGURAÇÃO ──────────────────────────────────────────────────────────────
URL_SEI = "https://sei4.pf.gov.br/sei/"
SESSION_FILE = "session.json"
NUMERO_DOC = "145514577"  # input("Número do documento modelo: ")
NUMERO_PROCESSO = "08455.007609/2026-11"  # input("Número do processo SEI: ")
ARQUIVO_XLSX = "DEAIN_procedimentos_20260409.xlsx"  # input("Nome do arquivo .xlsx (na pasta do projeto): ")
ARQUIVO_XLSX = os.path.join(os.path.dirname(os.path.abspath(__file__)), ARQUIVO_XLSX)


def ler_planilha(caminho):
    wb = openpyxl.load_workbook(caminho)
    ws = wb.worksheets[0]
    linhas = []
    for row in ws.iter_rows(min_row=2, values_only=True):  # pula cabeçalho
        delegado = row[1]  # coluna 2 (índice 1)
        procedimentos = row[4]  # coluna 5 (índice 4)
        if delegado or procedimentos:
            linhas.append((str(delegado or ""), str(procedimentos or "")))
    return linhas


async def substituir_no_editor(doc_page, delegado, procedimentos):
    """Substitui os marcadores no editor CKEditor do SEI e salva."""

    # 1. Localiza o frame que contém o CKEditor
    editor_frame = None
    for _ in range(30):
        for f in doc_page.frames:
            try:
                tem_ckeditor = await f.evaluate(
                    "typeof CKEDITOR !== 'undefined' && Object.keys(CKEDITOR.instances).length > 0"
                )
                if tem_ckeditor:
                    editor_frame = f
                    break
            except Exception:
                continue
        if editor_frame:
            break
        await asyncio.sleep(0.5)

    if not editor_frame:
        raise RuntimeError("Não foi possível localizar o frame com o CKEditor.")

    print(f"  ✅ CKEditor localizado.")

    # 2. Prepara os procedimentos como parágrafos HTML separados
    # Aceita separadores comuns: vírgula, ponto-e-vírgula, quebra de linha
    import re

    itens = [p.strip() for p in re.split(r"[;,\n]+", procedimentos) if p.strip()]
    # Cada procedimento vira um parágrafo na mesma classe do parágrafo original
    procedimentos_html = "".join(
        f'<p class="Texto_Justificado_Recuo_P_Linh_Esp_Simples_Calibri">- {item}</p>'
        for item in itens
    )

    # 3. Faz a substituição
    resultado = await editor_frame.evaluate(
        """([delegado, procedimentosHtml]) => {
            const editorName = Object.keys(CKEDITOR.instances)[0];
            const editor = CKEDITOR.instances[editorName];
            let content = editor.getData();
            
            const antes = {
                temNome: content.includes('[Nome do destinatário]'),
                temProc: content.includes('-XXXXXX'),
                tamanho: content.length
            };
            
            // Substitui o nome do destinatário
            content = content.split('[Nome do destinatário]').join(delegado);
            
            // Substitui o parágrafo inteiro do -XXXXXX pelos parágrafos dos procedimentos
            // Tenta primeiro substituir o parágrafo completo (mais limpo)
            const paragrafoRegex = /<p[^>]*>-XXXXXX<\\/p>/;
            if (paragrafoRegex.test(content)) {
                content = content.replace(paragrafoRegex, procedimentosHtml);
            } else {
                // Fallback: substitui só o texto
                content = content.split('-XXXXXX').join(procedimentosHtml);
            }
            
            editor.setData(content);
            
            return {
                antes,
                depois: {
                    aindaTemNome: content.includes('[Nome do destinatário]'),
                    aindaTemProc: content.includes('-XXXXXX'),
                    tamanho: content.length
                }
            };
        }""",
        [delegado, procedimentos_html],
    )

    print(f"  [DEBUG] Antes: {resultado['antes']}")
    print(f"  [DEBUG] Depois: {resultado['depois']}")

    if resultado["depois"]["aindaTemNome"] or resultado["depois"]["aindaTemProc"]:
        print("  ⚠️  Atenção: algum marcador não foi substituído!")

    await asyncio.sleep(1)

    # 4. Salva via comando do CKEditor
    await editor_frame.evaluate(
        """() => {
            const editorName = Object.keys(CKEDITOR.instances)[0];
            const editor = CKEDITOR.instances[editorName];
            editor.execCommand('save');
        }"""
    )

    await doc_page.wait_for_load_state("networkidle")
    await asyncio.sleep(1)


# ── FLUXO PRINCIPAL ───────────────────────────────────────────────────────────
async def main():
    if not os.path.exists(SESSION_FILE):
        print(f"Arquivo de sessão '{SESSION_FILE}' não encontrado.")
        print("Execute 'save_session.py' primeiro para salvar o login.")
        return

    linhas = ler_planilha(ARQUIVO_XLSX)
    linhas = linhas[:1]  # 🔧 TESTE: processa só a primeira linha
    print(f"  {len(linhas)} linha(s) encontrada(s) na planilha.\n")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=500)
        context = await browser.new_context(storage_state=SESSION_FILE)
        page = await context.new_page()

        # ── PASSO 1: ABRIR SEI COM SESSÃO SALVA ──────────────────────────────
        print("[1] Abrindo SEI com sessão salva...")
        await page.goto(URL_SEI)
        await page.wait_for_load_state("networkidle")
        print("  ✅ SEI aberto.")

        for i, (delegado, procedimentos) in enumerate(linhas, start=1):
            print(
                f"\n══ Linha {i}/{len(linhas)}: Delegado='{delegado}' | Procedimentos='{procedimentos}' ══"
            )

            # ── PASSO 2: ABRIR PROCESSO ───────────────────────────────────────
            print(f"  Buscando processo {NUMERO_PROCESSO}...")
            main_frame = page.frame("ifrConteudo") or next(
                (f for f in page.frames if "ifrConteudo" in (f.name or "")), page
            )

            await main_frame.wait_for_selector(
                'input[id="txtPesquisaRapida"], input[name="txtPesquisaRapida"]',
                state="visible",
            )
            await main_frame.fill(
                'input[id="txtPesquisaRapida"], input[name="txtPesquisaRapida"]',
                NUMERO_PROCESSO,
            )
            await main_frame.press(
                'input[id="txtPesquisaRapida"], input[name="txtPesquisaRapida"]',
                "Enter",
            )
            await page.wait_for_load_state("networkidle")

            # ── PASSO 3: INCLUIR DOCUMENTO ────────────────────────────────────
            print("  Incluindo documento...")
            main_frame = page.frame("ifrConteudo") or next(
                (f for f in page.frames if "ifrConteudo" in (f.name or "")), page
            )

            btn_selector = (
                'a[title="Incluir Documento"], '
                'img[title="Incluir Documento"], '
                'img[alt="Incluir Documento"]'
            )

            target_frame = main_frame
            try:
                await main_frame.wait_for_selector(
                    btn_selector, state="visible", timeout=5000
                )
            except Exception:
                for f in page.frames:
                    try:
                        await f.wait_for_selector(
                            btn_selector, state="visible", timeout=2000
                        )
                        target_frame = f
                        break
                    except Exception:
                        continue

            await target_frame.click(btn_selector)
            await page.wait_for_load_state("networkidle")

            # Seleciona tipo "Ofício"
            tipo_frame = None
            for f in page.frames:
                try:
                    await f.wait_for_selector(
                        "#frmDocumentoEscolherTipo", state="attached", timeout=3000
                    )
                    tipo_frame = f
                    break
                except Exception:
                    continue
            if tipo_frame is None:
                tipo_frame = page

            await tipo_frame.evaluate("escolher(11)")
            await page.wait_for_load_state("networkidle")

            # Preenche formulário
            doc_frame = None
            for f in page.frames:
                try:
                    await f.wait_for_selector(
                        'label[for="optProtocoloDocumentoTextoBase"]',
                        state="attached",
                        timeout=3000,
                    )
                    doc_frame = f
                    break
                except Exception:
                    continue
            if doc_frame is None:
                doc_frame = main_frame

            await doc_frame.click('label[for="optProtocoloDocumentoTextoBase"]')
            await doc_frame.wait_for_selector(
                'input[id="txtProtocoloDocumentoTextoBase"], input[name="txtProtocoloDocumentoTextoBase"]',
                state="visible",
            )
            await doc_frame.fill(
                'input[id="txtProtocoloDocumentoTextoBase"], input[name="txtProtocoloDocumentoTextoBase"]',
                NUMERO_DOC,
            )

            await doc_frame.click('label[for="optRestrito"]')
            await doc_frame.wait_for_selector(
                "#selHipoteseLegal", state="visible", timeout=5000
            )
            await doc_frame.select_option("#selHipoteseLegal", value="1251")
            print(f"  ✅ Formulário preenchido.")

            # Salva e captura a janela de edição
            async with context.expect_page() as new_page_info:
                await doc_frame.click("#btnSalvar")
            doc_page = await new_page_info.value
            await doc_page.wait_for_load_state("networkidle")
            await doc_page.bring_to_front()
            print("  ✅ Documento salvo. Janela de edição aberta.")

            # ── PASSO 4: SUBSTITUIR MARCADORES E SALVAR ───────────────────────
            print(f"  Substituindo marcadores no texto...")
            await substituir_no_editor(doc_page, delegado, procedimentos)
            print("  ✅ Substituição concluída e documento salvo.")

            # Fecha a janela de edição e volta para a janela principal
            await doc_page.close()
            await page.bring_to_front()
            print(f"  ✅ Linha {i} concluída.")

        print(f"\n{'═'*60}")
        print(f"  ✅ Todas as {len(linhas)} linha(s) processadas com sucesso!")
        input("\nPressione Enter para fechar o navegador...")
        await browser.close()


asyncio.run(main())

import asyncio
from playwright.async_api import async_playwright

URL_SEI_INTERNO = (
    "https://sei4.pf.gov.br/sei/" "controlador.php?acao=procedimento_workspace"
)

SESSION_FILE = "session.json"


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()

        # ⚠️ ENTRA DIRETO EM UMA AÇÃO INTERNA DO SEI
        await page.goto(URL_SEI_INTERNO)

        print("\n🔐 Faça login manualmente no navegador.")
        print("✅ Aguarde ATÉ aparecer a tela principal do SEI (workspace/lista).")
        print("👉 Somente então pressione Enter aqui no terminal...\n")

        input()

        # ✅ SALVA SESSÃO JÁ DENTRO DO SEI
        await context.storage_state(path=SESSION_FILE)

        print(f"✅ Sessão do SEI salva com sucesso em '{SESSION_FILE}'.")
        await browser.close()


asyncio.run(main())

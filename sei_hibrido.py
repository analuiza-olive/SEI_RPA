"""
02_sei_hibrido.py
─────────────────────────────────────────────────────────────────────────────
Abordagem híbrida para automação do SEI:

  • Passos 1-3 (navegar, criar documento, preencher formulário)
    → HTTP direto com requests + cookies da sessão salva
    → Sem browser, sem frames, sem flakiness

  • Passo 4 (editar CKEditor + salvar)
    → Playwright abre SOMENTE a página do editor já conhecida
    → Injeta JS, salva, fecha

Fluxo por linha da planilha:
  1. POST para criar o documento a partir do modelo (igual ao formulário do SEI)
  2. Extrai a URL do editor do HTML de resposta
  3. Playwright abre essa URL, substitui marcadores no CKEditor e salva

Requisitos:
    pip install playwright openpyxl requests lxml beautifulsoup4
    playwright install chromium

Sessão:
    O script lê o session.json salvo pelo script de login (Playwright).
    Veja a função carregar_cookies() para o formato esperado.
─────────────────────────────────────────────────────────────────────────────
"""

import asyncio
import json
import os
import re
import sys
from typing import Optional

import openpyxl
import requests
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

# ═══════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÕES
# ═══════════════════════════════════════════════════════════════════════════

BASE_URL      = "https://sei4.pf.gov.br/sei"
SESSION_FILE  = "session.json"

# Número do processo SEI destino (formatado)
NUMERO_PROCESSO = "08455.007609/2026-11"

# Número do documento modelo (usado como texto-base do ofício)
NUMERO_DOC_MODELO = "145514577"

# Hipótese legal de restrição (value do <select> no formulário SEI)
HIPOTESE_LEGAL = "1251"

# ID do tipo de documento "Ofício" no SEI
# → Se não souber, rode uma vez com DESCOBRIR_IDS = True
TIPO_DOCUMENTO_ID = "11"          # valor passado em escolher(N) no SEI

# Planilha
ARQUIVO_XLSX = "DEAIN_procedimentos_20260409.xlsx"
ARQUIVO_XLSX = os.path.join(os.path.dirname(os.path.abspath(__file__)), ARQUIVO_XLSX)

COL_DELEGADO      = 1   # índice base-0 (coluna B)
COL_PROCEDIMENTOS = 4   # índice base-0 (coluna E)

# Marcadores no documento modelo que serão substituídos
MARCADOR_DELEGADO      = "[Nome do destinatário]"
MARCADOR_PROCEDIMENTOS = "-XXXXXX"

HEADLESS = False   # True para rodar sem janela visível no passo do CKEditor

# ── Modo debug ────────────────────────────────────────────────────────────
# True  → imprime status HTTP, URL final e primeiros 3000 chars do HTML de
#         cada resposta; salva HTMLs em arquivos debug_*.html para inspeção
# False → execução silenciosa normal
DEBUG = True

# ═══════════════════════════════════════════════════════════════════════════
# DEBUG
# ═══════════════════════════════════════════════════════════════════════════

def debug_resp(label: str, resp: requests.Response) -> None:
    """Imprime diagnóstico de uma resposta HTTP e salva o HTML em arquivo."""
    if not DEBUG:
        return
    print(f"\n  ┌─ DEBUG: {label}")
    print(f"  │  status  : {resp.status_code}")
    print(f"  │  url     : {resp.url}")
    print(f"  │  redirect: {[r.url for r in resp.history] or '(nenhum)'}")
    ct = resp.headers.get("Content-Type", "")
    print(f"  │  content : {ct}")
    # Imprime os primeiros 3000 chars do HTML
    snippet = resp.text[:3000].replace("\n", " ").replace("\r", "")
    print(f"  │  html    : {snippet}")
    print(f"  └{'─'*60}")
    # Salva HTML completo em arquivo para inspeção no browser
    slug = re.sub(r"\W+", "_", label)[:40]
    path = f"debug_{slug}.html"
    with open(path, "w", encoding="utf-8") as f:
        f.write(resp.text)
    print(f"  💾 HTML salvo em '{path}'")


# ═══════════════════════════════════════════════════════════════════════════
# SESSÃO HTTP  (requests)
# ═══════════════════════════════════════════════════════════════════════════

def carregar_cookies(session_file: str) -> dict:
    """
    Lê o session.json gerado pelo Playwright e devolve um dict de cookies
    prontos para o requests.

    O Playwright salva storage_state com esta estrutura:
      {
        "cookies": [
          {"name": "...", "value": "...", "domain": "...", ...},
          ...
        ],
        "origins": [...]
      }
    """
    with open(session_file, encoding="utf-8") as f:
        state = json.load(f)

    cookies = {}
    for c in state.get("cookies", []):
        cookies[c["name"]] = c["value"]
    return cookies


def criar_sessao_http(session_file: str) -> requests.Session:
    """Cria uma requests.Session autenticada com os cookies do Playwright."""
    sess = requests.Session()
    sess.verify = False          # ajuste se o cert SSL da PF for válido
    sess.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        "Referer": BASE_URL + "/",
    })
    cookies = carregar_cookies(session_file)
    sess.cookies.update(cookies)
    return sess


# ═══════════════════════════════════════════════════════════════════════════
# PASSO 1 — Descobrir o IdProcedimento a partir do número formatado
# ═══════════════════════════════════════════════════════════════════════════

def obter_id_procedimento(sess: requests.Session, numero_processo: str) -> Optional[str]:
    """
    Faz a pesquisa rápida pelo número do processo e extrai o IdProcedimento
    da URL de redirecionamento.

    O SEI redireciona para:
      .../sei/?acao=procedimento_trabalhar&id_procedimento=XXXXXXX&...
    """
    url = BASE_URL + "/"
    params = {
        "acao": "pesquisa_rapida",
        "txtPesquisaRapida": numero_processo,
    }
    resp = sess.get(url, params=params, allow_redirects=True, timeout=20)
    debug_resp("obter_id_procedimento", resp)

    # Tenta extrair id_procedimento da URL final após redirecionamentos
    id_proc = _extrair_id_da_url(resp.url)
    if id_proc:
        return id_proc

    # Fallback: procura no HTML (às vezes o SEI retorna a página sem redirecionar)
    return _extrair_id_do_html(resp.text)


def _extrair_id_da_url(url: str) -> Optional[str]:
    m = re.search(r"id_procedimento=(\d+)", url)
    return m.group(1) if m else None


def _extrair_id_do_html(html: str) -> Optional[str]:
    """Procura id_procedimento em qualquer link da página de resultado."""
    m = re.search(r"id_procedimento=(\d+)", html)
    return m.group(1) if m else None


# ═══════════════════════════════════════════════════════════════════════════
# PASSO 2 — Criar o documento (formulário "Incluir Documento")
# ═══════════════════════════════════════════════════════════════════════════

def obter_token_formulario(sess: requests.Session, id_procedimento: str) -> Optional[str]:
    """
    Abre a página do formulário de novo documento para extrair o token
    anti-CSRF (campo oculto 'hdnToken' ou similar), se existir.
    """
    url = BASE_URL + "/"
    params = {
        "acao": "documento_escolher_tipo",
        "id_procedimento": id_procedimento,
    }
    resp = sess.get(url, params=params, timeout=20)
    debug_resp("obter_token_formulario", resp)
    soup = BeautifulSoup(resp.text, "lxml")

    # Tenta localizar token oculto — o SEI pode usar 'hdnToken' ou 'token'
    for name in ("hdnToken", "token", "_token"):
        tag = soup.find("input", {"name": name})
        if tag:
            return tag.get("value", "")
    return ""   # SEI nem sempre usa CSRF token; retorna string vazia se não achar


def criar_documento(
    sess: requests.Session,
    id_procedimento: str,
    delegado: str,
) -> Optional[str]:
    """
    Submete o formulário de criação de documento (equivalente a clicar em
    'Incluir Documento' → escolher tipo → preencher e salvar).

    Retorna a URL do editor CKEditor se bem-sucedido, ou None.
    """
    token = obter_token_formulario(sess, id_procedimento)

    url = BASE_URL + "/"
    data = {
        "acao":                          "documento_cadastrar",
        "id_procedimento":               id_procedimento,
        "sin_tipo_documento":            "G",          # G = gerado
        "id_serie":                      TIPO_DOCUMENTO_ID,
        "optTipoDocumento":              "T",          # T = texto
        "optProtocoloDocumentoTextoBase": "S",         # usar documento modelo
        "txtProtocoloDocumentoTextoBase": NUMERO_DOC_MODELO,
        "optNivelAcesso":                "1",          # restrito
        "selHipoteseLegal":              HIPOTESE_LEGAL,
        "txtDescricao":                  f"Ofício - {delegado}",
        "hdnToken":                      token,
        "sbmSalvar":                     "Salvar",
    }

    if DEBUG:
        print(f"\n  ┌─ DEBUG: criar_documento — payload POST")
        for k, v in data.items():
            print(f"  │  {k} = {v!r}")
        print(f"  └{'─'*60}")

    resp = sess.post(url, data=data, allow_redirects=True, timeout=30)
    debug_resp("criar_documento", resp)

    # O SEI responde com a página do editor ou com redirect para ela
    url_editor = _extrair_url_editor(resp)
    if url_editor:
        return url_editor

    # Tenta extrair de dentro do HTML (alguns SEIs embarcam a URL num script)
    return _extrair_url_editor_do_html(resp.text)


def _extrair_url_editor(resp: requests.Response) -> Optional[str]:
    """Procura a URL do CKEditor na URL final após redirecionamentos."""
    if "editor" in resp.url or "documento_visualizar" in resp.url:
        return resp.url
    return None


def _extrair_url_editor_do_html(html: str) -> Optional[str]:
    """
    Procura padrões como:
      window.open('...editor...', ...)
      location.href = '...editor...'
      <a href="...editor...">
    """
    padroes = [
        r"window\.open\(['\"]([^'\"]+editor[^'\"]*)['\"]",
        r"location\.href\s*=\s*['\"]([^'\"]+editor[^'\"]*)['\"]",
        r'href=["\']([^"\']+acao=editor[^"\']*)["\']',
    ]
    for p in padroes:
        m = re.search(p, html, re.IGNORECASE)
        if m:
            url = m.group(1)
            # Garante URL absoluta
            if url.startswith("http"):
                return url
            return BASE_URL + "/" + url.lstrip("/")
    return None


# ═══════════════════════════════════════════════════════════════════════════
# PASSO 3 — Editar CKEditor + salvar  (único passo que usa Playwright)
# ═══════════════════════════════════════════════════════════════════════════

async def editar_ckeditor(
    context,               # Playwright BrowserContext já autenticado
    url_editor: str,
    delegado: str,
    procedimentos: str,
) -> bool:
    """
    Abre a URL do editor, substitui os marcadores no CKEditor e salva.
    Retorna True se salvou com sucesso.
    """
    page = await context.new_page()
    try:
        await page.goto(url_editor, wait_until="domcontentloaded", timeout=30000)

        # Aguarda CKEditor inicializar
        await page.wait_for_function(
            "typeof CKEDITOR !== 'undefined' && "
            "Object.keys(CKEDITOR.instances).length > 0 && "
            "Object.values(CKEDITOR.instances)[0].status === 'ready'",
            timeout=20000,
        )

        if DEBUG:
            # Mostra quantas instâncias do CKEditor foram encontradas
            instancias = await page.evaluate(
                "Object.keys(CKEDITOR.instances)"
            )
            print(f"  [DEBUG CKEditor] instâncias encontradas: {instancias}")
            conteudo_atual = await page.evaluate(
                "Object.values(CKEDITOR.instances)[0].getData()"
            )
            print(f"  [DEBUG CKEditor] conteúdo atual (200 chars): "
                  f"{conteudo_atual[:200]!r}")

        # Substitui marcadores via JS
        await page.evaluate(
            """([marcDelegado, delegado, marcProc, procedimentos]) => {
                const editor = Object.values(CKEDITOR.instances)[0];
                let html = editor.getData();
                html = html.split(marcDelegado).join(delegado);
                html = html.split(marcProc).join(procedimentos);
                editor.setData(html);
            }""",
            [MARCADOR_DELEGADO, delegado, MARCADOR_PROCEDIMENTOS, procedimentos],
        )

        # Salva — tenta botão, fallback para atalho
        salvo = False
        for sel in ["#btnSalvar", 'button[title="Salvar"]', 'input[value="Salvar"]']:
            try:
                await page.wait_for_selector(sel, state="visible", timeout=3000)
                await page.click(sel)
                salvo = True
                break
            except PWTimeout:
                continue

        if not salvo:
            await page.keyboard.press("Control+Alt+s")

        await page.wait_for_load_state("networkidle", timeout=15000)
        return True

    except Exception as exc:
        print(f"    ✗ Erro no CKEditor: {exc}")
        return False
    finally:
        await page.close()


# ═══════════════════════════════════════════════════════════════════════════
# PLANILHA
# ═══════════════════════════════════════════════════════════════════════════

def ler_planilha(caminho: str) -> list[tuple[str, str]]:
    wb = openpyxl.load_workbook(caminho)
    ws = wb.worksheets[0]
    linhas = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        delegado      = row[COL_DELEGADO]
        procedimentos = row[COL_PROCEDIMENTOS]
        if delegado or procedimentos:
            linhas.append((str(delegado or ""), str(procedimentos or "")))
    return linhas


# ═══════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════

async def main():
    # ── Valida pré-requisitos ─────────────────────────────────────────────
    if not os.path.exists(SESSION_FILE):
        print(f"✗ Sessão não encontrada: '{SESSION_FILE}'")
        print("  Execute o script de login primeiro para gerar o arquivo de sessão.")
        sys.exit(1)

    if not os.path.exists(ARQUIVO_XLSX):
        print(f"✗ Planilha não encontrada: '{ARQUIVO_XLSX}'")
        sys.exit(1)

    # ── Carrega dados ─────────────────────────────────────────────────────
    linhas = ler_planilha(ARQUIVO_XLSX)
    print(f"  {len(linhas)} linha(s) encontrada(s) na planilha.")

    if DEBUG:
        print("\n  ⚠️  DEBUG=True — processando apenas a PRIMEIRA linha.")
        print("     Os HTMLs de resposta serão salvos em debug_*.html")
        print("     Ajuste os parâmetros do formulário e defina DEBUG=False")
        print("     para rodar em produção.\n")
        linhas = linhas[:1]
    else:
        print()

    # ── Sessão HTTP (requests) ────────────────────────────────────────────
    print("[HTTP] Criando sessão autenticada...")
    sess = criar_sessao_http(SESSION_FILE)

    # ── Descobre IdProcedimento uma única vez ─────────────────────────────
    print(f"[HTTP] Buscando IdProcedimento de '{NUMERO_PROCESSO}'...")
    id_proc = obter_id_procedimento(sess, NUMERO_PROCESSO)
    if not id_proc:
        print("  ✗ Não foi possível obter o IdProcedimento. "
              "Verifique o número do processo e a sessão.")
        sys.exit(1)
    print(f"  ✅ IdProcedimento = {id_proc}\n")

    # ── Playwright (só para o CKEditor) ──────────────────────────────────
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=HEADLESS)
        context = await browser.new_context(storage_state=SESSION_FILE)

        sucessos = 0
        falhas   = 0

        for i, (delegado, procedimentos) in enumerate(linhas, start=1):
            print(f"══ Linha {i}/{len(linhas)}: "
                  f"Delegado='{delegado}' | Procedimentos='{procedimentos}' ══")

            # Passo 2: cria documento via HTTP
            print("  [HTTP] Criando documento...")
            url_editor = criar_documento(sess, id_proc, delegado)

            if not url_editor:
                print("  ✗ Não foi possível obter a URL do editor. "
                      "Pulando linha.")
                falhas += 1
                continue

            print(f"  ✅ Documento criado. Editor: {url_editor[:80]}...")

            # Passo 3: edita CKEditor via Playwright
            print("  [Playwright] Editando CKEditor...")
            ok = await editar_ckeditor(context, url_editor, delegado, procedimentos)

            if ok:
                print(f"  ✅ Linha {i} concluída.")
                sucessos += 1
            else:
                print(f"  ✗ Linha {i} falhou no CKEditor.")
                falhas += 1

        print(f"\n{'═'*60}")
        print(f"  Concluído: {sucessos} criado(s), {falhas} falha(s).")

        await browser.close()


if __name__ == "__main__":
    # Suprime warnings de SSL (requests com verify=False)
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    asyncio.run(main())

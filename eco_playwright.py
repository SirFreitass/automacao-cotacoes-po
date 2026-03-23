"""
eco_playwright.py
-----------------
Automação do ECO Requisition para criação de POs via Playwright.

Parâmetro `confirmar`:
  - callable(titulo, mensagem) → True (prosseguir) / False (cancelar)
  - None = modo autônomo sem pausas (usar após testes aprovados)
"""

import asyncio
import json
import os
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

from config import SHIP_VIA_MAP, GL_CODE_PLANILHA

ECO_URL         = "https://requisition.chouest.com"
TIMEOUT         = 30_000
SHORT_WAIT      = 5_000
VENDOR_MAP_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vendor_map.json")


# ─────────────────────────────────────────────────────────────────────
# Memória de fornecedores — persiste entre execuções
# ─────────────────────────────────────────────────────────────────────

def _carregar_vendor_map() -> dict:
    """Lê o mapa de fornecedores salvo (nome_buscado → texto_opção_ECO)."""
    try:
        with open(VENDOR_MAP_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _salvar_vendor_map(vendor_map: dict):
    """Persiste o mapa de fornecedores em disco."""
    try:
        with open(VENDOR_MAP_FILE, "w", encoding="utf-8") as f:
            json.dump(vendor_map, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────

def _numero_cotacao(analise: dict) -> str:
    for forn in analise.get("ranking_preco", []):
        nc = forn.get("numero_cotacao")
        if nc:
            return str(nc)
    return ""


def _termo_busca_vendor(nome: str) -> str:
    """
    Retorna o termo de busca para o autocomplete de vendor no ECO.
    Usa apenas as primeiras 2 palavras com 3+ caracteres para maximizar
    as chances de encontrar autocomplete (ex: 'Master Control Systems Inc'
    → 'Master Control').
    """
    if not nome:
        return ""
    palavras_sig = [w for w in nome.split() if len(w) >= 3]
    if not palavras_sig:
        return nome
    return " ".join(palavras_sig[:2])


def _carregar_vessels() -> dict:
    """
    Lê a planilha 'Brazil Vessels - GL CODE.xlsx' e retorna
    {NOME_EMBARCAÇÃO_MAIUSCULO: gl_code_str}.
    A planilha tem múltiplos grupos de colunas lado a lado:
      Grupo 1 : col 0 = nome, col 1 = GL
      Grupos N: col 3N = nome, col 3N+2 = GL  (N = 1, 2, 3, ...)
    """
    from openpyxl import load_workbook
    vessels = {}
    try:
        wb = load_workbook(GL_CODE_PLANILHA, read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            # Grupo 1: col A = nome, col B = GL
            if row[0] and row[1]:
                vessels[str(row[0]).upper().strip()] = str(int(row[1]))
            # Grupos seguintes: col 3, 6, 9, ... = nome; col 5, 8, 11, ... = GL
            col = 3
            while col + 2 < len(row):
                nome = row[col]
                gl   = row[col + 2]
                if nome and gl:
                    vessels[str(nome).upper().strip()] = str(int(gl))
                col += 3
        wb.close()
    except Exception:
        pass
    return vessels


async def _ok(confirmar, titulo: str, mensagem: str) -> bool:
    """
    Pausa para confirmação do usuário. Se confirmar=None, retorna True direto.
    Roda o callback bloqueante em executor para não travar o event loop.
    """
    if confirmar is None:
        return True
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, confirmar, titulo, mensagem)


async def _escolher_async(escolher, titulo: str, opcoes: list) -> int:
    """
    Exibe lista de opções numeradas e retorna o índice (0-based) escolhido.
    Se escolher=None (modo autônomo), retorna 0 (primeira opção).
    Retorna -1 se o usuário cancelar.
    """
    if escolher is None:
        return 0
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, escolher, titulo, opcoes)


# ─────────────────────────────────────────────────────────────────────
# Login — confirmação em cada sub-passo
# ─────────────────────────────────────────────────────────────────────

async def _login(page, usuario: str, senha: str, confirmar):

    # 1. Navegar para login
    if not await _ok(confirmar,
                     "LOGIN — Passo 1/5: Navegar",
                     f"Vai abrir a URL:\n{ECO_URL}/login\n\nProsseguir?"):
        raise RuntimeError("Cancelado pelo usuário (navegar para login)")
    await page.goto(f"{ECO_URL}/login")

    # 2. Aguardar campo usuário
    if not await _ok(confirmar,
                     "LOGIN — Passo 2/5: Aguardar campo usuário",
                     "Aguardando o campo '#username' ficar visível.\n\nProsseguir?"):
        raise RuntimeError("Cancelado pelo usuário (aguardar campo usuário)")
    await page.wait_for_selector("#username", timeout=TIMEOUT)

    # 3. Preencher usuário
    if not await _ok(confirmar,
                     "LOGIN — Passo 3/5: Preencher usuário",
                     f"Vai clicar no campo '#username' e digitar:\n'{usuario}'\n\nProsseguir?"):
        raise RuntimeError("Cancelado pelo usuário (preencher usuário)")
    await page.fill("#username", usuario)

    # 4. Preencher senha
    # O campo de senha é o input dentro do SEGUNDO kendo-formfield da página
    if not await _ok(confirmar,
                     "LOGIN — Passo 4/5: Preencher senha",
                     "Vai localizar o campo de senha via seletor:\n"
                     "'kendo-formfield' → nth(1) → 'input' (segundo bloco de formulário)\n\n"
                     "e preencher com a senha informada.\n\nProsseguir?"):
        raise RuntimeError("Cancelado pelo usuário (preencher senha)")
    pw = page.locator("kendo-formfield").nth(1).locator("input").first
    await pw.fill(senha)

    # 5. Clicar no botão de login
    if not await _ok(confirmar,
                     "LOGIN — Passo 5/5: Clicar em Entrar",
                     "Vai clicar no botão de login via seletor:\n"
                     "'button[type=\"submit\"]' (primeiro)\n\n"
                     "Verifique no navegador se o botão correto está destacado.\n\nProsseguir?"):
        raise RuntimeError("Cancelado pelo usuário (clicar login)")
    await page.locator("button[type='submit']").first.click()

    # Aguarda redirecionamento
    await page.wait_for_url(f"{ECO_URL}/**", timeout=TIMEOUT)
    await page.wait_for_load_state("networkidle", timeout=TIMEOUT)

    if not await _ok(confirmar,
                     "LOGIN — Concluído",
                     f"Login realizado. URL atual:\n{page.url}\n\n"
                     "A tela carregada está correta?"):
        raise RuntimeError("Usuário indicou que o login não funcionou corretamente")


# ─────────────────────────────────────────────────────────────────────
# Criação de PO — confirmação em cada sub-passo
# ─────────────────────────────────────────────────────────────────────

async def _criar_po_par(page, par: dict, vessels: dict, confirmar, escolher, vendor_map: dict) -> dict:
    analise       = par["analise"]
    po_data       = analise.get("po", {})
    melhor        = analise.get("melhor_preco") or {}
    numero_cot    = _numero_cotacao(analise)
    numero_po     = po_data.get("numero_po") or "?"
    centro_custo  = (po_data.get("centro_de_custo") or po_data.get("solicitante") or "").strip()
    forn_extraido = (po_data.get("fornecedor_escolhido_comentario") or "").strip()
    # NÃO usar melhor.get("nome") — pode ser fornecedor diferente do item atual
    fornecedor_eco = forn_extraido or ""
    itens_po      = po_data.get("itens") or []

    if not numero_cot:
        return {"po": numero_po, "status": "ERRO",
                "mensagem": "Número de cotação não encontrado na análise"}

    try:
        # ── A. Navegar para o histórico ─────────────────────────────────
        if not await _ok(confirmar,
                         f"PO {numero_po} — A: Navegar para histórico",
                         f"Vai abrir:\n{ECO_URL}/requisition/history\n\nProsseguir?"):
            return {"po": numero_po, "status": "CANCELADO", "mensagem": "Cancelado (navegar histórico)"}
        await page.goto(f"{ECO_URL}/requisition/history")

        # ── B. Aguardar tabela carregar ──────────────────────────────────
        if not await _ok(confirmar,
                         f"PO {numero_po} — B: Aguardar tabela",
                         "Aguardando a tabela de histórico carregar.\n"
                         "Seletor: 'kendo-grid-list table tbody tr'\n\nProsseguir?"):
            return {"po": numero_po, "status": "CANCELADO", "mensagem": "Cancelado (aguardar tabela)"}
        search_box = page.locator("kendo-textbox input").first
        await search_box.wait_for(state="visible", timeout=TIMEOUT)
        await page.wait_for_selector("kendo-grid-list table tbody tr", timeout=TIMEOUT)
        await page.wait_for_timeout(1500)

        # ── C. Preencher campo de busca ──────────────────────────────────
        # Garante o formato XXXX.XXXXXX — insere ponto na posição 4 SOMENTE
        # para números de exatamente 10 dígitos (ex: "2025039982" → "2025.039982").
        # Outros formatos (ex: ECO REQ com 7 dígitos) são usados como estão.
        if "." not in numero_cot and len(numero_cot) == 10 and numero_cot.isdigit():
            numero_busca = numero_cot[:4] + "." + numero_cot[4:]
        else:
            numero_busca = numero_cot

        if not await _ok(confirmar,
                         f"PO {numero_po} — C: Pesquisar cotação",
                         f"Número extraído   : {numero_cot}\n"
                         f"Número formatado  : {numero_busca}\n\n"
                         "Seletor: 'kendo-textbox input' (primeiro)\n\nProsseguir?"):
            return {"po": numero_po, "status": "CANCELADO", "mensagem": "Cancelado (preencher busca)"}
        # Usa press_sequentially (digita caractere a caractere) em vez de fill()
        # para garantir que o ponto seja digitado corretamente no Angular
        await search_box.click()
        await search_box.press("Control+a")   # seleciona qualquer texto existente
        await search_box.press_sequentially(numero_busca, delay=50)

        # Aguarda a grade mostrar pelo menos uma célula com o número pesquisado
        # (não usa timeout fixo — só avança quando o ECO realmente filtrou)
        if not await _ok(confirmar,
                         f"PO {numero_po} — C: Aguardando resultados",
                         f"Campo preenchido com '{numero_busca}'.\n\n"
                         "Aguardando a tabela filtrar e exibir resultados contendo este número.\n"
                         f"Seletor: td:has-text('{numero_busca}')\n\nProsseguir para aguardar?"):
            return {"po": numero_po, "status": "CANCELADO", "mensagem": "Cancelado (aguardar filtro)"}

        try:
            await page.wait_for_selector(
                f"kendo-grid-list table tbody tr td:has-text('{numero_busca}')",
                timeout=TIMEOUT,
            )
        except PWTimeout:
            return {
                "po": numero_po, "status": "ERRO",
                "mensagem": f"Nenhum resultado encontrado para '{numero_busca}' no histórico do ECO"
            }

        # ── D. Ler resultados ────────────────────────────────────────────
        result_cells = page.locator("td[role='gridcell'][aria-colindex='3']")
        await result_cells.first.wait_for(state="visible", timeout=TIMEOUT)
        count = await result_cells.count()

        linhas = []
        for idx in range(min(count, 8)):
            try:
                txt = (await result_cells.nth(idx).inner_text()).strip()
                linhas.append(f"  [{idx+1}] {txt}")
            except Exception:
                linhas.append(f"  [{idx+1}] (erro ao ler)")

        if not await _ok(confirmar,
                         f"PO {numero_po} — D: Resultados da busca",
                         f"Cotação pesquisada: {numero_cot}\n"
                         f"Centro de custo: {centro_custo or '(não informado)'}\n\n"
                         f"{count} resultado(s) encontrado(s):\n" +
                         "\n".join(linhas) +
                         ("\n  ..." if count > 8 else "") +
                         "\n\nProsseguir para abrir a requisição?"):
            return {"po": numero_po, "status": "CANCELADO", "mensagem": "Cancelado (ver resultados)"}

        # ── E. Selecionar a requisição ───────────────────────────────────
        if count > 1 and centro_custo:
            abriu = False
            linha_escolhida = 0
            for idx in range(count):
                txt = await result_cells.nth(idx).inner_text()
                if centro_custo.upper() in txt.upper():
                    linha_escolhida = idx + 1
                    if not await _ok(confirmar,
                                     f"PO {numero_po} — E: Abrir requisição",
                                     f"Encontrado resultado com centro de custo '{centro_custo}' "
                                     f"na linha {linha_escolhida}.\n\n"
                                     "Vai clicar no botão desta linha.\n\nProsseguir?"):
                        return {"po": numero_po, "status": "CANCELADO", "mensagem": "Cancelado (abrir req)"}
                    row = result_cells.nth(idx).locator("xpath=..")
                    await row.locator("button").click()
                    abriu = True
                    break
            if not abriu:
                if not await _ok(confirmar,
                                 f"PO {numero_po} — E: Abrir requisição (fallback)",
                                 f"Centro de custo '{centro_custo}' não encontrado nos resultados.\n\n"
                                 "Vai clicar no botão da PRIMEIRA linha como fallback.\n"
                                 "Seletor: 'td[role=gridcell][aria-colindex=8] button' (primeiro)\n\nProsseguir?"):
                    return {"po": numero_po, "status": "CANCELADO", "mensagem": "Cancelado (abrir req fallback)"}
                await page.locator("td[role='gridcell'][aria-colindex='8'] button").first.click()
        else:
            if not await _ok(confirmar,
                             f"PO {numero_po} — E: Abrir requisição",
                             f"{count} resultado(s). Vai clicar no botão da PRIMEIRA linha.\n"
                             "Seletor: 'td[role=gridcell][aria-colindex=8] button' (primeiro)\n\nProsseguir?"):
                return {"po": numero_po, "status": "CANCELADO", "mensagem": "Cancelado (abrir req)"}
            await page.locator("td[role='gridcell'][aria-colindex='8'] button").first.click()

        # ── F. Aguardar botões Order ─────────────────────────────────────
        if not await _ok(confirmar,
                         f"PO {numero_po} — F: Aguardar formulário",
                         "Aguardando os botões 'Order' carregarem.\n"
                         "Seletor: 'button.order.grn.card-1'\n\nProsseguir?"):
            return {"po": numero_po, "status": "CANCELADO", "mensagem": "Cancelado (aguardar Order buttons)"}
        order_btns = page.locator("button.order.grn.card-1")
        await order_btns.first.wait_for(state="visible", timeout=TIMEOUT)
        await page.wait_for_timeout(1500)

        total_btns = await order_btns.count()

        if not await _ok(confirmar,
                         f"PO {numero_po} — F: Botões encontrados",
                         f"{total_btns} botão(ões) 'Order' encontrado(s) na página.\n\n"
                         "A quantidade está correta?\nProsseguir para processar cada um?"):
            return {"po": numero_po, "status": "CANCELADO", "mensagem": "Cancelado (conferir botões Order)"}

        itens_criados = 0
        po_gerado = numero_po  # valor padrão; atualizado pelo popup de confirmação

        for k in range(total_btns):
            # Os botões NÃO desaparecem do DOM após serem processados — apenas mudam
            # de texto ("Order" → "Order (1)"). Por isso usamos nth(k) para acessar
            # o botão correto de cada linha, independente do estado dos anteriores.
            btn = order_btns.nth(k)
            texto = (await btn.inner_text()).strip()
            ativo = await btn.is_enabled()

            if texto != "Order" or not ativo:
                # Botão já processado ("Order (1)") ou desabilitado — pular silenciosamente
                await _ok(confirmar,
                          f"PO {numero_po} — Botão {k+1}/{total_btns}: Pular",
                          f"Botão {k+1}: texto='{texto}'\n\n"
                          "Este item já possui PO gerada ou está desabilitado.\n"
                          "Pulando para o próximo. Clique OK para continuar.")
                continue

            # Variáveis de estado do item — inicializadas aqui para evitar
            # UnboundLocalError se algum passo intermediário for pulado.
            vendor_input = None

            # ── G. Clicar em Order ───────────────────────────────────────
            if not await _ok(confirmar,
                             f"PO {numero_po} — G: Clicar Order {k+1}/{total_btns}",
                             f"Vai rolar até o botão {k+1} e clicar nele.\n"
                             f"Fornecedor que será usado: '{fornecedor_eco}'\n\nProsseguir?"):
                continue   # pula este item, continua os outros

            await btn.scroll_into_view_if_needed()
            await page.wait_for_timeout(800)
            await btn.click()
            await page.wait_for_timeout(1000)

            # ── H. Popup de confirmação (opcional) ───────────────────────
            try:
                confirm_popup = page.locator("kendo-popup button").first
                await confirm_popup.wait_for(state="visible", timeout=SHORT_WAIT)
                popup_txt = (await confirm_popup.inner_text()).strip()
                if not await _ok(confirmar,
                                 f"PO {numero_po} — H: Popup de confirmação",
                                 f"Apareceu um popup com botão: '{popup_txt}'\n"
                                 "Seletor: 'kendo-popup button' (primeiro)\n\n"
                                 "Vai clicar neste botão. Prosseguir?"):
                    continue
                await confirm_popup.click()
                await page.wait_for_timeout(1000)
            except PWTimeout:
                if not await _ok(confirmar,
                                 f"PO {numero_po} — H: Sem popup",
                                 "Nenhum popup de confirmação apareceu (normal em alguns casos).\n\n"
                                 "Prosseguir para o formulário do item?"):
                    continue

            await page.wait_for_timeout(1500)

            # ── H2. Identificar item pelo campo Description do formulário ─
            # Lê a descrição já preenchida no modal para encontrar o item
            # correspondente em itens_po e usar seu fornecedor_item específico.
            fornecedor_eco_item = fornecedor_eco  # fallback = vendor geral da PO
            try:
                desc_field = page.get_by_label("Description", exact=False).first
                desc_valor = (await desc_field.input_value()).strip()
                if desc_valor:
                    desc_norm = desc_valor.lower()
                    for it in itens_po:
                        desc_it = (it.get("descricao") or "").lower()
                        pn_it   = (it.get("pn_fornecedor") or it.get("pn") or "").lower()
                        if (desc_it and desc_it[:30] in desc_norm) or (pn_it and pn_it in desc_norm):
                            forn_it = (it.get("fornecedor_item") or "").strip()
                            if forn_it:
                                fornecedor_eco_item = forn_it
                            break
            except Exception:
                pass  # usa fornecedor_eco_item = fornecedor_eco

            # ── I. Preencher fornecedor ──────────────────────────────────
            if not await _ok(confirmar,
                             f"PO {numero_po} — I: Campo fornecedor",
                             f"Vai localizar o campo de fornecedor:\n"
                             f"Seletor: 'req-vendor input' (primeiro)\n\n"
                             f"e digitar caractere a caractere: '{fornecedor_eco_item}'\n\nProsseguir?"):
                continue

            vendor_input = page.locator("req-vendor input").first
            await vendor_input.wait_for(state="visible", timeout=TIMEOUT)
            # Usa termo reduzido (primeiras 2 palavras sig.) para o autocomplete ECO.
            # O nome completo é usado como chave do vendor_map para persistência.
            termo_busca = _termo_busca_vendor(fornecedor_eco_item)
            # press_sequentially digita char a char, acionando o autocomplete Angular
            await vendor_input.click()
            await vendor_input.press("Control+a")
            await vendor_input.press_sequentially(termo_busca, delay=80)

            # ── J. Aguardar e selecionar autocomplete ────────────────────
            chave_forn = fornecedor_eco_item.lower().strip()
            try:
                await page.wait_for_selector(
                    "mat-option[role='option']", timeout=SHORT_WAIT
                )
                opts = page.locator("mat-option[role='option']")
                n_opts = await opts.count()
                if n_opts > 0:
                    # Lê os textos de todas as opções
                    opcoes_txt = []
                    for oi in range(n_opts):
                        try:
                            opcoes_txt.append((await opts.nth(oi).inner_text()).strip())
                        except Exception:
                            opcoes_txt.append(f"(opção {oi+1})")

                    opcao_salva = vendor_map.get(chave_forn)
                    idx_escolhido = -1

                    if opcao_salva and opcao_salva in opcoes_txt:
                        # Já temos a escolha salva — usa diretamente
                        idx_escolhido = opcoes_txt.index(opcao_salva)
                        await _ok(confirmar,
                                  f"PO {numero_po} — J: Fornecedor (memória)",
                                  f"Usando seleção salva para '{fornecedor_eco_item}':\n\n"
                                  f"  [{idx_escolhido+1}] {opcao_salva}\n\n"
                                  "Clique OK para confirmar.")
                    elif n_opts == 1:
                        # Apenas uma opção — seleciona sem perguntar
                        idx_escolhido = 0
                        await _ok(confirmar,
                                  f"PO {numero_po} — J: Autocomplete (única opção)",
                                  f"Uma única sugestão encontrada:\n  {opcoes_txt[0]}\n\n"
                                  "Vai selecionar esta opção. Clique OK.")
                    else:
                        # Múltiplas opções sem mapeamento salvo — pergunta ao usuário
                        idx_escolhido = await _escolher_async(
                            escolher,
                            f"PO {numero_po} — Selecionar fornecedor",
                            opcoes_txt,
                        )
                        if idx_escolhido >= 0:
                            # Salva a escolha para próximas execuções
                            vendor_map[chave_forn] = opcoes_txt[idx_escolhido]
                            _salvar_vendor_map(vendor_map)

                    if idx_escolhido >= 0:
                        await opts.nth(idx_escolhido).click()
                        await page.wait_for_timeout(500)
                    else:
                        await _ok(confirmar,
                                  f"PO {numero_po} — J: Seleção cancelada",
                                  "Nenhuma opção selecionada. Continuando sem fornecedor.")
                else:
                    await _ok(confirmar,
                              f"PO {numero_po} — J: Sem autocomplete",
                              "Nenhuma sugestão apareceu para este fornecedor.\n"
                              "Verifique se o nome está cadastrado no ECO.\n\n"
                              "Clique OK para continuar sem selecionar.")
            except PWTimeout:
                await _ok(confirmar,
                          f"PO {numero_po} — J: Autocomplete não apareceu",
                          f"Timeout aguardando sugestões para '{fornecedor_eco_item}'.\n"
                          "O nome pode estar incorreto ou não cadastrado.\n\n"
                          "Clique OK para continuar sem selecionar.")

            # ── K. GL Code ───────────────────────────────────────────────
            # Pequena pausa após vendor para o Angular processar o formulário
            await page.wait_for_timeout(800)
            gl_sel = page.locator(
                "mat-select[role='combobox'][formcontrolname='itemCodeId']"
            )
            try:
                await gl_sel.wait_for(state="visible", timeout=3000)
                gl_txt = (await gl_sel.inner_text()).strip()
                if gl_txt:
                    await _ok(confirmar,
                              f"PO {numero_po} — K: GL Code já preenchido",
                              f"O campo GL Code já contém: '{gl_txt}'\nNão será alterado.\n\nClique OK para continuar.")
                else:
                    gl_code = vessels.get(centro_custo.upper()) if centro_custo else None
                    if gl_code:
                        if not await _ok(confirmar,
                                         f"PO {numero_po} — K: Preencher GL Code",
                                         f"GL Code está vazio.\n"
                                         f"Centro de custo: '{centro_custo}'\n"
                                         f"GL Code encontrado: '{gl_code}'\n\n"
                                         "Vai clicar no campo GL Code e digitar o código.\n\nProsseguir?"):
                            pass  # continua sem preencher
                        else:
                            await gl_sel.click()
                            await page.wait_for_timeout(500)
                            await page.keyboard.type(str(gl_code))
                            await page.wait_for_timeout(500)
                            await page.keyboard.press("Enter")
                            await page.wait_for_timeout(500)
                    else:
                        await _ok(confirmar,
                                  f"PO {numero_po} — K: GL Code não encontrado",
                                  f"GL Code vazio e centro de custo '{centro_custo}' não "
                                  "encontrado na aba 'vessels'.\nCampo será deixado em branco.\n\nClique OK para continuar.")
            except PWTimeout:
                pass

            # ── L. Preço ─────────────────────────────────────────────────
            preco_preenchido = ""
            item_casado = None  # inicializa aqui para uso em L2 (Qty)
            price_input = page.locator("input[formcontrolname='price']")
            try:
                await price_input.wait_for(state="visible", timeout=3000)
                current = await price_input.input_value()
                if current:
                    await _ok(confirmar,
                              f"PO {numero_po} — L: Preço já preenchido",
                              f"O campo de preço já contém: '{current}'\nNão será alterado.\n\nClique OK para continuar.")
                    preco_preenchido = current
                else:
                    desc_input = page.locator("input[formcontrolname='description']")
                    desc_eco = (await desc_input.input_value()).lower()
                    preco_encontrado = None
                    item_casado = None
                    for item in itens_po:
                        desc_po = (item.get("descricao") or "").lower()
                        pn_po   = (item.get("pn") or "").lower()
                        if (desc_eco[:20] in desc_po or desc_po[:20] in desc_eco or
                                (pn_po and pn_po in desc_eco)):
                            preco_encontrado = item.get("preco_unitario")
                            item_casado = item
                            break
                    if preco_encontrado is not None:
                        if not await _ok(confirmar,
                                         f"PO {numero_po} — L: Preencher preço",
                                         f"Campo de preço está vazio.\n"
                                         f"Descrição no ECO: '{desc_eco[:60]}'\n"
                                         f"Item casado (PO): '{(item_casado.get('descricao') or '')[:60]}'\n"
                                         f"Preço a preencher: {preco_encontrado}\n\n"
                                         "Vai digitar este preço. Prosseguir?"):
                            pass
                        else:
                            await price_input.fill(str(preco_encontrado))
                            preco_preenchido = str(preco_encontrado)
                    else:
                        await _ok(confirmar,
                                  f"PO {numero_po} — L: Preço não encontrado",
                                  f"Não foi possível casar nenhum item da PO com a descrição:\n'{desc_eco[:80]}'\n\n"
                                  "O campo de preço será deixado em branco.\nClique OK para continuar.")
            except PWTimeout:
                await _ok(confirmar,
                          f"PO {numero_po} — L: Campo preço não encontrado",
                          "O campo 'input[formcontrolname=price]' não apareceu.\n\n"
                          "Clique OK para continuar sem preencher o preço.")

            # ── L2. Quantidade ───────────────────────────────────────────
            qtd_preenchida = ""
            # Tenta primeiro por label (mais robusto com Angular)
            qty_input = None
            try:
                by_label = page.get_by_label("Qty", exact=False)
                await by_label.first.wait_for(state="visible", timeout=2000)
                qty_input = by_label.first
            except PWTimeout:
                pass

            # Fallback: tenta formcontrolnames conhecidos
            if qty_input is None:
                for _sel in ["input[formcontrolname='qty']",
                             "input[formcontrolname='quantity']",
                             "input[formcontrolname='qtyOrdered']",
                             "input[formcontrolname='requestedQty']",
                             "input[formcontrolname='orderedQty']"]:
                    try:
                        candidate = page.locator(_sel)
                        await candidate.wait_for(state="visible", timeout=800)
                        qty_input = candidate
                        break
                    except PWTimeout:
                        continue

            if qty_input:
                current_qty = await qty_input.input_value()
                qtd_esperada = None
                if item_casado:
                    q = item_casado.get("quantidade")
                    qtd_esperada = str(int(q)) if q is not None else None

                if qtd_esperada and current_qty != qtd_esperada:
                    if not await _ok(confirmar,
                                     f"PO {numero_po} — L2: Quantidade divergente",
                                     f"Campo Qty atual : {current_qty or '(vazio)'}\n"
                                     f"Quantidade da PO: {qtd_esperada}\n\n"
                                     "Vai substituir pelo valor da PO. Prosseguir?"):
                        qtd_preenchida = current_qty
                    else:
                        await qty_input.triple_click()
                        await qty_input.fill(qtd_esperada)
                        qtd_preenchida = qtd_esperada
                else:
                    qtd_preenchida = current_qty
                    await _ok(confirmar,
                              f"PO {numero_po} — L2: Quantidade OK",
                              f"Campo Qty: '{current_qty}'"
                              + (f" — bate com PO ({qtd_esperada})" if qtd_esperada else " — sem referência")
                              + "\n\nClique OK para continuar.")
            # Se nenhum seletor encontrou o campo, continua sem alterar (não bloqueia)

            # ── L3. Ship VIA ─────────────────────────────────────────────
            tipo_freight   = (melhor.get("tipo_freight") or "").strip()
            ship_via_alvo  = SHIP_VIA_MAP.get(tipo_freight.lower())

            if ship_via_alvo:
                if not await _ok(confirmar,
                                 f"PO {numero_po} — L3: Ship VIA",
                                 f"Tipo de frete (cotação): '{tipo_freight}'\n"
                                 f"Ship VIA a selecionar  : '{ship_via_alvo}'\n\n"
                                 "Vai localizar o campo Ship VIA e selecionar esta opção.\n\nProsseguir?"):
                    pass  # pula, não bloqueia
                else:
                    try:
                        # 1ª tentativa: get_by_label (mais robusto, independe de estrutura)
                        sv = page.get_by_label("Ship VIA", exact=False)
                        await sv.wait_for(state="visible", timeout=2000)
                        tag = await sv.evaluate("el => el.tagName.toLowerCase()")
                        if tag == "select":
                            await sv.select_option(label=ship_via_alvo)
                        else:
                            # mat-select: clica no trigger e escolhe mat-option
                            await sv.click()
                            await page.wait_for_timeout(400)
                            opt = page.locator("mat-option").filter(has_text=ship_via_alvo).first
                            await opt.wait_for(state="visible", timeout=3000)
                            await opt.click()
                        await page.wait_for_timeout(400)
                    except Exception:
                        try:
                            # Fallback: mat-form-field com label "Ship VIA"
                            sv_field = page.locator("mat-form-field").filter(
                                has=page.locator("mat-label, label", has_text="Ship VIA")
                            ).first
                            sv_trigger = sv_field.locator("mat-select, select").first
                            await sv_trigger.wait_for(state="visible", timeout=3000)
                            tag2 = await sv_trigger.evaluate("el => el.tagName.toLowerCase()")
                            if tag2 == "select":
                                await sv_trigger.select_option(label=ship_via_alvo)
                            else:
                                await sv_trigger.click()
                                await page.wait_for_timeout(400)
                                opt2 = page.locator("mat-option").filter(has_text=ship_via_alvo).first
                                await opt2.wait_for(state="visible", timeout=3000)
                                await opt2.click()
                            await page.wait_for_timeout(400)
                        except Exception as _sv_err:
                            await _ok(confirmar,
                                      f"PO {numero_po} — L3: Ship VIA não preenchido",
                                      f"Não foi possível selecionar '{ship_via_alvo}'.\n"
                                      f"Erro: {str(_sv_err)[:120]}\n\n"
                                      "Preencha manualmente e clique OK para continuar.")
            else:
                await _ok(confirmar,
                          f"PO {numero_po} — L3: Ship VIA — preencher manualmente",
                          f"Tipo de frete '{tipo_freight or '(não informado)'}' não tem mapeamento "
                          "automático para Ship VIA.\n\n"
                          "Opções disponíveis no ECO:\n"
                          "  • ECO UPS ACCT# 707185\n"
                          "  • Free delivery\n"
                          "  • Runner Pick up\n"
                          "  • Supplier Ship\n\n"
                          "Preencha manualmente no navegador e clique OK para continuar.")

            # ── M. Confirmar e submeter ──────────────────────────────────
            vendor_atual = fornecedor_eco_item
            try:
                if vendor_input is not None:
                    vendor_atual = await vendor_input.input_value()
            except Exception:
                pass  # mantém fornecedor_eco como fallback

            if not await _ok(confirmar,
                             f"PO {numero_po} — M: SUBMETER item {k+1}/{total_btns}",
                             f"RESUMO DO ITEM A SUBMETER:\n\n"
                             f"  Fornecedor : {vendor_atual or '(vazio)'}\n"
                             f"  Quantidade : {qtd_preenchida or '(não alterada)'}\n"
                             f"  Preço      : {preco_preenchido or '(vazio / já preenchido)'}\n"
                             f"  Centro     : {centro_custo or '(não informado)'}\n\n"
                             "Clique SIM para salvar ou NÃO para fechar sem salvar."):
                # Fecha o modal sem salvar (tecla Escape ou botão ×)
                try:
                    await page.keyboard.press("Escape")
                    await page.wait_for_timeout(1000)
                except Exception:
                    pass
                continue

            submit = page.get_by_role("button", name="Save")
            await submit.wait_for(state="visible", timeout=TIMEOUT)
            await submit.click()

            # Aguarda popup de confirmação "purchase order have been generated successfully"
            # e clica em Close para fechar
            try:
                close_btn = page.get_by_role("button", name="Close")
                await close_btn.wait_for(state="visible", timeout=10000)
                # Tenta capturar o número da PO do popup
                try:
                    popup_txt = await page.locator("text=/PO Number:/").inner_text(timeout=2000)
                    po_gerado_popup = popup_txt.split("PO Number:")[-1].strip().split()[0]
                    if po_gerado_popup:
                        po_gerado = po_gerado_popup
                except Exception:
                    pass
                await _ok(confirmar,
                          f"PO {numero_po} — Item {k+1} salvo",
                          f"PO gerada com sucesso!\n\n"
                          f"  PO Number : {po_gerado}\n\n"
                          "Clique OK para fechar o popup e continuar.")
                await close_btn.click()
                await page.wait_for_timeout(1000)
            except PWTimeout:
                # Popup não apareceu — continua normalmente
                await page.wait_for_timeout(1500)

            itens_criados += 1

        # ── N. Capturar número da PO gerada ─────────────────────────────
        try:
            po_span = page.locator("app-requisition-header-toggle h1 span").first
            await po_span.wait_for(state="visible", timeout=5000)
            po_gerado = (await po_span.inner_text()).strip()
        except PWTimeout:
            po_gerado = numero_po

        await _ok(confirmar,
                  f"PO {numero_po} — Concluído",
                  f"Processamento finalizado!\n\n"
                  f"  PO gerada  : {po_gerado}\n"
                  f"  Itens OK   : {itens_criados}\n\n"
                  "Clique OK para continuar para a próxima PO.")

        return {
            "po": po_gerado,
            "status": "OK",
            "mensagem": f"{itens_criados} item(ns) processado(s)"
        }

    except Exception as exc:
        return {
            "po": numero_po,
            "status": "ERRO",
            "mensagem": str(exc)[:200]
        }


# ─────────────────────────────────────────────────────────────────────
# Runner principal
# ─────────────────────────────────────────────────────────────────────

async def _run(usuario: str, senha: str, lote: list, callback, confirmar, escolher) -> list:
    vessels    = _carregar_vessels()
    vendor_map = _carregar_vendor_map()
    resultados = []

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            slow_mo=100,
        )
        context = await browser.new_context()
        page    = await context.new_page()
        page.set_default_timeout(TIMEOUT)

        callback("Fazendo login no ECO Requisition...")
        await _login(page, usuario, senha, confirmar)
        callback("Login OK. Iniciando criação de POs...")

        for i, par in enumerate(lote):
            numero_po = (par["analise"].get("po") or {}).get("numero_po") or f"Par {i+1}"
            callback(f"[{i+1}/{len(lote)}] Processando PO {numero_po}...")
            resultado = await _criar_po_par(page, par, vessels, confirmar, escolher, vendor_map)
            resultados.append(resultado)

        await browser.close()

    return resultados


def criar_pos(usuario: str, senha: str, lote: list, callback,
              confirmar=None, escolher=None) -> list:
    """
    Ponto de entrada síncrono chamado pelo tkinter em thread separada.

    confirmar: callable(titulo, mensagem) → bool
        Confirmações passo a passo durante testes. None = autônomo.

    escolher: callable(titulo, opcoes: list[str]) → int
        Exibe lista numerada e retorna índice escolhido (0-based), -1 = cancelar.
        None = autônomo (seleciona primeira opção ou usa vendor_map).
    """
    return asyncio.run(_run(usuario, senha, lote, callback, confirmar, escolher))

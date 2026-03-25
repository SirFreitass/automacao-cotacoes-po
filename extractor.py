"""
extractor.py
------------
Extrai dados estruturados de PDFs de cotação e PO usando a API do Google Gemini.
Gratuito até 1.500 requisições/dia. Suporta PDFs digitais e escaneados.
"""

import json
import re
import time
import pdfplumber
from google import genai
from google.genai import types
from config import GOOGLE_API_KEY, GEMINI_MODEL


PROMPT_COTACAO = """
Você é um assistente especializado em análise de documentos de compras industriais.

Analise este PDF que contém cotações de fornecedores e extraia as informações em JSON.
Se houver múltiplos fornecedores no documento, extraia cada um separadamente.

Retorne SOMENTE o JSON abaixo, sem texto adicional, sem markdown:

{
  "fornecedores": [
    {
      "nome": "Nome COMPLETO do fornecedor — OBRIGATÓRIO, nunca deixar vazio. Procure no cabeçalho, logotipo, assinatura, rodapé, 'From:', 'Vendor:', 'Supplier:', 'Company:'. Se o nome aparecer de formas diferentes no documento (ex: 'Puckett CAT' e 'Puckett Machinery Company'), use o nome mais completo/formal.",
      "contato": "Email ou telefone se disponível",
      "itens": [
        {
          "pn": "Part Number exato como aparece",
          "descricao": "Descrição completa do item",
          "quantidade": 1,
          "uom": "Unidade de medida do item tal como aparece no documento (ex: each, ft, lot, box, lb, gal, meter, set, pair, dozen). Se não informada, retornar 'each'.",
          "preco_unitario": 0.00,
          "preco_total_item": 0.00,
          "item_identico_ao_solicitado": true,
          "observacao_item": "Se for item similar ou substituto, descreva aqui"
        }
      ],
      "preco_total": 0.00,
      "moeda": "USD",
      "prazo_entrega": "Ex: 3-5 business days, In stock, 2 weeks",
      "prazo_entrega_dias": 0,
      "tipo_freight": "OBRIGATÓRIO — classifique em uma destas 4 opções: 'Supplier Ship' (quando há custo de frete, prepaid and add, FOB, freight collect, best way, ground, ou qualquer cobrança de envio), 'Free Delivery' (sem custo de frete, free shipping, no charge, included in price), 'Runner Pick up' (ECO Runner, coleta, pick up pela ECO), 'UPS Account' (UPS, conta UPS da ECO). Se houver custo de frete > 0, SEMPRE use 'Supplier Ship'. Na dúvida, use 'Supplier Ship'.",
      "custo_freight": 0.00,
      "forma_pagamento": "Ex: Net 30, Credit Card, COD",
      "data_cotacao": "Data de emissão da cotação no formato YYYY-MM-DD. Se não houver, null.",
      "validade_cotacao": "Data de validade no formato YYYY-MM-DD. Se mencionar dias (ex: 'valid for 30 days', 'válido por 30 dias'), calcule somando à data de emissão. Tente sempre retornar uma data concreta no formato YYYY-MM-DD.",
      "validade_dias": 30,
      "numero_cotacao": "Número da cotação — procure em todo o documento por padrões como 2025.XXXXXX ou 2026.XXXXXX onde X é dígito (ex: 2025.039982, 2026.008941). Verifique cabeçalho, rodapé, assunto do e-mail, referências e campos como 'Quote #', 'Quotation No', 'Ref:'. Se não encontrar este padrão, use o número de cotação que aparecer.",
      "numero_eco_req": "Número da REQ ECO no formato numérico (ex: 031326015461). Procure por 'REQ#', 'REQ:', 'Requisition' ou sequências longas de números que identifiquem a solicitação.",
      "observacoes": "Qualquer informação relevante adicional"
    }
  ]
}

Regras importantes:
- prazo_entrega_dias: converta para número inteiro estimado de dias úteis (ex: "2 weeks" = 10, "In stock" = 1, "3-5 days" = 5)
- preco_total: some itens + freight se aplicável
- item_identico_ao_solicitado: false se for substituto, similar ou número de parte diferente
- Se algum campo não estiver disponível, use null
- nome do fornecedor: NUNCA retorne vazio. Use o nome formal/completo que aparece no documento.
- uom: normalize para inglês (ex: "unidade"/"un" → "each", "pé"/"pés" → "feet", "caixa" → "box"). Use sempre o termo em inglês.
- Se o mesmo fornecedor aparece com nomes ligeiramente diferentes (ex: "Louisiana CAT" vs "Louisiana Cat"), use a versão EXATA como aparece no CABEÇALHO ou logotipo da cotação.

--- EXEMPLO DE SAÍDA CORRETA ---
Dado um PDF com cotação da "Power Specialties" para uma válvula, a saída esperada seria:

{"fornecedores": [{"nome": "Power Specialties", "contato": "sales@powerspec.com", "itens": [{"pn": "V-2541-B", "descricao": "2in Ball Valve 316SS 150# Flanged", "quantidade": 2, "uom": "each", "preco_unitario": 485.00, "preco_total_item": 970.00, "item_identico_ao_solicitado": true, "observacao_item": null}], "preco_total": 1015.00, "moeda": "USD", "prazo_entrega": "2-3 weeks ARO", "prazo_entrega_dias": 15, "tipo_freight": "Supplier Ship", "custo_freight": 45.00, "forma_pagamento": "Net 30", "data_cotacao": "2026-03-10", "validade_cotacao": "2026-04-09", "validade_dias": 30, "numero_cotacao": "2026.010582", "numero_eco_req": "031326015461", "observacoes": "FOB Origin, Freight Prepaid and Add"}]}

Notas sobre o exemplo:
- numero_cotacao segue o padrão 20XX.XXXXXX (encontrado no cabeçalho do e-mail)
- tipo_freight é "Supplier Ship" porque há custo de frete ($45.00)
- validade_cotacao foi calculada: data_cotacao + 30 dias
- preco_total = 970.00 (itens) + 45.00 (freight) = 1015.00
--- FIM DO EXEMPLO ---
"""

PROMPT_PO = """
Você é um assistente especializado em análise de documentos de compras industriais.

Analise este PDF de Purchase Order (PO) e extraia as informações em JSON.

Retorne SOMENTE o JSON abaixo, sem texto adicional, sem markdown:

{
  "po": {
    "numero_po": "Número da PO",
    "data": "Data da PO",
    "fornecedor_selecionado": "Nome do fornecedor para quem a PO foi emitida",
    "solicitante": "Nome do solicitante/requisitante se disponível",
    "fornecedor_escolhido_comentario": "Nome do fornecedor/fabricante REAL mencionado explicitamente nos comentários do comprador. REGRAS CRÍTICAS: (1) NUNCA retornar 'Nautical Ventures', 'ECO', 'ECO Purchasing' — estes são a própria empresa emissora da PO, jamais o fornecedor. (2) O padrão mais comum nos comentários é: 'ECO REQ#:XXXXXXXX - NOME_EMBARCAÇÃO - NOME_FORNECEDOR' — extraia apenas o NOME_FORNECEDOR (última parte após o segundo hífen). Exemplos: 'ECO Req#:022526011514 - baker' → 'baker'; 'ECO REQ#:030226012600 - KUDU - baker' → 'baker'; 'ECO REQ#:030326012826 - BRAM SPIRIT - brt marine' → 'brt marine'. (3) Outros formatos: 'purchasing from Power Specialties' → 'Power Specialties'; 'vendor: TNG Telecom' → 'TNG Telecom'. (4) Se não houver menção EXPLÍCITA de fornecedor real, retornar null.",
    "numero_eco_req": "Número da REQ ECO no formato numérico (ex: 031326015461). Procure por 'REQ#', 'REQ:', 'Requisition' ou sequências longas de números que identifiquem a solicitação nos comentários ou cabeçalho da PO.",
    "numero_cotacao_ref": "Número de cotação referenciado na PO — procure por padrões como 2025.XXXXXX ou 2026.XXXXXX (ex: 2025.039982) em qualquer campo da PO (comentários, descrição, referências). Se não encontrar este padrão, retornar null.",
    "centro_de_custo": "Nome do centro de custo/embarcação extraído da seção 'Cost Center Apportionment'. O formato é '(CÓDIGO) NOME - USD VALOR'. Retornar apenas o NOME. Ex: '(0185) C-ADMIRAL - USD 3.500,00' → 'C-ADMIRAL'. Se houver múltiplos centros de custo, retornar o nome do primeiro. Se não houver, retornar null.",
    "itens": [
      {
        "pn": "Código interno REAL do produto na PO — geralmente no formato XX.XXXXXX (ex: 10.711325, 90259010). ATENÇÃO: números sequenciais como '000010', '000020', '000030' são NÚMEROS DE LINHA do SAP, NÃO são part numbers reais. Nesse caso, procure o PN real na descrição do item (formato XX.XXXXXX) e use esse.",
        "pn_fornecedor": "PN do fornecedor extraído da descrição do item — geralmente aparece entre parênteses no final da descrição. Ex: 'Antenna GPS (GPS-ANT-001)' → 'GPS-ANT-001'. Se não houver parênteses com PN, retornar null.",
        "descricao": "Descrição completa do item",
        "quantidade": 1,
        "uom": "Unidade de medida do item (ex: each, ft, lot, box, lb, gal, meter, set, pair, dozen). Se não informada, retornar 'each'.",
        "preco_unitario": 0.00,
        "preco_total_item": 0.00,
        "fornecedor_item": "Nome do fornecedor/fabricante específico para ESTE item, se mencionado nos comentários ou notações individuais do item. Ignorar o fornecedor geral da PO — apenas capturar se houver menção explícita por item. Ex: nota do item diz 'purchasing from Bruce Kay' → 'Bruce Kay'. Se não houver, retornar null."
      }
    ],
    "subtotal": 0.00,
    "custo_freight": 0.00,
    "preco_total": 0.00,
    "moeda": "USD",
    "forma_pagamento": "Forma de pagamento",
    "prazo_entrega": "Prazo de entrega solicitado",
    "observacoes": "Observações gerais da PO incluindo comentários do comprador"
  }
}

Regras:
- O fornecedor no cabeçalho da PO pode ser um agente/broker (ex: Nautical Ventures). O fornecedor real escolhido geralmente está nos comentários do comprador.
- O campo pn pode ser um código interno da empresa — o PN real do fornecedor está entre parênteses no fim da descrição.
- Se o freight não estiver discriminado na PO, use 0.00 e registre em observacoes.
- Se algum campo não estiver disponível, use null.

--- EXEMPLO DE SAÍDA CORRETA ---
Dado um PDF de PO emitido para "Nautical Ventures" (broker) onde o comprador escreveu nos comentários "purchasing from Power Specialties - Quote 2026.010582 - REQ# 031326015461":

{"po": {"numero_po": "PO-2026-04521", "data": "2026-03-15", "fornecedor_selecionado": "Nautical Ventures", "solicitante": "John Smith", "fornecedor_escolhido_comentario": "Power Specialties", "numero_eco_req": "031326015461", "numero_cotacao_ref": "2026.010582", "centro_de_custo": "C-ADMIRAL", "itens": [{"pn": "90259010", "pn_fornecedor": "V-2541-B", "descricao": "10.710081 2in Ball Valve 316SS (V-2541-B)", "quantidade": 2, "uom": "each", "preco_unitario": 485.00, "preco_total_item": 970.00, "fornecedor_item": null}], "subtotal": 970.00, "custo_freight": 45.00, "preco_total": 1015.00, "moeda": "USD", "forma_pagamento": "Net 30", "prazo_entrega": "2-3 weeks", "observacoes": "purchasing from Power Specialties - Quote 2026.010582 - REQ# 031326015461"}}

Notas sobre o exemplo:
- fornecedor_selecionado = "Nautical Ventures" (cabeçalho da PO — é o broker)
- fornecedor_escolhido_comentario = "Power Specialties" (extraído dos comentários — é o fornecedor REAL)
- pn_fornecedor = "V-2541-B" (extraído dos parênteses na descrição)
- numero_cotacao_ref = "2026.010582" (padrão 20XX.XXXXXX encontrado nos comentários)
- centro_de_custo = "C-ADMIRAL" (apenas o nome, sem código nem valor)
- fornecedor_item = null (não há menção de fornecedor específico por item)
--- FIM DO EXEMPLO ---
"""


def _extrair_texto_pdf(caminho_pdf: str) -> str:
    """Extrai texto de todas as páginas do PDF usando pdfplumber."""
    try:
        partes = []
        with pdfplumber.open(caminho_pdf) as pdf:
            for i, page in enumerate(pdf.pages, 1):
                texto = page.extract_text() or ""
                if texto.strip():
                    partes.append(f"--- Página {i} ---\n{texto}")
        return "\n\n".join(partes)
    except Exception:
        return ""


# ── Pré-extração de campos estruturados via regex ──────────────────────

def _pre_extrair_campos(texto_pdf: str) -> dict:
    """Extrai campos estruturados do texto do PDF via regex ANTES de enviar ao Gemini."""
    campos = {}
    if not texto_pdf:
        return campos

    # Quotation codes (20XX.XXXXXX)
    codigos_raw = re.findall(r'202[4-9]\s*[.\s]\s*\d{6,}', texto_pdf)
    if codigos_raw:
        codigos = []
        for c in codigos_raw:
            limpo = re.sub(r'\s+', '', c)
            if '.' not in limpo:
                limpo = limpo[:4] + '.' + limpo[4:]  # type: ignore
            codigos.append(limpo)
        campos["quotation_codes"] = list(set(codigos))

    # ECO REQ numbers (10+ digit sequences, often after REQ# or REQ:)
    reqs = re.findall(r'(?:REQ\s*#?\s*:?\s*)(\d{10,})', texto_pdf, re.IGNORECASE)
    if reqs:
        campos["eco_req_numbers"] = list(set(reqs))

    # Part numbers (XX.XXXXXX format — código interno)
    pns = re.findall(r'\b(\d{2}\.\d{6})\b', texto_pdf)
    if pns:
        campos["part_numbers"] = list(set(pns))

    # Dollar amounts
    valores = re.findall(r'\$\s*([\d,]+\.\d{2})', texto_pdf)
    if valores:
        campos["dollar_amounts"] = valores[:10]  # type: ignore

    # Dates (various formats)
    datas_us = re.findall(r'\b(\d{1,2}/\d{1,2}/\d{2,4})\b', texto_pdf)
    datas_iso = re.findall(r'\b(\d{4}-\d{2}-\d{2})\b', texto_pdf)
    todas_datas = datas_us + datas_iso
    if todas_datas:
        campos["dates_found"] = todas_datas[:10]  # type: ignore

    # Supplier hints (From:, Vendor:, Supplier:, Company:, Quoted by:)
    vendor_patterns = re.findall(
        r'(?:From|Vendor|Supplier|Company|Quoted\s+by|Prepared\s+by)\s*[:]\s*(.+)',
        texto_pdf, re.IGNORECASE
    )
    if vendor_patterns:
        campos["vendor_hints"] = [v.strip()[:80] for v in vendor_patterns[:5]]  # type: ignore

    # Buyer comment vendor mentions (common in POs)
    buyer_vendors = re.findall(
        r'(?:FORN\.?\s*|purchasing\s+from\s+|vendor\s*:?\s*|buying\s+from\s+)([A-Z][A-Za-z\s&.,]+)',
        texto_pdf, re.IGNORECASE
    )
    if buyer_vendors:
        campos["buyer_vendor_mentions"] = [v.strip().rstrip('.,') for v in buyer_vendors[:5]]  # type: ignore

    # Payment terms
    pay_patterns = re.findall(
        r'(Net\s+\d+|COD|Credit\s+Card|Due\s+on\s+Receipt|Prepaid|C\.?O\.?D\.?)',
        texto_pdf, re.IGNORECASE
    )
    if pay_patterns:
        campos["payment_terms"] = list(set(pay_patterns))

    # Freight/shipping hints
    freight_patterns = re.findall(
        r'(FOB\s+\w+|Freight\s+(?:Prepaid|Collect|Included)|Free\s+Shipping|'
        r'Prepaid\s+and\s+Add|UPS\s+Ground|FedEx|Best\s+Way|No\s+Charge)',
        texto_pdf, re.IGNORECASE
    )
    if freight_patterns:
        campos["freight_hints"] = list(set(freight_patterns))[:5]  # type: ignore

    # Cost center (format: (XXXX) NAME - USD)
    cc = re.findall(r'\(\d{4}\)\s+([A-Z][\w\s-]+?)\s*[-–]', texto_pdf)
    if cc:
        campos["cost_centers"] = [c.strip() for c in cc[:3]]  # type: ignore

    return campos


def _formatar_pre_extracoes(campos: dict) -> str:
    """Formata os campos pré-extraídos como seção de referência para o prompt."""
    if not campos:
        return ""

    partes = [
        "--- DADOS PRÉ-IDENTIFICADOS NO DOCUMENTO (REFERÊNCIA OBRIGATÓRIA) ---",
        "Os seguintes dados foram encontrados no texto do documento via análise automatizada.",
        "USE estes dados como referência prioritária. Se o dado pré-identificado estiver correto, USE-O.",
        "",
    ]

    mapa = {
        "quotation_codes":      "CÓDIGOS DE COTAÇÃO encontrados",
        "eco_req_numbers":      "NÚMEROS ECO REQ encontrados",
        "part_numbers":         "PART NUMBERS (formato XX.XXXXXX) encontrados",
        "vendor_hints":         "POSSÍVEIS FORNECEDORES (cabeçalho/assinatura)",
        "buyer_vendor_mentions":"FORNECEDORES mencionados em comentários do comprador",
        "payment_terms":        "TERMOS DE PAGAMENTO encontrados",
        "freight_hints":        "REFERÊNCIAS DE FRETE/ENVIO",
        "dates_found":          "DATAS encontradas no documento",
        "dollar_amounts":       "VALORES EM DÓLAR encontrados",
        "cost_centers":         "CENTROS DE CUSTO encontrados",
    }

    for chave, rotulo in mapa.items():
        if chave in campos:
            valores = campos[chave]
            partes.append(f"• {rotulo}: {', '.join(str(v) for v in valores)}")

    partes.append("--- FIM DOS DADOS PRÉ-IDENTIFICADOS ---")
    return "\n".join(partes)


def _chamar_gemini(caminho_pdf: str, prompt: str, tentativas: int = 3) -> dict:
    """
    Envia PDF + texto extraído + dados pré-identificados ao Gemini.
    Usa response_mime_type='application/json' para forçar saída JSON completa.
    Retry automático em caso de erro 429 (quota).
    """
    client = genai.Client(api_key=GOOGLE_API_KEY)

    # Abre como bytes para evitar erro de encoding com nomes acentuados (Windows)
    with open(caminho_pdf, "rb") as f:
        arquivo = client.files.upload(
            file=f,
            config=types.UploadFileConfig(display_name="documento.pdf", mime_type="application/pdf"),
        )

    # Extrai texto do PDF como fonte complementar
    texto_pdf = _extrair_texto_pdf(caminho_pdf)

    # Pré-extrai campos estruturados via regex
    campos_pre = _pre_extrair_campos(texto_pdf)
    pre_extracoes = _formatar_pre_extracoes(campos_pre)

    # Monta prompt enriquecido: prompt base + dados pré-identificados + texto PDF
    prompt_completo = prompt
    if pre_extracoes:
        prompt_completo += f"\n\n{pre_extracoes}"
    if texto_pdf:
        prompt_completo += (
            "\n\n--- TEXTO EXTRAÍDO DO PDF (use como referência complementar) ---\n"
            f"{texto_pdf}\n"
            "--- FIM DO TEXTO EXTRAÍDO ---"
        )

    resposta = None
    try:
        for tentativa in range(1, tentativas + 1):
            try:
                resposta = client.models.generate_content(
                    model=GEMINI_MODEL,
                    contents=[arquivo, prompt_completo],
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                    ),
                )
                break  # sucesso
            except Exception as e:
                msg = str(e)
                if "429" in msg or "RESOURCE_EXHAUSTED" in msg:
                    if tentativa < tentativas:
                        time.sleep(60)
                        continue
                raise
    finally:
        try:
            client.files.delete(name=arquivo.name)
        except Exception:
            pass

    if resposta is None:
        raise RuntimeError("Gemini não retornou resposta após todas as tentativas.")

    texto = resposta.text.strip()  # type: ignore

    # Remove blocos markdown se o modelo os incluir
    texto = re.sub(r"^```(?:json)?\s*", "", texto)
    texto = re.sub(r"\s*```$", "", texto)

    return json.loads(texto)


FREIGHT_VALIDOS = {"Supplier Ship", "Free Delivery", "Runner Pick up", "UPS Account"}
RE_QUOTATION = re.compile(r'^202[4-9]\.\d{6,}$')
RE_DATA = re.compile(r'^\d{4}-\d{2}-\d{2}$')


def _validar_cotacao(dados: dict) -> list[str]:
    """Valida dados extraídos de cotação e retorna lista de problemas encontrados."""
    problemas = []
    fornecedores = dados.get("fornecedores", [])

    if not fornecedores:
        problemas.append("Nenhum fornecedor foi extraído. Revise o documento inteiro.")
        return problemas

    for i, f in enumerate(fornecedores, 1):
        nome = f.get("nome")
        if not nome or nome == "null":
            problemas.append(f"Fornecedor {i}: nome está vazio. Procure no cabeçalho, logotipo, assinatura ou rodapé.")

        # Quotation code
        qc = f.get("numero_cotacao")
        if not qc or not RE_QUOTATION.match(str(qc)):
            problemas.append(
                f"Fornecedor {i} ({nome}): numero_cotacao ausente ou fora do padrão 20XX.XXXXXX. "
                "Procure em TODAS as páginas: cabeçalhos, rodapés, assunto de e-mail, campos 'Quote #', 'Ref:'."
            )

        # Freight
        tf = f.get("tipo_freight")
        if tf not in FREIGHT_VALIDOS:
            problemas.append(
                f"Fornecedor {i} ({nome}): tipo_freight '{tf}' inválido. "
                f"Use uma das opções: {', '.join(FREIGHT_VALIDOS)}."
            )

        # Preços
        itens = f.get("itens", [])
        if not itens:
            problemas.append(f"Fornecedor {i} ({nome}): nenhum item extraído.")
        for j, item in enumerate(itens, 1):
            pu = item.get("preco_unitario")
            if pu is None or (isinstance(pu, (int, float)) and pu <= 0):
                problemas.append(f"Fornecedor {i} ({nome}), item {j}: preco_unitario ausente ou zero.")
            # PN vazio
            pn = (item.get("pn") or "").strip()
            if not pn:
                problemas.append(
                    f"Fornecedor {i} ({nome}), item {j}: pn (part number) está vazio. "
                    "Procure na linha do item, coluna 'Part #', 'P/N', 'Item #' ou no catálogo."
                )

        # Data
        dc = f.get("data_cotacao")
        if dc and not RE_DATA.match(str(dc)):
            problemas.append(f"Fornecedor {i} ({nome}): data_cotacao '{dc}' não está no formato YYYY-MM-DD.")
        if not dc:
            problemas.append(f"Fornecedor {i} ({nome}): data_cotacao está vazia. Procure 'Date', 'Issued', 'Quote Date'.")

        # Prazo de entrega
        pe = f.get("prazo_entrega")
        if not pe or pe == "null":
            problemas.append(f"Fornecedor {i} ({nome}): prazo_entrega está vazio. Procure 'Lead Time', 'Delivery', 'Ship Date', 'ETA', 'ARO'.")

        # Forma de pagamento
        fp = f.get("forma_pagamento")
        if not fp or fp == "null":
            problemas.append(f"Fornecedor {i} ({nome}): forma_pagamento está vazio. Procure 'Terms', 'Payment', 'Net 30'.")

    return problemas


def _validar_po(dados: dict) -> list[str]:
    """Valida dados extraídos de PO e retorna lista de problemas encontrados."""
    problemas = []
    po = dados.get("po", {})

    if not po:
        problemas.append("Nenhum dado de PO foi extraído. Revise o documento inteiro.")
        return problemas

    if not po.get("numero_po"):
        problemas.append("numero_po está vazio.")

    # Centro de custo
    cc = po.get("centro_de_custo")
    if not cc or cc == "null":
        problemas.append(
            "centro_de_custo está vazio. Procure na seção 'Cost Center Apportionment' "
            "ou 'Ship To'. Formato: '(XXXX) NOME - USD'. Retorne apenas o NOME."
        )

    # Fornecedor dos comentários
    fec = po.get("fornecedor_escolhido_comentario")
    if not fec or fec == "null":
        problemas.append(
            "fornecedor_escolhido_comentario está vazio. Procure nos 'Buyer comments', "
            "'Observações', 'Notes' por menções como 'Vendor:', 'FORN.', 'purchasing from'."
        )

    # Itens
    itens = po.get("itens", [])
    if not itens:
        problemas.append("Nenhum item extraído da PO.")
    for j, item in enumerate(itens, 1):
        pu = item.get("preco_unitario")
        if pu is None or (isinstance(pu, (int, float)) and pu <= 0):
            problemas.append(f"Item {j}: preco_unitario ausente ou zero.")
        # Detecta PNs falsos (números de linha SAP: 000010, 000020, etc.)
        pn = str(item.get("pn") or "").strip()
        if re.match(r'^0{2,}\d{1,2}$', pn):
            problemas.append(
                f"Item {j}: pn '{pn}' parece ser número de linha SAP, NÃO um part number real. "
                "Procure o PN real na descrição do item (formato XX.XXXXXX como 10.711325)."
            )
        # PN completamente vazio
        if not pn:
            problemas.append(
                f"Item {j}: pn está vazio. Procure o código do produto na PO — "
                "geralmente no formato XX.XXXXXX (ex: 10.711325, 90259010)."
            )
        # Descrição vazia
        desc = (item.get("descricao") or "").strip()
        if not desc:
            problemas.append(f"Item {j}: descricao está vazia.")

    # Quotation ref
    ref = po.get("numero_cotacao_ref")
    if ref and not RE_QUOTATION.match(str(ref)):
        problemas.append(
            f"numero_cotacao_ref '{ref}' fora do padrão 20XX.XXXXXX. "
            "Procure em comentários, descrições e referências da PO."
        )

    return problemas


def extrair_cotacoes(caminho_pdf: str) -> dict:
    """
    Extrai dados de cotações do PDF.
    Valida o resultado e re-extrai uma vez com instruções corretivas se houver problemas.
    Aplica fallbacks com pdfplumber para campos críticos ainda ausentes.
    """
    dados = _chamar_gemini(caminho_pdf, PROMPT_COTACAO)
    problemas = _validar_cotacao(dados)

    if problemas:
        correcao = (
            f"{PROMPT_COTACAO}\n\n"
            "--- ATENÇÃO: CORREÇÕES NECESSÁRIAS ---\n"
            "Uma extração anterior retornou os seguintes problemas:\n"
            + "\n".join(f"- {p}" for p in problemas) + "\n"
            "Por favor, corrija estes problemas na nova extração.\n"
            "--- FIM DAS CORREÇÕES ---"
        )
        dados_corrigidos = _chamar_gemini(caminho_pdf, correcao)
        if dados_corrigidos.get("fornecedores"):
            dados = dados_corrigidos

    # ── Fallbacks com pdfplumber para campos críticos ──────────────────
    texto_pdf = _extrair_texto_pdf(caminho_pdf)
    campos_pre = _pre_extrair_campos(texto_pdf)
    codigos_pre = campos_pre.get("quotation_codes", [])

    for forn in dados.get("fornecedores", []):
        # Quotation code
        qc = forn.get("numero_cotacao")
        if (not qc or not RE_QUOTATION.match(str(qc))) and codigos_pre:
            forn["numero_cotacao"] = codigos_pre[0]

        # ECO REQ
        req = forn.get("numero_eco_req")
        if not req or req == "null":
            reqs_pre = campos_pre.get("eco_req_numbers", [])
            if reqs_pre:
                forn["numero_eco_req"] = reqs_pre[0]

    return dados


def _extrair_quotation_code_pdfplumber(caminho_pdf: str) -> str | None:
    """
    Extrai o Quotation Code (formato 20XX.XXXXXX) diretamente do texto do PDF
    usando pdfplumber. Fallback robusto independente do Gemini.
    Tenta múltiplos padrões para cobrir variações de formatação.
    """
    try:
        texto_completo = ""
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                texto_completo += (page.extract_text() or "") + "\n"

        if not texto_completo.strip():
            return None

        # Ano válido: 2024-2029 (evita capturar números como 2032.600687 do ECO REQ)
        ANO = r'(202[4-9])'

        # 1. Padrão exato: 2026.010070
        match = re.search(ANO + r'\.(\d{6,})', texto_completo)
        if match:
            return f"{match.group(1)}.{match.group(2)}"

        # 2. Com espaços ao redor do ponto: "2026 . 010070"
        match = re.search(ANO + r'\s*\.\s*(\d{6,})', texto_completo)
        if match:
            return f"{match.group(1)}.{match.group(2)}"

        # 3. Sem ponto mas com separação clara: "2026 010070"
        match = re.search(ANO + r'\s+(\d{6,})', texto_completo)
        if match:
            return f"{match.group(1)}.{match.group(2)}"

    except Exception:
        pass
    return None


def extrair_po(caminho_pdf: str) -> dict:
    """
    Extrai dados da PO do PDF.
    Valida o resultado e re-extrai uma vez com instruções corretivas se houver problemas.
    Garante que numero_cotacao_ref seja extraído via pdfplumber como fallback.
    """
    dados = _chamar_gemini(caminho_pdf, PROMPT_PO)
    problemas = _validar_po(dados)

    if problemas:
        correcao = (
            f"{PROMPT_PO}\n\n"
            "--- ATENÇÃO: CORREÇÕES NECESSÁRIAS ---\n"
            "Uma extração anterior retornou os seguintes problemas:\n"
            + "\n".join(f"- {p}" for p in problemas) + "\n"
            "Por favor, corrija estes problemas na nova extração.\n"
            "--- FIM DAS CORREÇÕES ---"
        )
        dados_corrigidos = _chamar_gemini(caminho_pdf, correcao)
        if dados_corrigidos.get("po"):
            dados = dados_corrigidos

    po = dados.get("po", {})

    # ── Fallbacks com pdfplumber para campos críticos ──────────────────
    texto_pdf = _extrair_texto_pdf(caminho_pdf)
    campos_pre = _pre_extrair_campos(texto_pdf)

    # Quotation code
    ref = po.get("numero_cotacao_ref")
    if not ref or not re.match(r'202[4-9]\.\d{6,}', str(ref)):
        code = _extrair_quotation_code_pdfplumber(caminho_pdf)
        if code:
            po["numero_cotacao_ref"] = code

    # Part numbers via regex se itens estão sem PN
    pns_encontrados = campos_pre.get("part_numbers", [])
    itens = po.get("itens", [])
    if pns_encontrados and itens:
        for item in itens:
            pn = (item.get("pn") or "").strip()
            # Se PN está vazio ou é número de linha SAP
            if not pn or re.match(r'^0{2,}\d{1,2}$', pn):
                desc = (item.get("descricao") or "").upper()
                # Tenta encontrar um PN do regex que apareça na descrição
                for pn_regex in pns_encontrados:
                    if pn_regex in desc or pn_regex.replace(".", "") in desc.replace(".", ""):  # type: ignore
                        item["pn"] = pn_regex
                        break

    # Fornecedor dos comentários via regex
    fec = po.get("fornecedor_escolhido_comentario")
    if not fec or fec == "null":
        mencoes = campos_pre.get("buyer_vendor_mentions", [])
        if mencoes:
            po["fornecedor_escolhido_comentario"] = mencoes[0]

    # Centro de custo via regex
    cc = po.get("centro_de_custo")
    if not cc or cc == "null":
        centros = campos_pre.get("cost_centers", [])
        if centros:
            po["centro_de_custo"] = centros[0]

    dados["po"] = po
    return dados

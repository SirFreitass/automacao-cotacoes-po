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


# ═══════════════════════════════════════════════════════════════════════
# COTAÇÃO — 2 passadas: P1 (itens) + P2 (termos/metadados)
# ═══════════════════════════════════════════════════════════════════════

PROMPT_COTACAO_P1 = """
Analise este PDF de cotação de fornecedor e extraia ITENS E PREÇOS em JSON.
Foco: nome do fornecedor, lista de itens, preços e totais.

{
  "fornecedores": [
    {
      "nome": "Nome COMPLETO do fornecedor (cabeçalho, logotipo, assinatura, 'From:', 'Vendor:'). OBRIGATÓRIO.",
      "contato": "Email ou telefone",
      "itens": [
        {
          "pn": "Part Number exato",
          "descricao": "Descrição completa do item",
          "quantidade": 1,
          "uom": "Unidade de medida em inglês (each, ft, box, lb, gal, meter, set, pair, dozen). Default: 'each'.",
          "preco_unitario": 0.00,
          "preco_total_item": 0.00,
          "item_identico_ao_solicitado": true,
          "observacao_item": "Se substituto/similar, descreva aqui"
        }
      ],
      "preco_total": 0.00,
      "moeda": "USD"
    }
  ]
}

Regras:
- preco_total: soma dos itens (sem freight nesta passada)
- item_identico_ao_solicitado: false se substituto/similar
- nome: NUNCA vazio. Use o nome formal do CABEÇALHO ou logotipo.
- Se múltiplos fornecedores, extraia cada um separadamente.

Exemplo: {"fornecedores": [{"nome": "Power Specialties", "contato": "sales@powerspec.com", "itens": [{"pn": "V-2541-B", "descricao": "2in Ball Valve 316SS", "quantidade": 2, "uom": "each", "preco_unitario": 485.00, "preco_total_item": 970.00, "item_identico_ao_solicitado": true, "observacao_item": null}], "preco_total": 970.00, "moeda": "USD"}]}
"""

PROMPT_COTACAO_P2 = """
Analise este PDF de cotação e extraia TERMOS COMERCIAIS E REFERÊNCIAS em JSON.
Foco: frete, pagamento, datas, prazos, número da cotação.

{
  "termos": [
    {
      "nome_fornecedor": "Nome do fornecedor (para associar aos itens extraídos anteriormente)",
      "tipo_freight": "OBRIGATÓRIO — uma destas 4 opções: 'Supplier Ship' (custo de frete, prepaid and add, FOB, freight collect, best way, ground), 'Free Delivery' (sem custo, free shipping, no charge, included), 'Runner Pick up' (ECO Runner, coleta, pick up), 'UPS Account' (UPS, conta UPS). Custo > 0 → SEMPRE 'Supplier Ship'. Na dúvida → 'Supplier Ship'.",
      "custo_freight": 0.00,
      "forma_pagamento": "Ex: Net 30, Credit Card, COD. Procure 'Terms', 'Payment Terms'.",
      "prazo_entrega": "Texto original. Ex: '3-5 business days', 'In stock', '2 weeks ARO'",
      "prazo_entrega_dias": 0,
      "data_cotacao": "YYYY-MM-DD. Procure 'Date', 'Quote Date', 'Issued'.",
      "validade_cotacao": "YYYY-MM-DD. Se diz '30 days', calcule: data_cotacao + 30 dias.",
      "validade_dias": 30,
      "numero_cotacao": "Procure padrão 20XX.XXXXXX (ex: 2026.010582) em cabeçalho, rodapé, 'Quote #', 'Ref:', assunto de e-mail. Se não encontrar este padrão, use qualquer número de cotação.",
      "numero_eco_req": "Formato numérico longo (ex: 031326015461). Procure 'REQ#', 'REQ:', 'Requisition'.",
      "observacoes": "Informações adicionais relevantes (condições, FOB, etc.)"
    }
  ]
}

Regras:
- prazo_entrega_dias: converta para dias úteis ("2 weeks" = 10, "In stock" = 1, "3-5 days" = 5)
- Se campo não disponível, use null. NUNCA deixe tipo_freight vazio.

Exemplo: {"termos": [{"nome_fornecedor": "Power Specialties", "tipo_freight": "Supplier Ship", "custo_freight": 45.00, "forma_pagamento": "Net 30", "prazo_entrega": "2-3 weeks ARO", "prazo_entrega_dias": 15, "data_cotacao": "2026-03-10", "validade_cotacao": "2026-04-09", "validade_dias": 30, "numero_cotacao": "2026.010582", "numero_eco_req": "031326015461", "observacoes": "FOB Origin, Freight Prepaid and Add"}]}
"""

# ═══════════════════════════════════════════════════════════════════════
# PO — 2 passadas: P1 (itens) + P2 (metadados/comentários)
# ═══════════════════════════════════════════════════════════════════════

PROMPT_PO_P1 = """
Analise este PDF de Purchase Order (PO) e extraia ITENS E VALORES em JSON.
Foco: número da PO, itens, preços, totais.

{
  "po": {
    "numero_po": "Número da PO",
    "data": "Data da PO",
    "fornecedor_selecionado": "Nome do fornecedor no cabeçalho da PO",
    "itens": [
      {
        "pn": "Código interno do produto (formato XX.XXXXXX como 10.711325). ATENÇÃO: '000010', '000020' são números de LINHA SAP, não PNs. Procure o PN real na descrição.",
        "pn_fornecedor": "PN entre parênteses na descrição. Ex: '... (GPS-ANT-001)' → 'GPS-ANT-001'. Se não houver, null.",
        "descricao": "Descrição completa do item",
        "quantidade": 1,
        "uom": "Unidade de medida em inglês (each, ft, box, lb). Default: 'each'.",
        "preco_unitario": 0.00,
        "preco_total_item": 0.00
      }
    ],
    "subtotal": 0.00,
    "custo_freight": 0.00,
    "preco_total": 0.00,
    "moeda": "USD"
  }
}

Exemplo: {"po": {"numero_po": "PO-2026-04521", "data": "2026-03-15", "fornecedor_selecionado": "Nautical Ventures", "itens": [{"pn": "90259010", "pn_fornecedor": "V-2541-B", "descricao": "10.710081 2in Ball Valve 316SS (V-2541-B)", "quantidade": 2, "uom": "each", "preco_unitario": 485.00, "preco_total_item": 970.00}], "subtotal": 970.00, "custo_freight": 45.00, "preco_total": 1015.00, "moeda": "USD"}}
"""

PROMPT_PO_P2 = """
Analise este PDF de Purchase Order (PO) e extraia METADADOS E COMENTÁRIOS em JSON.
Foco: fornecedor real, referências, centro de custo, termos.

{
  "po_meta": {
    "fornecedor_escolhido_comentario": "Fornecedor REAL nos comentários do comprador. REGRAS: (1) NUNCA 'Nautical Ventures', 'ECO', 'ECO Purchasing' — são a empresa emissora. (2) Padrão comum: 'ECO REQ#:XXXXXXXX - EMBARCAÇÃO - FORNECEDOR' — extraia APENAS o FORNECEDOR (última parte). Ex: 'ECO REQ#:030326012826 - BRAM SPIRIT - brt marine' → 'brt marine'. (3) Outros: 'purchasing from Power Specialties' → 'Power Specialties'. (4) Se não há menção explícita → null.",
    "numero_eco_req": "Número ECO REQ (ex: 031326015461). Procure 'REQ#', 'REQ:', 'Requisition'.",
    "numero_cotacao_ref": "Padrão 20XX.XXXXXX (ex: 2026.010582) em comentários, descrição, referências. Se não encontrar → null.",
    "centro_de_custo": "Nome do centro de custo em 'Cost Center Apportionment'. Formato '(XXXX) NOME - USD'. Retorne SÓ o NOME. Ex: '(0185) C-ADMIRAL - USD 3500' → 'C-ADMIRAL'.",
    "solicitante": "Nome do solicitante/requisitante",
    "forma_pagamento": "Termos de pagamento (Net 30, COD, etc.)",
    "prazo_entrega": "Prazo de entrega solicitado",
    "observacoes": "Todos os comentários do comprador, buyer notes, observações.",
    "fornecedor_item": [
      {"pn": "PN do item", "fornecedor": "Fornecedor específico para este item, se mencionado. Se não → null."}
    ]
  }
}

Exemplo: {"po_meta": {"fornecedor_escolhido_comentario": "Power Specialties", "numero_eco_req": "031326015461", "numero_cotacao_ref": "2026.010582", "centro_de_custo": "C-ADMIRAL", "solicitante": "John Smith", "forma_pagamento": "Net 30", "prazo_entrega": "2-3 weeks", "observacoes": "purchasing from Power Specialties - Quote 2026.010582 - REQ# 031326015461", "fornecedor_item": []}}
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


def _chamar_gemini(caminho_pdf: str, prompt: str, tentativas: int = 3,
                   texto_pdf: str = None, campos_pre: dict = None) -> dict:
    """
    Envia PDF + texto extraído + dados pré-identificados ao Gemini.
    Usa response_mime_type='application/json' para forçar saída JSON completa.
    Retry automático em caso de erro 429 (quota).

    texto_pdf e campos_pre podem ser passados para evitar re-leitura do PDF.
    """
    client = genai.Client(api_key=GOOGLE_API_KEY)

    # Abre como bytes para evitar erro de encoding com nomes acentuados (Windows)
    with open(caminho_pdf, "rb") as f:
        arquivo = client.files.upload(
            file=f,
            config=types.UploadFileConfig(display_name="documento.pdf", mime_type="application/pdf"),
        )

    # Usa texto/campos cacheados ou extrai pela primeira vez
    if texto_pdf is None:
        texto_pdf = _extrair_texto_pdf(caminho_pdf)
    if campos_pre is None:
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


def _merge_cotacao(p1: dict, p2: dict) -> dict:
    """Combina resultados da Passada 1 (itens) e Passada 2 (termos) da cotação."""
    fornecedores = p1.get("fornecedores", [])
    termos = p2.get("termos", [])

    # Indexa termos por nome normalizado do fornecedor
    termos_idx = {}
    for t in termos:
        nome = (t.get("nome_fornecedor") or "").lower().strip()
        if nome:
            termos_idx[nome] = t

    for forn in fornecedores:
        nome = (forn.get("nome") or "").lower().strip()
        # Busca termos: match exato ou substring
        termo = termos_idx.get(nome)
        if not termo:
            termo = next((t for k, t in termos_idx.items()
                         if k in nome or nome in k), None)
        # Se só há 1 fornecedor e 1 termo, associa direto
        if not termo and len(fornecedores) == 1 and len(termos) == 1:
            termo = termos[0]

        if termo:
            forn["tipo_freight"] = termo.get("tipo_freight")
            forn["custo_freight"] = termo.get("custo_freight")
            forn["forma_pagamento"] = termo.get("forma_pagamento")
            forn["prazo_entrega"] = termo.get("prazo_entrega")
            forn["prazo_entrega_dias"] = termo.get("prazo_entrega_dias")
            forn["data_cotacao"] = termo.get("data_cotacao")
            forn["validade_cotacao"] = termo.get("validade_cotacao")
            forn["validade_dias"] = termo.get("validade_dias", 30)
            forn["numero_cotacao"] = termo.get("numero_cotacao")
            forn["numero_eco_req"] = termo.get("numero_eco_req")
            forn["observacoes"] = termo.get("observacoes")
            # Soma freight ao preco_total se aplicável
            freight = termo.get("custo_freight") or 0
            if freight and forn.get("preco_total"):
                forn["preco_total"] = forn["preco_total"] + float(freight)

    return {"fornecedores": fornecedores}


def extrair_cotacoes(caminho_pdf: str) -> dict:
    """
    Extrai dados de cotações do PDF em 2 passadas focadas:
      P1: itens, preços, fornecedor (dados concretos)
      P2: freight, pagamento, datas, referências (metadados)
    Valida, re-extrai se necessário, e aplica fallbacks pdfplumber.
    """
    # Extrai texto UMA vez — reutilizado em todas as chamadas
    texto_pdf = _extrair_texto_pdf(caminho_pdf)
    campos_pre = _pre_extrair_campos(texto_pdf)

    # ── Passada 1: itens e preços ──────────────────────────────────────
    p1 = _chamar_gemini(caminho_pdf, PROMPT_COTACAO_P1,
                        texto_pdf=texto_pdf, campos_pre=campos_pre)

    # ── Passada 2: termos e referências ────────────────────────────────
    p2 = _chamar_gemini(caminho_pdf, PROMPT_COTACAO_P2,
                        texto_pdf=texto_pdf, campos_pre=campos_pre)

    # ── Merge dos resultados ───────────────────────────────────────────
    dados = _merge_cotacao(p1, p2)

    # ── Validação + retry focado ───────────────────────────────────────
    problemas = _validar_cotacao(dados)
    if problemas:
        # Retry com prompt combinado (menor que o antigo) + lista de problemas
        correcao = (
            f"{PROMPT_COTACAO_P1}\n\n{PROMPT_COTACAO_P2}\n\n"
            "--- CORREÇÕES NECESSÁRIAS ---\n"
            + "\n".join(f"- {p}" for p in problemas) + "\n"
            "Corrija estes problemas. Retorne JSON com chave 'fornecedores'.\n"
        )
        dados_corrigidos = _chamar_gemini(caminho_pdf, correcao,
                                          texto_pdf=texto_pdf, campos_pre=campos_pre)
        if dados_corrigidos.get("fornecedores"):
            dados = dados_corrigidos

    # ── Fallbacks pdfplumber (já cacheados) ────────────────────────────
    codigos_pre = campos_pre.get("quotation_codes", [])

    for forn in dados.get("fornecedores", []):
        qc = forn.get("numero_cotacao")
        if (not qc or not RE_QUOTATION.match(str(qc))) and codigos_pre:
            forn["numero_cotacao"] = codigos_pre[0]

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


def _merge_po(p1: dict, p2: dict) -> dict:
    """Combina resultados da Passada 1 (itens) e Passada 2 (metadados) da PO."""
    po = p1.get("po", {})
    meta = p2.get("po_meta", {})

    # Mescla metadados na PO
    po["fornecedor_escolhido_comentario"] = meta.get("fornecedor_escolhido_comentario")
    po["numero_eco_req"] = meta.get("numero_eco_req")
    po["numero_cotacao_ref"] = meta.get("numero_cotacao_ref")
    po["centro_de_custo"] = meta.get("centro_de_custo")
    po["solicitante"] = meta.get("solicitante") or po.get("solicitante")
    po["forma_pagamento"] = meta.get("forma_pagamento")
    po["prazo_entrega"] = meta.get("prazo_entrega")
    po["observacoes"] = meta.get("observacoes")

    # Fornecedor por item (se P2 retornou)
    forn_itens = {fi.get("pn"): fi.get("fornecedor")
                  for fi in meta.get("fornecedor_item", []) if fi.get("pn")}
    if forn_itens:
        for item in po.get("itens", []):
            pn = item.get("pn") or ""
            if pn in forn_itens:
                item["fornecedor_item"] = forn_itens[pn]

    return {"po": po}


def extrair_po(caminho_pdf: str) -> dict:
    """
    Extrai dados da PO em 2 passadas focadas:
      P1: número, itens, preços (dados concretos)
      P2: fornecedor real, ECO REQ, cotação, centro de custo (metadados)
    Valida, re-extrai se necessário, e aplica fallbacks pdfplumber.
    """
    # Extrai texto UMA vez
    texto_pdf = _extrair_texto_pdf(caminho_pdf)
    campos_pre = _pre_extrair_campos(texto_pdf)

    # ── Passada 1: itens e valores ─────────────────────────────────────
    p1 = _chamar_gemini(caminho_pdf, PROMPT_PO_P1,
                        texto_pdf=texto_pdf, campos_pre=campos_pre)

    # ── Passada 2: metadados e comentários ─────────────────────────────
    p2 = _chamar_gemini(caminho_pdf, PROMPT_PO_P2,
                        texto_pdf=texto_pdf, campos_pre=campos_pre)

    # ── Merge dos resultados ───────────────────────────────────────────
    dados = _merge_po(p1, p2)

    # ── Validação + retry focado ───────────────────────────────────────
    problemas = _validar_po(dados)
    if problemas:
        correcao = (
            f"{PROMPT_PO_P1}\n\n{PROMPT_PO_P2}\n\n"
            "--- CORREÇÕES NECESSÁRIAS ---\n"
            + "\n".join(f"- {p}" for p in problemas) + "\n"
            "Corrija estes problemas. Retorne JSON com chave 'po' contendo TODOS os campos.\n"
        )
        dados_corrigidos = _chamar_gemini(caminho_pdf, correcao,
                                          texto_pdf=texto_pdf, campos_pre=campos_pre)
        if dados_corrigidos.get("po"):
            dados = dados_corrigidos

    po = dados.get("po", {})

    # ── Fallbacks pdfplumber (já cacheados) ────────────────────────────

    # Quotation code
    ref = po.get("numero_cotacao_ref")
    codigos_pre = campos_pre.get("quotation_codes", [])
    if not ref or not re.match(r'202[4-9]\.\d{6,}', str(ref)):
        if codigos_pre:
            po["numero_cotacao_ref"] = codigos_pre[0]

    # Part numbers via regex
    pns_encontrados = campos_pre.get("part_numbers", [])
    itens = po.get("itens", [])
    if pns_encontrados and itens:
        for item in itens:
            pn = (item.get("pn") or "").strip()
            if not pn or re.match(r'^0{2,}\d{1,2}$', pn):
                desc = (item.get("descricao") or "").upper()
                for pn_regex in pns_encontrados:
                    if pn_regex in desc or pn_regex.replace(".", "") in desc.replace(".", ""):
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

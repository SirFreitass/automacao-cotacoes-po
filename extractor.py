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
      "nome": "Nome do fornecedor",
      "contato": "Email ou telefone se disponível",
      "itens": [
        {
          "pn": "Part Number exato como aparece",
          "descricao": "Descrição completa do item",
          "quantidade": 1,
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
    "fornecedor_escolhido_comentario": "Nome do fornecedor/fabricante REAL mencionado explicitamente nos comentários do comprador, observações ou notas da PO — NÃO o fornecedor do cabeçalho (que é o agente/broker). Exemplos de como costuma aparecer: 'purchasing from Power Specialties', 'vendor: TNG Telecom', 'buying from Bruce Kay', 'purchasing process made by ECO - TNG Telecom'. Retorne APENAS o nome do fornecedor real, sem texto adicional. Se não houver menção EXPLÍCITA de fornecedor real nos comentários, retornar null — NUNCA invente ou suponha um nome.",
    "numero_eco_req": "Número da REQ ECO no formato numérico (ex: 031326015461). Procure por 'REQ#', 'REQ:', 'Requisition' ou sequências longas de números que identifiquem a solicitação nos comentários ou cabeçalho da PO.",
    "numero_cotacao_ref": "Número de cotação referenciado na PO — procure por padrões como 2025.XXXXXX ou 2026.XXXXXX (ex: 2025.039982) em qualquer campo da PO (comentários, descrição, referências). Se não encontrar este padrão, retornar null.",
    "centro_de_custo": "Nome do centro de custo/embarcação extraído da seção 'Cost Center Apportionment'. O formato é '(CÓDIGO) NOME - USD VALOR'. Retornar apenas o NOME. Ex: '(0185) C-ADMIRAL - USD 3.500,00' → 'C-ADMIRAL'. Se houver múltiplos centros de custo, retornar o nome do primeiro. Se não houver, retornar null.",
    "itens": [
      {
        "pn": "Código interno da PO (pode ser código numérico interno da empresa)",
        "pn_fornecedor": "PN do fornecedor extraído da descrição do item — geralmente aparece entre parênteses no final da descrição. Ex: 'Antenna GPS (GPS-ANT-001)' → 'GPS-ANT-001'. Se não houver parênteses com PN, retornar null.",
        "descricao": "Descrição completa do item",
        "quantidade": 1,
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
"""


def _chamar_gemini(caminho_pdf: str, prompt: str, tentativas: int = 3) -> dict:
    """
    Envia PDF ao Gemini e retorna JSON extraído.
    Retry automático em caso de erro 429 (quota), com espera entre tentativas.
    """
    client = genai.Client(api_key=GOOGLE_API_KEY)

    # Abre como bytes para evitar erro de encoding com nomes de arquivo acentuados (Windows)
    with open(caminho_pdf, "rb") as f:
        arquivo = client.files.upload(
            file=f,
            config=types.UploadFileConfig(display_name="documento.pdf", mime_type="application/pdf"),
        )

    resposta = None
    try:
        for tentativa in range(1, tentativas + 1):
            try:
                resposta = client.models.generate_content(
                    model=GEMINI_MODEL,
                    contents=[arquivo, prompt],
                )
                break  # sucesso — sai do loop
            except Exception as e:
                msg = str(e)
                if "429" in msg or "RESOURCE_EXHAUSTED" in msg:
                    if tentativa < tentativas:
                        # Aguarda 60 s antes de tentar novamente (limite por minuto)
                        time.sleep(60)
                        continue
                raise  # outros erros ou última tentativa — propaga
    finally:
        try:
            client.files.delete(name=arquivo.name)
        except Exception:
            pass

    if resposta is None:
        raise RuntimeError("Gemini não retornou resposta após todas as tentativas.")

    texto = resposta.text.strip()

    # Remove blocos markdown se o modelo os incluir
    texto = re.sub(r"^```(?:json)?\s*", "", texto)
    texto = re.sub(r"\s*```$", "", texto)

    return json.loads(texto)


def extrair_cotacoes(caminho_pdf: str) -> dict:
    """
    Extrai dados de cotações do PDF.
    Retorna dict com chave 'fornecedores'.
    """
    return _chamar_gemini(caminho_pdf, PROMPT_COTACAO)


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
    Retorna dict com chave 'po'.
    Garante que numero_cotacao_ref seja extraído via pdfplumber como fallback.
    """
    dados = _chamar_gemini(caminho_pdf, PROMPT_PO)
    po = dados.get("po", {})

    # Se o Gemini não extraiu o quotation code, busca direto no texto do PDF
    ref = po.get("numero_cotacao_ref")
    if not ref or not re.match(r'202[4-9]\.\d{6,}', str(ref)):
        code = _extrair_quotation_code_pdfplumber(caminho_pdf)
        if code:
            po["numero_cotacao_ref"] = code
            dados["po"] = po

    return dados

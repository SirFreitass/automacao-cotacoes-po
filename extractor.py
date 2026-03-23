"""
extractor.py
------------
Extrai dados estruturados de PDFs de cotação e PO usando a API do Google Gemini.
Gratuito até 1.500 requisições/dia. Suporta PDFs digitais e escaneados.
"""

import json
import re
import time
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
      "tipo_freight": "Prepaid and Add | Free Delivery | ECO Runner | UPS Account | outro",
      "custo_freight": 0.00,
      "forma_pagamento": "Ex: Net 30, Credit Card, COD",
      "validade_cotacao": "Data de validade se mencionada",
      "numero_cotacao": "Número da cotação — procure padrões como 2025.XXXXXX ou 2026.XXXXXX (ex: 2025.123456). Se não encontrar este padrão, use o número que aparecer.",
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
    "fornecedor_escolhido_comentario": "Nome do fornecedor/fabricante real mencionado nos comentários do comprador ou observações da PO. Ex: se o comentário diz 'tng telecom' ou 'purchasing process made by eco - tng telecom', retornar 'tng telecom'. Se não houver menção de fornecedor nos comentários, retornar null.",
    "numero_eco_req": "Número da REQ ECO no formato numérico (ex: 031326015461). Procure por 'REQ#', 'REQ:', 'Requisition' ou sequências longas de números que identifiquem a solicitação nos comentários ou cabeçalho da PO.",
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


def extrair_po(caminho_pdf: str) -> dict:
    """
    Extrai dados da PO do PDF.
    Retorna dict com chave 'po'.
    """
    return _chamar_gemini(caminho_pdf, PROMPT_PO)

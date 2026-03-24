"""
utils.py
--------
Funções utilitárias compartilhadas entre os módulos do sistema.
"""

import re


def norm_pn(pn: str) -> str:
    """Normaliza PN removendo separadores para comparação robusta."""
    return re.sub(r'[^A-Z0-9]', '', (pn or "").upper())


def norm_vendor(nome: str) -> str:
    """Normaliza nome de fornecedor para lookup: minúsculo, sem separadores."""
    return re.sub(r'[^a-z0-9]', '', (nome or "").lower())


_RE_QUOTATION_CODE = re.compile(r'202[4-9]\.\d{6,}')


def numero_cotacao(analise: dict) -> str | None:
    """Pega o número de cotação do primeiro fornecedor que tiver."""
    for forn in analise.get("ranking_preco", []):
        nc = forn.get("numero_cotacao")
        if nc:
            return str(nc)
    return None


def quotation_code(analise: dict) -> str:
    """
    Retorna o Quotation Code no formato 20XX.XXXXXX extraído da PO.
    Retorna string vazia se não encontrar no formato válido.
    """
    po = analise.get("po", {})
    ref = po.get("numero_cotacao_ref")
    if ref and _RE_QUOTATION_CODE.match(str(ref)):
        return str(ref)
    return ""


def normalizar_freight(tipo: str, custo) -> str:
    """
    Normaliza o tipo de freight para o padrão ECO.
    Sempre retorna uma das 4 opções válidas:
      "UPS Account", "Runner Pick up", "Free Delivery", "Supplier Ship"
    """
    t = (tipo or "").lower()
    if "ups" in t:
        return "UPS Account"
    if "eco runner" in t or "runner" in t or "coleta" in t or "pick up" in t:
        return "Runner Pick up"
    if "free" in t or "no charge" in t or "no freight" in t or "no cost" in t or "included" in t:
        return "Free Delivery"
    # Qualquer outro caso → Supplier Ship
    return "Supplier Ship"


def normalizar_freight_robo(tipo: str, custo) -> str:
    """
    Converte o tipo de freight para o vocabulário do Req-o-matic (pos sheet col G).
    Sempre retorna uma das 4 opções válidas:
      "UPS Account", "Runner Pick up", "Free Delivery", "Supplier Ship"
    """
    t = (tipo or "").lower()
    if "ups" in t:
        return "UPS Account"
    if "runner" in t or "eco runner" in t or "coleta" in t or "pick up" in t:
        return "Runner Pick up"
    if "free" in t or "no charge" in t or "no freight" in t or "no cost" in t or "included" in t:
        return "Free Delivery"
    # Qualquer outro caso (custo > 0, texto desconhecido, vazio) → Supplier Ship
    return "Supplier Ship"

"""
utils.py
--------
Funções utilitárias compartilhadas entre os módulos do sistema.
"""

import re


def norm_pn(pn: str) -> str:
    """Normaliza PN removendo separadores para comparação robusta."""
    return re.sub(r'[^A-Z0-9]', '', (pn or "").upper())


def numero_cotacao(analise: dict) -> str | None:
    """Pega o número de cotação do primeiro fornecedor que tiver."""
    for forn in analise.get("ranking_preco", []):
        nc = forn.get("numero_cotacao")
        if nc:
            return str(nc)
    return None


def normalizar_freight(tipo: str, custo) -> str:
    """
    Normaliza o tipo de freight para o padrão ECO:
    - UPS Account  → mantém (exceção — usa conta UPS da ECO)
    - ECO Runner   → mantém (coleta feita pela ECO)
    - Free Delivery → mantém (sem custo de frete)
    - Qualquer outro com custo incluído → "Supplier Ship"
    """
    t = (tipo or "").lower()
    if "ups" in t:
        return tipo or ""
    if "eco runner" in t or "runner" in t or "coleta" in t:
        return tipo or ""
    if "free" in t or "no charge" in t or "no freight" in t:
        return tipo or ""
    if custo or "prepaid" in t or "add" in t or "include" in t or "ship" in t or "freight" in t:
        return "Supplier Ship"
    return tipo or ""


def normalizar_freight_robo(tipo: str, custo) -> str:
    """
    Converte o tipo de freight para o vocabulário do Req-o-matic (pos sheet col G).
    Mapeamento:
      UPS Account  → "UPS Account"
      ECO Runner   → "Runner Pick up"
      Free/No charge → "Free Delivery"
      Supplier Ship / com custo → "Supplier Ship"
    """
    t = (tipo or "").lower()
    if "ups" in t:
        return "UPS Account"
    if "runner" in t or "eco runner" in t or "coleta" in t:
        return "Runner Pick up"
    if "free" in t or "no charge" in t or "no freight" in t:
        return "Free Delivery"
    if custo or "prepaid" in t or "add" in t or "include" in t or "ship" in t or "freight" in t:
        return "Supplier Ship"
    return tipo or ""

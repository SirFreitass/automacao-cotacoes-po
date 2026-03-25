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
    Retorna o Quotation Code no formato 20XX.XXXXXX.
    Cadeia de busca:
      1. PO → numero_cotacao_ref
      2. ranking_preco → primeiro fornecedor com numero_cotacao
      3. resumo_fornecedores → primeiro fornecedor com numero_cotacao
    """
    # 1. Da PO (extraído dos comentários/referências)
    po = analise.get("po", {})
    ref = po.get("numero_cotacao_ref")
    if ref and _RE_QUOTATION_CODE.match(str(ref)):
        return str(ref)

    # 2. Do ranking de preços (fornecedores da cotação)
    for forn in analise.get("ranking_preco", []):
        nc = forn.get("numero_cotacao")
        if nc and _RE_QUOTATION_CODE.match(str(nc)):
            return str(nc)

    # 3. Do resumo de fornecedores
    for forn in analise.get("resumo_fornecedores", []):
        nc = forn.get("numero_cotacao")
        if nc and _RE_QUOTATION_CODE.match(str(nc)):
            return str(nc)

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


# Alias para compatibilidade — mesma lógica de normalizar_freight
normalizar_freight_robo = normalizar_freight


# ── Normalização de UOM — mapa canônico único ──────────────────────────
# Chaves em lowercase; valores EXATOS como aparecem no dropdown do ECO.
# Importado por excel_exporter.py e eco_playwright.py.
UOM_MAP = {
    "each": "each", "ea": "each", "pc": "each", "pcs": "each", "piece": "each",
    "unit": "each", "un": "each", "und": "each", "units": "each",
    "box": "box", "bx": "box",
    "case": "case", "cs": "case",
    "cm": "cm",
    "cu yd": "cu yd", "cubic yard": "cu yd",
    "day": "day", "days": "day",
    "dm": "dm",
    "dozen": "dozen", "dz": "dozen", "doz": "dozen",
    "drum": "drum",
    "feet": "feet", "ft": "feet", "foot": "feet",
    "gal": "gal", "gallon": "gal", "gallons": "gal",
    "hour": "hour", "hr": "hour", "hrs": "hour",
    "lb": "lb", "lbs": "lb", "pound": "lb", "pounds": "lb",
    "liter": "liter", "ltr": "liter", "l": "liter", "litre": "liter",
    "meter": "meter", "m": "meter", "mtr": "meter", "metre": "meter",
    "miles": "miles", "mi": "miles", "mile": "miles",
    "month": "month", "mo": "month", "months": "month",
    "oz": "oz", "ounce": "oz", "ounces": "oz",
    "pack": "pack", "pk": "pack",
    "pail": "pail",
    "pair": "pair", "pr": "pair", "pairs": "pair",
    "quart": "quart", "qt": "quart",
    "set": "set", "kit": "set", "lot": "set",
    "sq ft": "sq ft", "sqft": "sq ft", "square foot": "sq ft", "square feet": "sq ft",
    "ton": "ton", "tons": "ton",
    "week": "week", "wk": "week", "weeks": "week",
}


def normalizar_uom(uom_raw: str) -> str:
    """Normaliza UOM extraída para o nome padrão do ECO Requisition."""
    u = (uom_raw or "each").strip().lower()
    return UOM_MAP.get(u, u)

# =============================================================
# CONFIGURAÇÕES DO SISTEMA — EXEMPLO
# =============================================================
# Copie este arquivo para config.py e preencha com seus dados.
# NUNCA commite o config.py no Git (ele está no .gitignore).
# =============================================================

# Chave da API do Google Gemini (gratuita em https://aistudio.google.com/apikey)
GOOGLE_API_KEY = "SUA_CHAVE_AQUI"

# Modelo Gemini a ser usado
GEMINI_MODEL = "gemini-2.5-flash"

# Caminho completo para a planilha Req-o-matic (ajuste para sua máquina)
ROBO_PLANILHA = r"C:\Caminho\Para\Req-o-matic v3.6.94 - POs.xlsm"

# Planilha de GL Codes por embarcação (ajuste para sua máquina)
GL_CODE_PLANILHA = r"C:\Caminho\Para\Brazil Vessels - GL CODE.xlsx"

# Tipos de freight reconhecidos
FREIGHT_TYPES = [
    "Prepaid and Add",
    "Free Delivery",
    "ECO Runner",
    "UPS Account",
]

# Mapeamento de tipo_freight (cotação) → opção "Ship VIA" no ECO Requisition
SHIP_VIA_MAP = {
    "prepaid and add": "Supplier Ship",
    "free delivery":   "Free delivery",
    "eco runner":      "Runner Pick up",
    "ups account":     "ECO UPS ACCT# 707185",
}

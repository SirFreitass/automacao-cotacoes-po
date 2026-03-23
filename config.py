# =============================================================
# CONFIGURAÇÕES DO SISTEMA
# =============================================================
# Coloque sua chave da API do Google (Gemini) abaixo.
# Para obter sua chave GRATUITA: https://aistudio.google.com/apikey
# 1. Acesse o link acima e faça login com sua conta Google
# 2. Clique em "Create API Key"
# 3. Copie a chave e cole abaixo
# =============================================================

GOOGLE_API_KEY = "AIzaSyArQqhDZS1YHA3i3eV9L_MWFr6-EH6qQik"

# Modelo Gemini a ser usado (gemini-2.0-flash = gratuito)
GEMINI_MODEL = "gemini-2.5-flash"

# Caminho para a planilha Req-o-matic (usada para lookup de fornecedores na Tabela Forn)
# Ajuste se a planilha estiver em outro local
ROBO_PLANILHA = r"C:\Users\freit\Automação - Analise de Cotações + PO\Req-o-matic v3.6.94 - POs.xlsm"

# Planilha de GL Codes por embarcação (Brazil Vessels - GL CODE.xlsx)
GL_CODE_PLANILHA = r"C:\Users\freit\Automação - Analise de Cotações + PO\Brazil Vessels - GL CODE.xlsx"

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

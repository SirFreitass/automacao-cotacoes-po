# CLAUDE.md — Automação Cotações & PO (ECO Purchasing)

## O que é este projeto
App desktop (Tkinter + Python) que:
1. **Analisa** PDFs de cotações e POs usando Google Gemini (IA)
2. **Exporta** planilha Excel com ranking de fornecedores e dados extraídos
3. **Automatiza** o lançamento de POs no sistema ECO Requisition via Playwright (browser automation)

## Arquivos principais

| Arquivo | Função |
|---|---|
| `main.py` | Interface Tkinter, orquestra análise e automação |
| `extractor.py` | Extrai dados de PDFs via Gemini (2 passadas: itens + metadados) |
| `analyzer.py` | Compara cotações, monta ranking de preços, valida itens |
| `excel_exporter.py` | Gera planilha Excel de resultado e lê planilha ROBO |
| `eco_playwright.py` | Robô Playwright: preenche ECO Requisition no browser |
| `utils.py` | Funções compartilhadas: normalização de vendor, UOM, freight, quotation code |
| `config.py` | **NUNCA commitar** — chaves API e caminhos locais (está no .gitignore) |
| `config.example.py` | Exemplo público de config.py |
| `vendor_map.json` | Cache de aprendizado: nome extraído → nome exato no ECO |

## Regras de negócio críticas

### Quotation Code
- Formato obrigatório: `20[2-9][0-9]\.\d{6,}` (ex: `2024.123456`)
- Anos válidos: 2024–2029 apenas
- Regex em `utils.py`: `_RE_QUOTATION_CODE = re.compile(r'202[4-9]\.\d{6,}')`
- Nunca usar números de REQ ou PO como cotação

### Vendor (Fornecedor)
- **Nautical Ventures** é broker/emissora — NUNCA é o fornecedor real. Está na blacklist.
- Cadeia de resolução: comentários da PO → observações → cotação (melhor preço)
- Padrão buyer comments: `ECO REQ#:XXXXXXX - VESSEL - vendor`
- `vendor_map.json` é cache de aprendizado automático (runtime, não commitar valores errados)
- `norm_vendor()` em utils.py: minúsculo + só alfanumérico para chave de lookup

### Freight — 4 opções válidas no ECO
1. `Supplier Ship`
2. `Free Delivery`
3. `Runner Pick up`
4. `ECO UPS ACCT# 707185`
- Qualquer outro valor deve ser mapeado por `normalizar_freight()` em utils.py
- Nunca criar novas opções de freight fora dessas 4

### Quantidades
- O robô **NUNCA edita quantidades** — apenas lê e alerta se houver divergência
- Quantidade é informativa, não deve ser alterada automaticamente

### GL Code
- Campo no ECO usa **Kendo UI** (não Material Design)
- Seletor correto: `kendo-combobox` (não `mat-select`)

## Fluxo do robô ECO (eco_playwright.py)

```
1. Busca ECO REQ# no campo de pesquisa
2. Entra na requisição
3. Approve (se botão visível) → reload → aguarda Angular
4. Check out (se botão visível) — usa get_by_role("button", name="Check out", exact=True)
5. Preenche campos: vendor, Ship VIA, GL Code, Unit Price, UOM
6. Order (submete)
```

**Atenção:**
- Após Approve: sempre recarregar a página (`page.reload()`) e aguardar `domcontentloaded` + 2s para Angular re-renderizar
- Check out: usar `get_by_role("button", name="Check out", exact=True)` — evita clicar em "Uncheck out"
- `networkidle` NÃO usar — ECO tem polling contínuo que nunca "settle"

## Segurança
- `config.py` está no `.gitignore` e **NUNCA deve ser commitado**
- Contém `GOOGLE_API_KEY` — se exposta, revogar imediatamente no Google AI Studio
- Usar `config.example.py` como referência pública

## Stack técnica
- Python 3.11+
- `google-genai` (Gemini 2.5 Flash) — extração de PDFs
- `playwright` async — automação ECO
- `openpyxl` — leitura/escrita Excel
- `pdfplumber` — fallback de extração de texto de PDF
- `tkinter` — interface desktop

## Colunas da planilha de saída (aba ROBO)
A=Quotation Code, B=Produto, C=Description, D=Unit Price, E=Cost Center Desc,
F=Supplier, G=Freight, H=OBS, I=ECO REQ, J=Observação PO, K=Forn. Extraido,
L=Forn Extraído ECO, M=PO

- Detecção de colunas por **header normalizado** (não por posição fixa) — suporta formatos antigo (23 cols) e novo (14 cols)
- Vendor para o robô vem da coluna **F (Supplier)**

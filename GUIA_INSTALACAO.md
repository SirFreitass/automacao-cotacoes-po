# Guia de Instalação e Uso

## Requisitos
- Windows 10/11
- Conexão com a internet (para instalação e uso da API Claude)

---

## Passo 1 — Instalar o Python

1. Acesse: https://www.python.org/downloads/
2. Clique em **"Download Python 3.11.x"**
3. Execute o instalador
4. ⚠️ **IMPORTANTE**: Marque a opção **"Add Python to PATH"** antes de clicar em Install
5. Clique em **"Install Now"**

---

## Passo 2 — Instalar as dependências

1. Abra a pasta do programa
2. Clique com o botão direito em um espaço vazio da pasta
3. Selecione **"Abrir no Terminal"** (ou "Open in Terminal")
4. Digite o comando abaixo e pressione Enter:

```
pip install -r requirements.txt
```

Aguarde a instalação terminar.

---

## Passo 3 — Obter a chave da API Google Gemini (GRATUITA)

1. Acesse: https://aistudio.google.com/apikey
2. Faça login com sua conta Google
3. Clique em **"Create API Key"**
4. Copie a chave gerada
5. Abra o arquivo **`config.py`** com o Bloco de Notas
6. Substitua `COLE-SUA-CHAVE-AQUI` pela sua chave
7. Salve o arquivo

> **Custo**: GRATUITO — o plano free cobre até 1.500 análises/dia.
> Seu uso estimado (~40/dia) usa apenas 3% do limite gratuito.

---

## Passo 4 — Usar o programa

1. Dê duplo clique em **`main.py`**
   - Ou clique com botão direito → Abrir com → Python
2. Na janela que abrir:
   - Clique em **"Selecionar PDF..."** ao lado de "Cotações" e escolha o PDF com as cotações
   - Clique em **"Selecionar PDF..."** ao lado de "PO" e escolha o PDF da Purchase Order
   - Escolha onde salvar o Excel (opcional)
3. Clique em **"ANALISAR"**
4. Aguarde (10–30 segundos dependendo do tamanho dos PDFs)
5. O Excel será aberto automaticamente com o resultado

---

## O que o programa analisa

### Cotações (comparativo entre fornecedores):
- Melhor preço total
- Melhor prazo de entrega
- Tipo de freight (Prepaid and Add, Free Delivery, ECO Runner, UPS Account)
- Forma de pagamento
- Itens com substitutos ou similares (alerta)

### PO (conformidade com melhor cotação):
- Preço total coincide com a cotação selecionada
- Custo de freight incluído corretamente
- Part Numbers (PN) coincidindo
- Preços unitários por item coincidindo

---

## Resultado — Planilha Excel gerada

| Aba | Conteúdo |
|-----|----------|
| **Resumo Análise** | Tabela comparativa colorida (verde = melhor preço, amarelo = alerta) |
| **Alertas PO** | Lista de divergências encontradas entre PO e cotação |
| **Dados VBA** | Dados estruturados prontos para sua planilha de emissão de PO |

---

## Problemas comuns

| Problema | Solução |
|----------|---------|
| "Módulo não encontrado" | Execute o Passo 2 novamente |
| "API Key inválida" | Verifique se colou a chave correta no config.py |
| "Erro de conexão" | Verifique sua internet |
| PDF não lido corretamente | Envie um PDF de exemplo para ajuste do sistema |

---

## Suporte
Em caso de dúvidas, abra o programa e peça ajuda ao Claude Code.

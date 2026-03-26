"""
analyzer.py
-----------
Compara cotações entre fornecedores e valida a PO contra o fornecedor escolhido.
"""
import re
from datetime import date, timedelta

from utils import norm_pn as _norm_pn, aprender_vendor as _aprender_vendor


def _parse_data(s) -> date | None:
    """Converte string YYYY-MM-DD em objeto date. Retorna None se inválido."""
    if not s:
        return None
    try:
        return date.fromisoformat(str(s)[:10])
    except ValueError:
        return None


def _checar_validade(forn: dict, hoje: date) -> dict:
    """
    Determina status de validade da cotação.
    Prioridade: validade_cotacao (ISO) → data_cotacao + 30 dias.
    Retorna dict com validade_data_iso, validade_dias_restantes, validade_vencida.
    """
    validade = _parse_data(forn.get("validade_cotacao"))
    if validade is None:
        emissao = _parse_data(forn.get("data_cotacao"))
        if emissao:
            validade = emissao + timedelta(days=30)

    if validade is None:
        return {"validade_data_iso": None, "validade_dias_restantes": None, "validade_vencida": None}

    dias = (validade - hoje).days
    return {
        "validade_data_iso": validade.isoformat(),
        "validade_dias_restantes": dias,
        "validade_vencida": dias < 0,
    }


from utils import normalizar_freight as _normalizar_freight


def analisar(dados_cotacao: dict, dados_po: dict) -> dict:
    """
    Recebe dados extraídos pelo extractor.py e retorna análise completa.

    Retorna:
    {
        "ranking_preco": [...],          # Fornecedores ordenados por preço
        "ranking_prazo": [...],          # Fornecedores ordenados por prazo
        "melhor_preco": {...},           # Fornecedor com menor preço total
        "melhor_prazo": {...},           # Fornecedor com menor prazo
        "alertas_itens": [...],          # Itens com substitutos / similares
        "alertas_po": [...],             # Divergências PO vs cotação
        "resumo_fornecedores": [...],    # Tabela completa para Excel
        "po": {...},                     # Dados da PO
    }
    """
    fornecedores = dados_cotacao.get("fornecedores", [])
    po = dados_po.get("po", {})

    resultado = {
        "ranking_preco": [],
        "ranking_prazo": [],
        "melhor_preco": None,
        "melhor_prazo": None,
        "alertas_itens": [],
        "alertas_po": [],
        "resumo_fornecedores": [],
        "fornecedor_resolvido": None,   # nome definitivo — resolvido UMA vez aqui
        "po": po,
    }

    if not fornecedores:
        return resultado

    # --- Ranking por preço ---
    fornecedores_com_preco = [f for f in fornecedores if f.get("preco_total") is not None]
    ranking_preco = sorted(fornecedores_com_preco, key=lambda f: f["preco_total"])
    resultado["ranking_preco"] = ranking_preco
    if ranking_preco:
        resultado["melhor_preco"] = ranking_preco[0]

    # --- Ranking por prazo ---
    fornecedores_com_prazo = [f for f in fornecedores if f.get("prazo_entrega_dias") is not None]
    ranking_prazo = sorted(fornecedores_com_prazo, key=lambda f: f["prazo_entrega_dias"])
    resultado["ranking_prazo"] = ranking_prazo
    if ranking_prazo:
        resultado["melhor_prazo"] = ranking_prazo[0]

    # --- Alertas de itens similares / substitutos ---
    for forn in fornecedores:
        for item in forn.get("itens", []):
            if item.get("item_identico_ao_solicitado") is False:
                resultado["alertas_itens"].append({
                    "fornecedor": forn.get("nome"),
                    "pn": item.get("pn"),
                    "descricao": item.get("descricao"),
                    "observacao": item.get("observacao_item"),
                })

    # --- Resumo para Excel ---
    melhor_preco_nome = (resultado["melhor_preco"] or {}).get("nome", "")
    melhor_prazo_nome = (resultado["melhor_prazo"] or {}).get("nome", "")
    hoje = date.today()

    for i, forn in enumerate(ranking_preco):
        posicao_prazo = next(
            (j + 1 for j, f in enumerate(ranking_prazo) if f.get("nome") == forn.get("nome")),
            "-"
        )
        tem_substituto = any(
            not item.get("item_identico_ao_solicitado", True)
            for item in forn.get("itens", [])
        )
        validade_info = _checar_validade(forn, hoje)
        resultado["resumo_fornecedores"].append({
            "posicao_preco": i + 1,
            "posicao_prazo": posicao_prazo,
            "nome": forn.get("nome"),
            "preco_total": forn.get("preco_total"),
            "moeda": forn.get("moeda", "USD"),
            "prazo_entrega": forn.get("prazo_entrega"),
            "prazo_entrega_dias": forn.get("prazo_entrega_dias"),
            "tipo_freight": _normalizar_freight(forn.get("tipo_freight"), forn.get("custo_freight")),
            "custo_freight": forn.get("custo_freight"),
            "forma_pagamento": forn.get("forma_pagamento"),
            "tem_item_substituto": "SIM" if tem_substituto else "Não",
            "numero_cotacao": forn.get("numero_cotacao"),
            "validade_cotacao": validade_info["validade_data_iso"] or forn.get("validade_cotacao"),
            "validade_vencida": validade_info["validade_vencida"],
            "validade_dias_restantes": validade_info["validade_dias_restantes"],
        })

    # --- Alertas de validade vencida ---
    for forn_resumo in resultado["resumo_fornecedores"]:
        if forn_resumo.get("validade_vencida"):
            dias = abs(forn_resumo.get("validade_dias_restantes") or 0)
            resultado["alertas_po"].append({
                "tipo": "VALIDADE",
                "severidade": "AVISO",
                "mensagem": (
                    f"Cotação de '{forn_resumo.get('nome')}' está vencida há {dias} dia(s) "
                    f"(validade: {forn_resumo.get('validade_cotacao')}). "
                    f"Confirme com o fornecedor se os preços ainda são válidos."
                ),
            })

    # --- Validação da PO ---
    if not po:
        return resultado

    alertas = resultado["alertas_po"]

    # Identifica o fornecedor escolhido pelo comprador (via comentários da PO)
    forn_comentario = (po.get("fornecedor_escolhido_comentario") or "").lower().strip()

    if forn_comentario:
        # Busca o fornecedor escolhido na lista de cotações
        forn_escolhido = next(
            (f for f in fornecedores
             if forn_comentario in (f.get("nome") or "").lower()
             or (f.get("nome") or "").lower() in forn_comentario),
            None
        )
        if forn_escolhido is None:
            alertas.append({
                "tipo": "FORNECEDOR",
                "severidade": "INFO",
                "mensagem": (
                    f"Fornecedor indicado nos comentários da PO ('{po.get('fornecedor_escolhido_comentario')}') "
                    f"não foi localizado nas cotações recebidas. Verifique se o nome corresponde."
                ),
            })
        else:
            # Verifica se o escolhido é o de melhor preço
            melhor = resultado["melhor_preco"]
            if melhor and forn_escolhido.get("nome") != melhor.get("nome"):
                preco_escolhido = forn_escolhido.get("preco_total") or 0
                preco_melhor = melhor.get("preco_total") or 0
                diferenca = preco_escolhido - preco_melhor
                pct = (diferenca / preco_melhor * 100) if preco_melhor else 0
                alertas.append({
                    "tipo": "FORNECEDOR",
                    "severidade": "INFO",
                    "mensagem": (
                        f"Comprador escolheu '{forn_escolhido.get('nome')}' (${preco_escolhido:,.2f}), "
                        f"mas o menor preço é de '{melhor.get('nome')}' (${preco_melhor:,.2f}) — "
                        f"diferença de ${diferenca:,.2f} ({pct:+.1f}%). "
                        f"Verifique se a escolha está justificada."
                    ),
                })
    else:
        # Sem fornecedor nos comentários: usa o de melhor preço como referência
        forn_escolhido = resultado["melhor_preco"]

    # Usa o fornecedor escolhido como referência para validação
    referencia = forn_escolhido or resultado["melhor_preco"]

    # ── Fornecedor resolvido — UMA única vez, propagado a todos os módulos ──
    if referencia:
        resultado["fornecedor_resolvido"] = referencia.get("nome")
        # Aprende mapeamento fornecedor comentário → nome na cotação
        if forn_comentario and referencia.get("nome"):
            _aprender_vendor(forn_comentario, referencia.get("nome"))
    elif forn_comentario:
        resultado["fornecedor_resolvido"] = po.get("fornecedor_escolhido_comentario")
    if not referencia:
        return resultado

    # Freight: verifica apenas se o fornecedor escolhido cobra freight
    tipo_freight = _normalizar_freight(
        referencia.get("tipo_freight"), referencia.get("custo_freight")
    ).lower()
    freight_cotacao = referencia.get("custo_freight") or 0
    freight_po = po.get("custo_freight") or 0

    if "prepaid and add" in tipo_freight and freight_po == 0:
        alertas.append({
            "tipo": "FREIGHT",
            "severidade": "ALERTA",
            "mensagem": (
                f"Cotação de '{referencia.get('nome')}' tem freight 'Prepaid and Add' "
                f"(${freight_cotacao:,.2f}), mas a PO não inclui custo de freight. Verifique!"
            ),
        })
    elif freight_cotacao > 0 and freight_po == 0:
        alertas.append({
            "tipo": "FREIGHT",
            "severidade": "AVISO",
            "mensagem": (
                f"Cotação de '{referencia.get('nome')}' inclui freight de ${freight_cotacao:,.2f}, "
                f"mas não foi identificado freight na PO."
            ),
        })

    # Comparação item a item: usa pn_fornecedor como chave primária
    itens_ref = {}
    for item in referencia.get("itens", []):
        pn = (item.get("pn") or "").upper()
        if pn:
            itens_ref[pn] = item
    # Índice normalizado (sem hífens/espaços) para match robusto
    itens_ref_norm = {_norm_pn(k): v for k, v in itens_ref.items()}

    for item_po in po.get("itens", []):
        # Usa pn_fornecedor (extraído dos parênteses da descrição) como PN primário
        pn_busca = (item_po.get("pn_fornecedor") or item_po.get("pn") or "").upper()
        if not pn_busca:
            continue

        pn_busca_norm = _norm_pn(pn_busca)

        item_ref = (
            itens_ref.get(pn_busca)
            or next((v for k, v in itens_ref.items() if pn_busca in k or k in pn_busca), None)
            or (itens_ref_norm.get(pn_busca_norm) if pn_busca_norm else None)
            or next((v for k, v in itens_ref_norm.items()
                     if pn_busca_norm and (pn_busca_norm in k or k in pn_busca_norm)), None)
        )

        if item_ref is None:
            alertas.append({
                "tipo": "PART NUMBER",
                "severidade": "ALERTA",
                "mensagem": (
                    f"PN '{pn_busca}' da PO não encontrado na cotação de "
                    f"'{referencia.get('nome')}'. Verifique se o item está correto."
                ),
            })

    return resultado

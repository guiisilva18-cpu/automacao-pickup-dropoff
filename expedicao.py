"""Expedição: pacotes expedidos, veículos e % de ocupação por PA, extraído
direto da tabela `expedicao_pa` no TiDB (mesmo banco do App Ponto de Apoio,
alimentada pela planilha Google Sheets "CONTROLE DIARIO" via
App Ponto de Apoio/mysql_source.sincronizar_expedicao_pa). Pedido do
Guilherme (2026-08-18): slide de Expedição na apresentação semanal, falando
de veículos.

Mesma fórmula de % de ocupação já usada no App Ponto de Apoio
(dashboard_logic.montar_kpis_expedicao): pacotes_expedidos / (qtde_veiculos
* capacidade_kg_do_perfil) -- fica None quando o perfil do veículo não é
reconhecido (combinado tipo "TOCO E TRUCK", ou "DSR" de folga). Tabela de
capacidade copiada de App Ponto de Apoio/config.py -- só muda se o
Guilherme atualizar lá também.
"""
from collections import defaultdict

from dotenv import load_dotenv

import gravar_mysql

load_dotenv()

# Mesmas 19 PAs ativas de App Ponto de Apoio/config.PAS_ATIVAS (corte de
# 01/08/2026, Guilherme) -- expedicao_pa ainda tem PA MELI-SMR-SP e
# PA PEGAKI-VCP-SP no histórico bruto, que saíram do escopo e não devem
# entrar nos totais.
PAS_ATIVAS = {
    "PA MELI-CJM-SP", "PA MELI-BRE-SP", "PA MELI-CJM 14-SP", "PA MELI-CJM 04-SP",
    "PA MELI-CJM 02-SP", "PA INFRA-SP", "PA VIA-SP", "PA NESTLE-SP",
    "PA RENNER-CAB-SP", "PA MANDAE-SP", "PA SATELITAL-SP", "PA CEA-SP",
    "PA OLIST-SP", "PA ESTOCA-VGP-SP", "PA SATELITAL-VGP-SP", "PA CUBBO-EMB-SP",
    "PA WEPINK-ITP 02-SP", "PA WEPINK-ITP-SP", "PA SHOPEE-BRE-SP",
}

CAPACIDADE_VEICULO_KG = {
    "CARRETA": 15000,
    "CAVALO": 15000,
    "TRUCK": 4800,
    "TOCO": 3600,
    "TOCO P": 3600,
    "3/4": 2400,
    "VUC": 1680,
    "VAN": 960,
    "FURGAO": 960,
    "FURGÃO": 960,
    "HR-FURGAO": 600,
    "HR-FURGÃO": 600,
    "ONIBUS": 480,
    "ÔNIBUS": 480,
    "CAMIONETE": 360,
    "FIORINO": 360,
    "UTILITARIO": 360,
    "UTILITÁRIO": 360,
    "PASSEIO": 240,
    "CORREIOS": 120,
    "MOTO": 48,
}


def buscar_expedicao(inicio: str, fim: str) -> dict:
    conexao = gravar_mysql._conectar()
    try:
        with conexao.cursor() as cur:
            cur.execute(
                "SELECT base_remetente, pacotes_expedidos, perfil_veiculo, qtde_veiculos "
                "FROM expedicao_pa WHERE data_referencia BETWEEN %s AND %s",
                (inicio, fim),
            )
            linhas = cur.fetchall()
    finally:
        conexao.close()

    por_pa = defaultdict(lambda: {"pacotes": 0, "veiculos": 0, "ultimo_perfil": None, "perfis": defaultdict(int)})
    veiculos_por_perfil = defaultdict(int)
    for base, pacotes, perfil, qtde in linhas:
        if base not in PAS_ATIVAS:
            continue
        info = por_pa[base]
        info["pacotes"] += pacotes or 0
        info["veiculos"] += qtde or 0
        if perfil:
            info["ultimo_perfil"] = perfil  # mesma regra do App Ponto de Apoio: "last" por PA
        if perfil and perfil != "DSR" and qtde:
            veiculos_por_perfil[perfil] += qtde
            info["perfis"][perfil] += qtde

    pas = []
    for base, info in por_pa.items():
        perfil = info["ultimo_perfil"]
        capacidade = CAPACIDADE_VEICULO_KG.get(perfil) if perfil else None
        pacotes, qtde = info["pacotes"], info["veiculos"]
        media = round(pacotes / qtde, 1) if qtde else None
        pct_ocupacao = round(pacotes / (qtde * capacidade) * 100, 1) if qtde and capacidade else None

        # Veículo de fato mais usado pela PA no período (maior soma de
        # qtde_veiculos por perfil) -- diferente de `perfil_veiculo` acima,
        # que é só o último perfil visto (mesma regra do App Ponto de
        # Apoio). Pedido do Guilherme (2026-08-18): slide com o carro mais
        # usado por PA de verdade, não só o do último dia.
        perfil_mais_usado, qtde_perfil_mais_usado = (None, None)
        if info["perfis"]:
            perfil_mais_usado, qtde_perfil_mais_usado = max(info["perfis"].items(), key=lambda kv: kv[1])

        pas.append({
            "pa": base,
            "pacotes_expedidos": pacotes,
            "qtde_veiculos": qtde,
            "perfil_veiculo": perfil,
            "capacidade": capacidade,
            "pct_ocupacao": pct_ocupacao,
            "media_por_veiculo": media,
            "perfil_mais_usado": perfil_mais_usado,
            "qtde_perfil_mais_usado": qtde_perfil_mais_usado,
        })
    pas.sort(key=lambda r: -r["pacotes_expedidos"])

    total_pacotes = sum(r["pacotes_expedidos"] for r in pas)
    total_veiculos = sum(r["qtde_veiculos"] for r in pas)
    perfil_ranking = sorted(veiculos_por_perfil.items(), key=lambda kv: -kv[1])

    return {
        "periodo": (inicio, fim),
        "pas": pas,
        "total_pacotes_expedidos": total_pacotes,
        "total_veiculos": total_veiculos,
        "media_por_veiculo": round(total_pacotes / total_veiculos, 1) if total_veiculos else None,
        "veiculos_por_perfil": [{"perfil": p, "total": q} for p, q in perfil_ranking],
    }

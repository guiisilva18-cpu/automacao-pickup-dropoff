"""Grava o snapshot diário (D-1) de PICKUP/DROPOFF no mesmo TiDB que
alimenta o dashboard App Ponto de Apoio (tela "Bases", gráficos Pickup e
Dropoff, 2026-08-12) — só o dia roda aqui; o consolidado mensal completo
continua indo só por e-mail pro Vanilson, não fica no banco.

Mesmas variáveis de ambiente (DB_HOST/DB_PORT/DB_USER/DB_PASSWORD/DB_NAME/
DB_SSL) e mesmo padrão de conexão do App Ponto de Apoio (config.py +
mysql_source.py) — TiDB Cloud exige TLS.
"""
import logging
import os
import ssl

import pymysql

log = logging.getLogger(__name__)


def _conectar():
    config = {
        "host": os.environ["DB_HOST"],
        "port": int(os.environ.get("DB_PORT", 3306)),
        "user": os.environ["DB_USER"],
        "password": os.environ["DB_PASSWORD"],
        "database": os.environ["DB_NAME"],
        "charset": "utf8mb4",
    }
    if os.environ.get("DB_SSL", "false").lower() == "true":
        config["ssl"] = ssl.create_default_context()
    return pymysql.connect(**config, autocommit=True)


def gravar_pickup(dia: str, registros: list[dict]):
    conexao = _conectar()
    try:
        with conexao.cursor() as cur:
            for r in registros:
                cur.execute(
                    "INSERT INTO pickup_diario "
                    "(data_referencia, base_remetente, qtd_a_coletar, qtd_coletada_no_prazo, "
                    "taxa_real_pct, taxa_com_tentativas_pct, taxa_poc_pct) "
                    "VALUES (%s, %s, %s, %s, %s, %s, %s) "
                    "ON DUPLICATE KEY UPDATE "
                    "qtd_a_coletar = VALUES(qtd_a_coletar), "
                    "qtd_coletada_no_prazo = VALUES(qtd_coletada_no_prazo), "
                    "taxa_real_pct = VALUES(taxa_real_pct), "
                    "taxa_com_tentativas_pct = VALUES(taxa_com_tentativas_pct), "
                    "taxa_poc_pct = VALUES(taxa_poc_pct)",
                    (
                        dia, r["base"], r["qtd_a_coletar"], r["qtd_coletada_no_prazo"],
                        r["taxa_real_pct"] or 0, r["taxa_com_tentativas_pct"] or 0,
                        r["taxa_poc_pct"] or 0,
                    ),
                )
    finally:
        conexao.close()
    log.info("pickup_diario gravado no banco para %s (%s bases)", dia, len(registros))


def gravar_dropoff(dia: str, registros: list[dict]):
    conexao = _conectar()
    try:
        with conexao.cursor() as cur:
            for r in registros:
                cur.execute(
                    "INSERT INTO dropoff_diario "
                    "(data_referencia, base_remetente, pendente, coletado, total, taxa_pct) "
                    "VALUES (%s, %s, %s, %s, %s, %s) "
                    "ON DUPLICATE KEY UPDATE "
                    "pendente = VALUES(pendente), coletado = VALUES(coletado), "
                    "total = VALUES(total), taxa_pct = VALUES(taxa_pct)",
                    (dia, r["base"], r["pendente"], r["coletado"], r["total"], r["taxa_pct"] or 0),
                )
    finally:
        conexao.close()
    log.info("dropoff_diario gravado no banco para %s (%s bases)", dia, len(registros))

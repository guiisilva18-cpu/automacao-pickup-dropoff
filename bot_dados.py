"""Dados do resumo diário do bot do Feishu (Pickup, Dropoff, Transferência,
Expedição, previsões e assertividade). Tudo sai do que as automações já
gravam no TiDB + consultas ao vivo no JMS -- nada aqui grava dado de
negócio, só as tabelas próprias do bot (previsao_coleta_diaria,
resumo_feishu_envio).

Linha do tempo (rodando no dia D, de manhã): resultados são de D-1 (a
Transferência usa a janela D-2 a D-1, a mesma do e-mail); previsões são pro
próprio dia D; assertividade compara a previsão feita em D-1 com o
realizado de D-1.
"""
import logging
import re
import time
from collections import defaultdict
from datetime import date, datetime
from io import BytesIO
from zoneinfo import ZoneInfo

import requests
from openpyxl import load_workbook

import bot_regras as regras
import expedicao as exp
import extrair_dropoff as ed
import extrair_pickup as ep
import extrair_transferencia as et
import gravar_mysql

log = logging.getLogger(__name__)

FUSO = ZoneInfo("America/Sao_Paulo")

# Regra de previsão dos P.As pedida pelo Guilherme (24/09/2026): coletado no
# dia anterior + 300 pacotes, por P.A, valendo pra todos (Meli incluídos --
# a previsão oficial do hub_coleta ficou de fora, é manual e o bot é 100%
# automático).
INCREMENTO_PA = 300

# Planilha "CONTROLE DIARIO" de Expedição (mesma de App Ponto de Apoio/
# config.URL_EXPEDICAO_XLSX; leitura pública, sem login).
ID_PLANILHA_EXPEDICAO = "1o4CAdCFfT7QLXUn5Qu2MdD_XMPYF3aoFL-1VZ-H8ftw"
APELIDOS_PA = {"PA CUBBO-BEM-SP": "PA CUBBO-EMB-SP", "PA WEPINK-ITP2-SP": "PA WEPINK-ITP 02-SP"}

PAS_ATIVAS = exp.PAS_ATIVAS
BASES = ep.BASES_PICKUP

DDL = [
    """CREATE TABLE IF NOT EXISTS previsao_coleta_diaria (
        data_previsao DATE NOT NULL,
        tipo VARCHAR(10) NOT NULL,
        nome VARCHAR(150) NOT NULL,
        previsto INT NOT NULL,
        coletado_no_envio INT NOT NULL DEFAULT 0,
        fonte VARCHAR(60) NULL,
        gerado_em TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (data_previsao, tipo, nome)
    )""",
    """CREATE TABLE IF NOT EXISTS resumo_feishu_envio (
        data_envio DATE NOT NULL PRIMARY KEY,
        iniciado_em TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        status VARCHAR(20) NOT NULL
    )""",
]


def hoje_sp() -> date:
    return datetime.now(FUSO).date()


def _linhas(conn, sql: str, params: tuple = ()) -> list[dict]:
    with conn.cursor() as cur:
        cur.execute(sql, params)
        cols = [c[0] for c in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]


def _f(v):
    return float(v) if v is not None else None


# ---------------------------------------------------------------- banco / lock
def garantir_tabelas(conn):
    with conn.cursor() as cur:
        for ddl in DDL:
            cur.execute(ddl)


def tentar_lock(conn, dia: date) -> bool:
    """Garante 1 envio por dia mesmo com o workflow disparado 3x (uma por
    automação que termina). Lock 'iniciado' há mais de 45 min é considerado
    de uma execução que morreu e pode ser retomado."""
    with conn.cursor() as cur:
        cur.execute("INSERT IGNORE INTO resumo_feishu_envio (data_envio, status) VALUES (%s, 'iniciado')", (dia,))
        if cur.rowcount == 1:
            return True
        cur.execute(
            "UPDATE resumo_feishu_envio SET iniciado_em = NOW() "
            "WHERE data_envio = %s AND status = 'iniciado' AND iniciado_em < NOW() - INTERVAL 45 MINUTE",
            (dia,),
        )
        return cur.rowcount == 1


def liberar_lock(conn, dia: date):
    with conn.cursor() as cur:
        cur.execute("DELETE FROM resumo_feishu_envio WHERE data_envio = %s AND status = 'iniciado'", (dia,))


def marcar_enviado(conn, dia: date):
    with conn.cursor() as cur:
        cur.execute(
            "INSERT INTO resumo_feishu_envio (data_envio, status) VALUES (%s, 'enviado') "
            "ON DUPLICATE KEY UPDATE status = 'enviado'",
            (dia,),
        )


def aguardar_resumo_pa(conn, d1: date, limite_min: int = 20) -> list[str]:
    """resumo_pa de D-1 (fechamento por P.A) é base da previsão e da
    assertividade dos P.As; vem da carga diária do JMS (~06:00 BRT)."""
    fim = time.time() + limite_min * 60
    while True:
        n = _linhas(conn, "SELECT COUNT(*) AS n FROM resumo_pa WHERE data_referencia = %s", (d1,))[0]["n"]
        if n > 0:
            return []
        if time.time() >= fim:
            return [f"Fechamento dos P.As de {d1:%d/%m} (resumo_pa) ainda não carregou — previsão dos P.As pode estar vazia."]
        time.sleep(60)


def obter_pickup_dropoff_d1(conn, d1: date, d2: date, limite_min: int = 50, passo_s: int = 300):
    """Pickup e Dropoff de D-1 ao vivo no JMS, sem esperar a carga das 09:00.
    O relatório de Indicadores de Negócios (bigdataReport) só fecha D-1 por
    volta das 09:00 BRT (rodando às 05h/06h devolve tudo zerado -- achado de
    03/09/2026), então o Pickup só é aceito quando o total de D-1 chega a pelo
    menos 50% do de D-2 (referência no banco); antes disso espera `passo_s` e
    consulta de novo, até `limite_min`. Vencido o prazo, cai no que a carga
    das 09:00 já gravou no banco e, se ainda estiver vazio, segue com aviso.
    Devolve (pickup, dropoff, avisos)."""
    avisos = []
    ref = int(_linhas(conn, "SELECT COALESCE(SUM(qtd_a_coletar), 0) AS t FROM pickup_diario WHERE data_referencia = %s", (d2,))[0]["t"])
    minimo = ref * 0.5
    fim = time.time() + limite_min * 60

    pickup, fechado = None, False
    while True:
        try:
            pickup = pickup_d1_ao_vivo(d1)
        except Exception:
            log.exception("Falha ao buscar o Pickup de %s ao vivo", d1)
            pickup = None
        total = sum(r["qtd_a_coletar"] for r in pickup) if pickup else 0
        fechado = pickup is not None and total > 0 and total >= minimo
        if fechado or time.time() >= fim:
            break
        log.info("Pickup de %s ainda parece aberto (total %s, referência D-2 %s); nova consulta em %ss", d1, total, ref, passo_s)
        time.sleep(passo_s)
    if not fechado:
        do_banco = pickup_d1(conn, d1)
        if do_banco and sum(r["qtd_a_coletar"] for r in do_banco) >= minimo:
            pickup = do_banco
        else:
            avisos.append(f"Pickup de {d1:%d/%m} ainda não fechou no JMS — números podem estar incompletos.")

    dropoff = None
    try:
        dropoff = dropoff_d1_ao_vivo(d1)
    except Exception:
        log.exception("Falha ao buscar o Dropoff de %s ao vivo", d1)
    if not dropoff:
        dropoff = dropoff_d1(conn, d1)
        if not dropoff:
            avisos.append(f"Dropoff de {d1:%d/%m} ainda não fechou no JMS.")
    return pickup, dropoff, avisos


# ---------------------------------------------------------------- D-1: Pickup / Dropoff
def pickup_d1(conn, d1: date) -> list[dict]:
    linhas = _linhas(
        conn,
        "SELECT base_remetente AS base, qtd_a_coletar, qtd_coletada_no_prazo, taxa_real_pct, "
        "taxa_com_tentativas_pct, taxa_poc_pct FROM pickup_diario WHERE data_referencia = %s",
        (d1,),
    )
    soma_exata = {}
    try:
        for r in _linhas(
            conn, "SELECT base_remetente, timely_try_taking_num FROM taxa_tentativa_coleta WHERE data_referencia = %s", (d1,)
        ):
            soma_exata[r["base_remetente"]] = r["timely_try_taking_num"]
    except Exception:
        log.warning("taxa_tentativa_coleta indisponível; 'soma coletados + tentativas' será derivada da taxa")
    saida = []
    for r in linhas:
        qtd = int(r["qtd_a_coletar"] or 0)
        if qtd <= 0:
            continue
        tent = _f(r["taxa_com_tentativas_pct"]) or 0.0
        saida.append({
            "data": d1,
            "base": r["base"],
            "qtd_a_coletar": qtd,
            "coletada_no_prazo": int(r["qtd_coletada_no_prazo"] or 0),
            "soma_tentativas": int(soma_exata.get(r["base"], round(qtd * tent / 100))),
            "taxa_real": _f(r["taxa_real_pct"]) or 0.0,
            "taxa_tentativas": tent,
            "taxa_poc": _f(r["taxa_poc_pct"]) or 0.0,
        })
    saida.sort(key=lambda x: -x["taxa_tentativas"])
    return saida


def pickup_d1_ao_vivo(d1: date) -> list[dict]:
    """Mesma extração do e-mail do Pickup (extrair_pickup.buscar_pickup)."""
    saida = []
    for r in ep.buscar_pickup(d1.isoformat()):
        qtd = int(r["qtd_a_coletar"] or 0)
        if qtd <= 0:
            continue
        saida.append({
            "data": d1, "base": r["base"], "qtd_a_coletar": qtd,
            "coletada_no_prazo": int(r["qtd_coletada_no_prazo"] or 0),
            "soma_tentativas": int(r["soma_coletados_tentativas"] or 0),
            "taxa_real": float(r["taxa_real_pct"] or 0), "taxa_tentativas": float(r["taxa_com_tentativas_pct"] or 0),
            "taxa_poc": float(r["taxa_poc_pct"] or 0),
        })
    saida.sort(key=lambda x: -x["taxa_tentativas"])
    return saida


def dropoff_d1_ao_vivo(d1: date) -> list[dict]:
    """Mesma extração do e-mail do Dropoff (extrair_dropoff.buscar_dropoff)."""
    saida = [
        {"data": d1, "base": r["base"], "pendente": int(r["pendente"] or 0), "coletado": int(r["coletado"] or 0),
         "total": int(r["total"] or 0), "taxa": float(r["taxa_pct"] or 0)}
        for r in ed.buscar_dropoff(d1.isoformat()) if int(r["total"] or 0) > 0
    ]
    saida.sort(key=lambda x: -x["taxa"])
    return saida


def dropoff_d1(conn, d1: date) -> list[dict]:
    linhas = _linhas(
        conn,
        "SELECT base_remetente AS base, pendente, coletado, total, taxa_pct FROM dropoff_diario WHERE data_referencia = %s",
        (d1,),
    )
    saida = [
        {"data": d1, "base": r["base"], "pendente": int(r["pendente"] or 0), "coletado": int(r["coletado"] or 0),
         "total": int(r["total"] or 0), "taxa": _f(r["taxa_pct"]) or 0.0}
        for r in linhas if int(r["total"] or 0) > 0
    ]
    saida.sort(key=lambda x: -x["taxa"])
    return saida


# ---------------------------------------------------------------- Transferência (JMS ao vivo)
def transferencia(d2: date, d1: date) -> list[dict]:
    """Mesma janela do e-mail (D-2 a D-1). Linhas sem entrega no período
    (taxa None) ficam de fora, como no print-modelo; ordem = taxa desc,
    estável (PAs antes das bases em empate).

    extrair_transferencia.buscar_transferencia consulta as 19 PAs +
    14 bases franquia hardcoded lá (BASES_TRANSFERENCIA/
    BASES_FRANQUIA_TRANSFERENCIA), listas próprias que não acompanham
    automaticamente os cortes de roster feitos em config.PAS_ATIVAS/
    extrair_pickup.BASES_PICKUP -- por isso filtra aqui pelas ATIVAS atuais
    antes de devolver (achado 07/10/2026, Guilherme: "tire o que não é
    mais" -- a tabela ainda trazia INFRA/MANDAE/CEA/OLIST/COT-SP/OSC-SP/
    CARAP-SP etc., todas já fora do escopo)."""
    regs = et.buscar_transferencia(d2.isoformat(), d1.isoformat())
    regs = [r for r in regs if r["taxa_pct"] is not None]
    ativos = PAS_ATIVAS | set(ep.BASES_PICKUP)
    regs = [r for r in regs if r["base"] in ativos]
    regs.sort(key=lambda r: -r["taxa_pct"])
    return regs


# ---------------------------------------------------------------- Expedição
def _capacidade(perfil, qtde):
    """Mesma regra de dashboard_logic._capacidade_perfil_expedicao."""
    if not perfil:
        return None, None
    tabela = exp.CAPACIDADE_VEICULO_KG
    if perfil in tabela:
        cap = tabela[perfil]
        return cap, (cap * qtde if qtde else None)
    partes = [p.strip() for p in re.split(r"\s+E\s+|(?<!\d)/(?!\d)", perfil) if p.strip()]
    if len(partes) < 2:
        return None, None
    caps = [tabela.get(p) for p in partes]
    if any(c is None for c in caps):
        return None, None
    return sum(caps), sum(caps)


def _num_celula(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return int(v) if v == v else None
    try:
        return int(float(str(v).strip().replace(",", ".")))
    except ValueError:
        return None


def sincronizar_expedicao(conn) -> bool:
    """Atualiza expedicao_pa a partir da planilha dos líderes, igual ao
    App Ponto de Apoio (mysql_source.sincronizar_expedicao_pa), mas sem
    depender do dashboard estar aberto. Nunca levanta: se falhar, vale o
    que já estiver no banco."""
    try:
        resp = requests.get(
            f"https://docs.google.com/spreadsheets/d/{ID_PLANILHA_EXPEDICAO}/export?format=xlsx", timeout=30
        )
        resp.raise_for_status()
        wb = load_workbook(BytesIO(resp.content), data_only=True, read_only=True)
    except Exception:
        log.exception("Falha ao baixar a planilha de Expedição; usando o que já está no banco")
        return False

    hoje = hoje_sp()
    linhas = []
    for ws in wb.worksheets:
        nome = ws.title.strip()
        data_aba = None
        for texto in (nome, f"{nome}.{hoje.year}"):
            try:
                data_aba = datetime.strptime(texto, "%d.%m.%Y").date()
                break
            except ValueError:
                continue
        if data_aba is None or data_aba > hoje:
            continue
        linhas_aba = list(ws.iter_rows(values_only=True))
        if len(linhas_aba) < 3:
            continue
        cab = [str(c).strip() if c is not None else "" for c in linhas_aba[1]]
        low = [c.lower() for c in cab]
        try:
            i_pa = next(i for i, c in enumerate(cab) if c.upper() == "PA")
        except StopIteration:
            continue
        i_lider = next((i for i, c in enumerate(low) if "lider" in c or "líder" in c), None)
        i_vol = next((i for i, c in enumerate(low) if "vol" in c and "di" in c), None)
        i_perfil = next((i for i, c in enumerate(low) if "perfil" in c), None)
        i_qtde = next((i for i, c in enumerate(low) if "qtde" in c and "ve" in c), None)
        for linha in linhas_aba[2:]:
            if i_pa >= len(linha) or linha[i_pa] is None or not str(linha[i_pa]).strip():
                continue
            pa = APELIDOS_PA.get(str(linha[i_pa]).strip(), str(linha[i_pa]).strip())
            lider = None
            if i_lider is not None and i_lider < len(linha) and linha[i_lider] is not None:
                lider = str(linha[i_lider]).strip() or None
            pacotes = _num_celula(linha[i_vol]) if i_vol is not None and i_vol < len(linha) else None
            perfil = None
            if i_perfil is not None and i_perfil < len(linha) and linha[i_perfil] is not None:
                bruto = linha[i_perfil]
                if isinstance(bruto, (datetime, date)):
                    perfil = f"{bruto.day}/{bruto.month}"  # "3/4" vira data no Google Sheets
                else:
                    perfil = str(bruto).strip().upper()
                    if perfil in ("", "0", "0.0", "NAN"):
                        perfil = None
            qtde = _num_celula(linha[i_qtde]) if i_qtde is not None and i_qtde < len(linha) else None
            linhas.append((data_aba, pa, lider, pacotes, perfil, qtde))

    if not linhas:
        log.warning("Planilha de Expedição sem abas no formato DD.MM; nada sincronizado")
        return False

    # Às vezes 2 líderes preenchem o mesmo PA no mesmo dia em linhas
    # separadas (ex: um só com pacotes, outro só com veículo) -- sem
    # agregar, o ON DUPLICATE KEY UPDATE abaixo faz a última linha apagar
    # os pacotes/veículos da primeira (achado em auditoria de 25/09/2026,
    # sumia 81 mil pacotes e 7 carretas do PA MELI-CJM-SP em 24/09).
    agrupado: dict[tuple, dict] = {}
    for data_aba, pa, lider, pacotes, perfil, qtde in linhas:
        acc = agrupado.setdefault((data_aba, pa), {"lideres": [], "pacotes": None, "perfis": [], "qtde": None})
        if lider and lider not in acc["lideres"]:
            acc["lideres"].append(lider)
        if pacotes is not None:
            acc["pacotes"] = (acc["pacotes"] or 0) + pacotes
        if perfil and perfil not in acc["perfis"]:
            acc["perfis"].append(perfil)
        if qtde is not None:
            acc["qtde"] = (acc["qtde"] or 0) + qtde
    linhas = [
        (data_aba, pa, " / ".join(acc["lideres"]) or None, acc["pacotes"],
         " E ".join(acc["perfis"]) or None, acc["qtde"])
        for (data_aba, pa), acc in agrupado.items()
    ]

    with conn.cursor() as cur:
        cur.executemany(
            "INSERT INTO expedicao_pa (data_referencia, base_remetente, lider, pacotes_expedidos, "
            "perfil_veiculo, qtde_veiculos) VALUES (%s, %s, %s, %s, %s, %s) "
            "ON DUPLICATE KEY UPDATE lider = VALUES(lider), pacotes_expedidos = VALUES(pacotes_expedidos), "
            "perfil_veiculo = VALUES(perfil_veiculo), qtde_veiculos = VALUES(qtde_veiculos)",
            linhas,
        )
    log.info("Expedição sincronizada da planilha: %s linhas", len(linhas))
    return True


def expedicao_d1(conn, d1: date) -> dict | None:
    """Mesma tabela/cards da tela Expedição do dash (dashboard_logic.
    montar_kpis_expedicao) pra 1 dia."""
    rows = [
        r for r in _linhas(
            conn,
            "SELECT base_remetente, lider, pacotes_expedidos, perfil_veiculo, qtde_veiculos "
            "FROM expedicao_pa WHERE data_referencia = %s",
            (d1,),
        )
        if r["base_remetente"] in PAS_ATIVAS
    ]
    if not rows:
        return None

    por_pa = {}
    for r in rows:
        por_pa[r["base_remetente"]] = r  # 1 linha por PA/dia (chave primária)
    linhas = []
    for pa, r in por_pa.items():
        pacotes = int(r["pacotes_expedidos"]) if r["pacotes_expedidos"] is not None else None
        qtde = int(r["qtde_veiculos"]) if r["qtde_veiculos"] is not None else None
        capacidade, cap_total = _capacidade(r["perfil_veiculo"], qtde)
        linhas.append({
            "pa": pa,
            "lider": r["lider"],
            "pacotes": pacotes,
            "perfil": r["perfil_veiculo"],
            "qtde_veiculos": qtde,
            "capacidade": capacidade,
            "pct_ocupacao": round(pacotes / cap_total * 100, 1) if pacotes is not None and cap_total else None,
            "media": round(pacotes / qtde, 1) if pacotes is not None and qtde else None,
        })
    linhas.sort(key=lambda x: -(x["pacotes"] if x["pacotes"] is not None else -1))
    # Só mostra P.A com veículo ou ocupação preenchidos (Guilherme, 24/09/2026:
    # "ocupação 0 e veículo 0, aí não mande") -- os totais abaixo seguem os do dash.
    linhas = [r for r in linhas if (r["qtde_veiculos"] or 0) > 0 or (r["pct_ocupacao"] or 0) > 0]

    total_pacotes = sum(r["pacotes_expedidos"] or 0 for r in rows)
    total_veiculos = sum(r["qtde_veiculos"] or 0 for r in rows)
    por_perfil = defaultdict(int)
    for r in rows:
        if r["perfil_veiculo"] and r["perfil_veiculo"] != "DSR" and r["qtde_veiculos"] is not None:
            por_perfil[r["perfil_veiculo"]] += int(r["qtde_veiculos"])
    mais_usado = max(por_perfil.items(), key=lambda kv: kv[1]) if por_perfil else None
    return {
        "linhas": linhas,
        "total_pacotes": int(total_pacotes),
        "total_veiculos": int(total_veiculos),
        "media_geral": round(total_pacotes / total_veiculos, 1) if total_veiculos else None,
        "mais_usado": mais_usado,
    }


# ---------------------------------------------------------------- previsão: fontes ao vivo
def snapshot_pickup_ao_vivo(dia: date):
    """Uma única consulta ao JMS (taxa de coleta no prazo, dia = hoje):
    Bases -> só TikTok (deveria/coletado); P.As -> coletado sem filtro de
    origem (cada P.A é de 1 cliente, não é tudo TikTok)."""
    brutos = ep._buscar_registros_brutos(dia.isoformat())
    bases = {b: {"deveria": 0, "coletado": 0} for b in BASES}
    pas = defaultdict(int)
    for r in brutos:
        nome = (r.get("pickNetworkName") or "").strip()
        deveria = r.get("shouldTakingNum", 0) or 0
        coletado = r.get("timelyTakingNum", 0) or 0
        if nome in bases and r.get("orderSourcename") == ep.ORIGEM_PEDIDO_FILTRO:
            bases[nome]["deveria"] += deveria
            bases[nome]["coletado"] += coletado
        if nome in PAS_ATIVAS:
            pas[nome] += coletado
    return bases, dict(pas)


def dropoff_pendente_ao_vivo(dias: list[date]) -> dict[str, dict[date, int]]:
    """Pedidos TikTok que já entraram no Yoyi e aguardam coleta (orderType=3)
    por dia de entrada -- só a busca de PENDENTE (a de COLETADO é pesada e
    não é necessária pra previsão)."""
    saida = {b: {d: 0 for d in dias} for b in BASES}
    for d in dias:
        for r in ed._buscar_registros_brutos(d.isoformat(), ed.ORDER_TYPE_PENDENTE):
            nome = (r.get("pickNetworkName") or "").strip()
            if nome in saida and r.get("orderSourceName") == ed.ORIGEM_PEDIDO_FILTRO:
                saida[nome][d] += 1
    return saida


# ---------------------------------------------------------------- previsão: montagem
def coletados_d1_por_pa(conn, d1: date) -> dict[str, int]:
    rows = _linhas(
        conn,
        "SELECT base_remetente, COALESCE(SUM(coletados), 0) AS c FROM resumo_pa WHERE data_referencia = %s GROUP BY base_remetente",
        (d1,),
    )
    return {r["base_remetente"]: int(r["c"]) for r in rows if r["base_remetente"] in PAS_ATIVAS}


def previsao_pas(conn, d1: date, coletado_ao_vivo: dict[str, int]) -> dict:
    """previsto = coletados de D-1 (fechamento, resumo_pa) + INCREMENTO_PA; já
    coletado = posição ao vivo no JMS. Ficam de fora os P.As Meli (pedido do
    Guilherme, 24/09/2026) e os P.As sem movimento ontem nem hoje (não faz
    sentido prever 300 pra quem está parado)."""
    col_d1 = coletados_d1_por_pa(conn, d1)
    linhas, sem_movimento = [], []
    for pa in sorted(PAS_ATIVAS):
        if regras.eh_meli(pa):
            continue
        anterior = col_d1.get(pa, 0)
        vivo = coletado_ao_vivo.get(pa, 0)
        if anterior == 0 and vivo == 0:
            sem_movimento.append(pa)
            continue
        linhas.append({"pa": pa, "previsto": int(anterior + INCREMENTO_PA), "coletado": int(vivo),
                       "coletado_d1": int(anterior), "fonte": f"D-1 + {INCREMENTO_PA}"})
    total_prev = sum(r["previsto"] for r in linhas)
    total_col = sum(r["coletado"] for r in linhas)
    return {
        "linhas": linhas,
        "sem_movimento": sem_movimento,
        "total_previsto": total_prev,
        "total_coletado": total_col,
        "total_pendente": max(total_prev - total_col, 0),
    }


def previsao_bases(snapshot_bases: dict) -> list[dict]:
    linhas = []
    for base, v in snapshot_bases.items():
        pendente = max(v["deveria"] - v["coletado"], 0)
        linhas.append({"base": base, "deveria": v["deveria"], "coletado": v["coletado"], "pendente": pendente})
    linhas.sort(key=lambda x: -x["pendente"])
    return linhas


def previsao_dropoff(pendente_por_base: dict, d1: date, hoje: date) -> list[dict]:
    linhas = []
    for base, por_dia in pendente_por_base.items():
        p1, p0 = por_dia.get(d1, 0), por_dia.get(hoje, 0)
        linhas.append({"base": base, "pendente_d1": p1, "pendente_hoje": p0, "previsto": p1 + p0})
    linhas.sort(key=lambda x: -x["previsto"])
    return linhas


# ---------------------------------------------------------------- persistência + assertividade
def gravar_previsao(conn, hoje: date, bases: list[dict], pas: dict):
    registros = [(hoje, "BASE", b["base"], b["pendente"], b["coletado"], "Pendente Pickup") for b in bases]
    registros += [(hoje, "PA", p["pa"], p["previsto"], 0, p["fonte"]) for p in pas["linhas"]]
    with conn.cursor() as cur:
        cur.executemany(
            "INSERT INTO previsao_coleta_diaria (data_previsao, tipo, nome, previsto, coletado_no_envio, fonte) "
            "VALUES (%s, %s, %s, %s, %s, %s) "
            "ON DUPLICATE KEY UPDATE previsto = VALUES(previsto), coletado_no_envio = VALUES(coletado_no_envio), fonte = VALUES(fonte)",
            registros,
        )


def assertividade(conn, d1: date, final_base: dict[str, int] | None = None) -> dict | None:
    """Compara a previsão feita em D-1 (guardada no envio daquele dia) com o
    realizado de D-1. P.A: realizado = coletados do fechamento (resumo_pa).
    Base: previsto era o pendente do Pickup no momento do envio; realizado =
    quanto foi coletado no prazo DEPOIS do envio (coletada_no_prazo final de
    D-1 menos o que já estava coletado no envio). "Passou" = realizado >=
    previsto."""
    prev = _linhas(
        conn,
        "SELECT tipo, nome, previsto, coletado_no_envio FROM previsao_coleta_diaria WHERE data_previsao = %s",
        (d1,),
    )
    if not prev:
        return None
    final_pa = coletados_d1_por_pa(conn, d1)
    if final_base is None:
        final_base = {
            r["base_remetente"]: int(r["qtd_coletada_no_prazo"] or 0)
            for r in _linhas(conn, "SELECT base_remetente, qtd_coletada_no_prazo FROM pickup_diario WHERE data_referencia = %s", (d1,))
        }
    bases, pas = [], []
    for r in prev:
        previsto = int(r["previsto"])
        if r["tipo"] == "BASE":
            if previsto <= 0 or r["nome"] not in final_base:
                continue
            realizado = max(final_base[r["nome"]] - int(r["coletado_no_envio"]), 0)
            destino = bases
        else:
            if previsto <= 0 or r["nome"] not in final_pa:
                continue
            realizado = final_pa[r["nome"]]
            destino = pas
        destino.append({
            "nome": r["nome"], "previsto": previsto, "realizado": realizado,
            "diferenca": realizado - previsto, "pct": realizado / previsto * 100, "passou": realizado >= previsto,
        })
    bases.sort(key=lambda x: -x["previsto"])
    pas.sort(key=lambda x: -x["previsto"])
    if not bases and not pas:
        return None
    return {"bases": bases, "pas": pas}

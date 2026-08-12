"""Extrai PENDENTE/COLETADO/TOTAL/TAXA por base (JMS > "Relatório Yoyi",
tela de status de pedidos no dropoff) para o mesmo conjunto fixo de 14
bases do PICKUP, gera uma planilha diária (snapshot do dia, sem histórico
acumulado) e manda por e-mail.

Endpoint capturado por rede (Guilherme, 2026-08-11):
POST /businessindicator/bigdataReport/detailDir/datacabin/detail/sdp_post_report_detail
Cada registro é 1 pedido, com `pickNetworkName` (base), `orderSourceName`
(origem) e `orderType`/`orderTypeExport` (status). Testado por tentativa:
- orderType=3 -> "已入库待揽收" = PENDENTE (aguardando coleta)
- orderType=4 -> "已揽收" = COLETADO
O payload não tem filtro de base/origem — filtra tudo do lado de cá
(TikTok + BASES_PICKUP), igual os outros scripts deste projeto.
`timeType` é sempre "enterTime" (= "YoYi入库时间", horário de entrada no
YoYi, único horário que essa tela oferece).
"""
import json
import logging
import os
import smtplib
import ssl
import sys
from collections import defaultdict
from datetime import date, timedelta
from email.message import EmailMessage
from io import BytesIO
from pathlib import Path

import requests
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import gravar_mysql

load_dotenv()

PASTA_BASE = Path(__file__).parent
PASTA_LOGS = PASTA_BASE / "logs"
PASTA_OUTPUT = PASTA_BASE / "output"
PASTA_DADOS = PASTA_BASE / "dados"
PASTA_LOGS.mkdir(exist_ok=True)
PASTA_OUTPUT.mkdir(exist_ok=True)
PASTA_DADOS.mkdir(exist_ok=True)

# Histórico acumulado do mês (Guilherme, 2026-08-11): a busca de COLETADO é
# pesada (pode passar de 100 mil linhas/dia), então o painel mensal NÃO
# pode reprocessar o mês inteiro a cada execução — cada dia roda 1x, salva
# o resultado aqui, e os dias seguintes só leem esse arquivo de volta (sem
# bater na API de novo) pra montar o Consolidado/Geral Diário.
HISTORICO_PATH = PASTA_DADOS / "historico_dropoff.json"


def carregar_historico() -> dict[str, list[dict]]:
    if not HISTORICO_PATH.exists():
        return {}
    return json.loads(HISTORICO_PATH.read_text(encoding="utf-8"))


def salvar_historico(historico: dict[str, list[dict]]):
    HISTORICO_PATH.write_text(
        json.dumps(historico, ensure_ascii=False, indent=2), encoding="utf-8"
    )

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(PASTA_LOGS / "extrair_dropoff.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)

URL_DROPOFF = (
    "https://gw.jtjms-br.com/businessindicator/bigdataReport/detailDir/"
    "datacabin/detail/sdp_post_report_detail"
)

HEADERS = {
    "Accept": "application/json, text/plain, */*",
    "Content-Type": "application/json;charset=UTF-8",
    "lang": "PT",
    "langType": "PT",
}

TAMANHO_PAGINA = 1000
ORIGEM_PEDIDO_FILTRO = "TikTok"

ORDER_TYPE_PENDENTE = 3  # "已入库待揽收"
ORDER_TYPE_COLETADO = 4  # "已揽收"

# Mesmas 14 bases do relatório PICKUP (Guilherme, 2026-08-11).
BASES_PICKUP = [
    "CARAP 02-SP",
    "CARAP-SP",
    "CHM-SP",
    "CLP-SP",
    "COT-SP",
    "F JND-SP",
    "F S-JRG-SP",
    "ITUP-SP",
    "JND 02-SP",
    "JND-SP",
    "OSC 02-SP",
    "OSC-SP",
    "S-CSVD-SP",
    "S-FREG-SP",
]


class TokenExpiradoError(Exception):
    pass


def _buscar_registros_brutos(dia: str, order_type: int) -> list[dict]:
    token = os.environ["JMS_TOKEN_INDICADORES"]
    headers = {**HEADERS, "authToken": token}

    registros = []
    current = 1
    while True:
        payload = {
            "current": current,
            "size": TAMANHO_PAGINA,
            "timeType": "enterTime",
            "orderType": order_type,
            "startTime": f"{dia} 00:00:00",
            "endTime": f"{dia} 23:59:59",
            "countryId": "1",
        }
        resp = requests.post(URL_DROPOFF, headers=headers, json=payload, timeout=30)
        if resp.status_code in (401, 403):
            raise TokenExpiradoError(
                "JMS_TOKEN_INDICADORES expirado ou sem sessão ativa. "
                "Faça login no portal (conta do Vanilson) e atualize JMS_TOKEN_INDICADORES no .env."
            )
        resp.raise_for_status()
        resultado = resp.json()
        if resultado.get("code") != 1:
            raise RuntimeError(f"Erro ao buscar relatório Yoyi (orderType={order_type}): {resultado}")

        dados = resultado.get("data") or {}
        pagina = dados.get("records", [])
        registros.extend(pagina)
        total = dados.get("total", 0)
        if not pagina or len(registros) >= total:
            break
        current += 1
        if current > 300:
            # COLETADO (orderType=4) é cumulativo desde sempre pra quem
            # entrou naquele dia, não fecha nunca — em dias de pico pode
            # passar de 100 mil linhas (rede inteira, não só as 14 bases).
            log.warning(
                "orderType=%s: parou de paginar em 300 páginas / %s registros (total esperado: %s)",
                order_type, len(registros), total,
            )
            break

    return registros


def buscar_dropoff(dia: str) -> list[dict]:
    """Busca PENDENTE (orderType=3) e COLETADO (orderType=4) separadamente,
    filtra por ORIGEM_PEDIDO_FILTRO e BASES_PICKUP, e monta 1 linha por
    base com PENDENTE/COLETADO/TOTAL/TAXA."""
    brutos_pendente = _buscar_registros_brutos(dia, ORDER_TYPE_PENDENTE)
    brutos_coletado = _buscar_registros_brutos(dia, ORDER_TYPE_COLETADO)

    contagem: dict[str, dict[str, int]] = defaultdict(lambda: {"pendente": 0, "coletado": 0})
    for r in brutos_pendente:
        base = (r.get("pickNetworkName") or "").strip()
        if base in BASES_PICKUP and r.get("orderSourceName") == ORIGEM_PEDIDO_FILTRO:
            contagem[base]["pendente"] += 1
    for r in brutos_coletado:
        base = (r.get("pickNetworkName") or "").strip()
        if base in BASES_PICKUP and r.get("orderSourceName") == ORIGEM_PEDIDO_FILTRO:
            contagem[base]["coletado"] += 1

    faltando = sorted(set(BASES_PICKUP) - set(contagem.keys()))
    if faltando:
        log.warning("Bases da lista sem nenhum registro (pendente ou coletado) em %s: %s", dia, ", ".join(faltando))

    registros = []
    for base in BASES_PICKUP:
        v = contagem.get(base, {"pendente": 0, "coletado": 0})
        total = v["pendente"] + v["coletado"]
        taxa_pct = round(v["coletado"] / total * 100, 2) if total else None
        registros.append({
            "base": base,
            "pendente": v["pendente"],
            "coletado": v["coletado"],
            "total": total,
            "taxa_pct": taxa_pct,
        })
    return registros


CABECALHO = [
    "Data 日期",
    "Base 基地",
    "Pendente 待取件",
    "Coletado 取件成功",
    "Total 已扫描总数",
    "Taxa 取件率",
]

COR_HEADER = "C00000"
PREENCHIMENTO_HEADER = PatternFill(start_color=COR_HEADER, end_color=COR_HEADER, fill_type="solid")
FONTE_HEADER = Font(bold=True, color="FFFFFF")


def _formatar_aba(ws, num_linhas: int, num_colunas: int, col_taxa: int):
    for cel in ws[1]:
        cel.font = FONTE_HEADER
        cel.fill = PREENCHIMENTO_HEADER
        cel.alignment = Alignment(wrap_text=True, vertical="center")
    for row in ws.iter_rows(min_row=2, min_col=col_taxa, max_col=col_taxa):
        row[0].number_format = "0.00%"
    ultima_coluna = get_column_letter(num_colunas)
    ws.auto_filter.ref = f"A1:{ultima_coluna}{num_linhas + 1}"
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 30


def _linha_registro(r: dict) -> list:
    return [
        r["pendente"],
        r["coletado"],
        r["total"],
        (r["taxa_pct"] / 100) if r["taxa_pct"] is not None else None,
    ]


def montar_planilha(dia: str, registros: list[dict]) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Taxa de Coletas por Base"
    ws.append(CABECALHO)
    for r in registros:
        ws.append([dia, r["base"], *_linha_registro(r)])
    _formatar_aba(ws, len(registros), len(CABECALHO), col_taxa=6)
    for i, largura in enumerate([12, 14, 12, 12, 14, 12], start=1):
        ws.column_dimensions[get_column_letter(i)].width = largura

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


CABECALHO_CONSOLIDADO = [
    "Base 基地",
    "Pendente 待取件",
    "Coletado 取件成功",
    "Total 已扫描总数",
    "Taxa 取件率",
]


def _consolidar_por_base(dados: list[tuple[str, list[dict]]]) -> list[dict]:
    """Soma pendente/coletado de cada base ao longo dos dias em `dados` e
    recalcula a taxa a partir dos totais (peso pelo volume, não média das
    taxas diárias)."""
    somas: dict[str, dict[str, int]] = defaultdict(lambda: {"pendente": 0, "coletado": 0})
    for _dia, registros in dados:
        for r in registros:
            base = r["base"]
            somas[base]["pendente"] += r["pendente"]
            somas[base]["coletado"] += r["coletado"]

    consolidado = []
    for base in BASES_PICKUP:
        v = somas.get(base, {"pendente": 0, "coletado": 0})
        total = v["pendente"] + v["coletado"]
        taxa_pct = round(v["coletado"] / total * 100, 2) if total else None
        consolidado.append({
            "base": base,
            "pendente": v["pendente"],
            "coletado": v["coletado"],
            "total": total,
            "taxa_pct": taxa_pct,
        })
    return consolidado


def montar_painel_mensal(dados: list[tuple[str, list[dict]]]) -> BytesIO:
    """Painel mensal: aba 'Consolidado' (totais do período, 1 linha por
    base) + aba 'Geral Diário' (1 linha por dia x base, com autofiltro)."""
    wb = Workbook()

    ws_consolidado = wb.active
    ws_consolidado.title = "Consolidado"
    ws_consolidado.append(CABECALHO_CONSOLIDADO)
    consolidado = _consolidar_por_base(dados)
    for r in consolidado:
        ws_consolidado.append([r["base"], *_linha_registro(r)])
    _formatar_aba(ws_consolidado, len(consolidado), len(CABECALHO_CONSOLIDADO), col_taxa=5)
    for i, largura in enumerate([14, 12, 12, 14, 12], start=1):
        ws_consolidado.column_dimensions[get_column_letter(i)].width = largura

    ws_diario = wb.create_sheet(title="Geral Diário")
    ws_diario.append(CABECALHO)
    total_linhas = 0
    for dia, registros in dados:
        for r in registros:
            ws_diario.append([dia, r["base"], *_linha_registro(r)])
            total_linhas += 1
    _formatar_aba(ws_diario, total_linhas, len(CABECALHO), col_taxa=6)
    for i, largura in enumerate([12, 14, 12, 12, 14, 12], start=1):
        ws_diario.column_dimensions[get_column_letter(i)].width = largura

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def salvar_copia_local(nome_arquivo: str, conteudo: bytes) -> Path:
    caminho = PASTA_OUTPUT / nome_arquivo
    caminho.write_bytes(conteudo)
    return caminho


def enviar_email(dia: str, anexos: list[tuple[str, bytes]]):
    remetente = os.environ["EMAIL_SENDER"]
    senha_app = os.environ["EMAIL_APP_PASSWORD"]
    destinatarios = [e.strip() for e in os.environ["EMAIL_TO_DROPOFF"].split(",") if e.strip()]

    msg = EmailMessage()
    msg["Subject"] = f"Taxa de Coletas por Base (DROPOFF) - {dia}"
    msg["From"] = remetente
    msg["To"] = ", ".join(destinatarios)
    msg.set_content(
        f"Segue em anexo a Taxa de Coletas por Base do dia {dia} "
        "(pendente/coletado/total, origem TikTok, horário de entrada no YoYi) "
        "e o painel consolidado do mês até essa data."
    )
    for nome_arquivo, conteudo in anexos:
        msg.add_attachment(
            conteudo,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=nome_arquivo,
        )

    with smtplib.SMTP_SSL("smtp.feishu.cn", 465, context=ssl.create_default_context()) as smtp:
        smtp.login(remetente, senha_app)
        smtp.send_message(msg)


def main():
    dia = sys.argv[1] if len(sys.argv) > 1 else (date.today() - timedelta(days=1)).isoformat()
    dia_dt = date.fromisoformat(dia)
    inicio_mes = dia_dt.replace(day=1)

    log.info("Buscando Taxa de Coletas por Base (DROPOFF) para %s (só esse dia)...", dia)
    registros_dia = buscar_dropoff(dia)

    # Alimenta o gráfico "Dropoff" na tela Bases do App Ponto de Apoio
    # (2026-08-12) — só o snapshot do dia, não bloqueia o e-mail se falhar.
    try:
        gravar_mysql.gravar_dropoff(dia, registros_dia)
    except Exception:
        log.exception("Falha ao gravar dropoff_diario no banco (e-mail segue normalmente)")

    historico = carregar_historico()
    historico[dia] = registros_dia
    # Descarta dias de meses anteriores — o painel é sempre "mês corrente
    # até D-1", não precisa carregar histórico de meses passados.
    historico = {d: r for d, r in historico.items() if d >= inicio_mes.isoformat()}
    salvar_historico(historico)

    dados_mes = [(d, historico[d]) for d in sorted(historico.keys())]

    conteudo_diario = montar_planilha(dia, registros_dia).getvalue()
    conteudo_painel = montar_painel_mensal(dados_mes).getvalue()

    caminho_diario = salvar_copia_local("DROPOFF_diario.xlsx", conteudo_diario)
    caminho_painel = salvar_copia_local("DROPOFF_painel_mensal.xlsx", conteudo_painel)
    log.info("Planilhas salvas em %s e %s", caminho_diario, caminho_painel)

    nome_anexo_diario = f"DROPOFF_{dia}.xlsx"
    nome_anexo_painel = f"DROPOFF_Painel_{inicio_mes.isoformat()[:7]}.xlsx"
    enviar_email(dia, [(nome_anexo_diario, conteudo_diario), (nome_anexo_painel, conteudo_painel)])
    log.info("E-mail enviado para %s", os.environ["EMAIL_TO_DROPOFF"])


if __name__ == "__main__":
    try:
        main()
    except TokenExpiradoError as e:
        log.error(str(e))
        sys.exit(1)

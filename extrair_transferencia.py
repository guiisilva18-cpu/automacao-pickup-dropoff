"""Extrai a "Taxa de Transferência" (JMS > Indicadores de Negócios > tela com
o filtro "Nome da Base Remetente") só pras nossas 19 PAs ativas (mesma lista
de config.PAS_ATIVAS do App Ponto de Apoio) e manda por e-mail.

Descoberta do endpoint (2026-08-14, sessão com o Guilherme):
- URL real: .../composite_waybill/delivery_ontime_sum -- sempre devolve 1
  registro só (é um "sum", soma agregada), mesmo sem filtro nenhum (nesse
  caso soma a rede inteira). O campo que realmente filtra por base é
  networkId/networkCode/networkName -- os campos agentId/agentName/agentCode
  que aparecem junto no payload são fixos (conta do Vanilson) e não fazem
  filtro nenhum, foram descartados.
- Não existe endpoint que devolva todas as bases de uma vez com breakdown
  (diferente do Pickup) -- é 1 chamada por base mesmo. Os `networkId`/
  `networkCode` de cada base foram obtidos 1x via
  businessindicator/getNetworksDecideByType/pageNetworksByType (endpoint que
  alimenta o dropdown "Nome da Base Remetente" no portal) e ficam fixos
  abaixo -- só muda se a base for renomeada/recriada no JMS.
- Janela de data fixa: D-2 até D-1 (pedido do Guilherme, 2026-08-14) --
  timeType=0 = "Estatísticas por data de chegada planejada" (selecionado no
  portal), a estatística de D-1 ainda não fecha completamente por isso o
  atraso de 2 dias.
"""
import logging
import os
import smtplib
import ssl
import sys
from datetime import date, timedelta
from email.message import EmailMessage
from io import BytesIO
from pathlib import Path

import requests
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

load_dotenv()

PASTA_BASE = Path(__file__).parent
PASTA_LOGS = PASTA_BASE / "logs"
PASTA_OUTPUT = PASTA_BASE / "output"
PASTA_LOGS.mkdir(exist_ok=True)
PASTA_OUTPUT.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(PASTA_LOGS / "extrair_transferencia.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)

URL_DELIVERY_ONTIME = (
    "https://gw.jtjms-br.com/businessindicator/bigdataReport/detailDir/"
    "datacabin/composite_waybill/delivery_ontime_sum"
)

HEADERS = {
    "Accept": "application/json, text/plain, */*",
    "Content-Type": "application/json;charset=UTF-8",
    "lang": "PT",
    "langType": "PT",
}

# Mesmas 19 PAs de config.PAS_ATIVAS (App Ponto de Apoio) -- networkId/
# networkCode obtidos via getNetworksDecideByType/pageNetworksByType,
# 2026-08-14 (ver docstring do módulo).
BASES_TRANSFERENCIA = {
    "PA CEA-SP": (1954, "311185"),
    "PA CUBBO-EMB-SP": (3112, "311466"),
    "PA ESTOCA-VGP-SP": (776, "311045"),
    "PA INFRA-SP": (2088, "311184"),
    "PA MANDAE-SP": (1424, "311111"),
    "PA MELI-BRE-SP": (2192, "311210"),
    "PA MELI-CJM 02-SP": (2320, "316356"),
    "PA MELI-CJM 04-SP": (2327, "316362"),
    "PA MELI-CJM 14-SP": (2352, "316366"),
    "PA MELI-CJM-SP": (2178, "311058"),
    "PA NESTLE-SP": (1974, "311189"),
    "PA OLIST-SP": (2126, "311951"),
    "PA RENNER-CAB-SP": (2138, "311202"),
    "PA SATELITAL-SP": (1777, "311055"),
    "PA SATELITAL-VGP-SP": (3110, "311464"),
    "PA SHOPEE-BRE-SP": (911, "311109"),
    "PA VIA-SP": (1875, "311117"),
    "PA WEPINK-ITP 02-SP": (2965, "311456"),
    "PA WEPINK-ITP-SP": (2581, "311398"),
}

# As 14 bases franquia do relatório Pickup/Dropoff (mesma lista de
# extrair_dropoff.BASES_PICKUP) -- networkId/networkCode obtidos em
# 26/08/2026 via businessindicator/getNetworksDecideByType/pageNetworksByType
# (payload precisa de type=336, parentNetworkId=[129], searchKey=<nome>;
# sem esses campos o endpoint devolve erro genérico de "chamada entre
# serviços"). Pedido do Guilherme, 26/08/2026: incluir as bases também na
# Taxa de Transferência, além das 19 PAs.
BASES_FRANQUIA_TRANSFERENCIA = {
    "CARAP 02-SP": (3057, "311459"),
    "CARAP-SP": (1644, "319129"),
    "CHM-SP": (673, "311131"),
    "CLP-SP": (2632, "311407"),
    "COT-SP": (1475, "319118"),
    "F JND-SP": (2170, "311104"),
    "F S-JRG-SP": (2522, "311392"),
    "ITUP-SP": (2057, "311191"),
    "JND 02-SP": (2780, "311439"),
    "JND-SP": (1014, "311115"),
    "OSC 02-SP": (2694, "311423"),
    "OSC-SP": (1140, "311138"),
    "S-CSVD-SP": (989, "311103"),
    "S-FREG-SP": (1419, "311147"),
}

# Em cópia nos e-mails de Pickup/Dropoff/Transferência (pedido do
# Guilherme, 21/08/2026).
EMAILS_COPIA = [
    "patricia.hora@jtexpress.com.br",
    "ingrid.merces@jtexpress.com.br",
    "fernando.santos@jtexpress.com.br",
]


class TokenExpiradoError(Exception):
    pass


def _buscar_registro_base(base: str, network_id: int, network_code: str, inicio: str, fim: str) -> dict:
    token = os.environ["JMS_TOKEN_INDICADORES"]
    headers = {**HEADERS, "authToken": token}
    payload = {
        "current": 1,
        "size": 20,
        "countryId": "1",
        "startTime": f"{inicio} 00:00:00",
        "endTime": f"{fim} 23:59:59",
        "timeType": 0,
        "networkId": network_id,
        "networkCode": network_code,
        "networkName": base,
    }
    resp = requests.post(URL_DELIVERY_ONTIME, headers=headers, json=payload, timeout=30)
    if resp.status_code in (401, 403):
        raise TokenExpiradoError(
            "JMS_TOKEN_INDICADORES expirado ou sem sessão ativa. "
            "Faça login no portal (conta do Vanilson) e atualize JMS_TOKEN_INDICADORES no .env."
        )
    resp.raise_for_status()
    resultado = resp.json()
    if resultado.get("code") != 1:
        raise RuntimeError(f"Erro ao buscar transferência da base {base}: {resultado}")

    registros = (resultado.get("data") or {}).get("records") or []
    return registros[0] if registros else {}


def buscar_transferencia(inicio: str, fim: str) -> list[dict]:
    """1 chamada por site (esse endpoint não devolve todas de uma vez).
    Site sem nenhuma entrega no período vem com os totais zerados. Junta
    as 19 PAs (BASES_TRANSFERENCIA) com as 14 bases franquia
    (BASES_FRANQUIA_TRANSFERENCIA) num relatório só, marcando o tipo de
    cada linha -- pedido do Guilherme, 26/08/2026."""
    saida = []
    todos = (
        [(nome, "PA", ids) for nome, ids in BASES_TRANSFERENCIA.items()]
        + [(nome, "Base", ids) for nome, ids in BASES_FRANQUIA_TRANSFERENCIA.items()]
    )
    for nome, tipo, (network_id, network_code) in todos:
        bruto = _buscar_registro_base(nome, network_id, network_code, inicio, fim)
        total = bruto.get("deliveryTotal", 0) or 0
        no_prazo = bruto.get("deliveryOntimeTotal", 0) or 0
        fora_prazo = bruto.get("deliveryNotontimeTotal", 0) or 0
        taxa_pct = round(no_prazo / total * 100, 2) if total else None
        saida.append({
            "base": nome,
            "tipo": tipo,
            "entregas_total": total,
            "entregas_no_prazo": no_prazo,
            "entregas_fora_prazo": fora_prazo,
            "taxa_pct": taxa_pct,
        })
        if not total:
            log.warning("%s %s sem entregas no período %s a %s", tipo, nome, inicio, fim)
    return saida


CABECALHO = [
    "Base",
    "Tipo",
    "Total de Entregas",
    "Entregues no Prazo",
    "Fora do Prazo",
    "Taxa de Transferência",
]
COR_HEADER = "C00000"
PREENCHIMENTO_HEADER = PatternFill(start_color=COR_HEADER, end_color=COR_HEADER, fill_type="solid")
FONTE_HEADER = Font(bold=True, color="FFFFFF")


def montar_planilha(inicio: str, fim: str, registros: list[dict]) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Taxa de Transferência"
    ws.append(CABECALHO)
    for r in registros:
        ws.append([
            r["base"],
            r["tipo"],
            r["entregas_total"],
            r["entregas_no_prazo"],
            r["entregas_fora_prazo"],
            (r["taxa_pct"] / 100) if r["taxa_pct"] is not None else None,
        ])

    for cel in ws[1]:
        cel.font = FONTE_HEADER
        cel.fill = PREENCHIMENTO_HEADER
        cel.alignment = Alignment(wrap_text=True, vertical="center")
    for row in ws.iter_rows(min_row=2, min_col=6, max_col=6, max_row=len(registros) + 1):
        row[0].number_format = "0.00%"
    ws.auto_filter.ref = f"A1:F{len(registros) + 1}"
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 22
    for i, largura in enumerate([22, 10, 18, 18, 16, 20], start=1):
        ws.column_dimensions[get_column_letter(i)].width = largura

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def salvar_copia_local(nome_arquivo: str, conteudo: bytes) -> Path:
    caminho = PASTA_OUTPUT / nome_arquivo
    caminho.write_bytes(conteudo)
    return caminho


def enviar_email(inicio: str, fim: str, nome_anexo: str, conteudo: bytes):
    remetente = os.environ["EMAIL_SENDER"]
    senha_app = os.environ["EMAIL_APP_PASSWORD"]
    destinatarios = [e.strip() for e in os.environ["EMAIL_TO_TRANSFERENCIA"].split(",") if e.strip()]

    msg = EmailMessage()
    msg["Subject"] = f"Taxa de Transferência - {inicio} a {fim}"
    msg["From"] = remetente
    msg["To"] = ", ".join(destinatarios)
    msg["Cc"] = ", ".join(EMAILS_COPIA)
    msg.set_content(
        f"Segue em anexo a taxa de transferência (por PA e por base franquia, "
        f"data de chegada planejada) do período de {inicio} a {fim}."
    )
    msg.add_attachment(
        conteudo,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=nome_anexo,
    )

    with smtplib.SMTP_SSL("smtp.feishu.cn", 465, context=ssl.create_default_context()) as smtp:
        smtp.login(remetente, senha_app)
        smtp.send_message(msg)


def main():
    if len(sys.argv) > 2:
        inicio, fim = sys.argv[1], sys.argv[2]
    else:
        hoje = date.today()
        inicio = (hoje - timedelta(days=2)).isoformat()
        fim = (hoje - timedelta(days=1)).isoformat()

    log.info("Buscando Taxa de Transferência de %s a %s...", inicio, fim)
    registros = buscar_transferencia(inicio, fim)

    conteudo = montar_planilha(inicio, fim, registros).getvalue()
    caminho = salvar_copia_local("TRANSFERENCIA_atual.xlsx", conteudo)
    log.info("Planilha salva em %s", caminho)

    nome_anexo = f"TRANSFERENCIA_{inicio}_a_{fim}.xlsx"
    enviar_email(inicio, fim, nome_anexo, conteudo)
    log.info("E-mail enviado para %s", os.environ["EMAIL_TO_TRANSFERENCIA"])


if __name__ == "__main__":
    try:
        main()
    except TokenExpiradoError as e:
        log.error(str(e))
        sys.exit(1)

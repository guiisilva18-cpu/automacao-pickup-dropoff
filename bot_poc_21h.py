"""Resumo noturno (21h) da Taxa de Transferência POC por base -- % de
pedidos com FOTO de comprovação de coleta (relatório EPOP do JMS),
confirmado pelo Guilherme 07/10/2026 ("é a transferencia poc mesmo, que
vem as fotos e tudo mais"). Dado de HOJE (fechamento parcial às 21h, "do
mesmo dia" como a tela de referência mostra -- não D-1).

Dois envios, mesma rodada:
1. Cartão resumido (Base + Quantidade de POC + Taxa) pro webhook do
   Feishu, do melhor pro pior (Guilherme: "mostrar primeiro quem ta
   melhor... e o ultimo pior").
2. E-mail com o relatório DETALHADO (1 linha por loja/base, vindo direto
   do Relatório de Monitoramento EPOP), só com as nossas bases
   (BASES_PICKUP) -- webhook de grupo não aceita arquivo, por isso vai
   por e-mail (Guilherme, 07/10/2026: "mande o arquivo por email
   kuan.chen@jtexpress.com.br").

Webhook, palavra-chave e e-mail PRÓPRIOS dessa automação (diferentes do
bot da manhã): FEISHU_WEBHOOK_URL_POC_21H / FEISHU_KEYWORD_POC_21H /
EMAIL_TO_POC. Reusa EMAIL_SENDER/EMAIL_APP_PASSWORD (mesma conta de
e-mail que extrair_pickup.py já usa).

Uso:  python bot_poc_21h.py              envia (cartão + e-mail)
      python bot_poc_21h.py --dry-run    só mostra o que mandaria, não envia nada
"""
import logging
import os
import smtplib
import ssl
import sys
from datetime import date
from email.message import EmailMessage
from io import BytesIO
from pathlib import Path

import bot_dados as bd
import feishu_api
from bot_cards import poc
from dotenv import load_dotenv
from extrair_pickup import BASES_PICKUP, ORIGEM_PEDIDO_FILTRO, _buscar_registros_epop_brutos
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

load_dotenv()

LOG_DIR = Path(__file__).parent / "logs"
LOG_DIR.mkdir(exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / "bot_poc_21h.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)

PREENCHIMENTO_HEADER = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
FONTE_HEADER = Font(bold=True, color="FFFFFF")

CABECALHO_EPOP = [
    "Base", "Código da base", "Loja", "Endereço da loja",
    "Deveria coletar", "Coletado", "Não coletado", "Tem foto POC",
]


def montar_linhas(hoje: date) -> list[dict]:
    """Pega o Pickup de HOJE ao vivo (dia em andamento, pickup_diario do
    banco só fecha D-1 de madrugada) e extrai taxa_poc/pendente por base
    (EPOP -- foto de comprovação de coleta). Ordenado do MELHOR pro PIOR
    (taxa_poc desc)."""
    pickup = bd.pickup_d1_ao_vivo(hoje)

    linhas = []
    for r in pickup:
        deveria = r["qtd_a_coletar"]
        if not deveria:
            continue
        pendente = max(0, r["epop_total"] - r["epop_com_imagem"])
        linhas.append({"base": r["base"], "deveria": deveria, "pendente": pendente, "taxa_poc": r["taxa_poc"]})
    linhas.sort(key=lambda r: -r["taxa_poc"])
    return linhas


def montar_excel_epop(hoje: date) -> bytes:
    """Relatório detalhado (1 linha por loja/base) do dia, só com as
    nossas bases (BASES_PICKUP) e origem TikTok -- mesmo filtro bruto que
    extrair_pickup.buscar_pickup já aplica pra agregar, só que aqui não
    agrega, exporta linha a linha (pedido do Guilherme: "o relatório todo
    em baixo... somente com nossas bases")."""
    brutos = _buscar_registros_epop_brutos(hoje.isoformat())
    linhas = [
        r for r in brutos
        if (r.get("pickNetworkName") or "").strip() in BASES_PICKUP
        and r.get("orderSourceName") == ORIGEM_PEDIDO_FILTRO
    ]
    linhas.sort(key=lambda r: (r.get("pickNetworkName") or "", r.get("shopName") or ""))

    wb = Workbook()
    ws = wb.active
    ws.title = "POC detalhado"
    ws.append(CABECALHO_EPOP)
    for cel in ws[1]:
        cel.font = FONTE_HEADER
        cel.fill = PREENCHIMENTO_HEADER
    for r in linhas:
        ws.append([
            r.get("pickNetworkName"), r.get("pickNetworkCode"), r.get("shopName"), r.get("shopAddress"),
            r.get("shouldTakeQty"), r.get("takenQty"), r.get("notTakenQty"),
            "Sim" if r.get("signaturePictureUrl") else "Não",
        ])
    ws.auto_filter.ref = f"A1:{get_column_letter(len(CABECALHO_EPOP))}{len(linhas) + 1}"
    larguras = [14, 14, 30, 50, 16, 12, 14, 14]
    for i, largura in enumerate(larguras, start=1):
        ws.column_dimensions[get_column_letter(i)].width = largura

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def enviar_email_excel(hoje: date, conteudo: bytes):
    remetente = os.environ["EMAIL_SENDER"]
    senha_app = os.environ["EMAIL_APP_PASSWORD"]
    destinatarios = [e.strip() for e in os.environ["EMAIL_TO_POC"].split(",") if e.strip()]

    msg = EmailMessage()
    msg["Subject"] = f"POC detalhado - {hoje:%d/%m/%Y}"
    msg["From"] = remetente
    msg["To"] = ", ".join(destinatarios)
    msg.set_content(
        f"Segue em anexo o relatório detalhado de POC (Transferência -- foto de "
        f"comprovação de coleta) do dia {hoje:%d/%m/%Y}, por loja/base (só as "
        "nossas bases, origem TikTok)."
    )
    msg.add_attachment(
        conteudo, maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"POC_detalhado_{hoje:%Y-%m-%d}.xlsx",
    )

    with smtplib.SMTP_SSL("smtp.feishu.cn", 465, context=ssl.create_default_context()) as smtp:
        smtp.login(remetente, senha_app)
        smtp.send_message(msg)


def main():
    dry_run = "--dry-run" in sys.argv
    hoje = date.today()

    linhas = montar_linhas(hoje)
    if not linhas:
        log.info("Nenhuma base com dado de Pickup pra %s -- nada pra mandar", hoje)
        return

    total_pendente = sum(r["pendente"] for r in linhas)
    log.info("%s: %d bases, %d pedido(s) sem foto POC ainda no total", hoje, len(linhas), total_pendente)

    titulo = f"Taxa de Transferência POC -- {hoje:%d/%m/%Y}"
    rodape = f"Resumo automático POC às 21h · {total_pendente} pedido(s) sem foto ainda"
    cartoes = poc(linhas, titulo, rodape)

    excel = montar_excel_epop(hoje)
    log.info("Excel detalhado montado: %d bytes", len(excel))

    if dry_run:
        for c in cartoes:
            log.info("CARTAO: %s", c["header"]["title"]["content"])
            for el in c["elements"]:
                if el.get("tag") == "column_set":
                    log.info("  %s", " | ".join(col["elements"][0]["content"] for col in el["columns"]))
        print(f"[dry-run] {len(cartoes)} cartao(oes), {len(linhas)} base(s), {total_pendente} pendente(s), excel {len(excel)} bytes")
        return

    webhook = feishu_api.Webhook(
        url=os.environ["FEISHU_WEBHOOK_URL_POC_21H"],
        palavra=os.environ.get("FEISHU_KEYWORD_POC_21H", "POC"),
    )
    for c in cartoes:
        webhook.enviar_cartao(c)
    log.info("Enviado: %d cartao(oes) pro webhook de 21h", len(cartoes))

    enviar_email_excel(hoje, excel)
    log.info("E-mail com Excel detalhado enviado pra %s", os.environ["EMAIL_TO_POC"])


if __name__ == "__main__":
    main()

"""Resumo noturno (21h) de "Tentativa de coleta fora do prazo" por base --
pedido do Guilherme, 07/10/2026, a partir da tela JMS "Indicadores de
Negócios > Prazo > Taxa de coleta no prazo > Lista" (Regional=SPS,
Origem do Pedido=TikTok, Horário de término do prazo de coleta,
Tentativa de coleta dentro do prazo=N, data=HOJE/"do mesmo dia").

Endpoint e payload capturados por rede pelo próprio Guilherme
(DevTools > Network, 07/10/2026) -- ver `buscar_tentativa_fora_prazo`.
Chave: `isOutTime` é o nome real do campo "Tentativa de coleta dentro do
prazo" na API (valor "N" = fora do prazo, o mesmo que a tela filtra).
`orderSourceCode: ["D67"]` = Origem "TikTok". `agentId/agentName/
agentCode` = Regional "SPS" (fixo na tela, não é selecionável).

"quantidade" = contagem de pedidos com isOutTime=N por base (1 linha por
pedido no retorno da API, soma direta -- NÃO é mais derivado do relatório
agregado de Pickup como nas primeiras versões deste script). "taxa" =
quantidade / qtd_a_coletar do dia (mesmo "deveria coletar" que
bot_dados.pickup_d1_ao_vivo já busca, reaproveitado como denominador pra
não precisar de uma segunda busca paginada gigante sem o filtro
isOutTime). Ordenado do MELHOR pro PIOR (Guilherme: "mostrar primeiro
quem ta melhor... e o ultimo pior").

Dois envios, mesma rodada:
1. Cartão resumido (Base + Quantidade + Taxa) pro webhook do Feishu.
2. E-mail com o relatório DETALHADO (1 linha por pedido fora do prazo,
   com Base/Motorista/Loja/Pedido), só com as nossas bases
   (BASES_PICKUP) -- webhook de grupo não aceita arquivo, por isso vai
   por e-mail (Guilherme, 07/10/2026: "mande o arquivo por email
   kuan.chen@jtexpress.com.br").

Webhook, palavra-chave e e-mail PRÓPRIOS (FEISHU_WEBHOOK_URL_POC_21H /
FEISHU_KEYWORD_POC_21H / EMAIL_TO_POC), separados do bot da manhã.

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
import requests
from bot_cards import poc
from dotenv import load_dotenv
from extrair_pickup import BASES_PICKUP
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

URL_TENTATIVA_FORA_PRAZO = (
    "https://gw.jtjms-br.com/businessindicator/bigdataReport/detail/timely_collection_rate_detail_new"
)
HEADERS = {
    "Accept": "application/json, text/plain, */*",
    "Content-Type": "application/json;charset=UTF-8",
    "lang": "PT",
    "langType": "PT",
}
TAMANHO_PAGINA = 1000
# Regional "SPS" -- campo travado na tela (não é selecionável), capturado
# por rede direto da requisição real (Guilherme, 07/10/2026).
AGENT_ID, AGENT_NAME, AGENT_CODE = 129, "SPS", "350000"
ORIGEM_PEDIDO_CODE = "D67"  # TikTok

PREENCHIMENTO_HEADER = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
FONTE_HEADER = Font(bold=True, color="FFFFFF")
CABECALHO_DETALHE = ["Estação de agendamento", "Comerciante ID", "Nome de comerciante", "Motorista associado", "Contagem de Número da Remessa"]


def buscar_tentativa_fora_prazo(dia: date) -> list[dict]:
    """Pagina o endpoint timely_collection_rate_detail_new pro dia inteiro,
    isOutTime="N" (tentativa de coleta fora do prazo), origem TikTok,
    SEM filtrar rede (traz todas as bases da rede SPS, filtra pra
    BASES_PICKUP depois) -- mesmo padrão de extrair_pickup._buscar_registros_epop_brutos."""
    token = os.environ["JMS_TOKEN_INDICADORES"]
    headers = {**HEADERS, "authToken": token}
    registros = []
    current = 1
    while True:
        payload = {
            "current": current, "size": TAMANHO_PAGINA,
            "agentId": AGENT_ID, "agentName": AGENT_NAME, "agentCode": AGENT_CODE,
            "countryId": "1",
            "startTime": f"{dia.isoformat()} 00:00:00", "endTime": f"{dia.isoformat()} 23:59:59",
            "timeType": "2", "isOutTime": "N", "orderSourceCode": [ORIGEM_PEDIDO_CODE], "orderStatus": 0,
        }
        resp = requests.post(URL_TENTATIVA_FORA_PRAZO, headers=headers, json=payload, timeout=30)
        if resp.status_code in (401, 403):
            raise RuntimeError("JMS_TOKEN_INDICADORES expirado ou sem sessão ativa.")
        resp.raise_for_status()
        resultado = resp.json()
        if resultado.get("code") != 1:
            raise RuntimeError(f"Erro ao buscar tentativa fora do prazo: {resultado}")

        dados = resultado.get("data") or {}
        pagina = dados.get("records", [])
        registros.extend(pagina)
        total = dados.get("total", 0)
        if not pagina or len(registros) >= total:
            break
        current += 1
        if current > 300:  # trava de segurança (300k linhas), nao deveria chegar perto
            log.warning("Parou de paginar em 300 páginas (total esperado: %s)", total)
            break
    return registros


def montar_linhas(hoje: date, brutos: list[dict]) -> list[dict]:
    """Agrega os registros brutos (1 por pedido fora do prazo) por base,
    só BASES_PICKUP. taxa = quantidade / qtd_a_coletar (Pickup de hoje,
    já buscado em paralelo, mesmo "deveria coletar" do dia). Ordenado do
    MELHOR pro PIOR (taxa asc)."""
    por_base: dict[str, int] = {}
    for r in brutos:
        base = (r.get("pickNetworkName") or "").strip()
        if base not in BASES_PICKUP:
            continue
        por_base[base] = por_base.get(base, 0) + 1

    deveria_por_base = {r["base"]: r["qtd_a_coletar"] for r in bd.pickup_d1_ao_vivo(hoje)}

    linhas = []
    for base in BASES_PICKUP:
        quantidade = por_base.get(base, 0)
        deveria = deveria_por_base.get(base, 0)
        if not deveria and not quantidade:
            continue
        taxa_fora = round(quantidade / deveria * 100, 2) if deveria else 0.0
        linhas.append({"base": base, "pendente": quantidade, "taxa_fora": taxa_fora})
    linhas.sort(key=lambda r: r["taxa_fora"])
    return linhas


def montar_excel_detalhado(brutos: list[dict]) -> bytes:
    """Agrupado igual ao pivot que o Guilherme mostrou (Estação de
    agendamento/Comerciante ID/Nome de comerciante/Motorista associado +
    Contagem de Número da Remessa) -- 1 linha por combinação
    base+comerciante+motorista, só nossas bases (BASES_PICKUP), com a
    contagem de pedidos fora do prazo daquele grupo."""
    grupos: dict[tuple, int] = {}
    nomes: dict[tuple, tuple] = {}
    for r in brutos:
        base = (r.get("pickNetworkName") or "").strip()
        if base not in BASES_PICKUP:
            continue
        chave = (base, r.get("merchantId"), r.get("pickStaffName") or "")
        grupos[chave] = grupos.get(chave, 0) + 1
        nomes[chave] = (r.get("merchantName"),)

    linhas = sorted(grupos.items(), key=lambda item: (item[0][0], item[0][2] or ""))

    wb = Workbook()
    ws = wb.active
    ws.title = "Fora do prazo"
    ws.append(CABECALHO_DETALHE)
    for cel in ws[1]:
        cel.font = FONTE_HEADER
        cel.fill = PREENCHIMENTO_HEADER
    for (base, comerciante_id, motorista), contagem in linhas:
        (nome_comerciante,) = nomes[(base, comerciante_id, motorista)]
        ws.append([base, comerciante_id, nome_comerciante, motorista, contagem])
    ws.auto_filter.ref = f"A1:{get_column_letter(len(CABECALHO_DETALHE))}{len(linhas) + 1}"
    for i, largura in enumerate([18, 22, 30, 36, 24], start=1):
        ws.column_dimensions[get_column_letter(i)].width = largura

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def enviar_email_excel(hoje: date, conteudo: bytes, total_linhas: int):
    remetente = os.environ["EMAIL_SENDER"]
    senha_app = os.environ["EMAIL_APP_PASSWORD"]
    destinatarios = [e.strip() for e in os.environ["EMAIL_TO_POC"].split(",") if e.strip()]

    msg = EmailMessage()
    msg["Subject"] = f"Tentativa de coleta fora do prazo - {hoje:%d/%m/%Y}"
    msg["From"] = remetente
    msg["To"] = ", ".join(destinatarios)
    msg.set_content(
        f"Segue em anexo o relatório de pedidos com tentativa de coleta fora do "
        f"prazo do dia {hoje:%d/%m/%Y} ({total_linhas} pedido(s) no total), agrupado "
        "por base/comerciante/motorista com a contagem de remessas -- só nossas "
        "bases (Regional SPS, origem TikTok)."
    )
    msg.add_attachment(
        conteudo, maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"Fora_do_prazo_{hoje:%Y-%m-%d}.xlsx",
    )

    with smtplib.SMTP_SSL("smtp.feishu.cn", 465, context=ssl.create_default_context()) as smtp:
        smtp.login(remetente, senha_app)
        smtp.send_message(msg)


def main():
    dry_run = "--dry-run" in sys.argv
    hoje = date.today()

    brutos = buscar_tentativa_fora_prazo(hoje)
    log.info("%s: %d pedido(s) fora do prazo na rede SPS inteira (antes de filtrar pras nossas bases)", hoje, len(brutos))

    linhas = montar_linhas(hoje, brutos)
    if not linhas:
        log.info("Nenhuma das nossas bases com dado hoje -- nada pra mandar")
        return

    total_nossas_bases = sum(r["pendente"] for r in linhas)
    log.info("Nossas bases: %d pedido(s) fora do prazo no total", total_nossas_bases)

    titulo = f"POC -- Tentativa de coleta fora do prazo -- {hoje:%d/%m/%Y}"
    rodape = f"Resumo automático POC às 21h · {total_nossas_bases} pedido(s) fora do prazo no total"
    cartoes = poc(linhas, titulo, rodape)

    excel = montar_excel_detalhado(brutos)
    log.info("Excel detalhado montado: %d bytes", len(excel))

    if dry_run:
        for c in cartoes:
            log.info("CARTAO: %s", c["header"]["title"]["content"])
            for el in c["elements"]:
                if el.get("tag") == "column_set":
                    log.info("  %s", " | ".join(col["elements"][0]["content"] for col in el["columns"]))
        print(f"[dry-run] {len(cartoes)} cartao(oes), {len(linhas)} base(s), {total_nossas_bases} fora do prazo, excel {len(excel)} bytes")
        return

    webhook = feishu_api.Webhook(
        url=os.environ["FEISHU_WEBHOOK_URL_POC_21H"],
        palavra=os.environ.get("FEISHU_KEYWORD_POC_21H", "POC"),
    )
    for c in cartoes:
        webhook.enviar_cartao(c)
    log.info("Enviado: %d cartao(oes) pro webhook de 21h", len(cartoes))

    enviar_email_excel(hoje, excel, total_nossas_bases)
    log.info("E-mail com Excel detalhado enviado pra %s", os.environ["EMAIL_TO_POC"])


if __name__ == "__main__":
    main()

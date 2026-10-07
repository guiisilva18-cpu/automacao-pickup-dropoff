"""Resumo noturno (21h) da Taxa de Transferência POC por base -- % de
pedidos com FOTO de comprovação de coleta (relatório EPOP do JMS),
confirmado pelo Guilherme 07/10/2026 ("é a transferencia poc mesmo, que
vem as fotos e tudo mais"). Dado de HOJE (fechamento parcial às 21h, "do
mesmo dia" como a tela de referência mostra -- não D-1).

taxa_poc já vem calculada por base em extrair_pickup.buscar_pickup /
bot_dados.pickup_d1_ao_vivo (razão epop_com_imagem/epop_total do
Relatório de Monitoramento EPOP) -- reusa esse campo, sem chamada nova.
"pendente" = epop_total - epop_com_imagem (pedidos EPOP ainda sem foto).

Ordenado do MELHOR pro PIOR (Guilherme: "mostrar primeiro quem ta
melhor... e o ultimo pior").

Webhook e palavra-chave PRÓPRIOS (diferentes do bot da manhã,
FEISHU_WEBHOOK_URL_POC_21H / FEISHU_KEYWORD_POC_21H).

Uso:  python bot_poc_21h.py              envia
      python bot_poc_21h.py --dry-run    só mostra o que mandaria, não envia
"""
import logging
import os
import sys
from datetime import date
from pathlib import Path

import bot_dados as bd
import feishu_api
from bot_cards import poc
from dotenv import load_dotenv

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

    if dry_run:
        for c in cartoes:
            log.info("CARTAO: %s", c["header"]["title"]["content"])
            for el in c["elements"]:
                if el.get("tag") == "column_set":
                    log.info("  %s", " | ".join(col["elements"][0]["content"] for col in el["columns"]))
        print(f"[dry-run] {len(cartoes)} cartao(oes), {len(linhas)} base(s), {total_pendente} pendente(s)")
        return

    webhook = feishu_api.Webhook(
        url=os.environ["FEISHU_WEBHOOK_URL_POC_21H"],
        palavra=os.environ.get("FEISHU_KEYWORD_POC_21H", "POC"),
    )
    for c in cartoes:
        webhook.enviar_cartao(c)
    log.info("Enviado: %d cartao(oes) pro webhook de 21h", len(cartoes))


if __name__ == "__main__":
    main()

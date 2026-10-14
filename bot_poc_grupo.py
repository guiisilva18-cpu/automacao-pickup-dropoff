"""Bot do grupo Feishu "POC" (webhook proprio, palavra-chave POC) -- pedido
do Guilherme, 14/10/2026:
  - 09:00  resumo do dia (so a Taxa de Transferencia POC/EPOP por base e
    quanto ficou pendente de foto)
  - 18:00 e 20:00  aviso/lembrete pra fazer o POC, com o pendente por base

Fonte do dado: bot_dados.pickup_d1_ao_vivo (extracao ao vivo do JMS, mesma
usada no e-mail do Pickup -- tem epop_total/epop_com_imagem, que o
pickup_diario do banco nao guarda).

Uso:
  python bot_poc_grupo.py resumo   (09h)
  python bot_poc_grupo.py aviso    (18h / 20h)
  python bot_poc_grupo.py teste    (manda o resumo agora, pra conferir o webhook/palavra-chave)
"""
import logging
import sys
from datetime import date

sys.stdout.reconfigure(encoding="utf-8")

import bot_cards
import bot_dados
from feishu_api import Webhook

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

WEBHOOK_URL = "https://open.feishu.cn/open-apis/bot/v2/hook/50e3f6b5-eaf0-4120-ab2a-3defa6c9b77d"
PALAVRA_CHAVE = "POC"


def montar_e_enviar(modo: str):
    hoje = date.today()
    linhas = bot_dados.pickup_d1_ao_vivo(hoje)
    webhook = Webhook(url=WEBHOOK_URL, palavra=PALAVRA_CHAVE)

    if modo in ("resumo", "teste"):
        titulo = f"Resumo POC — {hoje:%d/%m/%Y}"
    else:  # aviso
        titulo = f"⚠️ Aviso: Fazer o POC — {hoje:%d/%m/%Y}"

    cartoes = bot_cards.poc(linhas, titulo, webhook.rodape)
    if not cartoes:
        log.warning("Sem dados de pickup hoje (%s) -- nada enviado.", hoje)
        return
    for c in cartoes:
        webhook.enviar_cartao(c)
    log.info("%s enviado(s) %d cartao(oes) pro grupo POC.", modo, len(cartoes))


def main():
    modo = sys.argv[1] if len(sys.argv) > 1 else "teste"
    if modo not in ("resumo", "aviso", "teste"):
        print("Uso: python bot_poc_grupo.py resumo|aviso|teste")
        return
    montar_e_enviar(modo)


if __name__ == "__main__":
    main()

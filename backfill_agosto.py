"""Backfill pontual: busca PICKUP e DROPOFF de 01 a 09/08/2026 (dias
anteriores ao início da automação, 10-11/08) direto no JMS, e mescla no
mesmo histórico que a automação diária já mantém (dados/historico_*.json).
Pedido do Guilherme (2026-08-13) pra montar apresentação com o mês
completo até o dia disponível."""
import logging
from datetime import date, timedelta
from pathlib import Path

import extrair_dropoff
import extrair_pickup

LOG_DIR = Path(__file__).parent / "logs"
LOG_DIR.mkdir(exist_ok=True)
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

DIAS = [f"2026-08-{d:02d}" for d in range(1, 10)]


def main():
    hist_pickup = extrair_pickup.carregar_historico()
    hist_dropoff = extrair_dropoff.carregar_historico()

    for dia in DIAS:
        if dia not in hist_pickup:
            log.info("PICKUP %s...", dia)
            hist_pickup[dia] = extrair_pickup.buscar_pickup(dia)
            extrair_pickup.salvar_historico(hist_pickup)
        else:
            log.info("PICKUP %s já no histórico, pulando", dia)

        if dia not in hist_dropoff:
            log.info("DROPOFF %s...", dia)
            hist_dropoff[dia] = extrair_dropoff.buscar_dropoff(dia)
            extrair_dropoff.salvar_historico(hist_dropoff)
        else:
            log.info("DROPOFF %s já no histórico, pulando", dia)

    log.info("Backfill concluído.")


if __name__ == "__main__":
    main()

"""Carga noturna (23h) da tela "POC" do App Ponto de Apoio -- detalhe por
Base/Comerciante/Motorista de pedidos com TENTATIVA DE COLETA fora do
prazo (campo `isOutTime="N"`, endpoint capturado por rede pelo próprio
Guilherme, DevTools > Network, 07/10/2026) -- ver
`buscar_tentativa_fora_prazo`/`agrupar_fora_prazo`. `orderSourceCode:
["D67"]` = Origem "TikTok". `agentId/agentName/agentCode` = Regional
"SPS" (fixo na tela, não é selecionável). Filtrado só pras nossas bases
(BASES_PICKUP). Dado de HOJE (fechamento parcial às 23h, não D-1).

Histórico: antes mandava um cartão no Feishu (Taxa de Transferência
POC/EPOP) e um Excel por e-mail -- os dois foram removidos em
09/10/2026 (Guilherme: "Tire do email o excel e do feishu vamos focar
nessa tela"), ficando só a gravação em `poc_fora_prazo_detalhe` (TiDB),
que alimenta a tela "POC" do App Ponto de Apoio.

Uso:  python bot_poc_21h.py              grava no banco
      python bot_poc_21h.py --dry-run    só mostra o que gravaria, não grava nada
"""
import logging
import os
import sys
from datetime import date
from pathlib import Path

import gravar_mysql
import requests
from dotenv import load_dotenv
from extrair_pickup import BASES_PICKUP

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


def agrupar_fora_prazo(brutos: list[dict]) -> list[dict]:
    """Agrupa os registros brutos de tentativa de coleta fora do prazo por
    base/comerciante/motorista, só pras nossas bases (BASES_PICKUP) -- 1
    linha por combinação, com a contagem de pedidos fora do prazo daquele
    grupo. Alimenta a tabela poc_fora_prazo_detalhe (tela "POC" do App
    Ponto de Apoio, pedido do Guilherme, 09/10/2026)."""
    grupos: dict[tuple, int] = {}
    nomes: dict[tuple, str] = {}
    for r in brutos:
        base = (r.get("pickNetworkName") or "").strip()
        if base not in BASES_PICKUP:
            continue
        chave = (base, r.get("merchantId"), r.get("pickStaffName") or "")
        grupos[chave] = grupos.get(chave, 0) + 1
        nomes[chave] = r.get("merchantName")

    linhas = sorted(grupos.items(), key=lambda item: (item[0][0], item[0][2] or ""))
    return [
        {
            "base": base,
            "comerciante_id": comerciante_id,
            "nome_comerciante": nomes[(base, comerciante_id, motorista)],
            "motorista": motorista or None,
            "contagem": contagem,
        }
        for (base, comerciante_id, motorista), contagem in linhas
    ]


def main():
    dry_run = "--dry-run" in sys.argv
    hoje = date.today()

    brutos = buscar_tentativa_fora_prazo(hoje)
    log.info("%s: %d pedido(s) fora do prazo na rede SPS inteira (antes de filtrar pras nossas bases)", hoje, len(brutos))
    linhas_detalhe = agrupar_fora_prazo(brutos)
    total_fora_prazo = sum(r["contagem"] for r in linhas_detalhe)
    log.info("Nossas bases: %d linha(s), %d pedido(s) fora do prazo no total", len(linhas_detalhe), total_fora_prazo)

    if dry_run:
        print(f"[dry-run] {len(linhas_detalhe)} linha(s), {total_fora_prazo} fora do prazo no total")
        return

    gravar_mysql.gravar_poc_fora_prazo_detalhe(hoje.isoformat(), linhas_detalhe)
    log.info("Gravado em poc_fora_prazo_detalhe pra %s", hoje)


if __name__ == "__main__":
    main()

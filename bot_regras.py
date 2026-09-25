"""Regras de faixa (semáforo) do resumo diário -- separadas do desenho pra
valerem igual nas imagens (bot_imagens) e nos cartões (bot_cards).

Pedido do Guilherme, 25/09/2026 (atualiza a regra de 24/09):
  Pickup (coluna "com tentativas"): >=98,99 verde, 95 a <98,99 laranja, <95 vermelho.
  Dropoff: >=95 verde, 90 a <95 laranja, <90 vermelho.
  Transferência: mantém a regra antiga do Pickup, 95/90 (print-modelo não mudou).
  Expedição, % de ocupação: >=100 verde, abaixo vermelho (regra do dash).
"""

VERDE, LARANJA, VERMELHO = "verde", "laranja", "vermelho"


def faixa_pickup(taxa: float) -> str:
    if taxa >= 98.99:
        return VERDE
    if taxa >= 95:
        return LARANJA
    return VERMELHO


def faixa_dropoff(taxa: float) -> str:
    if taxa >= 95:
        return VERDE
    if taxa >= 90:
        return LARANJA
    return VERMELHO


def faixa_transferencia(taxa: float) -> str:
    if taxa >= 95:
        return VERDE
    if taxa >= 90:
        return LARANJA
    return VERMELHO


def faixa_ocupacao(pct: float) -> str:
    return VERDE if pct >= 100 else VERMELHO


def eh_meli(pa: str) -> bool:
    """P.As Meli ficam fora da previsão (Guilherme, 24/09/2026)."""
    return pa.upper().startswith("PA MELI")

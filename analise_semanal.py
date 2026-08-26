"""Agrega PICKUP, DROPOFF e TRANSFERÊNCIA em blocos de 7 dias a partir de
01/08 (Semana 1 = 01-07/08, Semana 2 = 08-14/08, Semana 3 = 15-.../08,
parcial até o último dia disponível) usando os mesmos _consolidar_por_base
já existentes (soma bruta, taxa recalculada pelo total -- não é média de
%)."""
import json
from datetime import date, timedelta

import extrair_dropoff
import extrair_pickup

INICIO_MES = date(2026, 8, 1)


def carregar_historico_transferencia() -> dict[str, list[dict]]:
    caminho = "dados/historico_transferencia.json"
    try:
        with open(caminho, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}


def montar_dados(historico: dict, dias: list[str]) -> list[tuple[str, list[dict]]]:
    return [(d, historico[d]) for d in dias if d in historico]


def _semanas_disponiveis(ultimo_dia: str) -> list[tuple[str, list[str]]]:
    """Quebra 01/08 até `ultimo_dia` em blocos de 7 dias completos -- uma
    semana em andamento (incompleta) não entra como referência de
    comparação (pedido do Guilherme, 2026-08-18: "não vamos usar como
    referência"). Assim que a semana fechar 7 dias ela aparece sozinha na
    próxima geração."""
    fim = date.fromisoformat(ultimo_dia)
    semanas = []
    inicio_semana = INICIO_MES
    n = 1
    while inicio_semana + timedelta(days=6) <= fim:
        fim_semana = inicio_semana + timedelta(days=6)
        dias = [(inicio_semana + timedelta(days=i)).isoformat() for i in range(7)]
        semanas.append((f"Semana {n} ({inicio_semana.day:02d}-{fim_semana.day:02d}/08)", dias))
        inicio_semana = fim_semana + timedelta(days=1)
        n += 1
    return semanas


def _consolidar_transferencia_por_base(dados: list[tuple[str, list[dict]]]) -> list[dict]:
    from collections import defaultdict

    somas: dict[str, dict[str, int]] = defaultdict(lambda: {"total": 0, "no_prazo": 0, "fora_prazo": 0})
    for _dia, registros in dados:
        for r in registros:
            base = r["base"]
            somas[base]["total"] += r["entregas_total"]
            somas[base]["no_prazo"] += r["entregas_no_prazo"]
            somas[base]["fora_prazo"] += r["entregas_fora_prazo"]

    consolidado = []
    for base, v in somas.items():
        taxa_pct = round(v["no_prazo"] / v["total"] * 100, 2) if v["total"] else None
        consolidado.append({
            "base": base,
            "entregas_total": v["total"],
            "entregas_no_prazo": v["no_prazo"],
            "entregas_fora_prazo": v["fora_prazo"],
            "taxa_pct": taxa_pct,
        })
    consolidado.sort(key=lambda r: -r["entregas_total"])
    return consolidado


def main():
    hist_pickup = extrair_pickup.carregar_historico()
    hist_dropoff = extrair_dropoff.carregar_historico()
    hist_transferencia = carregar_historico_transferencia()

    ultimo_dia = max(set(hist_pickup) | set(hist_dropoff))
    semanas = _semanas_disponiveis(ultimo_dia)

    saida = {}
    for nome_semana, dias in semanas:
        dados_pickup = montar_dados(hist_pickup, dias)
        dados_dropoff = montar_dados(hist_dropoff, dias)
        dias_com_dado = sorted(set(d for d, _ in dados_pickup) | set(d for d, _ in dados_dropoff))

        pickup_consolidado = extrair_pickup._consolidar_por_base(dados_pickup) if dados_pickup else []
        dropoff_consolidado = extrair_dropoff._consolidar_por_base(dados_dropoff) if dados_dropoff else []

        saida[nome_semana] = {
            "dias_com_dado": dias_com_dado,
            "pickup": pickup_consolidado,
            "dropoff": dropoff_consolidado,
        }

    # Transferência: período completo 01/08 até o último dia disponível (não
    # quebrado por semana -- é 1 KPI agregado, pedido do Guilherme, 2026-08-18).
    dias_transferencia = sorted(hist_transferencia.keys())
    dados_transferencia = montar_dados(hist_transferencia, dias_transferencia)
    saida["Transferência (período completo)"] = {
        "dias_com_dado": dias_transferencia,
        "por_base": _consolidar_transferencia_por_base(dados_transferencia) if dados_transferencia else [],
    }

    # Data mais recente com dado de verdade (independe de semana fechada ou
    # não) -- usada pelos slides de Expedição, que continuam mostrando o
    # período completo mesmo com a semana em andamento fora da comparação
    # semanal acima.
    saida["ultimo_dia_disponivel"] = ultimo_dia
    saida["primeiro_dia_disponivel"] = INICIO_MES.isoformat()

    with open("analise_semanal.json", "w", encoding="utf-8") as f:
        json.dump(saida, f, ensure_ascii=False, indent=2)
    print("Salvo em analise_semanal.json")
    print("Semanas:", [n for n, _ in semanas])
    print("Transferência dias:", len(dias_transferencia))


if __name__ == "__main__":
    main()

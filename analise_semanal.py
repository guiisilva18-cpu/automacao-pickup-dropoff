"""Agrega PICKUP e DROPOFF em blocos de 7 dias a partir de 01/08 (Semana 1 =
01-07/08, Semana 2 = 08-14/08) usando os mesmos _consolidar_por_base já
existentes (soma bruta, taxa recalculada pelo total -- não é média de %)."""
import json

import extrair_dropoff
import extrair_pickup

SEMANA_1 = [f"2026-08-{d:02d}" for d in range(1, 8)]
SEMANA_2 = [f"2026-08-{d:02d}" for d in range(8, 15)]


def montar_dados(historico: dict, dias: list[str]) -> list[tuple[str, list[dict]]]:
    return [(d, historico[d]) for d in dias if d in historico]


def main():
    hist_pickup = extrair_pickup.carregar_historico()
    hist_dropoff = extrair_dropoff.carregar_historico()

    saida = {}
    for nome_semana, dias in [("Semana 1 (01-07/08)", SEMANA_1), ("Semana 2 (08-14/08)", SEMANA_2)]:
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

    with open("analise_semanal.json", "w", encoding="utf-8") as f:
        json.dump(saida, f, ensure_ascii=False, indent=2)
    print("Salvo em analise_semanal.json")


if __name__ == "__main__":
    main()

"""Bot do Feishu: resumo diário das operações pro Guilherme.

Roda no GitHub Actions (workflow "RESUMO FEISHU diario"), disparado quando
PICKUP/DROPOFF/TRANSFERENCIA diario terminam; o lock no banco garante 1
envio por dia. Modelo da mensagem (pedido do Guilherme, 24/09/2026):

  Bom dia! Segue resumo das operações do dia D-1.
  Taxa de coleta Pickup do dia D-1.
  Taxa de coleta Dropoff do dia D-1.
  Taxa de transferência do dia D-2.
  Ocupação dos veículos expedidos do dia D-1.
  Previsão de coleta Pickup do dia D.           (só o não coletado das Bases)
  Previsão de coleta Dropoff do dia D.
  Previsão de volumes geral para coletar no dia D.  (só o não coletado dos P.As, sem Meli)
  Assertividade da previsão do dia D-1.         (a partir do 2º envio)

Dois modos de envio (feishu_api.modo()):
  cartao  -- webhook do grupo (FEISHU_WEBHOOK_URL + FEISHU_KEYWORD): cada bloco vira
             cartão do Feishu com tabela e bolinhas de cor (webhook não manda imagem).
  imagem  -- app do Feishu (FEISHU_APP_ID/SECRET/DESTINO): cada bloco vira um PNG.

Uso:  python bot_feishu.py              envia
      python bot_feishu.py --dry-run    só renderiza (PNGs + cartoes.json em saida_bot/), não envia nem grava
      python bot_feishu.py --teste      envia marcado como TESTE, sem lock e sem gravar a previsão
      python bot_feishu.py --blocos 1,5 envia só esses blocos (numeração da lista acima, 1-based)
      python bot_feishu.py --force      ignora o lock do dia (reenvio manual)
"""
import argparse
import json
import logging
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

import bot_cards as bc
import bot_dados as bd
import bot_imagens as bi
import feishu_api
import gravar_mysql

log = logging.getLogger("bot_feishu")
PASTA_SAIDA = Path(__file__).parent / "saida_bot"
PAUSA_ENTRE_MENSAGENS = 0.5  # webhook aceita ~5 msg/s; sobra folga


def _tentar(avisos: list[str], nome: str, fn):
    """Um bloco que falha não derruba o resumo: vira aviso no fim da mensagem."""
    try:
        return fn()
    except Exception as e:  # noqa: BLE001
        log.exception("Falha no bloco '%s'", nome)
        avisos.append(f"{nome}: indisponível agora ({type(e).__name__}).")
        return None


def coletar(conn, hoje, sem_espera: bool = False, sincronizar: bool = True) -> dict:
    d1, d2 = hoje - timedelta(days=1), hoje - timedelta(days=2)
    avisos: list[str] = []

    if sincronizar:
        _tentar(avisos, "Expedição (planilha)", lambda: bd.sincronizar_expedicao(conn))
    if not sem_espera:
        avisos += bd.aguardar_resumo_pa(conn, d1)

    # Pickup/Dropoff de D-1 ao vivo (o envio é às 08:30, antes da carga das 09:00);
    # em teste/simulação não espera o fechamento, só consulta uma vez.
    pk_dp = _tentar(avisos, "Pickup/Dropoff D-1",
                    lambda: bd.obter_pickup_dropoff_d1(conn, d1, d2, limite_min=0 if sem_espera else 50))
    pickup, dropoff = (pk_dp[0], pk_dp[1]) if pk_dp else (None, None)
    if pk_dp:
        avisos += pk_dp[2]

    snapshot = _tentar(avisos, "Posição ao vivo (JMS)", lambda: bd.snapshot_pickup_ao_vivo(hoje))
    hora = datetime.now(bd.FUSO).strftime("%H:%M")
    pend_dropoff = _tentar(avisos, "Dropoff aguardando coleta", lambda: bd.dropoff_pendente_ao_vivo([d1, hoje]))
    final_base = {r["base"]: r["coletada_no_prazo"] for r in pickup} if pickup else None
    return {
        "hoje": hoje, "d1": d1, "d2": d2, "hora": hora, "avisos": avisos,
        "pickup": pickup,
        "dropoff": dropoff,
        "transf": _tentar(avisos, "Transferência", lambda: bd.transferencia(d2, d1)),
        "exped": _tentar(avisos, "Expedição D-1", lambda: bd.expedicao_d1(conn, d1)),
        "prev_bases": bd.previsao_bases(snapshot[0]) if snapshot else None,
        "prev_pas": _tentar(avisos, "Previsão P.As", lambda: bd.previsao_pas(conn, d1, snapshot[1] if snapshot else {})),
        "prev_drop": bd.previsao_dropoff(pend_dropoff, d1, hoje) if pend_dropoff else None,
        "assert": _tentar(avisos, "Assertividade", lambda: bd.assertividade(conn, d1, final_base)),
    }


def legendas(D: dict) -> dict:
    fmt = "%d/%m/%Y"
    return {
        "pickup": f"Taxa de coleta Pickup do dia {D['d1']:{fmt}}",
        "dropoff": f"Taxa de coleta Dropoff do dia {D['d1']:{fmt}}",
        # Janela D-2->D-1: o JMS não fecha o mesmo intervalo pra toda base
        # (algumas só têm dado de D-2, outras só de D-1, outras somam os
        # dois) -- rotular como "do dia D-2" enganava (achado em auditoria
        # de 25/09/2026). Mostra o intervalo real.
        "transf": f"Taxa de transferência ({D['d2']:%d/%m} a {D['d1']:{fmt}})",
        "exped": f"Ocupação dos veículos expedidos do dia {D['d1']:{fmt}}",
        "prev_bases": f"Previsão de coleta Pickup do dia {D['hoje']:{fmt}}",
        "prev_drop": f"Previsão de coleta Dropoff do dia {D['hoje']:{fmt}}",
        "prev_pas": f"Previsão de volumes geral para coletar no dia {D['hoje']:{fmt}}",
        "assert": f"Assertividade da previsão do dia {D['d1']:{fmt}}",
    }


def _bases_com_volume(D):
    return [r for r in D["prev_bases"] if r["deveria"] > 0 and r["pendente"] > 0] if D["prev_bases"] else []


def _dropoff_com_previsao(D):
    return [r for r in D["prev_drop"] if r["previsto"] > 0] if D["prev_drop"] else []


def blocos_imagens(D: dict) -> list:
    """[(legenda, [png...], nota_se_vazio)]"""
    L, bi_ = legendas(D), bi
    bases, drop = _bases_com_volume(D), _dropoff_com_previsao(D)
    pas = D["prev_pas"]
    blocos = [
        (L["pickup"], [bi_.img_pickup(D["pickup"])] if D["pickup"] else [], "sem dados"),
        (L["dropoff"], [bi_.img_dropoff(D["dropoff"])] if D["dropoff"] else [], "sem dados"),
        (L["transf"], [bi_.img_transferencia(D["transf"])] if D["transf"] else [], "sem dados"),
        (L["exped"], [bi_.img_expedicao(D["exped"], D["d1"])] if D["exped"] else [], "sem dados de expedição"),
        (L["prev_bases"], [bi_.img_previsao_bases(bases, D["hora"])] if bases else [], "sem dados"),
        (L["prev_drop"], [bi_.img_previsao_dropoff(drop, D["d1"], D["hoje"])] if drop else [],
         "nenhum pedido aguardando coleta" if D["prev_drop"] is not None else "sem dados"),
        (L["prev_pas"], [bi_.img_previsao_pas(pas, D["hora"])] if pas and pas["linhas"] else [], "sem dados"),
    ]
    if D["assert"]:
        blocos.append((L["assert"], bi_.img_assertividade(D["assert"], D["d1"]), ""))
    return blocos


def blocos_cartoes(D: dict, rodape: str) -> list:
    """[(legenda, [cartao...], nota_se_vazio)]"""
    L = legendas(D)
    bases, drop = _bases_com_volume(D), _dropoff_com_previsao(D)
    pas = D["prev_pas"]
    blocos = [
        (L["pickup"], bc.pickup(D["pickup"], L["pickup"], rodape) if D["pickup"] else [], "sem dados"),
        (L["dropoff"], bc.dropoff(D["dropoff"], L["dropoff"], rodape) if D["dropoff"] else [], "sem dados"),
        (L["transf"], bc.transferencia(D["transf"], L["transf"], rodape) if D["transf"] else [], "sem dados"),
        (L["exped"], bc.expedicao(D["exped"], L["exped"], rodape) if D["exped"] else [], "sem dados de expedição"),
        (L["prev_bases"], bc.previsao_bases(bases, D["hora"], L["prev_bases"], rodape) if bases else [], "sem dados"),
        (L["prev_drop"], bc.previsao_dropoff(drop, D["d1"], D["hoje"], L["prev_drop"], rodape) if drop else [],
         "nenhum pedido aguardando coleta"),
        (L["prev_pas"], bc.previsao_pas(pas, D["hora"], L["prev_pas"], rodape) if pas and pas["linhas"] else [], "sem dados"),
    ]
    if D["assert"]:
        blocos.append((L["assert"], bc.assertividade(D["assert"], D["d1"], rodape), ""))
    return blocos


def saudacao(D: dict, teste: bool) -> str:
    dia = f"{D['d1']:%d/%m/%Y}"
    if teste:
        return f"TESTE do bot de resumo (pode ignorar).\nResumo das operações do dia {dia}."
    return f"Bom dia!\nSegue resumo das operações do dia {dia}."


def _filtrar(blocos: list, selecionados: set[int] | None) -> list:
    if not selecionados:
        return blocos
    return [b for i, b in enumerate(blocos, start=1) if i in selecionados]


def enviar_imagens(texto: str, blocos: list, avisos: list[str]):
    fs = feishu_api.Feishu()
    if texto:
        fs.enviar_texto(texto)
    for legenda, imagens, nota in blocos:
        if imagens:
            fs.enviar_texto(f"{legenda}.")
            for i, png in enumerate(imagens, start=1):
                fs.enviar_imagem(fs.subir_imagem(png, f"resumo_{i}.png"))
        else:
            log.info("Bloco sem nada a enviar, pulado: %s (%s)", legenda, nota)
    if avisos:
        fs.enviar_texto("Avisos do dia:\n" + "\n".join(f"- {a}" for a in avisos))


def enviar_cartoes(wh: "feishu_api.Webhook", texto: str, blocos: list, avisos: list[str]):
    if texto:
        wh.enviar_texto(texto)
    for legenda, cartoes, nota in blocos:
        if cartoes:
            for c in cartoes:
                time.sleep(PAUSA_ENTRE_MENSAGENS)
                wh.enviar_cartao(c)
        else:
            log.info("Bloco sem nada a enviar, pulado: %s (%s)", legenda, nota)
    if avisos:
        wh.enviar_texto("Avisos do dia:\n" + "\n".join(f"- {a}" for a in avisos))


def salvar_dry_run(D: dict):
    PASTA_SAIDA.mkdir(exist_ok=True)
    print(saudacao(D, teste=False))
    n = 0
    for legenda, imagens, nota in blocos_imagens(D):
        print(f"\n{legenda}" + ("" if imagens else f"  ({nota})"))
        for png in imagens:
            n += 1
            (PASTA_SAIDA / f"{n:02d}.png").write_bytes(png)
            print(f"   imagem {n:02d}.png ({len(png) // 1024} KB)")
    todos = []
    print("\n--- cartões (modo webhook) ---")
    for legenda, cartoes, nota in blocos_cartoes(D, "Resumo automático"):
        for c in cartoes:
            tam = len(json.dumps(c, ensure_ascii=False).encode("utf-8"))
            print(f"   {c['header']['title']['content']}  [{tam / 1024:.1f} KB]")
            todos.append(c)
    (PASTA_SAIDA / "cartoes.json").write_text(json.dumps(todos, ensure_ascii=False, indent=1), encoding="utf-8")
    if D["avisos"]:
        print("\nAvisos do dia:")
        for a in D["avisos"]:
            print(" -", a)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="só renderiza em saida_bot/, sem lock, sem envio, sem gravar")
    ap.add_argument("--teste", action="store_true", help="envia marcado como TESTE, sem lock e sem gravar a previsão")
    ap.add_argument("--blocos", default="", help="números dos blocos a enviar, ex.: 1,5 (padrão: todos)")
    ap.add_argument("--sem-saudacao", action="store_true", help="não manda a mensagem de saudação (útil em testes repetidos)")
    ap.add_argument("--force", action="store_true", help="ignora o lock do dia")
    ap.add_argument("--sem-espera", action="store_true", help="não espera o Pickup/Dropoff de D-1 serem gravados")
    args = ap.parse_args(argv)
    selecionados = {int(x) for x in args.blocos.split(",") if x.strip()} or None

    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    hoje = bd.hoje_sp()
    modo = feishu_api.modo()
    if not args.dry_run and modo is None:
        log.warning("Feishu não configurado (FEISHU_WEBHOOK_URL ou FEISHU_APP_ID/SECRET/DESTINO) — nada a fazer.")
        return 0

    conn = gravar_mysql._conectar()
    try:
        bd.garantir_tabelas(conn)
        if args.dry_run:
            salvar_dry_run(coletar(conn, hoje, sem_espera=True, sincronizar=False))
            return 0

        real = not args.teste and not selecionados  # teste/blocos parciais não travam o dia nem gravam a previsão
        if real and not args.force and not bd.tentar_lock(conn, hoje):
            log.info("Resumo de %s já enviado (ou em andamento) — saindo.", hoje)
            return 0
        try:
            D = coletar(conn, hoje, sem_espera=args.sem_espera or args.teste)
            texto = None if args.sem_saudacao else saudacao(D, args.teste)
            if modo == "cartao":
                wh = feishu_api.Webhook()
                enviar_cartoes(wh, texto, _filtrar(blocos_cartoes(D, wh.rodape), selecionados),
                               [] if selecionados else D["avisos"])
            else:
                enviar_imagens(texto, _filtrar(blocos_imagens(D), selecionados), [] if selecionados else D["avisos"])
            if real and D["prev_bases"] and D["prev_pas"] and not selecionados:
                bd.gravar_previsao(conn, hoje, D["prev_bases"], D["prev_pas"])
            if real:
                bd.marcar_enviado(conn, hoje)
            log.info("Resumo de %s enviado (modo %s%s).", hoje, modo, ", TESTE" if args.teste else "")
        except Exception:
            if real:
                bd.liberar_lock(conn, hoje)
            raise
    finally:
        conn.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())

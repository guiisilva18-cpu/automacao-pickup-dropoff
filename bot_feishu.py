"""Bot do Feishu: resumo diário das operações (prints) pro Guilherme.

Roda no GitHub Actions (workflow "RESUMO FEISHU diario"), disparado quando
PICKUP/DROPOFF/TRANSFERENCIA diario terminam; o lock no banco garante 1
envio por dia. Modelo da mensagem (pedido do Guilherme, 24/09/2026):

  Bom dia! Segue resumo das operações do dia D-1.
  Taxa de coleta Pickup do dia D-1.            [print]
  Taxa de coleta Dropoff do dia D-1.           [print]
  Taxa de transferência do dia D-2.            [print]
  Ocupação dos veículos expedidos do dia D-1.  [print]
  Previsão de coleta Pickup do dia D.          [print]
  Previsão de coleta Dropoff do dia D.         [print]
  Previsão de volumes geral pro dia D.         [print]
  Assertividade da previsão do dia D-1.        [print, a partir do 2º envio]

Uso:  python bot_feishu.py            envia (exige FEISHU_* no ambiente)
      python bot_feishu.py --dry-run  só desenha os PNGs em saida_bot/ e mostra o plano
      python bot_feishu.py --force    ignora o lock do dia (reenvio manual)
"""
import argparse
import logging
import sys
from datetime import datetime, timedelta
from pathlib import Path

import bot_dados as bd
import bot_imagens as bi
import feishu_api
import gravar_mysql

log = logging.getLogger("bot_feishu")
PASTA_SAIDA = Path(__file__).parent / "saida_bot"


def _tentar(avisos: list[str], nome: str, fn):
    """Um bloco que falha não derruba o resumo: vira aviso no fim da mensagem."""
    try:
        return fn()
    except Exception as e:  # noqa: BLE001
        log.exception("Falha no bloco '%s'", nome)
        avisos.append(f"{nome}: indisponível agora ({type(e).__name__}).")
        return None


def montar(conn, hoje, sem_espera: bool = False, sincronizar: bool = True):
    d1, d2 = hoje - timedelta(days=1), hoje - timedelta(days=2)
    avisos: list[str] = []

    if sincronizar:
        _tentar(avisos, "Expedição (planilha)", lambda: bd.sincronizar_expedicao(conn))
    if not sem_espera:
        avisos += bd.aguardar_fechamento(conn, d1)

    dados_pickup = _tentar(avisos, "Pickup D-1", lambda: bd.pickup_d1(conn, d1))
    dados_dropoff = _tentar(avisos, "Dropoff D-1", lambda: bd.dropoff_d1(conn, d1))
    dados_transf = _tentar(avisos, "Transferência", lambda: bd.transferencia(d2, d1))
    dados_exped = _tentar(avisos, "Expedição D-1", lambda: bd.expedicao_d1(conn, d1))
    snapshot = _tentar(avisos, "Posição ao vivo (JMS)", lambda: bd.snapshot_pickup_ao_vivo(hoje))
    hora = datetime.now(bd.FUSO).strftime("%H:%M")
    pend_dropoff = _tentar(avisos, "Dropoff aguardando coleta", lambda: bd.dropoff_pendente_ao_vivo([d1, hoje]))

    prev_bases = bd.previsao_bases(snapshot[0]) if snapshot else None
    prev_pas = _tentar(avisos, "Previsão P.As", lambda: bd.previsao_pas(conn, d1, snapshot[1] if snapshot else {}))
    prev_drop = bd.previsao_dropoff(pend_dropoff, d1, hoje) if pend_dropoff else None
    assert_ = _tentar(avisos, "Assertividade", lambda: bd.assertividade(conn, d1))

    fmt = "%d/%m/%Y"
    blocos = []  # (legenda, [png], nota_se_vazio)
    blocos.append((f"Taxa de coleta Pickup do dia {d1:{fmt}}.",
                   [bi.img_pickup(dados_pickup)] if dados_pickup else [], "sem dados"))
    blocos.append((f"Taxa de coleta Dropoff do dia {d1:{fmt}}.",
                   [bi.img_dropoff(dados_dropoff)] if dados_dropoff else [], "sem dados"))
    blocos.append((f"Taxa de transferência do dia {d2:{fmt}}.",
                   [bi.img_transferencia(dados_transf)] if dados_transf else [], "sem dados"))
    blocos.append((f"Ocupação dos veículos expedidos do dia {d1:{fmt}}.",
                   [bi.img_expedicao(dados_exped, d1)] if dados_exped else [], "sem dados de expedição"))
    bases_img = [r for r in prev_bases if r["deveria"] > 0] if prev_bases else []
    blocos.append((f"Previsão de coleta Pickup do dia {hoje:{fmt}}.",
                   [bi.img_previsao_bases(bases_img, hora)] if bases_img else [], "sem dados"))
    drop_img = [r for r in prev_drop if r["previsto"] > 0] if prev_drop else []
    blocos.append((f"Previsão de coleta Dropoff do dia {hoje:{fmt}}.",
                   [bi.img_previsao_dropoff(drop_img, d1, hoje)] if drop_img else [],
                   "nenhum pedido aguardando coleta" if prev_drop is not None else "sem dados"))
    blocos.append((f"Previsão de volumes geral para coletar no dia {hoje:{fmt}}.",
                   [bi.img_previsao_pas(prev_pas, hora)] if prev_pas and prev_pas["linhas"] else [], "sem dados"))
    if assert_:
        blocos.append((f"Assertividade da previsão do dia {d1:{fmt}}.", bi.img_assertividade(assert_, d1), ""))

    saudacao = f"Bom dia!\nSegue resumo das operações do dia {d1:{fmt}}."
    return saudacao, blocos, avisos, (prev_bases, prev_pas)


def enviar(saudacao, blocos, avisos):
    fs = feishu_api.Feishu()
    fs.enviar_texto(saudacao)
    for legenda, imagens, nota in blocos:
        if imagens:
            fs.enviar_texto(legenda)
            for i, png in enumerate(imagens, start=1):
                fs.enviar_imagem(fs.subir_imagem(png, f"resumo_{i}.png"))
        else:
            fs.enviar_texto(f"{legenda} ({nota})")
    if avisos:
        fs.enviar_texto("Avisos do dia:\n" + "\n".join(f"- {a}" for a in avisos))


def salvar_dry_run(saudacao, blocos, avisos):
    PASTA_SAIDA.mkdir(exist_ok=True)
    print(saudacao)
    n = 0
    for legenda, imagens, nota in blocos:
        print(f"\n{legenda}" + ("" if imagens else f"  ({nota})"))
        for png in imagens:
            n += 1
            caminho = PASTA_SAIDA / f"{n:02d}.png"
            caminho.write_bytes(png)
            print(f"   -> {caminho.name} ({len(png) // 1024} KB)")
    if avisos:
        print("\nAvisos do dia:")
        for a in avisos:
            print(" -", a)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="só desenha os PNGs em saida_bot/, sem lock, sem envio, sem gravar")
    ap.add_argument("--force", action="store_true", help="ignora o lock do dia")
    ap.add_argument("--sem-espera", action="store_true", help="não espera o Pickup/Dropoff de D-1 serem gravados")
    args = ap.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    hoje = bd.hoje_sp()

    if not args.dry_run and not feishu_api.configurado():
        log.warning("FEISHU_APP_ID/FEISHU_APP_SECRET/FEISHU_DESTINO não configurados — nada a fazer.")
        return 0

    conn = gravar_mysql._conectar()
    try:
        bd.garantir_tabelas(conn)
        if args.dry_run:
            saudacao, blocos, avisos, _ = montar(conn, hoje, sem_espera=True, sincronizar=False)
            salvar_dry_run(saudacao, blocos, avisos)
            return 0

        if not args.force and not bd.tentar_lock(conn, hoje):
            log.info("Resumo de %s já enviado (ou em andamento) — saindo.", hoje)
            return 0
        try:
            saudacao, blocos, avisos, (prev_bases, prev_pas) = montar(conn, hoje, sem_espera=args.sem_espera)
            enviar(saudacao, blocos, avisos)
            if prev_bases and prev_pas:
                bd.gravar_previsao(conn, hoje, prev_bases, prev_pas)
            bd.marcar_enviado(conn, hoje)
            log.info("Resumo de %s enviado.", hoje)
        except Exception:
            bd.liberar_lock(conn, hoje)
            raise
    finally:
        conn.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())

"""Cartões do Feishu pro resumo diário (envio pelo webhook do grupo).

O webhook de bot customizado não aceita imagem, então o "print" vai como
cartão: tabela em colunas, cabeçalho vermelho J&T e bolinha verde/laranja/
vermelha na frente de cada linha (mesmas faixas de bot_regras). Cartão v1
(config/header/elements), que todo bot customizado aceita. Limite do Feishu
é ~30 KB por mensagem; as tabelas são quebradas em vários cartões ("1/2",
"2/2") antes de chegar lá.
"""
import json

import bot_regras as regras

BOLA = {regras.VERDE: "🟢", regras.LARANJA: "🟠", regras.VERMELHO: "🔴"}
LIMITE_BYTES = 24000


def fmt_int(n) -> str:
    return f"{int(n):,}".replace(",", ".")


def fmt_pct(v) -> str:
    return f"{v:.2f}".replace(".", ",") + "%"


def fmt_dec(v) -> str:
    return f"{v:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _md(texto: str, alinhar: str = "left") -> dict:
    return {"tag": "markdown", "content": texto, "text_align": alinhar}


def _linha(celulas: list[str], pesos: list[float], alinhas: list[str], fundo: str = "default") -> dict:
    return {
        "tag": "column_set", "flex_mode": "none", "background_style": fundo,
        "columns": [
            {"tag": "column", "width": "weighted", "weight": max(1, min(5, int(p))), "vertical_align": "center", "elements": [_md(c, a)]}
            for c, p, a in zip(celulas, pesos, alinhas)
        ],
    }


def kpis(itens: list[tuple[str, str]]) -> dict:
    """Faixa de números grandes no topo: [(rótulo, valor)]."""
    return {
        "tag": "column_set", "flex_mode": "none", "background_style": "grey",
        "columns": [
            {"tag": "column", "width": "weighted", "weight": 1, "vertical_align": "center",
             "elements": [_md(f"**{valor}**\n<font color='grey'>{rotulo}</font>", "center")]}
            for rotulo, valor in itens
        ],
    }


def cartao(titulo: str, elementos: list[dict], rodape: str, nota: str | None = None, template: str = "red") -> dict:
    corpo = list(elementos)
    if nota:
        corpo.append(_md(f"<font color='grey'>{nota}</font>"))
    corpo.append({"tag": "hr"})
    corpo.append({"tag": "note", "elements": [{"tag": "plain_text", "content": rodape}]})
    return {
        "config": {"wide_screen_mode": True},
        "header": {"template": template, "title": {"tag": "plain_text", "content": titulo}},
        "elements": corpo,
    }


def _bytes(obj) -> int:
    return len(json.dumps(obj, ensure_ascii=False).encode("utf-8"))


def cartoes_tabela(titulo: str, colunas: list[tuple[str, float, str]], linhas: list[list[str]], rodape: str,
                   topo: dict | None = None, nota: str | None = None) -> list[dict]:
    """colunas = [(título em markdown, peso, alinhamento)]. Quebra em vários
    cartões (título com "1/2") quando o JSON passaria de LIMITE_BYTES."""
    pesos = [c[1] for c in colunas]
    alinhas = [c[2] for c in colunas]
    cabecalho = _linha([f"**{c[0]}**" for c in colunas], pesos, alinhas, "grey")
    fixos = _bytes(cabecalho) + (_bytes(topo) if topo else 0) + _bytes(cartao(titulo, [], rodape, nota)) + 400

    grupos, atual, tamanho = [], [], fixos
    for cel in linhas:
        el = _linha(cel, pesos, alinhas)
        b = _bytes(el) + 2
        if atual and tamanho + b > LIMITE_BYTES:
            grupos.append(atual)
            atual, tamanho = [], fixos
        atual.append(el)
        tamanho += b
    grupos.append(atual)

    cartoes = []
    for k, grupo in enumerate(grupos, start=1):
        t = titulo if len(grupos) == 1 else f"{titulo} ({k}/{len(grupos)})"
        elementos = ([topo] if topo and k == 1 else []) + [cabecalho] + grupo
        cartoes.append(cartao(t, elementos, rodape, nota if k == len(grupos) else None))
    return cartoes


def _com_bola(faixa: str, texto: str) -> str:
    return f"{BOLA[faixa]} {texto}"


# ---------------------------------------------------------------- blocos do resumo
def pickup(linhas: list[dict], titulo: str, rodape: str) -> list[dict]:
    colunas = [("Base 基地", 3, "left"), ("A coletar 应揽收", 2, "right"), ("No prazo 及时", 2, "right"),
               ("Coletados + tentativas", 3, "right"), ("Real 实际", 2, "right"),
               ("Com tentativas 含尝试", 3, "right"), ("POC 签收", 2, "right")]
    rows = [[_com_bola(regras.faixa_pickup(r["taxa_tentativas"]), r["base"]), str(r["qtd_a_coletar"]),
             str(r["coletada_no_prazo"]), str(r["soma_tentativas"]), fmt_pct(r["taxa_real"]),
             f"**{fmt_pct(r['taxa_tentativas'])}**", fmt_pct(r["taxa_poc"])] for r in linhas]
    return cartoes_tabela(titulo, colunas, rows, rodape,
                          nota="🟢 98,99% ou mais · 🟠 95% a 98,98% · 🔴 abaixo de 95% (coluna Com tentativas)")


def dropoff(linhas: list[dict], titulo: str, rodape: str) -> list[dict]:
    colunas = [("Base 基地", 3, "left"), ("Pendente 待取件", 2, "right"), ("Coletado 取件成功", 2, "right"),
               ("Total 已扫描", 2, "right"), ("Taxa 取件率", 2, "right")]
    rows = [[_com_bola(regras.faixa_dropoff(r["taxa"]), r["base"]), str(r["pendente"]), str(r["coletado"]),
             str(r["total"]), f"**{fmt_pct(r['taxa'])}**"] for r in linhas]
    return cartoes_tabela(titulo, colunas, rows, rodape, nota="🟢 95% ou mais · 🟠 90% a 94,99% · 🔴 abaixo de 90%")


def transferencia(linhas: list[dict], titulo: str, rodape: str) -> list[dict]:
    colunas = [("Base", 4, "left"), ("Total", 2, "right"), ("No prazo", 2, "right"),
               ("Fora do prazo", 2, "right"), ("Taxa", 2, "right")]
    rows = [[_com_bola(regras.faixa_transferencia(r["taxa_pct"]), f"{r['base']} ({r['tipo']})"),
             str(r["entregas_total"]), str(r["entregas_no_prazo"]), str(r["entregas_fora_prazo"]),
             f"**{fmt_pct(r['taxa_pct'])}**"] for r in linhas]
    return cartoes_tabela(titulo, colunas, rows, rodape,
                          nota="🟢 95% ou mais · 🟠 90% a 94,99% · 🔴 abaixo de 90%")


def expedicao(dados: dict, titulo: str, rodape: str) -> list[dict]:
    mu = dados["mais_usado"]
    topo = kpis([
        ("Volume total expedido", fmt_int(dados["total_pacotes"])),
        ("Total de veículos", str(dados["total_veiculos"])),
        ("Média por veículo", fmt_dec(dados["media_geral"]) if dados["media_geral"] is not None else "—"),
        ("Veículo mais usado", mu[0] if mu else "—"),
    ])
    colunas = [("PA", 4, "left"), ("Líder", 3, "left"), ("Volume", 2, "right"), ("Perfil", 2, "left"),
               ("Veíc.", 1, "right"), ("% Ocup.", 2, "right")]
    rows = []
    for r in dados["linhas"]:
        pct = r["pct_ocupacao"]
        nome = _com_bola(regras.faixa_ocupacao(pct), r["pa"]) if pct is not None else f"⚪ {r['pa']}"
        rows.append([nome, r["lider"] or "—", fmt_int(r["pacotes"]) if r["pacotes"] is not None else "—",
                     r["perfil"] or "—", str(r["qtde_veiculos"]) if r["qtde_veiculos"] is not None else "—",
                     f"**{fmt_pct(pct)}**" if pct is not None else "—"])
    return cartoes_tabela(titulo, colunas, rows, rodape, topo=topo, nota="🟢 100% ou mais · 🔴 abaixo de 100%")


def previsao_bases(linhas: list[dict], hora: str, titulo: str, rodape: str) -> list[dict]:
    total = sum(r["pendente"] for r in linhas)
    colunas = [("Base", 3, "left"), ("Não coletado", 2, "right")]
    rows = [[r["base"], f"**{fmt_int(r['pendente'])}**" if r["pendente"] else "0"] for r in linhas]
    rows.append(["**TOTAL**", f"**{fmt_int(total)}**"])
    return cartoes_tabela(titulo, colunas, rows, rodape, topo=kpis([("Não coletado (previsão)", fmt_int(total))]),
                          nota=f"Extração do JMS às {hora} (TikTok): quantidade prevista que ainda não foi coletada.")


def previsao_dropoff(linhas: list[dict], d1, hoje, titulo: str, rodape: str) -> list[dict]:
    total = sum(r["previsto"] for r in linhas)
    colunas = [("Base", 3, "left"), ("Aguardando coleta", 2, "right")]
    rows = [[r["base"], f"**{fmt_int(r['previsto'])}**"] for r in linhas]
    rows.append(["**TOTAL**", f"**{fmt_int(total)}**"])
    return cartoes_tabela(titulo, colunas, rows, rodape, topo=kpis([("Previsão de coleta", fmt_int(total))]),
                          nota=f"Pedidos TikTok que já entraram no Yoyi e aguardam coleta (entrada de {d1:%d/%m} e de {hoje:%d/%m}).")


def previsao_pas(dados: dict, hora: str, titulo: str, rodape: str) -> list[dict]:
    itens = [(r["pa"], max(r["previsto"] - r["coletado"], 0)) for r in dados["linhas"]]
    itens = sorted((i for i in itens if i[1] > 0), key=lambda x: -x[1])
    if not itens:
        return []
    colunas = [("P.A", 3, "left"), ("Não coletado", 2, "right")]
    rows = [[pa, f"**{fmt_int(pend)}**"] for pa, pend in itens]
    rows.append(["**TOTAL (líquido)**", f"**{fmt_int(dados['total_pendente'])}**"])
    nota = f"Previsão = coletado ontem + 300 por P.A, sem P.As Meli. Posição às {hora} (JMS)."
    if dados["sem_movimento"]:
        nota += " Sem movimento (fora): " + ", ".join(dados["sem_movimento"]) + "."
    return cartoes_tabela(titulo, colunas, rows, rodape, topo=kpis([("Não coletado (previsão)", fmt_int(dados["total_pendente"]))]),
                          nota=nota)


def assertividade(dados: dict, d1, rodape: str) -> list[dict]:
    saida = []
    for rotulo, chave in (("Bases", "bases"), ("P.As", "pas")):
        itens = dados[chave]
        if not itens:
            continue
        passaram = sum(1 for r in itens if r["passou"])
        colunas = [("Nome", 4, "left"), ("Previsto", 2, "right"), ("Realizado", 2, "right"),
                   ("%", 2, "right"), ("Situação", 3, "left")]
        rows = [[r["nome"], fmt_int(r["previsto"]), fmt_int(r["realizado"]),
                 f"{r['pct']:.1f}".replace(".", ",") + "%",
                 ("🟢 Passou" if r["passou"] else "🔴 Não passou")] for r in itens]
        saida += cartoes_tabela(f"Assertividade da previsão do dia {d1:%d/%m/%Y} — {rotulo}", colunas, rows, rodape,
                                topo=kpis([("Passaram da previsão", f"{passaram} de {len(itens)}")]),
                                nota="Passou = realizado igual ou acima do previsto.")
    return saida

"""Desenha os prints do resumo diário do bot do Feishu (PNG, matplotlib).

Visual copiado dos prints-modelo do Guilherme: cabeçalho vermelho J&T com
texto branco bilíngue (PT + 中文), linhas inteiras pintadas por faixa
(verde/laranja/vermelho) nas taxas, e a tela de Expedição no estilo do dash.
A fonte é sempre DejaVu Sans (vem junto do matplotlib) pra a imagem sair
igual no PC e no GitHub Actions; o chinês usa uma fonte CJK do sistema
(Noto Sans CJK no Actions, Microsoft YaHei no Windows) e, se não houver
nenhuma, o cabeçalho sai só em português.
"""
import io
import logging

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib import font_manager as fm  # noqa: E402
from matplotlib.patches import Rectangle  # noqa: E402

log = logging.getLogger(__name__)

# Faixas de cor (pedido do Guilherme, 24/09/2026). Cada regra devolve a cor
# da LINHA inteira pelo valor percentual.
VERDE = "#00B050"
LARANJA = "#F8CBAD"
VERMELHO = "#FF0000"
HEADER = "#C00000"
CINZA_FUNDO = "#F2F2F2"
TEXTO = "#222222"
TEXTO_SUAVE = "#666666"
VERDE_TXT = "#1B7F3B"
VERMELHO_TXT = "#C62828"


import bot_regras as regras  # noqa: E402

_HEX = {regras.VERDE: VERDE, regras.LARANJA: LARANJA, regras.VERMELHO: VERMELHO}


def cor_pickup(taxa: float) -> str:
    return _HEX[regras.faixa_pickup(taxa)]


def cor_dropoff(taxa: float) -> str:
    return _HEX[regras.faixa_dropoff(taxa)]


def cor_transferencia(taxa: float) -> str:
    return _HEX[regras.faixa_transferencia(taxa)]


def _achar_fonte_cjk():
    for nome in ("Noto Sans CJK SC", "Noto Sans CJK JP", "Noto Sans SC", "Microsoft YaHei", "SimHei", "WenQuanYi Zen Hei"):
        try:
            caminho = fm.findfont(fm.FontProperties(family=nome), fallback_to_default=False)
            return fm.FontProperties(fname=caminho)
        except Exception:
            continue
    log.warning("Nenhuma fonte CJK encontrada; cabeçalhos sairão só em português")
    return None


_CJK = _achar_fonte_cjk()


def _fp(tamanho, negrito=False, cjk=False):
    if cjk and _CJK is not None:
        fp = _CJK.copy()
        fp.set_size(tamanho)
        if negrito:
            fp.set_weight("bold")
        return fp
    return fm.FontProperties(family="DejaVu Sans", size=tamanho, weight="bold" if negrito else "normal")


class Tela:
    """Canvas em pixels (origem no canto superior esquerdo)."""

    def __init__(self, largura: int, altura: int, fundo: str = "white"):
        self.w, self.h = largura, altura
        self.fig = plt.figure(figsize=(largura / 100, altura / 100), dpi=100, facecolor=fundo)
        self.ax = self.fig.add_axes([0, 0, 1, 1])
        self.ax.set_xlim(0, largura)
        self.ax.set_ylim(altura, 0)
        self.ax.axis("off")

    def rect(self, x, y, w, h, cor, borda=None, lw=0):
        self.ax.add_patch(Rectangle((x, y), w, h, facecolor=cor, edgecolor=borda or cor, linewidth=lw))

    def texto(self, x, y, s, tam=13, cor=TEXTO, negrito=False, ha="center", va="center", cjk=False):
        self.ax.text(x, y, s, fontproperties=_fp(tam, negrito, cjk), color=cor, ha=ha, va=va)

    def png(self) -> bytes:
        buf = io.BytesIO()
        self.fig.savefig(buf, format="png", dpi=100, facecolor=self.fig.get_facecolor())
        plt.close(self.fig)
        return buf.getvalue()


def fmt_int(n) -> str:
    return f"{int(n):,}".replace(",", ".")


def fmt_pct(v) -> str:
    return f"{v:.2f}".replace(".", ",") + "%"


def fmt_dec(v) -> str:
    return f"{v:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _n_linhas_cab(pt, zh) -> int:
    return len(pt.split("\n")) + (1 if zh and _CJK is not None else 0)


def _cabecalho_dupla(tela: Tela, x, y, w, h, pt, zh, tam=13):
    """PT em 1+ linhas ("\\n" quebra) e, embaixo, o chinês (se houver fonte)."""
    tela.rect(x, y, w, h, HEADER, borda="white", lw=1)
    linhas = pt.split("\n")
    passo = h / (_n_linhas_cab(pt, zh) + 1)
    for i, linha in enumerate(linhas):
        tela.texto(x + w / 2, y + passo * (i + 1), linha, tam, "white", True)
    if zh and _CJK is not None:
        tela.texto(x + w / 2, y + passo * (len(linhas) + 1), zh, tam, "white", True, cjk=True)


def tabela_faixas(colunas, larguras, linhas, cores, alinhas=None, tam=14, altura_linha=38, negrito_col=None) -> bytes:
    """Tabela estilo Excel do modelo: cabeçalho vermelho bilíngue + linhas
    inteiras coloridas. colunas = [(pt, zh)], linhas = lista de listas de str."""
    alinhas = alinhas or ["c"] * len(colunas)
    altura_cab = 20 + 28 * max(_n_linhas_cab(pt, zh) for pt, zh in colunas)
    largura = sum(larguras)
    altura = altura_cab + altura_linha * len(linhas) + 2
    tela = Tela(largura, altura)
    x = 0
    for (pt, zh), w in zip(colunas, larguras):
        _cabecalho_dupla(tela, x, 0, w, altura_cab, pt, zh, tam=tam - 1)
        x += w
    for i, (linha, cor) in enumerate(zip(linhas, cores)):
        y = altura_cab + i * altura_linha
        x = 0
        for j, (valor, w) in enumerate(zip(linha, larguras)):
            tela.rect(x, y, w, altura_linha, cor, borda="white", lw=0.6)
            al = alinhas[j]
            px = x + (w / 2 if al == "c" else 12 if al == "l" else w - 12)
            tela.texto(px, y + altura_linha / 2, valor, tam, "black", negrito=(negrito_col == j),
                       ha={"c": "center", "l": "left", "r": "right"}[al])
            x += w
    return tela.png()


# ---------------------------------------------------------------- blocos do resumo
def img_pickup(linhas: list[dict]) -> bytes:
    colunas = [("Data", "日期"), ("Base", "基地"), ("Qtd a coletar", "应揽收数量"), ("Qtd coletada\nno prazo", "及时揽收数量"),
               ("Soma coletados\n+ tentativas", "已收订单总数+尝试次数"), ("Taxa de coleta\nReal", "实际揽收准点率"),
               ("Taxa de coleta\ncom tentativas", "含尝试揽收率"), ("Transferência\nPOC", "签收图片")]
    larguras = [150, 170, 170, 190, 220, 190, 210, 190]
    rows = [[r["data"].strftime("%d/%m/%Y"), r["base"], str(r["qtd_a_coletar"]), str(r["coletada_no_prazo"]),
             str(r["soma_tentativas"]), fmt_pct(r["taxa_real"]), fmt_pct(r["taxa_tentativas"]), fmt_pct(r["taxa_poc"])]
            for r in linhas]
    cores = [cor_pickup(r["taxa_tentativas"]) for r in linhas]
    return tabela_faixas(colunas, larguras, rows, cores)


def img_dropoff(linhas: list[dict]) -> bytes:
    colunas = [("Data", "日期"), ("Base", "基地"), ("Pendente", "待取件"), ("Coletado", "取件成功"),
               ("Total", "已扫描总数"), ("Taxa", "取件率")]
    larguras = [190, 200, 170, 190, 190, 170]
    rows = [[r["data"].strftime("%d/%m/%Y"), r["base"], str(r["pendente"]), str(r["coletado"]), str(r["total"]),
             fmt_pct(r["taxa"])] for r in linhas]
    cores = [cor_dropoff(r["taxa"]) for r in linhas]
    return tabela_faixas(colunas, larguras, rows, cores)


def img_transferencia(linhas: list[dict]) -> bytes:
    colunas = [("Base", None), ("Tipo", None), ("Total de Entregas", None), ("Entregues no Prazo", None),
               ("Fora do Prazo", None), ("Taxa de Transferência", None)]
    larguras = [300, 120, 220, 230, 190, 260]
    rows = [[r["base"], r["tipo"], str(r["entregas_total"]), str(r["entregas_no_prazo"]),
             str(r["entregas_fora_prazo"]), fmt_pct(r["taxa_pct"])] for r in linhas]
    cores = [cor_transferencia(r["taxa_pct"]) for r in linhas]
    return tabela_faixas(colunas, larguras, rows, cores, alinhas=["l", "c", "c", "c", "c", "c"])


def _cards(tela: Tela, y, itens, altura=86, margem=24):
    n = len(itens)
    w = (tela.w - margem * (n + 1)) / n
    for i, (rotulo, valor, cor_valor) in enumerate(itens):
        x = margem + i * (w + margem)
        tela.rect(x, y, w, altura, "white", borda="#DDDDDD", lw=1.2)
        tela.rect(x, y, w, 4, "#E8A0A0")
        tela.texto(x + 18, y + 26, rotulo, 11.5, TEXTO, True, ha="left")
        tela.texto(x + 18, y + 58, valor, 24, cor_valor, True, ha="left")


def img_expedicao(dados: dict, dia) -> bytes:
    linhas = dados["linhas"]
    largura, alt_linha, topo = 1500, 40, 200
    altura = topo + 44 + alt_linha * len(linhas) + 20
    tela = Tela(largura, altura, fundo="#F4F5F7")
    tela.rect(0, 0, largura, 62, "#0F1A2E")
    tela.texto(largura / 2, 31, f"EXPEDIÇÃO  {dia:%d/%m}", 17, "white", True)
    mu = dados["mais_usado"]
    _cards(tela, 78, [
        ("VOLUME TOTAL EXPEDIDO", fmt_int(dados["total_pacotes"]), "#1565A8"),
        ("TOTAL DE VEÍCULOS", str(dados["total_veiculos"]), TEXTO),
        ("MÉDIA GERAL POR VEÍCULO", fmt_dec(dados["media_geral"]) if dados["media_geral"] is not None else "—", TEXTO),
        ("VEÍCULO MAIS USADO", f"{mu[0]}" if mu else "—", VERDE_TXT),
    ])
    tela.rect(24, topo - 6, largura - 48, 50 + alt_linha * len(linhas) + 6, "white", borda="#E3E3E3", lw=1)
    cols = [("PA", 40, "l"), ("LÍDER", 350, "l"), ("VOLUME EXPEDIDO", 660, "r"), ("PERFIL VEÍCULO", 690, "l"),
            ("QTDE VEÍCULOS", 975, "r"), ("CAPACIDADE", 1125, "r"), ("% OCUPAÇÃO", 1275, "r"), ("MÉDIA/VEÍCULO", 1450, "r")]
    for nome, x, al in cols:
        tela.texto(x, topo + 14, nome, 10.5, TEXTO_SUAVE, False, ha={"l": "left", "r": "right"}[al])
    for i, r in enumerate(linhas):
        y = topo + 44 + i * alt_linha + alt_linha / 2
        tela.ax.plot([40, largura - 40], [y - alt_linha / 2, y - alt_linha / 2], color="#EEEEEE", lw=1)
        pct = r["pct_ocupacao"]
        valores = [
            r["pa"], r["lider"] or "—", fmt_int(r["pacotes"]) if r["pacotes"] is not None else "—",
            r["perfil"] or "—", str(r["qtde_veiculos"]) if r["qtde_veiculos"] is not None else "—",
            fmt_int(r["capacidade"]) if r["capacidade"] else "—",
            fmt_pct(pct) if pct is not None else "—", fmt_dec(r["media"]) if r["media"] is not None else "—",
        ]
        for (nome, x, al), v in zip(cols, valores):
            cor = TEXTO
            if nome == "% OCUPAÇÃO" and pct is not None:
                cor = VERDE_TXT if pct >= 100 else VERMELHO_TXT
            tela.texto(x, y, v, 12, cor, False, ha={"l": "left", "r": "right"}[al])
    return tela.png()


def _tabela_simples(tela: Tela, y0, colunas, linhas, cores_status=None, alt=36, tam=13):
    """Tabela limpa (cabeçalho vermelho, linhas zebradas). colunas = [(titulo, largura, alinhamento)]."""
    x = 24
    for titulo, w, _ in colunas:
        tela.rect(x, y0, w, 44, HEADER, borda="white", lw=1)
        tela.texto(x + w / 2, y0 + 22, titulo, tam - 1, "white", True)
        x += w
    for i, linha in enumerate(linhas):
        y = y0 + 44 + i * alt
        x = 24
        for j, ((_, w, al), v) in enumerate(zip(colunas, linha)):
            fundo = "white" if i % 2 == 0 else CINZA_FUNDO
            cor_txt = TEXTO
            if cores_status and cores_status[i] and j in cores_status[i]:
                cor_txt = cores_status[i][j]
            tela.rect(x, y, w, alt, fundo, borda="#E6E6E6", lw=0.6)
            px = x + (w / 2 if al == "c" else 12 if al == "l" else w - 12)
            tela.texto(px, y + alt / 2, v, tam, cor_txt, negrito=bool(cor_txt != TEXTO),
                       ha={"c": "center", "l": "left", "r": "right"}[al])
            x += w


def _previsao_simples(rotulo_col: str, rotulo_total: str, linhas: list[tuple[str, int]], total: int, notas: list[str]) -> bytes:
    """Print de previsão com só o NÃO COLETADO (pedido do Guilherme, 24/09/2026):
    card do total + tabela nome x não coletado."""
    largura, topo, alt = 900, 130, 36
    tela = Tela(largura, topo + 44 + alt * (len(linhas) + 1) + 20 + 20 * len(notas) + 16)
    _cards(tela, 20, [(rotulo_total, fmt_int(total), VERMELHO_TXT)], margem=24)
    cols = [(rotulo_col, 580, "l"), ("Não coletado", 272, "r")]
    rows = [[nome, fmt_int(v)] for nome, v in linhas] + [["TOTAL", fmt_int(total)]]
    status = [{1: VERMELHO_TXT} if v > 0 else None for _, v in linhas] + [{1: VERMELHO_TXT}]
    _tabela_simples(tela, topo, cols, rows, status, alt=alt)
    y = topo + 44 + alt * (len(linhas) + 1) + 22
    for k, nota in enumerate(notas):
        tela.texto(24, y + 20 * k, nota, 10, TEXTO_SUAVE, ha="left")
    return tela.png()


def img_previsao_bases(linhas: list[dict], hora: str) -> bytes:
    return _previsao_simples("Base", "NÃO COLETADO (PREVISÃO)", [(r["base"], r["pendente"]) for r in linhas],
                             sum(r["pendente"] for r in linhas),
                             [f"Extração do JMS às {hora} (TikTok): quantidade prevista que ainda não foi coletada."])


def img_previsao_dropoff(linhas: list[dict], d1, hoje) -> bytes:
    return _previsao_simples("Base", "PREVISÃO DE COLETA", [(r["base"], r["previsto"]) for r in linhas],
                             sum(r["previsto"] for r in linhas),
                             [f"Pedidos TikTok que já entraram no Yoyi e aguardam coleta (entrada de {d1:%d/%m} e de {hoje:%d/%m})."])


def img_previsao_pas(dados: dict, hora: str) -> bytes:
    linhas = sorted(dados["linhas"], key=lambda r: -max(r["previsto"] - r["coletado"], 0))
    notas = [f"Previsão = coletado ontem + 300 por P.A, sem P.As Meli. Posição às {hora} (JMS). Total líquido = previsto − coletado."]
    if dados["sem_movimento"]:
        notas.append("Sem movimento (fora da previsão): " + ", ".join(dados["sem_movimento"]) + ".")
    return _previsao_simples("P.A", "NÃO COLETADO (PREVISÃO)",
                             [(r["pa"], max(r["previsto"] - r["coletado"], 0)) for r in linhas
                              if r["previsto"] - r["coletado"] > 0],
                             dados["total_pendente"], notas)


def _bloco_assert(tela: Tela, y0, titulo, itens, alt=34):
    tela.texto(24, y0, titulo, 15, HEADER, True, ha="left")
    cols = [("Nome", 340, "l"), ("Previsto", 200, "r"), ("Realizado", 200, "r"), ("Diferença", 200, "r"),
            ("% do previsto", 190, "r"), ("Situação", 170, "c")]
    rows, status = [], []
    for r in itens:
        rows.append([r["nome"], fmt_int(r["previsto"]), fmt_int(r["realizado"]),
                     ("+" if r["diferenca"] > 0 else "") + fmt_int(r["diferenca"]).replace("-", "−") if r["diferenca"] else "0",
                     f"{r['pct']:.1f}".replace(".", ",") + "%", "Passou" if r["passou"] else "Não passou"])
        cor = VERDE_TXT if r["passou"] else VERMELHO_TXT
        status.append({3: cor, 4: cor, 5: cor})
    _tabela_simples(tela, y0 + 16, cols, rows, status, alt=alt, tam=12.5)
    return y0 + 16 + 44 + alt * len(itens)


def img_assertividade(dados: dict, d1) -> list[bytes]:
    """Um print por grupo (Bases e P.As) pra manter legível."""
    saida = []
    for titulo, chave in (("Bases (Pickup — não coletado previsto x coletado depois do envio)", "bases"),
                          ("P.As (previsto x coletado no dia)", "pas")):
        itens = dados[chave]
        if not itens:
            continue
        passaram = sum(1 for r in itens if r["passou"])
        tela = Tela(1350, 60 + 16 + 44 + 34 * len(itens) + 50)
        tela.texto(24, 24, f"Assertividade da previsão de {d1:%d/%m/%Y} — {passaram} de {len(itens)} passaram da previsão",
                   13, TEXTO_SUAVE, ha="left")
        _bloco_assert(tela, 60, titulo, itens)
        saida.append(tela.png())
    return saida

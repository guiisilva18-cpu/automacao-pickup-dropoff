"""Apresentação de performance PICKUP + DROPOFF + TRANSFERÊNCIA, agosto/2026,
quebrada por semana (blocos de 7 dias a partir de 01/08 -- pedido do
Guilherme, 2026-08-13; estendida até 17/08 e com o KPI de Transferência
adicionado -- pedido do Guilherme, 2026-08-18). Estilo visual segue o
padrão J&T já usado em JT_Apresentacao_Performance_Julho2026-novo.pptx
(vermelho C00000/branco, cards de KPI, gráficos nativos, "Destaques" e
pontos de atenção).

Fonte dos dados: analise_semanal.json (gerado por analise_semanal.py a
partir de dados/historico_pickup.json, dados/historico_dropoff.json e
dados/historico_transferencia.json). "Destaques" e "Plano de Ação" são
calculados a partir dos números reais (maiores quedas/melhoras de taxa
entre semanas) em vez de texto fixo, pra continuar válido conforme mais
dias forem entrando no histórico.
"""
import json
import os
import sys
from datetime import date, timedelta

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_DATA_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu, Inches, Pt

LOGO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo_jt.jpg")

VERMELHO = RGBColor(0xC0, 0x00, 0x00)
VERMELHO_ESCURO = RGBColor(0x8C, 0x00, 0x00)
VERMELHO_CLARO = RGBColor(0xE0, 0x8A, 0x8A)
BRANCO = RGBColor(0xFF, 0xFF, 0xFF)
CINZA_TEXTO = RGBColor(0x33, 0x33, 0x33)
CINZA_CLARO = RGBColor(0xF2, 0xF2, 0xF2)
VERDE = RGBColor(0x2E, 0x7D, 0x32)
AMARELO = RGBColor(0xF9, 0xA8, 0x25)

SAIDA = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\guilherme.farias\Desktop\Apresentações\Performance Pickup e Dropoff - Agosto 2026.pptx"

with open("analise_semanal.json", encoding="utf-8") as f:
    DADOS = json.load(f)

NOMES_SEMANAS = sorted(
    (k for k in DADOS if k.startswith("Semana")),
    key=lambda k: int(k.split()[1]),
)
SEMANAS = [DADOS[k] for k in NOMES_SEMANAS]
S_PRIMEIRA, S_ULTIMA = SEMANAS[0], SEMANAS[-1]
S_PENULTIMA = SEMANAS[-2] if len(SEMANAS) > 1 else None
TRANSF = DADOS.get("Transferência (período completo)", {"dias_com_dado": [], "por_base": []})

# Expedição usa o período completo até o último dia com dado de verdade,
# independente de a semana em andamento entrar ou não na comparação
# semanal (Resumo Executivo/Pickup/Dropoff só usam semana fechada -- pedido
# do Guilherme, 2026-08-18).
DIA_INICIAL_EXPED = DADOS.get("primeiro_dia_disponivel", S_PRIMEIRA["dias_com_dado"][0])
DIA_FINAL_EXPED = DADOS.get("ultimo_dia_disponivel", S_ULTIMA["dias_com_dado"][-1])

# Bases "que a gente cuida" -- as 14 bases franquia do relatório PICKUP/DROPOFF
# (não inclui as PAs, que só entram nos indicadores de Transferência e Expedição).
import extrair_dropoff
import expedicao
BASES_ORDEM = extrair_dropoff.BASES_PICKUP


def mapa_por_base(registros):
    return {r["base"]: r for r in registros}


def totais_pickup(registros):
    deveria = sum(r["qtd_a_coletar"] for r in registros)
    no_prazo = sum(r["qtd_coletada_no_prazo"] for r in registros)
    taxa = round(no_prazo / deveria * 100, 2) if deveria else None
    return deveria, no_prazo, taxa


def totais_dropoff(registros):
    pendente = sum(r["pendente"] for r in registros)
    coletado = sum(r["coletado"] for r in registros)
    total = pendente + coletado
    taxa = round(coletado / total * 100, 2) if total else None
    return pendente, coletado, taxa

def rotulo_semana(nome_semana: str, dias_com_dado: list[str]) -> str:
    qtd = len(dias_com_dado)
    esperado = 7
    sufixo = f"({qtd} dias)" if qtd == esperado else f"(parcial, {qtd} dias)"
    return f"{nome_semana.split('(')[0].strip()}: {dias_com_dado[0][8:]}–{dias_com_dado[-1][8:]}/08 {sufixo}" if dias_com_dado else nome_semana

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
LARG, ALT = prs.slide_width, prs.slide_height
BRANCO_LAYOUT = prs.slide_layouts[6]

def nova_slide():
    return prs.slides.add_slide(BRANCO_LAYOUT)

def set_fill(shape, cor):
    shape.fill.solid()
    shape.fill.fore_color.rgb = cor
    shape.line.fill.background()

def caixa_texto(slide, left, top, width, height, texto, tamanho=14, cor=CINZA_TEXTO,
                 negrito=False, alinhamento=PP_ALIGN.LEFT, fonte="Calibri"):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = alinhamento
    r = p.add_run()
    r.text = texto
    r.font.size = Pt(tamanho)
    r.font.bold = negrito
    r.font.color.rgb = cor
    r.font.name = fonte
    return tb

def cabecalho(slide, kicker, titulo, subtitulo=None):
    barra = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, LARG, Inches(0.12))
    set_fill(barra, VERMELHO)
    caixa_texto(slide, Inches(0.5), Inches(0.24), Inches(10), Inches(0.3),
                kicker.upper(), tamanho=13, cor=VERMELHO, negrito=True)
    caixa_texto(slide, Inches(0.5), Inches(0.55), Inches(12.3), Inches(0.5),
                titulo, tamanho=19, cor=CINZA_TEXTO, negrito=True)
    if subtitulo:
        caixa_texto(slide, Inches(0.5), Inches(1.02), Inches(12.3), Inches(0.55),
                    subtitulo, tamanho=12, cor=RGBColor(0x66, 0x66, 0x66))

def rodape(slide, numero):
    barra = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, ALT - Inches(0.32), LARG, Inches(0.32))
    set_fill(barra, VERMELHO_ESCURO)
    caixa_texto(slide, Inches(0.5), ALT - Inches(0.31), Inches(9), Inches(0.3),
                "J&T EXPRESS  |  Performance Pickup & Dropoff — Agosto 2026",
                tamanho=10, cor=BRANCO)
    caixa_texto(slide, LARG - Inches(0.8), ALT - Inches(0.31), Inches(0.5), Inches(0.3),
                str(numero), tamanho=10, cor=BRANCO, alinhamento=PP_ALIGN.RIGHT)


def ativar_rotulos_pct(plot, tamanho=8, posicao=XL_DATA_LABEL_POSITION.OUTSIDE_END, cor=CINZA_TEXTO):
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format = '0"%"'
    dl.number_format_is_linked = False
    dl.font.size = Pt(tamanho)
    dl.font.bold = True
    dl.font.color.rgb = cor
    dl.position = posicao


def kpi_card(slide, left, top, width, height, rotulo, valor, nota=None, cor_valor=CINZA_TEXTO,
             tamanho_valor=24):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.adjustments[0] = 0.08
    set_fill(card, CINZA_CLARO)
    caixa_texto(slide, left + Inches(0.15), top + Inches(0.1), width - Inches(0.3), Inches(0.3),
                rotulo.upper(), tamanho=10.5, cor=RGBColor(0x77, 0x77, 0x77), negrito=True)
    caixa_texto(slide, left + Inches(0.15), top + Inches(0.42), width - Inches(0.3), Inches(0.5),
                valor, tamanho=tamanho_valor, cor=cor_valor, negrito=True)
    if nota:
        caixa_texto(slide, left + Inches(0.15), top + height - Inches(0.32), width - Inches(0.3), Inches(0.3),
                    nota, tamanho=9.5, cor=RGBColor(0x88, 0x88, 0x88))


totais_pk = [totais_pickup(s["pickup"]) for s in SEMANAS]  # (deveria, no_prazo, taxa) por semana
totais_dp = [totais_dropoff(s["dropoff"]) for s in SEMANAS]  # (pendente, coletado, taxa) por semana

taxa_transf_total = None
transf_total_entregas = sum(b["entregas_total"] for b in TRANSF["por_base"])
transf_total_no_prazo = sum(b["entregas_no_prazo"] for b in TRANSF["por_base"])
if transf_total_entregas:
    taxa_transf_total = round(transf_total_no_prazo / transf_total_entregas * 100, 2)

EXPED = expedicao.buscar_expedicao(DIA_INICIAL_EXPED, DIA_FINAL_EXPED)

# ---------------------------------------------------------------- Capa
slide = nova_slide()
fundo = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, LARG, ALT)
set_fill(fundo, VERMELHO)

logo_largura = Inches(2.6)
slide.shapes.add_picture(LOGO, (LARG - logo_largura) // 2, Inches(1.5), width=logo_largura)

caixa_texto(slide, Inches(1), Inches(4.25), Inches(11.33), Inches(0.7),
            "Performance Pickup & Dropoff", tamanho=36, cor=BRANCO, negrito=True,
            alinhamento=PP_ALIGN.CENTER)
caixa_texto(slide, Inches(1), Inches(4.95), Inches(11.33), Inches(0.4),
            "收派件绩效报告", tamanho=18, cor=RGBColor(0xFF, 0xC5, 0xC5), negrito=True,
            alinhamento=PP_ALIGN.CENTER)
caixa_texto(slide, Inches(1), Inches(5.5), Inches(11.33), Inches(0.5),
            f"Relatório — 01 a {DIA_FINAL_EXPED[8:]}/08/2026", tamanho=16, cor=BRANCO,
            alinhamento=PP_ALIGN.CENTER)
caixa_texto(slide, Inches(1), Inches(6.6), Inches(11.33), Inches(0.4),
            "J&T EXPRESS BRASIL", tamanho=12, cor=RGBColor(0xFF, 0xC5, 0xC5),
            negrito=True, alinhamento=PP_ALIGN.CENTER)

# ---------------------------------------------------------------- Slide 1 (Resumo Executivo)
slide = nova_slide()
subtitulo = "   |   ".join(rotulo_semana(n, s["dias_com_dado"]) for n, s in zip(NOMES_SEMANAS, SEMANAS))
cabecalho(slide, "Resumo Executivo 执行摘要", "Performance Pickup & Dropoff — Agosto 2026 收派件绩效 — 2026年8月", subtitulo)

n_semanas = len(SEMANAS)
gap_card = Inches(0.3)
largura_card = Emu(int((Inches(11.8) - gap_card * (n_semanas - 1)) / n_semanas))
x0 = Inches(0.5)
y_pickup = Inches(1.75)
y_dropoff = Inches(3.35)

caixa_texto(slide, x0, y_pickup - Inches(0.28), Inches(6), Inches(0.25),
            "TAXA DE PICKUP NO PRAZO", tamanho=11, cor=VERMELHO, negrito=True)
for i, (nome, (dev, prazo, taxa)) in enumerate(zip(NOMES_SEMANAS, totais_pk)):
    left = x0 + i * (largura_card + gap_card)
    cor_valor = CINZA_TEXTO
    nota = f"{prazo:,}".replace(",", ".") + " no prazo"
    if i > 0:
        anterior = totais_pk[i - 1][2]
        if taxa is not None and anterior is not None:
            delta = round(taxa - anterior, 2)
            nota = f"{'▼' if delta < 0 else '▲'} {abs(delta)} p.p. vs semana anterior"
            cor_valor = VERMELHO if delta < 0 else VERDE
    kpi_card(slide, left, y_pickup, largura_card, Inches(1.35),
             nome.split("(")[0].strip(), f"{taxa}%" if taxa is not None else "—", nota, cor_valor)

caixa_texto(slide, x0, y_dropoff - Inches(0.28), Inches(6), Inches(0.25),
            "TAXA DE DROPOFF (COLETADO)", tamanho=11, cor=VERMELHO, negrito=True)
for i, (nome, (pend, col, taxa)) in enumerate(zip(NOMES_SEMANAS, totais_dp)):
    left = x0 + i * (largura_card + gap_card)
    cor_valor = CINZA_TEXTO
    nota = f"{pend} pendente(s)"
    if i > 0:
        anterior = totais_dp[i - 1][2]
        if taxa is not None and anterior is not None:
            delta = round(taxa - anterior, 2)
            nota = f"{'▼' if delta < 0 else '▲'} {abs(delta)} p.p. vs semana anterior"
            cor_valor = VERMELHO if delta < 0 else VERDE
    kpi_card(slide, left, y_dropoff, largura_card, Inches(1.35),
             nome.split("(")[0].strip(), f"{taxa}%" if taxa is not None else "—", nota, cor_valor)

# --------- Destaques (calculados a partir dos números reais)
destaques = []

if S_PENULTIMA is not None:
    mp_prev = mapa_por_base(S_PENULTIMA["pickup"])
    mp_ult = mapa_por_base(S_ULTIMA["pickup"])
    deltas_pk = []
    for base in BASES_ORDEM:
        a, b = mp_prev.get(base, {}), mp_ult.get(base, {})
        ta, tb = a.get("taxa_real_pct"), b.get("taxa_real_pct")
        if ta is not None and tb is not None and a.get("qtd_a_coletar") and b.get("qtd_a_coletar"):
            deltas_pk.append((base, ta, tb, round(tb - ta, 2)))
    if deltas_pk:
        pior = min(deltas_pk, key=lambda x: x[3])
        melhor = max(deltas_pk, key=lambda x: x[3])
        if pior[3] < 0:
            destaques.append((
                f"{pior[0]} piorou no Pickup entre as semanas",
                f"{pior[1]:.2f}%\u2009→\u2009{pior[2]:.2f}% na taxa de coleta no prazo ({pior[3]:+.2f} p.p.) — maior queda entre as bases ativas.",
            ))
        if melhor[3] > 0 and melhor[0] != pior[0]:
            destaques.append((
                f"{melhor[0]} melhorou no Pickup entre as semanas",
                f"{melhor[1]:.2f}%\u2009→\u2009{melhor[2]:.2f}% ({melhor[3]:+.2f} p.p.) — maior avanço entre as bases ativas.",
            ))
    
pend_primeira = totais_dp[0][0]
pend_ultima = totais_dp[-1][0]
destaques.append((
    "Pendências de Dropoff",
    f"{pend_primeira} pendente(s) na {NOMES_SEMANAS[0].split('(')[0].strip()} para {pend_ultima} na {NOMES_SEMANAS[-1].split('(')[0].strip()}.",
))

mp_ult = mapa_por_base(S_ULTIMA["pickup"])
top_volume = sorted(
    (r for r in S_ULTIMA["pickup"] if r["qtd_a_coletar"]),
    key=lambda r: -r["qtd_a_coletar"],  
)[:2]
if top_volume:
    nomes = " e ".join(r["base"] for r in top_volume)
    faixas = ", ".join(f"{r['taxa_real_pct']:.2f}%" for r in top_volume)
    destaques.append((
        f"{nomes} — maior volume da rede",
        f"Taxa real na última semana: {faixas}. Maior impacto absoluto em qualquer melhora aqui.",
    ))

caixa_texto(slide, Inches(0.5), Inches(4.95), Inches(11), Inches(0.3),
            "DESTAQUES", tamanho=14, cor=VERMELHO, negrito=True)
y = Inches(5.32)
for titulo_d, corpo_d in destaques[:3]:
    marca = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.5), y + Inches(0.04), Inches(0.11), Inches(0.11))
    set_fill(marca, VERMELHO)
    caixa_texto(slide, Inches(0.75), y - Inches(0.04), Inches(11.2), Inches(0.27), titulo_d, tamanho=12.5, negrito=True)
    caixa_texto(slide, Inches(0.75), y + Inches(0.24), Inches(11.2), Inches(0.32), corpo_d, tamanho=11,
                cor=RGBColor(0x55, 0x55, 0x55))
    y += Inches(0.62)

rodape(slide, 1)

# ---- Slide 2 (Pickup)
slide = nova_slide()
cabecalho(slide, "Pickup 揽收", "Taxa de Coleta no Prazo por Base — Comparativo Semanal 各基地准点揽收率 — 每周对比",
          "Só bases franquia com volume em pelo menos uma semana. Taxa recalculada pelo total (não é média de %).")

mapas_pk = [mapa_por_base(s["pickup"]) for s in SEMANAS]
bases_com_volume = [b for b in BASES_ORDEM if any(m.get(b, {}).get("qtd_a_coletar", 0) for m in mapas_pk)]

chart_data = CategoryChartData()
chart_data.categories = bases_com_volume
cores_serie = [VERMELHO_CLARO, VERMELHO, VERMELHO_ESCURO]
for i, (nome, m) in enumerate(zip(NOMES_SEMANAS, mapas_pk)):
    chart_data.add_series(nome.split("(")[0].strip(), [m.get(b, {}).get("taxa_real_pct") or 0 for b in bases_com_volume])

grafico_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.8), chart_data
)
chart = grafico_frame.chart
chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM
chart.legend.include_in_layout = False
plot = chart.plots[0]
for i, serie in enumerate(plot.series):
    serie.format.fill.solid()
    serie.format.fill.fore_color.rgb = cores_serie[i % len(cores_serie)]
value_axis = chart.value_axis
value_axis.maximum_scale = 100
value_axis.minimum_scale = 0
ativar_rotulos_pct(plot)

caixa_texto(slide, Inches(0.5), Inches(6.75), Inches(12), Inches(0.4),
            "Meta de referência: taxa real acima de 90% (verde). A maioria das bases segue abaixo da meta.",
            tamanho=11, cor=RGBColor(0x77, 0x77, 0x77))

rodape(slide, 2)

# -Slide 3 (Dropoff)
slide = nova_slide()
cabecalho(slide, "Dropoff 自送", "Taxa de Coleta (Dropoff) por Base — Comparativo Semanal 各基地自送揽收率 — 每周对比",
          "Só bases franquia com volume em pelo menos uma semana.")

mapas_dp = [mapa_por_base(s["dropoff"]) for s in SEMANAS]
bases_dropoff = [b for b in BASES_ORDEM if any(m.get(b, {}).get("total", 0) for m in mapas_dp)]

chart_data2 = CategoryChartData()
chart_data2.categories = bases_dropoff
for i, (nome, m) in enumerate(zip(NOMES_SEMANAS, mapas_dp)):
    chart_data2.add_series(nome.split("(")[0].strip(), [m.get(b, {}).get("taxa_pct") or 0 for b in bases_dropoff])

grafico_frame2 = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.8), chart_data2
)
chart2 = grafico_frame2.chart
chart2.has_legend = True
chart2.legend.position = XL_LEGEND_POSITION.BOTTOM
chart2.legend.include_in_layout = False
plot2 = chart2.plots[0]
for i, serie in enumerate(plot2.series):
    serie.format.fill.solid()
    serie.format.fill.fore_color.rgb = cores_serie[i % len(cores_serie)]
chart2.value_axis.maximum_scale = 100
chart2.value_axis.minimum_scale = 0
ativar_rotulos_pct(plot2)

caixa_texto(slide, Inches(0.5), Inches(6.75), Inches(12), Inches(0.4),
            "Dropoff segue quase perfeito na maioria das bases — acompanhar as que passaram a ter volume pendente relevante.",
            tamanho=11, cor=RGBColor(0x77, 0x77, 0x77))

rodape(slide, 3)
# ---- Slide 4 (Transferência)#
slide = nova_slide()
periodo_transf = f"Período: {TRANSF['dias_com_dado'][0][8:]}–{TRANSF['dias_com_dado'][-1][8:]}/08 ({len(TRANSF['dias_com_dado'])} dias) — 19 PAs ativas" if TRANSF["dias_com_dado"] else "Sem dados"
cabecalho(slide, "Transferência 转运", "Taxa de Transferência — PAs Ativas 转运率 — 活跃网点", periodo_transf)

kpi_card(slide, Inches(0.5), Inches(1.75), Inches(3.3), Inches(1.5),
         "Taxa de Transferência Total", f"{taxa_transf_total}%" if taxa_transf_total is not None else "—",
         f"{transf_total_no_prazo:,}".replace(",", ".") + f" de {transf_total_entregas:,}".replace(",", ".") + " no prazo",
         cor_valor=(VERDE if (taxa_transf_total or 0) >= 90 else (AMARELO if (taxa_transf_total or 0) >= 70 else VERMELHO)),
         tamanho_valor=30)

piores = sorted((b for b in TRANSF["por_base"] if b["taxa_pct"] is not None), key=lambda b: b["taxa_pct"])[:3]
y_pior = Inches(1.85)
caixa_texto(slide, Inches(4.1), y_pior, Inches(8.2), Inches(0.28), "PIORES TAXAS NO PERÍODO", tamanho=11,
            cor=VERMELHO, negrito=True)
y_pior += Inches(0.35)
for b in piores:
    caixa_texto(slide, Inches(4.1), y_pior, Inches(8.2), Inches(0.35),
                f"{b['base']}: {b['taxa_pct']:.2f}%  ({b['entregas_no_prazo']:,}".replace(",", ".") +
                f" de {b['entregas_total']:,})".replace(",", "."),
                tamanho=12, cor=CINZA_TEXTO)
    y_pior += Inches(0.35)

bases_transf_ordenadas = sorted(TRANSF["por_base"], key=lambda b: -b["entregas_total"])
chart_data3 = CategoryChartData()
chart_data3.categories = [b["base"] for b in bases_transf_ordenadas]
chart_data3.add_series("Taxa de Transferência", [b["taxa_pct"] or 0 for b in bases_transf_ordenadas])

grafico_frame3 = slide.shapes.add_chart(
    XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.5), Inches(3.55), Inches(12.3), Inches(3.55), chart_data3
)
chart3 = grafico_frame3.chart
chart3.has_legend = False
plot3 = chart3.plots[0]
plot3.gap_width = 40
serie3 = plot3.series[0]
for i, b in enumerate(bases_transf_ordenadas):
    ponto = serie3.points[i]
    taxa = b["taxa_pct"] or 0
    ponto.format.fill.solid()
    ponto.format.fill.fore_color.rgb = VERDE if taxa >= 90 else (AMARELO if taxa >= 70 else VERMELHO)
chart3.value_axis.maximum_scale = 112
chart3.value_axis.minimum_scale = 0
chart3.category_axis.tick_labels.font.size = Pt(9)
ativar_rotulos_pct(plot3, tamanho=8.5)

rodape(slide, 4)

# ---------------------------------------------------------------- Slide 5 (Expedição)
slide = nova_slide()
qtd_dias_periodo = (date.fromisoformat(DIA_FINAL_EXPED) - date.fromisoformat(DIA_INICIAL_EXPED)).days + 1
periodo_exped = f"Período: {DIA_INICIAL_EXPED[8:]}–{DIA_FINAL_EXPED[8:]}/08 ({qtd_dias_periodo} dias) — 19 PAs ativas, planilha Controle Diário"
cabecalho(slide, "Expedição 发运", "Veículos e Ocupação — PAs Ativas 车辆及装载率 — 活跃网点", periodo_exped)

largura_card_exp = Inches(2.85)
gap_card_exp = Inches(0.2)
x0_exp = Inches(0.5)
y_cards_exp = Inches(1.6)

perfil_top = EXPED["veiculos_por_perfil"][0] if EXPED["veiculos_por_perfil"] else None
pct_frota_top = round(perfil_top["total"] / EXPED["total_veiculos"] * 100, 1) if perfil_top and EXPED["total_veiculos"] else None

cards_exp = [
    ("Total Pacotes Expedidos", f"{EXPED['total_pacotes_expedidos']:,}".replace(",", "."), None),
    ("Total de Veículos", f"{EXPED['total_veiculos']:,}".replace(",", "."), None),
    ("Média por Veículo", f"{EXPED['media_por_veiculo']:,}".replace(",", ".") if EXPED["media_por_veiculo"] else "—", "pacotes/veículo"),
    (
        "Veículo Mais Usado",
        perfil_top["perfil"] if perfil_top else "—",
        f"{perfil_top['total']} despachos ({pct_frota_top}% da frota)" if perfil_top else None,
    ),
]
for i, (rotulo, valor, nota) in enumerate(cards_exp):
    left = x0_exp + i * (largura_card_exp + gap_card_exp)
    kpi_card(slide, left, y_cards_exp, largura_card_exp, Inches(1.05), rotulo, valor, nota, tamanho_valor=22)

pas_com_capacidade = [p for p in EXPED["pas"] if p["pct_ocupacao"] is not None]
pas_ocupacao_ordenado = sorted(pas_com_capacidade, key=lambda p: -p["pct_ocupacao"])

caixa_texto(slide, Inches(0.5), Inches(2.78), Inches(11), Inches(0.24),
            "% DE OCUPAÇÃO POR PA (pacotes ÷ capacidade estimada do veículo)", tamanho=11,
            cor=VERMELHO, negrito=True)

chart_data4 = CategoryChartData()
chart_data4.categories = [p["pa"] for p in pas_ocupacao_ordenado]
chart_data4.add_series("% Ocupação", [p["pct_ocupacao"] for p in pas_ocupacao_ordenado])

grafico_frame4 = slide.shapes.add_chart(
    XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.5), Inches(3.05), Inches(12.3), Inches(3.85), chart_data4
)
chart4 = grafico_frame4.chart
chart4.has_title = False
chart4.has_legend = False
plot4 = chart4.plots[0]
plot4.gap_width = 30
serie4 = plot4.series[0]
for i, p in enumerate(pas_ocupacao_ordenado):
    ponto = serie4.points[i]
    ponto.format.fill.solid()
    taxa = p["pct_ocupacao"]
    ponto.format.fill.fore_color.rgb = VERMELHO if taxa > 115 else (AMARELO if taxa < 50 else VERDE)
chart4.value_axis.minimum_scale = 0
chart4.category_axis.tick_labels.font.size = Pt(8)
ativar_rotulos_pct(plot4, tamanho=8)

piores_ocup = pas_ocupacao_ordenado[:2]
melhores_ocup = pas_ocupacao_ordenado[-2:]


def _fmt_pa_pct(lista):
    return ", ".join(f"{p['pa']} {p['pct_ocupacao']}%" for p in lista)


nota_exp = (
    f"Acima de 115% (vermelho, possível sobrecarga): {_fmt_pa_pct(piores_ocup)}.  "
    f"Abaixo de 50% (amarelo, espaço ocioso): {_fmt_pa_pct(melhores_ocup)}."
)
caixa_texto(slide, Inches(0.5), Inches(6.98), Inches(12), Inches(0.2), nota_exp, tamanho=9,
            cor=RGBColor(0x77, 0x77, 0x77))

rodape(slide, 5)

# ---------------------------------------------------------------- Slide 6 (Expedição - detalhe por PA)
slide = nova_slide()
cabecalho(slide, "Expedição 发运", "Veículo Mais Usado e Média por PA 最常用车型及各网点平均值", periodo_exped)

# "Mais usado" = maior soma de qtde_veiculos por perfil dentro da própria
# PA no período inteiro (diferente do card "Veículo Mais Usado" da rede,
# que soma tudo junto) -- pedido do Guilherme, 2026-08-18.
pas_detalhe = sorted(EXPED["pas"], key=lambda r: -r["pacotes_expedidos"])

colunas_tabela = ["PA", "Vol. Expedido", "Veículo +Usado", "Total Veíc.", "Média Pac./Veíc."]
larguras_col = [Inches(1.75), Inches(1.15), Inches(1.15), Inches(0.85), Inches(1.2)]
largura_tabela = sum(larguras_col, Inches(0))

metade = (len(pas_detalhe) + 1) // 2
blocos = [pas_detalhe[:metade], pas_detalhe[metade:]]
x_tabelas = [Inches(0.5), Inches(0.5) + largura_tabela + Inches(0.3)]
y_tabela = Inches(1.75)
altura_linha = Inches(0.285)

for bloco, x in zip(blocos, x_tabelas):
    n_linhas = len(bloco) + 1
    tabela_shape = slide.shapes.add_table(n_linhas, len(colunas_tabela), x, y_tabela,
                                           largura_tabela, altura_linha * n_linhas)
    tabela = tabela_shape.table
    tabela.first_row = False
    for i, w in enumerate(larguras_col):
        tabela.columns[i].width = w
    tabela.rows[0].height = altura_linha
    for j, titulo_col in enumerate(colunas_tabela):
        cel = tabela.cell(0, j)
        cel.text = titulo_col
        cel.margin_top = cel.margin_bottom = Pt(1)
        cel.margin_left = cel.margin_right = Pt(3)
        cel.vertical_anchor = MSO_ANCHOR.MIDDLE
        cel.text_frame.word_wrap = False
        p_txt = cel.text_frame.paragraphs[0]
        p_txt.font.size = Pt(9)
        p_txt.font.bold = True
        p_txt.font.color.rgb = BRANCO
        cel.fill.solid()
        cel.fill.fore_color.rgb = VERMELHO

    for i, p in enumerate(bloco, start=1):
        tabela.rows[i].height = altura_linha
        valores = [
            p["pa"],
            f"{p['pacotes_expedidos']:,.0f}".replace(",", "."),
            p["perfil_mais_usado"] or "—",
            str(p["qtde_veiculos"]),
            f"{p['media_por_veiculo']:,.1f}".replace(",", ".") if p["media_por_veiculo"] else "—",
        ]
        for j, val in enumerate(valores):
            cel = tabela.cell(i, j)
            cel.text = val
            cel.margin_top = cel.margin_bottom = Pt(1)
            cel.margin_left = cel.margin_right = Pt(3)
            cel.vertical_anchor = MSO_ANCHOR.MIDDLE
            cel.text_frame.word_wrap = False
            p_txt = cel.text_frame.paragraphs[0]
            p_txt.font.size = Pt(9.5)
            p_txt.font.color.rgb = CINZA_TEXTO
            cel.fill.solid()
            cel.fill.fore_color.rgb = BRANCO if i % 2 else CINZA_CLARO

caixa_texto(slide, Inches(0.5), Inches(6.95), Inches(12), Inches(0.3),
            "\"Veículo + Usado\" considera todo o período (maior soma de veículos por perfil na PA) — pode diferir do último perfil visto no dia mais recente.",
            tamanho=9.5, cor=RGBColor(0x88, 0x88, 0x88))

rodape(slide, 6)

# ---------------------------------------------------------------- Slides 7-8 (Prints ao vivo do painel)
# Screenshot real da tela "Expedição" do App Ponto de Apoio (elemento
# #expedicao-export-area, mesma área que o botão "Exportar PNG" da própria
# tela já usa), capturado via Playwright -- pedido do Guilherme, 2026-08-18:
# ver o % de ocupação de ontem e hoje em telas separadas, com o print de
# verdade do painel (não um gráfico recriado). Os PNGs ficam em assets/ e
# precisam ser recapturados (script em scratchpad, sessão 2026-08-18) pra
# continuarem mostrando "ontem"/"hoje" de verdade em gerações futuras.
from PIL import Image

PASTA_ASSETS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
_HOJE_REAL = date.today()
_ONTEM_REAL = _HOJE_REAL - timedelta(days=1)
PRINTS_EXPEDICAO = [
    ("Ontem", _ONTEM_REAL.strftime("%d/%m/%Y"), "print_expedicao_ontem.png", 7),
    ("Hoje", _HOJE_REAL.strftime("%d/%m/%Y"), "print_expedicao_hoje.png", 8),
]

for rotulo_dia, data_str, nome_arquivo, numero_slide in PRINTS_EXPEDICAO:
    slide = nova_slide()
    cabecalho(slide, "Expedição 发运", f"Ocupação em Tempo Real — {rotulo_dia} ({data_str}) 实时装载率",
              "Print direto do painel App Ponto de Apoio, tela Expedição, filtrado para o dia.")

    caminho_imagem = os.path.join(PASTA_ASSETS, nome_arquivo)
    with Image.open(caminho_imagem) as img:
        largura_px, altura_px = img.size
    proporcao = altura_px / largura_px

    largura_max = Inches(12.3)
    altura_max = Inches(5.35)
    largura_img = largura_max
    altura_img = Emu(int(largura_max * proporcao))
    if altura_img > altura_max:
        altura_img = altura_max
        largura_img = Emu(int(altura_max / proporcao))

    left_img = (LARG - largura_img) // 2
    top_img = Inches(1.65)
    borda = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_img - Inches(0.03), top_img - Inches(0.03),
                                    largura_img + Inches(0.06), altura_img + Inches(0.06))
    set_fill(borda, RGBColor(0xE5, 0xE5, 0xE5))
    slide.shapes.add_picture(caminho_imagem, left_img, top_img, width=largura_img, height=altura_img)

    rodape(slide, numero_slide)

# ---------------------------------------------------------------- Slide 9 (Plano de ação)
slide = nova_slide()
cabecalho(slide, "Plano de Ação 行动计划", "Pontos de atenção para a próxima semana 下周关注要点")

acoes = []
if S_PENULTIMA is not None and deltas_pk:
    acoes.append((
        f"Investigar queda em {pior[0]}" if pior[3] < 0 else f"Sustentar o avanço de {melhor[0]}",
        f"Taxa de pickup variou {pior[3]:+.2f} p.p. entre as duas últimas semanas ({pior[1]:.2f}% → {pior[2]:.2f}%) — checar liderança, escala e processo de bipagem no local."
        if pior[3] < 0 else
        f"Taxa de pickup avançou {melhor[3]:+.2f} p.p. entre as duas últimas semanas — replicar o que funcionou pras demais bases.",
    ))
    piores_atual = sorted(
        [d for d in deltas_pk if d[2] is not None],
        key=lambda d: d[2],
    )[:2]
    if piores_atual:
        acoes.append((
            "Priorizar as bases com pior taxa de pickup atual",
            "; ".join(f"{b} ({tb:.2f}%)" for b, _, tb, _ in piores_atual) + " — piores taxas da última semana entre as bases ativas.",
        ))

acoes.append((
    "Investigar tendência de pendências no Dropoff",
    f"Pendente foi de {pend_primeira} ({NOMES_SEMANAS[0].split('(')[0].strip()}) para {pend_ultima} ({NOMES_SEMANAS[-1].split('(')[0].strip()}).",
))

if top_volume:
    nomes = " e ".join(r["base"] for r in top_volume)
    acoes.append((
        f"Acompanhar {nomes} de perto",
        "Maiores bases em volume de pickup da rede — qualquer melhora aqui tem o maior impacto absoluto.",
    ))

if piores:
    acoes.append((
        "Tratar as PAs com pior Taxa de Transferência",
        "; ".join(f"{b['base']} ({b['taxa_pct']:.2f}%)" for b in piores) + f" — abaixo da meta no período de {periodo_transf.split('—')[0].replace('Período: ', '').strip()}.",
    ))

y = Inches(1.75)
for i, (titulo_a, corpo_a) in enumerate(acoes[:5], start=1):
    numero = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.5), y, Inches(0.45), Inches(0.45))
    set_fill(numero, VERMELHO)
    tf = numero.text_frame
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    r = tf.paragraphs[0].add_run()
    r.text = str(i)
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = BRANCO
    caixa_texto(slide, Inches(1.2), y - Inches(0.02), Inches(11.2), Inches(0.32), titulo_a, tamanho=14, negrito=True)
    caixa_texto(slide, Inches(1.2), y + Inches(0.32), Inches(11.2), Inches(0.5), corpo_a, tamanho=11.5,
                cor=RGBColor(0x55, 0x55, 0x55))
    y += Inches(1.03)

rodape(slide, 9)

# ---------------------------------------------------------------- Slide 10 (Progresso vs Plano de Ação anterior)
# Compara os 5 pontos do PRIMEIRO plano de ação (gerado em 18/08/2026,
# período 01-17/08) contra o dado mais atual disponível -- pedido do
# Guilherme, 24/08/2026: "teve melhora ou não, faça uma análise". Os
# valores "ANTES" abaixo são um retrato fixo daquela geração (não
# recalculados a cada rodada, senão perderíamos a referência original) --
# só o "AGORA" é dinâmico, puxado da semana fechada mais recente/período
# de transferência atual.
slide = nova_slide()
cabecalho(slide, "Progresso 进展情况", "Plano de Ação Anterior — Melhorou ou Não? 之前行动计划的进展",
          "Comparado ao plano gerado em 18/08/2026 (período 01–17/08) contra o dado mais recente disponível.")


def _taxa_pickup_agora(base: str):
    r = mapa_por_base(S_ULTIMA["pickup"]).get(base)
    return r["taxa_real_pct"] if r and r.get("qtd_a_coletar") else None


def _pendente_dropoff_semana(semana_dados):
    return sum(r["pendente"] for r in semana_dados["dropoff"]) if semana_dados else None


def _transf_agora(base: str):
    for b in TRANSF["por_base"]:
        if b["base"] == base:
            return b
    return None


pendente_s2_original = 1503
pendente_agora = _pendente_dropoff_semana(S_ULTIMA)


def _bloco_verdict(slide, x, y, largura, rotulo, antes, rotulo_antes, agora, rotulo_agora,
                    maior_e_melhor=True, unidade="%", limiar_estavel=0.5):
    caixa_texto(slide, x, y, largura, Inches(0.3), rotulo, tamanho=13, negrito=True)
    tb = slide.shapes.add_textbox(x, y + Inches(0.34), largura, Inches(0.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    if antes is None or agora is None:
        r = p.add_run()
        r.text = "Sem dado suficiente pra comparar."
        r.font.size = Pt(11.5)
        r.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
        return
    fmt = (lambda v: f"{v:.2f}{unidade}") if unidade == "%" else (lambda v: f"{v:,}".replace(",", ".") + unidade)
    delta = round(agora - antes, 2)
    melhorou = (delta > 0) if maior_e_melhor else (delta < 0)
    estavel = abs(delta) < limiar_estavel
    cor_verdict = RGBColor(0x88, 0x88, 0x88) if estavel else (VERDE if melhorou else VERMELHO)
    verdict_txt = "ESTÁVEL" if estavel else ("MELHOROU" if melhorou else "PIOROU")
    sufixo_delta = "p.p." if unidade == "%" else unidade.strip() or "un."

    r1 = p.add_run()
    r1.text = f"{rotulo_antes}: {fmt(antes)}   →   {rotulo_agora}: {fmt(agora)}   "
    r1.font.size = Pt(12.5)
    r1.font.color.rgb = CINZA_TEXTO
    r2 = p.add_run()
    r2.text = f"{verdict_txt} ({delta:+,.0f} {sufixo_delta})".replace(",", ".") if unidade != "%" else f"{verdict_txt} ({delta:+.2f} p.p.)"
    r2.font.size = Pt(12.5)
    r2.font.bold = True
    r2.font.color.rgb = cor_verdict


y = Inches(1.85)
x_col = [Inches(0.5), Inches(6.65)]
altura_bloco = Inches(0.95)

_bloco_verdict(slide, x_col[0], y, Inches(5.9),
               "1. Queda em F S-JRG-SP (Pickup, vs Semana 1)",
               92.29, "Sem. 1", _taxa_pickup_agora("F S-JRG-SP"), NOMES_SEMANAS[-1].split("(")[0].strip())
_bloco_verdict(slide, x_col[1], y, Inches(5.9),
               "2. JND 02-SP — era a pior taxa de Pickup",
               73.82, "Sem. 2", _taxa_pickup_agora("JND 02-SP"), NOMES_SEMANAS[-1].split("(")[0].strip())
y += altura_bloco

_bloco_verdict(slide, x_col[0], y, Inches(5.9),
               "2b. F S-JRG-SP — também era pior taxa de Pickup",
               76.67, "Sem. 2", _taxa_pickup_agora("F S-JRG-SP"), NOMES_SEMANAS[-1].split("(")[0].strip())
_bloco_verdict(slide, x_col[1], y, Inches(5.9),
               "3. Pendências de Dropoff (tendência)",
               pendente_s2_original, "Sem. 2", pendente_agora, NOMES_SEMANAS[-1].split("(")[0].strip(),
               maior_e_melhor=False, unidade=" pendente(s)", limiar_estavel=1)
y += altura_bloco

cot_antes, cot_agora = 84.82, _taxa_pickup_agora("COT-SP")
chm_antes, chm_agora = 82.26, _taxa_pickup_agora("CHM-SP")
_bloco_verdict(slide, x_col[0], y, Inches(5.9),
               "4. COT-SP — maior volume de Pickup (vs Semana 1)",
               cot_antes, "Sem. 1", cot_agora, NOMES_SEMANAS[-1].split("(")[0].strip())
_bloco_verdict(slide, x_col[1], y, Inches(5.9),
               "4b. CHM-SP — 2ª maior volume de Pickup (vs Semana 1)",
               chm_antes, "Sem. 1", chm_agora, NOMES_SEMANAS[-1].split("(")[0].strip())
y += altura_bloco

transf_cjm02 = _transf_agora("PA MELI-CJM 02-SP")
transf_cjm14 = _transf_agora("PA MELI-CJM 14-SP")
transf_cubbo = _transf_agora("PA CUBBO-EMB-SP")
_bloco_verdict(slide, x_col[0], y, Inches(5.9),
               "5. PA MELI-CJM 02-SP — pior Taxa de Transferência",
               11.22, "01-17/08", transf_cjm02["taxa_pct"] if transf_cjm02 else None, "01-23/08")
_bloco_verdict(slide, x_col[1], y, Inches(5.9),
               "5b. PA CUBBO-EMB-SP — Taxa de Transferência",
               62.90, "01-17/08", transf_cubbo["taxa_pct"] if transf_cubbo else None, "01-23/08")
y += altura_bloco

nota_cjm14 = (
    "PA MELI-CJM 14-SP: taxa seguiu em 47.49% (74.602 de 157.087) — número IDÊNTICO ao período "
    "anterior (01-17/08), ou seja, sem nenhuma entrega nova registrada nos últimos dias. Vale checar "
    "se a base parou de operar ou se é falha na coleta do dado, antes de tratar como 'sem melhora'."
    if transf_cjm14 and transf_cjm14["entregas_total"] == 157087 else
    f"PA MELI-CJM 14-SP: {transf_cjm14['taxa_pct']:.2f}% no período atual." if transf_cjm14 else ""
)
caixa_texto(slide, Inches(0.5), y + Inches(0.05), Inches(12.3), Inches(0.5), nota_cjm14, tamanho=10.5,
            cor=RGBColor(0x88, 0x88, 0x88))

rodape(slide, 10)

# ---------------------------------------------------------------- Slide 11 (Melhoras reais no período)
# Slide dedicado só pros pontos que de fato melhoraram, com um gráfico
# maior/mais elaborado -- pedido do Guilherme, 24/08/2026: "mostre os que
# melhoraram" + "gráfico de barras bem mais elaborado". Cada item usa o
# período correto pro seu indicador (Pickup/Dropoff: Semana 2 -> Semana 3;
# Transferência: 01-17/08 -> 01-23/08), deixado explícito nas categorias
# do eixo pra não misturar bases de comparação sem dizer.
slide = nova_slide()
cabecalho(slide, "Progresso 进展情况", "Melhoras Reais no Período — O Que Avançou 本期真实进步",
          "Só indicadores com avanço confirmado nos dados. Pickup/Dropoff: Semana 2 → Semana 3. Transferência: 01–17/08 → 01–23/08.")

jnd_pk_agora = _taxa_pickup_agora("JND-SP")
jnd_dp_agora = None
carap_dp_agora = None
for r in S_ULTIMA["dropoff"]:
    if r["base"] == "JND-SP":
        jnd_dp_agora = r["taxa_pct"]
    if r["base"] == "CARAP-SP":
        carap_dp_agora = r["taxa_pct"]

melhoras_cats = [
    "JND-SP\n(Pickup)",
    "JND-SP\n(Dropoff)",
    "CARAP-SP\n(Dropoff)",
    "PA MELI-CJM 02-SP\n(Transferência)",
    "PA CUBBO-EMB-SP\n(Transferência)",
]
melhoras_antes = [86.09, 92.30, 93.27, 11.22, 62.90]
melhoras_agora = [
    jnd_pk_agora or 0,
    jnd_dp_agora or 0,
    carap_dp_agora or 0,
    transf_cjm02["taxa_pct"] if transf_cjm02 else 0,
    transf_cubbo["taxa_pct"] if transf_cubbo else 0,
]

chart_data_melhora = CategoryChartData()
chart_data_melhora.categories = melhoras_cats
chart_data_melhora.add_series("Antes", melhoras_antes)
chart_data_melhora.add_series("Agora", melhoras_agora)

grafico_melhora_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(1.75), Inches(12.3), Inches(4.4), chart_data_melhora
)
chart_melhora = grafico_melhora_frame.chart
chart_melhora.has_legend = True
chart_melhora.legend.position = XL_LEGEND_POSITION.BOTTOM
chart_melhora.legend.include_in_layout = False
plot_melhora = chart_melhora.plots[0]
for serie, cor in zip(plot_melhora.series, [RGBColor(0xBB, 0xBB, 0xBB), VERDE]):
    serie.format.fill.solid()
    serie.format.fill.fore_color.rgb = cor
chart_melhora.value_axis.maximum_scale = 100
chart_melhora.value_axis.minimum_scale = 0
ativar_rotulos_pct(plot_melhora, tamanho=10)
chart_melhora.category_axis.tick_labels.font.size = Pt(10.5)

deltas_melhora = [round(a - b, 2) for a, b in zip(melhoras_agora, melhoras_antes)]
maior_i = max(range(len(deltas_melhora)), key=lambda i: deltas_melhora[i])
destaque = (
    f"Maior avanço da rede: {melhoras_cats[maior_i].replace(chr(10), ' ')} "
    f"({melhoras_antes[maior_i]:.2f}% → {melhoras_agora[maior_i]:.2f}%, {deltas_melhora[maior_i]:+.2f} p.p.)."
)
caixa_texto(slide, Inches(0.5), Inches(6.35), Inches(12.3), Inches(0.4), destaque, tamanho=11.5,
            cor=VERDE, negrito=True)

rodape(slide, 11)

# ---------------------------------------------------------------- Slide 12 (Plano de ação -- pontos ainda críticos)
# Segundo plano de ação, agora olhando só pro que CONTINUA ruim depois da
# 1ª rodada (item por item, comparando com o slide "Progresso" acima) --
# pedido do Guilherme, 24/08/2026: "mais um plano de ação pros que estão
# ruim, todas que identificar".
slide = nova_slide()
cabecalho(slide, "Plano de Ação 行动计划", "Próximos Passos — O Que Ainda Precisa Melhorar 下一步：仍需改进的重点",
          "Itens do plano anterior que pioraram ou não avançaram, com base na comparação das páginas anteriores.")

acoes_criticas = [
    (
        "F S-JRG-SP segue em queda no Pickup",
        f"{92.29:.2f}% (Sem. 1) → {_taxa_pickup_agora('F S-JRG-SP'):.2f}% ({NOMES_SEMANAS[-1].split('(')[0].strip()}) "
        "— maior queda entre as bases acompanhadas nas últimas semanas. Prioridade máxima."
        if _taxa_pickup_agora("F S-JRG-SP") else "Sem dado suficiente pra comparar.",
    ),
    (
        "JND 02-SP continua com a pior taxa de Pickup",
        f"{73.82:.2f}% (Sem. 2) → {_taxa_pickup_agora('JND 02-SP'):.2f}% ({NOMES_SEMANAS[-1].split('(')[0].strip()}) "
        "— seguir cobrando plano de ação local (liderança, escala, bipagem)."
        if _taxa_pickup_agora("JND 02-SP") else "Sem dado suficiente pra comparar.",
    ),
    (
        "Pendências de Dropoff continuam crescendo",
        f"{pendente_s2_original} (Sem. 2) → {pendente_agora} ({NOMES_SEMANAS[-1].split('(')[0].strip()}) pendente(s) "
        "— tendência de piora precisa de ação imediata, não só monitoramento."
        if pendente_agora is not None else "Sem dado suficiente pra comparar.",
    ),
    (
        "COT-SP (maior volume de Pickup) também piorou",
        f"{cot_antes:.2f}% (Sem. 1) → {cot_agora:.2f}% ({NOMES_SEMANAS[-1].split('(')[0].strip()}) "
        "— maior base da rede em volume, qualquer queda aqui pesa muito no total."
        if cot_agora else "Sem dado suficiente pra comparar.",
    ),
    (
        "PA MELI-CJM 14-SP sem sinal de melhora na Transferência",
        "47.49% (74.602 de 157.087), número idêntico ao período anterior — confirmar com a base se "
        "parou de operar ou se é falha na coleta do dado antes de qualquer outra ação."
        if transf_cjm14 and transf_cjm14["entregas_total"] == 157087 else
        (f"{transf_cjm14['taxa_pct']:.2f}% no período atual, ainda abaixo da meta." if transf_cjm14 else "Sem dado suficiente pra comparar."),
    ),
]

y = Inches(1.75)
for i, (titulo_a, corpo_a) in enumerate(acoes_criticas[:5], start=1):
    numero = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.5), y, Inches(0.45), Inches(0.45))
    set_fill(numero, VERMELHO)
    tf = numero.text_frame
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    r = tf.paragraphs[0].add_run()
    r.text = str(i)
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = BRANCO
    caixa_texto(slide, Inches(1.2), y - Inches(0.02), Inches(11.2), Inches(0.32), titulo_a, tamanho=14, negrito=True)
    caixa_texto(slide, Inches(1.2), y + Inches(0.32), Inches(11.2), Inches(0.5), corpo_a, tamanho=11.5,
                cor=RGBColor(0x55, 0x55, 0x55))
    y += Inches(1.03)

rodape(slide, 12)

# ---------------------------------------------------------------- Encerramento
slide = nova_slide()
fundo_final = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, LARG, ALT)
set_fill(fundo_final, VERMELHO)

logo_largura_final = Inches(1.8)
slide.shapes.add_picture(LOGO, (LARG - logo_largura_final) // 2, Inches(2.2), width=logo_largura_final)

caixa_texto(slide, Inches(1), Inches(4.3), Inches(11.33), Inches(0.7),
            "Obrigado!", tamanho=40, cor=BRANCO, negrito=True, alinhamento=PP_ALIGN.CENTER)
caixa_texto(slide, Inches(1), Inches(5.3), Inches(11.33), Inches(0.5),
            "J&T EXPRESS  |  Performance Pickup & Dropoff — Agosto 2026", tamanho=13,
            cor=RGBColor(0xFF, 0xC5, 0xC5), alinhamento=PP_ALIGN.CENTER)

prs.save(SAIDA)
print(f"Salvo em {SAIDA}")

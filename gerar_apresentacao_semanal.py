"""Apresentação de performance PICKUP + DROPOFF, agosto/2026, quebrada por
semana (blocos de 7 dias a partir de 01/08 -- pedido do Guilherme,
2026-08-13). Estilo visual segue o padrão J&T já usado em
JT_Apresentacao_Performance_Julho2026-novo.pptx (vermelho C00000/branco,
cards de KPI, gráficos nativos, "Destaques" e pontos de atenção).

Fonte dos dados: analise_semanal.json (gerado por analise_semanal.py a
partir de dados/historico_pickup.json e dados/historico_dropoff.json).
"""
import json
import os
import sys

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu, Inches, Pt

LOGO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo_jt.jpg")

VERMELHO = RGBColor(0xC0, 0x00, 0x00)
VERMELHO_ESCURO = RGBColor(0x8C, 0x00, 0x00)
BRANCO = RGBColor(0xFF, 0xFF, 0xFF)
CINZA_TEXTO = RGBColor(0x33, 0x33, 0x33)
CINZA_CLARO = RGBColor(0xF2, 0xF2, 0xF2)
VERDE = RGBColor(0x2E, 0x7D, 0x32)
AMARELO = RGBColor(0xF9, 0xA8, 0x25)

SAIDA = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\guilherme.farias\Desktop\Performance Pickup e Dropoff - Agosto 2026.pptx"

with open("analise_semanal.json", encoding="utf-8") as f:
    DADOS = json.load(f)

S1 = DADOS["Semana 1 (01-07/08)"]
S2 = DADOS["Semana 2 (08-14/08)"]

BASES_ORDEM = [
    "CHM-SP", "COT-SP", "CARAP-SP", "F JND-SP", "F S-JRG-SP",
    "JND 02-SP", "JND-SP", "OSC-SP", "OSC 02-SP", "S-CSVD-SP",
]


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
    caixa_texto(slide, Inches(0.5), Inches(0.28), Inches(10), Inches(0.35),
                kicker.upper(), tamanho=13, cor=VERMELHO, negrito=True)
    caixa_texto(slide, Inches(0.5), Inches(0.62), Inches(11.5), Inches(0.55),
                titulo, tamanho=26, cor=CINZA_TEXTO, negrito=True)
    if subtitulo:
        caixa_texto(slide, Inches(0.5), Inches(1.15), Inches(11.5), Inches(0.4),
                    subtitulo, tamanho=13, cor=RGBColor(0x66, 0x66, 0x66))


def rodape(slide, numero):
    barra = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, ALT - Inches(0.32), LARG, Inches(0.32))
    set_fill(barra, VERMELHO_ESCURO)
    caixa_texto(slide, Inches(0.5), ALT - Inches(0.31), Inches(9), Inches(0.3),
                "J&T EXPRESS  |  Performance Pickup & Dropoff — Agosto 2026",
                tamanho=10, cor=BRANCO)
    caixa_texto(slide, LARG - Inches(0.8), ALT - Inches(0.31), Inches(0.5), Inches(0.3),
                str(numero), tamanho=10, cor=BRANCO, alinhamento=PP_ALIGN.RIGHT)


def kpi_card(slide, left, top, width, height, rotulo, valor, nota=None, cor_valor=CINZA_TEXTO):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.adjustments[0] = 0.08
    set_fill(card, CINZA_CLARO)
    caixa_texto(slide, left + Inches(0.15), top + Inches(0.12), width - Inches(0.3), Inches(0.35),
                rotulo.upper(), tamanho=11, cor=RGBColor(0x77, 0x77, 0x77), negrito=True)
    caixa_texto(slide, left + Inches(0.15), top + Inches(0.5), width - Inches(0.3), Inches(0.55),
                valor, tamanho=26, cor=cor_valor, negrito=True)
    if nota:
        caixa_texto(slide, left + Inches(0.15), top + height - Inches(0.4), width - Inches(0.3), Inches(0.35),
                    nota, tamanho=10, cor=RGBColor(0x88, 0x88, 0x88))


# ---------------------------------------------------------------- Capa
slide = nova_slide()
fundo = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, LARG, ALT)
set_fill(fundo, VERMELHO)

logo_largura = Inches(2.6)
slide.shapes.add_picture(LOGO, (LARG - logo_largura) // 2, Inches(1.5), width=logo_largura)

caixa_texto(slide, Inches(1), Inches(4.35), Inches(11.33), Inches(0.7),
            "Performance Pickup & Dropoff", tamanho=36, cor=BRANCO, negrito=True,
            alinhamento=PP_ALIGN.CENTER)
caixa_texto(slide, Inches(1), Inches(5.15), Inches(11.33), Inches(0.5),
            "Relatório Semanal — Agosto 2026", tamanho=18, cor=BRANCO,
            alinhamento=PP_ALIGN.CENTER)
caixa_texto(slide, Inches(1), Inches(6.6), Inches(11.33), Inches(0.4),
            "J&T EXPRESS BRASIL", tamanho=12, cor=RGBColor(0xFF, 0xC5, 0xC5),
            negrito=True, alinhamento=PP_ALIGN.CENTER)

# ---------------------------------------------------------------- Slide 1
slide = nova_slide()
cabecalho(slide, "Resumo Executivo", "Performance Pickup & Dropoff — Agosto 2026",
          f"Semana 1: {S1['dias_com_dado'][0][8:]}–{S1['dias_com_dado'][-1][8:]}/08 (7 dias)   |   "
          f"Semana 2: {S2['dias_com_dado'][0][8:]}–{S2['dias_com_dado'][-1][8:]}/08 (parcial, {len(S2['dias_com_dado'])} dias)")

dev1, prazo1, taxa1 = totais_pickup(S1["pickup"])
dev2, prazo2, taxa2 = totais_pickup(S2["pickup"])
pend1, col1, dtx1 = totais_dropoff(S1["dropoff"])
pend2, col2, dtx2 = totais_dropoff(S2["dropoff"])

delta_pickup = round(taxa2 - taxa1, 2)
delta_dropoff = round(dtx2 - dtx1, 2)

largura_card = Inches(2.75)
y_cards = Inches(1.9)
kpi_card(slide, Inches(0.5), y_cards, largura_card, Inches(1.5),
         "Taxa Pickup — Semana 1", f"{taxa1}%", f"{prazo1:,}".replace(",", ".") + " no prazo")
kpi_card(slide, Inches(3.45), y_cards, largura_card, Inches(1.5),
         "Taxa Pickup — Semana 2", f"{taxa2}%",
         f"{'▼' if delta_pickup < 0 else '▲'} {abs(delta_pickup)} p.p. vs Sem.1",
         cor_valor=(VERMELHO if delta_pickup < 0 else VERDE))
kpi_card(slide, Inches(7.1), y_cards, largura_card, Inches(1.5),
         "Taxa Dropoff — Semana 1", f"{dtx1}%", f"{pend1} pendente(s)")
kpi_card(slide, Inches(10.05), y_cards, largura_card, Inches(1.5),
         "Taxa Dropoff — Semana 2", f"{dtx2}%",
         f"{'▼' if delta_dropoff < 0 else '▲'} {abs(delta_dropoff)} p.p. vs Sem.1",
         cor_valor=(VERMELHO if delta_dropoff < 0 else VERDE))

caixa_texto(slide, Inches(0.5), Inches(3.7), Inches(11), Inches(0.35),
            "DESTAQUES", tamanho=15, cor=VERMELHO, negrito=True)

destaques = [
    ("F S-JRG-SP piorou forte no Pickup",
     "92,29% → 72,42% na taxa de coleta no prazo (-19,9 p.p.) — maior queda entre as bases ativas."),
    ("F JND-SP melhorou, mas segue crítica",
     "66,31% → 73,80% (+7,5 p.p.) — ainda a 2ª pior taxa, mesmo com a melhora."),
    ("Pendências de Dropoff dispararam",
     f"{pend1} pendente na Semana 1 para {pend2} na Semana 2 — mesmo em período mais curto (5 dias)."),
    ("COT-SP e CHM-SP (maior volume) estáveis, mas abaixo da meta",
     "Seguem na faixa de 82-85% de taxa real nas duas semanas — maior volume da rede, atenção contínua."),
]
y = Inches(4.15)
for titulo_d, corpo_d in destaques:
    marca = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.5), y + Inches(0.05), Inches(0.12), Inches(0.12))
    set_fill(marca, VERMELHO)
    caixa_texto(slide, Inches(0.75), y - Inches(0.05), Inches(11), Inches(0.3), titulo_d, tamanho=13, negrito=True)
    caixa_texto(slide, Inches(0.75), y + Inches(0.28), Inches(11), Inches(0.35), corpo_d, tamanho=11.5,
                cor=RGBColor(0x55, 0x55, 0x55))
    y += Inches(0.72)

rodape(slide, 1)

# ---------------------------------------------------------------- Slide 2 (Pickup)
slide = nova_slide()
cabecalho(slide, "Pickup", "Taxa de Coleta no Prazo por Base — Semana 1 vs Semana 2",
          "Só bases com volume em pelo menos uma das semanas. Taxa recalculada pelo total (não é média de %).")

mp1 = mapa_por_base(S1["pickup"])
mp2 = mapa_por_base(S2["pickup"])
bases_com_volume = [b for b in BASES_ORDEM if (mp1.get(b, {}).get("qtd_a_coletar", 0) or mp2.get(b, {}).get("qtd_a_coletar", 0))]

chart_data = CategoryChartData()
chart_data.categories = bases_com_volume
chart_data.add_series("Semana 1", [mp1.get(b, {}).get("taxa_real_pct") or 0 for b in bases_com_volume])
chart_data.add_series("Semana 2", [mp2.get(b, {}).get("taxa_real_pct") or 0 for b in bases_com_volume])

grafico_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.8), chart_data
)
chart = grafico_frame.chart
chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM
chart.legend.include_in_layout = False
plot = chart.plots[0]
plot.series[0].format.fill.solid()
plot.series[0].format.fill.fore_color.rgb = RGBColor(0xE0, 0x8A, 0x8A)
plot.series[1].format.fill.solid()
plot.series[1].format.fill.fore_color.rgb = VERMELHO
value_axis = chart.value_axis
value_axis.maximum_scale = 100
value_axis.minimum_scale = 0

caixa_texto(slide, Inches(0.5), Inches(6.75), Inches(12), Inches(0.4),
            "Meta de referência: taxa real acima de 90% (verde) — a maioria segue abaixo, com destaque negativo pras bases F JND-SP/F S-JRG-SP.",
            tamanho=11, cor=RGBColor(0x77, 0x77, 0x77))

rodape(slide, 2)

# ---------------------------------------------------------------- Slide 3 (Dropoff)
slide = nova_slide()
cabecalho(slide, "Dropoff", "Taxa de Coleta (Dropoff) por Base — Semana 1 vs Semana 2",
          "Só bases com volume em pelo menos uma das semanas.")

md1 = mapa_por_base(S1["dropoff"])
md2 = mapa_por_base(S2["dropoff"])
bases_dropoff = [b for b in BASES_ORDEM if (md1.get(b, {}).get("total", 0) or md2.get(b, {}).get("total", 0))]

chart_data2 = CategoryChartData()
chart_data2.categories = bases_dropoff
chart_data2.add_series("Semana 1", [md1.get(b, {}).get("taxa_pct") or 0 for b in bases_dropoff])
chart_data2.add_series("Semana 2", [md2.get(b, {}).get("taxa_pct") or 0 for b in bases_dropoff])

grafico_frame2 = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.8), chart_data2
)
chart2 = grafico_frame2.chart
chart2.has_legend = True
chart2.legend.position = XL_LEGEND_POSITION.BOTTOM
chart2.legend.include_in_layout = False
plot2 = chart2.plots[0]
plot2.series[0].format.fill.solid()
plot2.series[0].format.fill.fore_color.rgb = RGBColor(0xE0, 0x8A, 0x8A)
plot2.series[1].format.fill.solid()
plot2.series[1].format.fill.fore_color.rgb = VERMELHO
chart2.value_axis.maximum_scale = 100
chart2.value_axis.minimum_scale = 0

caixa_texto(slide, Inches(0.5), Inches(6.75), Inches(12), Inches(0.4),
            "Dropoff segue quase perfeito na maioria das bases — atenção pra CARAP-SP e JND-SP, que passaram a ter volume pendente relevante na Semana 2.",
            tamanho=11, cor=RGBColor(0x77, 0x77, 0x77))

rodape(slide, 3)

# ---------------------------------------------------------------- Slide 4 (Plano de ação)
slide = nova_slide()
cabecalho(slide, "Plano de Ação", "Pontos de atenção para a próxima semana")

acoes = [
    ("Investigar queda em F S-JRG-SP",
     "Taxa de pickup caiu quase 20 p.p. entre as semanas (92,29% → 72,42%) — checar liderança, escala e processo de bipagem no local."),
    ("Priorizar F JND-SP e F S-JRG-SP juntas",
     "As duas franquias seguem como as piores taxas de pickup da rede, mesmo com F JND-SP melhorando na Semana 2."),
    ("Investigar aumento de pendências no Dropoff",
     f"Pendente subiu de {pend1} para {pend2} unidades entre as semanas — mesmo com a Semana 2 ainda incompleta (5 dias)."),
    ("Acompanhar COT-SP e CHM-SP de perto",
     "São as duas maiores bases em volume de pickup e seguem abaixo de 90% de taxa real nas duas semanas — qualquer melhora aqui tem o maior impacto absoluto."),
]
y = Inches(1.9)
for i, (titulo_a, corpo_a) in enumerate(acoes, start=1):
    numero = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.5), y, Inches(0.5), Inches(0.5))
    set_fill(numero, VERMELHO)
    tf = numero.text_frame
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    r = tf.paragraphs[0].add_run()
    r.text = str(i)
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = BRANCO
    caixa_texto(slide, Inches(1.25), y - Inches(0.03), Inches(11), Inches(0.35), titulo_a, tamanho=15, negrito=True)
    caixa_texto(slide, Inches(1.25), y + Inches(0.35), Inches(11), Inches(0.5), corpo_a, tamanho=12,
                cor=RGBColor(0x55, 0x55, 0x55))
    y += Inches(1.15)

rodape(slide, 4)

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

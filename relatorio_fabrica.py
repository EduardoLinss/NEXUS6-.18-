"""
relatorio_fabrica.py - Relatorio em PDF do painel: a fabrica inteira.

O `relatorio_pdf.py` cuida de um setor por vez. Aqui o recorte e outro: uma
folha de rosto com os numeros do dia, o grafico comparando as funcoes no eixo
do relogio, um bloco por setor e a tabela de gargalos de toda a linha.

Reaproveita os estilos e utilitarios do relatorio de setor — cor por faixa de
carga, formatacao com virgula, rodape.
"""

from __future__ import annotations

import io
from datetime import datetime

import cpm_core
import painel as painel_mod
from relatorio_pdf import (
    AMBAR, CINZA, DISPONIVEL, LINHA, NAVY, NEUTRO, VERMELHO,
    _cor_carga, _estilo_tabela, _estilos, _num,
)

if DISPONIVEL:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        Flowable, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer,
        Table, TableStyle,
    )


class GraficoFabrica(Flowable if DISPONIVEL else object):
    """
    O mesmo grafico da tela: uma linha por funcao, no eixo do relogio.

    Na tela ele rola na horizontal; no papel precisa caber na largura, entao os
    pontos sao reescalados para a largura util da pagina.
    """

    ALTURA = 62 * mm
    MARGEM_ESQ = 16 * mm
    ALTURA_TRILHA = 3.2 * mm

    def __init__(self, curva):
        super().__init__()
        self.curva = curva

    def wrap(self, largura_disponivel, altura_disponivel):
        self.width = largura_disponivel
        self.height = (self.ALTURA + 6 * mm
                       + len(self.curva["faixas"]) * (self.ALTURA_TRILHA + 1.6 * mm))
        return self.width, self.height

    # -- conversao das coordenadas da tela para as da pagina --
    def _x(self, x_tela):
        origem = self.curva["margem_esq"]
        util_tela = max(self.curva["largura"] - origem, 1)
        util = self.width - self.MARGEM_ESQ
        return self.MARGEM_ESQ + (x_tela - origem) / util_tela * util

    def _y(self, y_tela):
        base = self.height - self.ALTURA
        return base + (1 - y_tela / self.curva["altura"]) * self.ALTURA

    def draw(self):
        canvas = self.canv
        curva = self.curva

        # grade horizontal e rotulos de percentual
        canvas.setFont("Helvetica", 6)
        for marca in curva["marcas"]:
            y = self._y(marca["y"])
            cem = marca["valor"] == 100
            canvas.setStrokeColor(colors.HexColor(VERMELHO if cem else "#DDE1EE"))
            canvas.setLineWidth(0.6 if cem else 0.3)
            if cem:
                canvas.setDash(3, 2)
            canvas.line(self.MARGEM_ESQ, y, self.width, y)
            canvas.setDash()
            canvas.setFillColor(colors.HexColor(VERMELHO if cem else CINZA))
            canvas.drawRightString(self.MARGEM_ESQ - 1.5 * mm, y - 1, f"{marca['valor']}%")

        # horas
        canvas.setFillColor(colors.HexColor(CINZA))
        canvas.setStrokeColor(colors.HexColor("#E8EBF4"))
        canvas.setLineWidth(0.3)
        for hora in curva["horas"]:
            x = self._x(hora["x"])
            canvas.line(x, self._y(curva["altura"]), x, self._y(0))
            canvas.drawCentredString(x, self._y(0) - 4 * mm, hora["rotulo"])

        # curvas
        for serie in curva["series"]:
            canvas.setStrokeColor(colors.HexColor(serie["cor"]))
            canvas.setLineWidth(1.1)
            caminho = canvas.beginPath()
            for posicao, ponto in enumerate(serie["pontos"].split()):
                x_tela, y_tela = (float(v) for v in ponto.split(","))
                x, y = self._x(x_tela), self._y(y_tela)
                caminho.moveTo(x, y) if posicao == 0 else caminho.lineTo(x, y)
            canvas.drawPath(caminho)

            canvas.setFillColor(colors.HexColor(serie["cor"]))
            canvas.circle(self._x(serie["x_pico"]), self._y(serie["y_pico"]),
                          1.5, stroke=0, fill=1)

        # trilhas de gargalo, uma por setor
        y = self.height - self.ALTURA - 6 * mm
        canvas.setFont("Helvetica", 5.5)
        for faixa in curva["faixas"]:
            y -= self.ALTURA_TRILHA + 1.6 * mm
            canvas.setFillColor(colors.HexColor(CINZA))
            canvas.drawRightString(self.MARGEM_ESQ - 1.5 * mm, y + 0.6 * mm,
                                   faixa["setor"].nome)
            canvas.setFillColor(colors.HexColor("#EDEFF6"))
            canvas.rect(self.MARGEM_ESQ, y, self.width - self.MARGEM_ESQ,
                        self.ALTURA_TRILHA, stroke=0, fill=1)
            for gargalo in faixa["gargalos"]:
                x1 = self._x(gargalo["x"])
                x2 = self._x(gargalo["x"] + gargalo["largura"])
                canvas.setFillColor(colors.HexColor(faixa["cor"]))
                canvas.setFillAlpha(gargalo["opacidade"])
                canvas.rect(x1, y, max(x2 - x1, 0.8), self.ALTURA_TRILHA,
                            stroke=0, fill=1)
            canvas.setFillAlpha(1)


def _capa(historia, estilos, resumo, contexto):
    turno = contexto.get("turno") or {}
    historia.append(Paragraph("Relatório da fábrica", estilos["titulo"]))
    historia.append(Paragraph(
        f"{turno.get('nome', '1º turno')} &nbsp;·&nbsp; produção "
        f"{contexto.get('producao', '—')} veículos &nbsp;·&nbsp; gerado em "
        f"{datetime.now():%d/%m/%Y %H:%M}", estilos["sub"]))

    numeros = [
        ["Setores analisados", str(resumo["setores_analisados"])],
        ["Gargalos", f"{resumo['total_gargalos']} "
                     f"({resumo['alta']} alta / {resumo['media']} média)"],
        ["Variantes em 100% ou mais", str(resumo["variantes_criticas"])],
        ["Postos ativos", f"{resumo['postos_ativos']} de {resumo['postos_totais']}"],
        ["Ausências", str(resumo["faltas"])],
    ]
    tabela = Table(numeros, colWidths=[60 * mm, 50 * mm])
    tabela.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor(CINZA)),
        ("FONTNAME", (1, 0), (1, -1), "Helvetica-Bold"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LINEBELOW", (0, 0), (-1, -2), 0.3, colors.HexColor(LINHA)),
    ]))
    historia.append(tabela)
    historia.append(Spacer(1, 10))

    for alerta in resumo.get("alertas", []):
        cor = VERMELHO if alerta["nivel"] == "critico" else AMBAR
        historia.append(Paragraph(
            f'<font color="{cor}"><b>{alerta["setor"].nome} · '
            f'{alerta["texto"]}</b></font> '
            f'<font size="7.5" color="{CINZA}">{alerta["detalhe"]}</font>',
            estilos["corpo"]))
    if resumo.get("alertas"):
        historia.append(Spacer(1, 8))


def _grafico(historia, estilos, curva):
    if not curva or curva.get("vazio"):
        return
    historia.append(Paragraph("Demanda ao longo do dia", estilos["secao"]))
    historia.append(GraficoFabrica(curva))
    historia.append(Spacer(1, 3))

    legenda = " &nbsp;·&nbsp; ".join(
        f'<font color="{s["cor"]}">■</font> {s["setor"].nome} '
        f'pico {s["pico"]:.0f}% às {s["hora_pico"]}'
        for s in curva["series"])
    historia.append(Paragraph(f'<font size="7">{legenda}</font>', estilos["pequeno"]))
    historia.append(Paragraph(
        f'<font size="7" color="{CINZA}">Cada linha é uma função, em % da '
        f'capacidade dela — 100% é o limite dos postos ativos. As barras abaixo '
        f'marcam quando as regras de gargalo dispararam.</font>',
        estilos["pequeno"]))
    historia.append(Spacer(1, 10))


def _setores(historia, estilos, linhas):
    historia.append(Paragraph("Carga por setor", estilos["secao"]))

    dados = [["Setor", "POPID", "Veículos", "Postos", "Gargalos", "Variantes"]]
    estilo = _estilo_tabela()
    for i, linha in enumerate(linhas, 1):
        resultado = linha.get("resultado")
        if not resultado:
            dados.append([linha["setor"].nome, linha.get("popid", "—"),
                          "—", "—", "—", linha.get("erro", "")])
            estilo.add("TEXTCOLOR", (5, i), (5, i), colors.HexColor(VERMELHO))
            continue

        cargas = []
        for variante in resultado.variantes.values():
            cor = _cor_carga(variante.total)
            cargas.append(f'<font color="{cor}">{variante.nome}: '
                          f'{_num(variante.total)}%</font>')

        resumo = linha["gargalos"]["resumo"]
        dados.append([
            linha["setor"].nome,
            resultado.popid,
            str(resultado.total_linhas),
            f"{resultado.postos_ativos}/{resultado.postos_originais}",
            f"{resumo['total']} ({resumo['alta']}A/{resumo['media']}M)",
            Paragraph(" &nbsp; ".join(cargas), estilos["pequeno"]),
        ])

    tabela = Table(dados, colWidths=[28 * mm, 22 * mm, 18 * mm, 18 * mm,
                                     26 * mm, 155 * mm])
    tabela.setStyle(estilo)
    historia.append(tabela)


def _momentos_criticos(historia, estilos, momentos):
    """
    A tabela de maior valor do relatorio: onde montagem complexa e falta de
    gente coincidem. So aparece quando algum setor tem horario de falta
    configurado — sem isso nao ha "momento" de falta para cruzar.
    """
    if not momentos:
        return

    historia.append(Paragraph(
        f"Momentos mais críticos do dia ({len(momentos)})", estilos["secao"]))
    historia.append(Paragraph(
        "Onde uma montagem complexa caiu numa janela com menos gente — o "
        "cruzamento entre gargalo e o horário de falta de cada setor. "
        "Ordenado pela pior sobrecarga primeiro.", estilos["corpo"]))
    historia.append(Spacer(1, 6))

    linhas = [["Setor", "Horário", "Sev.", "Regra", "Faltas",
              "Postos no trecho", "Sobrecarga"]]
    estilo = _estilo_tabela()
    for i, m in enumerate(momentos, 1):
        linhas.append([
            m["setor"].nome, m["hora"], m["severidade"],
            Paragraph(m["regra"], estilos["corpo"]),
            str(m["faltas"]),
            f"{m['postos_no_trecho']:.0f} de {m['postos_cheios']}",
            f"{m['sobrecarga']:.0f}%",
        ])
        cor = VERMELHO if m["severidade"] == "ALTA" else AMBAR
        estilo.add("TEXTCOLOR", (2, i), (2, i), colors.HexColor(cor))
        estilo.add("FONTNAME", (2, i), (2, i), "Helvetica-Bold")
        estilo.add("TEXTCOLOR", (6, i), (6, i),
                   colors.HexColor(VERMELHO if m["sobrecarga"] >= 100 else AMBAR))
        estilo.add("FONTNAME", (6, i), (6, i), "Helvetica-Bold")

    tabela = Table(linhas, colWidths=[24 * mm, 26 * mm, 14 * mm, 58 * mm,
                                      18 * mm, 28 * mm, 24 * mm])
    tabela.setStyle(estilo)
    historia.append(tabela)
    historia.append(Spacer(1, 10))


def _gargalos_da_linha(historia, estilos, resumo):
    historia.append(Paragraph(
        f"Gargalos da linha ({resumo['total_gargalos']})", estilos["secao"]))

    if not resumo["gargalos"]:
        historia.append(Paragraph(
            "Nenhuma regra disparou nos setores analisados.", estilos["corpo"]))
        return

    linhas = [["Setor", "Sev.", "Regra", "Posições", "Veículos", "O que fazer"]]
    estilo = _estilo_tabela()
    for i, gargalo in enumerate(resumo["gargalos"], 1):
        membros = ", ".join(
            f"{m['popid']}/{m['seq']}" if m["seq"] else m["popid"]
            for m in gargalo["membros"])
        linhas.append([
            gargalo["setor"].nome,
            gargalo["severidade"],
            Paragraph(gargalo["regra"], estilos["corpo"]),
            f"{gargalo['a']}–{gargalo['b']}",
            Paragraph(membros, estilos["pequeno"]),
            Paragraph(gargalo["detalhe"], estilos["pequeno"]),
        ])
        cor = VERMELHO if gargalo["severidade"] == "ALTA" else AMBAR
        estilo.add("TEXTCOLOR", (1, i), (1, i), colors.HexColor(cor))
        estilo.add("FONTNAME", (1, i), (1, i), "Helvetica-Bold")

    tabela = Table(linhas, colWidths=[24 * mm, 14 * mm, 48 * mm, 18 * mm,
                                      70 * mm, 93 * mm])
    tabela.setStyle(estilo)
    historia.append(tabela)


def _rodape(canvas, documento):
    canvas.saveState()
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.HexColor(CINZA))
    largura, _ = landscape(A4)
    canvas.drawRightString(largura - 15 * mm, 10 * mm,
                           f"NEXUS-6 · fábrica · página {documento.page}")
    canvas.setFillColor(colors.HexColor("#C20E17"))
    canvas.rect(15 * mm, 12 * mm, largura - 30 * mm, 0.6, stroke=0, fill=1)
    canvas.restoreState()


def gerar(linhas, resumo, curva, contexto=None, momentos=None) -> bytes:
    """PDF do painel. `linhas`, `resumo` e `curva` vêm do módulo painel."""
    if not DISPONIVEL:
        raise RuntimeError(
            "reportlab não está instalado. Rode: pip install reportlab")

    buffer = io.BytesIO()
    documento = SimpleDocTemplate(
        buffer, pagesize=landscape(A4),
        leftMargin=15 * mm, rightMargin=15 * mm,
        topMargin=14 * mm, bottomMargin=16 * mm,
        title="NEXUS-6 — relatório da fábrica", author="NEXUS-6",
    )

    estilos = _estilos()
    historia: list = []
    contexto = contexto or {}

    _capa(historia, estilos, resumo, contexto)
    _grafico(historia, estilos, curva)
    _momentos_criticos(historia, estilos, momentos or [])
    _setores(historia, estilos, linhas)
    historia.append(PageBreak())
    _gargalos_da_linha(historia, estilos, resumo)

    documento.build(historia, onFirstPage=_rodape, onLaterPages=_rodape)
    return buffer.getvalue()


def nome_arquivo(contexto=None) -> str:
    turno = (contexto or {}).get("turno") or {}
    sufixo = "turno2" if turno.get("id") == "2" else "turno1"
    return f"NEXUS6_fabrica_{sufixo}_{datetime.now():%Y%m%d_%H%M}.pdf"

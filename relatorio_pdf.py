"""
relatorio_pdf.py - Relatorio completo da analise em PDF.

Reune, num arquivo so: cabecalho do intervalo, carga por variante (com as
tarefas e o absenteismo), CPM MIX e os gargalos da sequencia.

Usa reportlab (`pip install reportlab`). Se a biblioteca nao estiver instalada,
`DISPONIVEL` fica False e a tela esconde o botao em vez de quebrar.
"""

from __future__ import annotations

import io
from datetime import datetime

from cpm_core import (
    FIM_TURNO_PADRAO, INICIO_TURNO_PADRAO, PAUSAS_PADRAO,
    SEGUNDOS_DISPONIVEIS_POR_UNIDADE, _pausas_em_segundos, relogio_apos,
    ritmo_do_turno, suavizar,
)

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_RIGHT
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        Flowable, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer,
        Table, TableStyle,
    )
    DISPONIVEL = True
except ImportError:  # reportlab ausente
    DISPONIVEL = False


NAVY = "#0F1642"
SCANIA = "#C20E17"
CINZA = "#666666"
LINHA = "#CCCCCC"
CABECALHO = "#1E2A75"

NEUTRO = "#D6D9E8"    # veiculo sem gargalo, na faixa andon

VERDE = "#1E7B34"
AMBAR = "#9A6B00"
VERMELHO = "#C0261C"


class FaixaAndon(Flowable if DISPONIVEL else object):
    """
    Linha do tempo da sequencia: um bloco por veiculo, na ordem de producao,
    colorido pela severidade. Quebra em varias linhas conforme a largura
    disponivel, com a posicao inicial marcada no comeco de cada linha.
    """

    LARGURA_BLOCO = 3.2 * mm
    ALTURA_BLOCO = 5.0 * mm
    ESPACO = 0.7 * mm
    ESPACO_LINHA = 3.6 * mm
    MARGEM_ROTULO = 12 * mm

    def __init__(self, unidades):
        super().__init__()
        self.unidades = unidades
        self._colunas = 1
        self._linhas = 1

    def wrap(self, largura_disponivel, altura_disponivel):
        util = max(10 * mm, largura_disponivel - self.MARGEM_ROTULO)
        passo = self.LARGURA_BLOCO + self.ESPACO
        self._colunas = max(1, int((util + self.ESPACO) // passo))
        self._linhas = max(1, -(-len(self.unidades) // self._colunas))
        self.width = largura_disponivel
        self.height = self._linhas * (self.ALTURA_BLOCO + self.ESPACO_LINHA)
        return self.width, self.height

    def draw(self):
        canvas = self.canv
        passo = self.LARGURA_BLOCO + self.ESPACO
        altura_linha = self.ALTURA_BLOCO + self.ESPACO_LINHA

        for indice, unidade in enumerate(self.unidades):
            linha, coluna = divmod(indice, self._colunas)
            x = self.MARGEM_ROTULO + coluna * passo
            y = self.height - (linha + 1) * altura_linha + self.ESPACO_LINHA

            if coluna == 0:
                canvas.setFont("Helvetica", 5.5)
                canvas.setFillColor(colors.HexColor(CINZA))
                canvas.drawRightString(self.MARGEM_ROTULO - 2 * mm,
                                       y + 1.4 * mm, str(indice + 1))

            heat = unidade.get("heat", "NENHUM")
            cor = {"ALTA": VERMELHO, "MEDIA": AMBAR}.get(heat, NEUTRO)
            canvas.setFillColor(colors.HexColor(cor))
            canvas.rect(x, y, self.LARGURA_BLOCO, self.ALTURA_BLOCO,
                        stroke=0, fill=1)


class LinhaDoTempo(Flowable if DISPONIVEL else object):
    """
    Linha do tempo do turno: eixo de horas, faixa de veiculos e a curva de
    demanda contra a capacidade dos postos ativos.

    A demanda de cada veiculo e a soma dos tempos das tarefas que se aplicam a
    ele, no mesmo percentual usado na carga por variante — por isso a media da
    curva bate exatamente com a carga total do setor.
    """

    ALTURA_FAIXA = 5.0 * mm
    ALTURA_GRAFICO = 42 * mm
    MARGEM_ESQ = 16 * mm
    ESPACO_EIXO = 9 * mm
    ESPACO_ENTRE = 4 * mm

    NIVEIS = (
        (0.60, "#7FC08C"),   # baixa
        (0.85, "#F0D66B"),   # media
        (1.00, "#F0A96B"),   # alta
        (99.0, "#E8756E"),   # acima da capacidade
    )

    def __init__(self, unidades, demanda, capacidade, gargalos,
                 inicio_turno=INICIO_TURNO_PADRAO, fim_turno=FIM_TURNO_PADRAO,
                 pausas=None):
        super().__init__()
        self.unidades = unidades
        self.fim_turno = fim_turno
        pausas = PAUSAS_PADRAO if pausas is None else pausas
        self.ritmo = ritmo_do_turno(len(unidades), inicio_turno, fim_turno, pausas)
        # horario de relogio de cada veiculo, ja empurrado pelas pausas
        self.horarios = [relogio_apos(i * self.ritmo, inicio_turno, fim_turno, pausas)
                         for i in range(len(unidades))]
        self.pausas = _pausas_em_segundos(inicio_turno, fim_turno, pausas)
        self.demanda = suavizar(demanda, 5)
        self.capacidade = max(capacidade, 1.0)
        self.gargalos = gargalos or []
        self.inicio_turno = inicio_turno
        pico = max(self.demanda or [0])
        # Se a demanda fica bem abaixo da capacidade, escalar pela capacidade
        # esmagaria a curva na base. Nesse caso a escala segue o pico e a linha
        # de capacidade e marcada como fora da escala.
        self.capacidade_fora_da_escala = pico < self.capacidade * 0.45
        if self.capacidade_fora_da_escala:
            self.teto = max(pico * 1.35, 1)
        else:
            self.teto = max(self.capacidade * 1.35, pico * 1.1, 1)

    def wrap(self, largura_disponivel, altura_disponivel):
        self.width = largura_disponivel
        self.height = (self.ESPACO_EIXO + self.ALTURA_FAIXA +
                       self.ESPACO_ENTRE + self.ALTURA_GRAFICO + 3 * mm)
        return self.width, self.height

    # -- geometria --
    def _passo(self):
        util = self.width - self.MARGEM_ESQ
        return util / max(len(self.unidades), 1)

    def _x(self, posicao):
        return self.MARGEM_ESQ + posicao * self._passo()

    def _y_grafico(self):
        return 3 * mm

    def _y_valor(self, valor):
        return self._y_grafico() + (valor / self.teto) * self.ALTURA_GRAFICO

    def _cor_nivel(self, valor):
        razao = valor / self.capacidade
        for limite, cor in self.NIVEIS:
            if razao < limite:
                return cor
        return VERMELHO

    # -- desenho --
    def draw(self):
        self._faixas_de_gargalo()
        self._eixo_horas()
        self._marcas_de_pausa()
        self._faixa_veiculos()
        self._curva()

    def _faixas_de_gargalo(self):
        """Marca os trechos de gargalo por tras do grafico, bem apagado: com
        muitas faixas, um rosa forte tomaria a leitura da curva."""
        canvas = self.canv
        canvas.setFillColor(colors.HexColor(VERMELHO))
        canvas.setFillAlpha(0.10)
        for gargalo in self.gargalos:
            x1 = self._x(gargalo["a"])
            x2 = self._x(gargalo["b"] + 1)
            canvas.rect(x1, self._y_grafico(), max(x2 - x1, 0.6),
                        self.height - self._y_grafico() - self.ESPACO_EIXO,
                        stroke=0, fill=1)
        canvas.setFillAlpha(1)

    def _posicao_do_horario(self, momento):
        """Onde um horario de relogio cai na faixa (posicao fracionaria)."""
        if not self.horarios or not (self.horarios[0] <= momento <= self.horarios[-1]):
            return None
        for i in range(len(self.horarios) - 1):
            atual, proximo = self.horarios[i], self.horarios[i + 1]
            if atual <= momento <= proximo:
                if proximo == atual:
                    return float(i)
                return i + (momento - atual) / (proximo - atual)
        return float(len(self.horarios) - 1)

    def _eixo_horas(self):
        """
        As marcas saem das horas cheias, nao das posicoes.

        Com ritmo de ~318 s por veiculo quase nenhum carro cai exatamente em um
        horario ":00" — varrer as posicoes procurando isso deixava o eixo quase
        vazio. Aqui a conta e invertida: para cada hora cheia, procura-se em que
        ponto da faixa ela cai. Como os horarios ja vem empurrados pelas pausas,
        as horas em volta de uma parada aparecem coladas — e o que se espera,
        porque nenhum veiculo passa nesse intervalo.
        """
        canvas = self.canv
        y = self.height - 3 * mm
        canvas.setFont("Helvetica", 6)

        marca = (int(self.horarios[0] // 3600) + 1) * 3600
        while marca <= self.horarios[-1]:
            posicao = self._posicao_do_horario(marca)
            if posicao is not None:
                x = self._x(posicao)
                canvas.setFillColor(colors.HexColor(CINZA))
                canvas.drawCentredString(x, y, "%02d:00" % (int(marca // 3600) % 24))
                canvas.setStrokeColor(colors.HexColor("#E4E7F2"))
                canvas.setLineWidth(0.3)
                canvas.line(x, self._y_grafico(), x, y - 1.5 * mm)
            marca += 3600

    def _marcas_de_pausa(self):
        """
        A pausa nao ocupa largura: nenhum veiculo passa durante ela. Por isso e
        uma linha, nao uma faixa — o salto de horario acontece exatamente ali.
        """
        canvas = self.canv
        topo = self._y_grafico() + self.ALTURA_GRAFICO + self.ESPACO_ENTRE
        for p1, p2 in self.pausas:
            posicao = self._posicao_do_horario(p1)
            if posicao is None:
                continue
            x = self._x(posicao)
            canvas.setStrokeColor(colors.HexColor("#7E88B5"))
            canvas.setLineWidth(0.7)
            canvas.setDash(1.5, 1.5)
            canvas.line(x, self._y_grafico(), x, topo + self.ALTURA_FAIXA)
            canvas.setDash()

            canvas.setFont("Helvetica", 5)
            canvas.setFillColor(colors.HexColor("#7E88B5"))
            canvas.drawCentredString(x, topo + self.ALTURA_FAIXA + 1.3 * mm,
                                     "pausa %dmin" % int((p2 - p1) // 60))

    def _faixa_veiculos(self):
        canvas = self.canv
        y = self._y_grafico() + self.ALTURA_GRAFICO + self.ESPACO_ENTRE
        passo = self._passo()
        largura = max(passo - 0.35 * mm, 0.5)

        canvas.setFont("Helvetica", 5.5)
        canvas.setFillColor(colors.HexColor(CINZA))
        canvas.drawRightString(self.MARGEM_ESQ - 2 * mm, y + 1.4 * mm, "Veículo")

        for posicao, unidade in enumerate(self.unidades):
            heat = unidade.get("heat", "NENHUM")
            cor = {"ALTA": VERMELHO, "MEDIA": AMBAR}.get(heat, NEUTRO)
            canvas.setFillColor(colors.HexColor(cor))
            canvas.rect(self._x(posicao), y, largura, self.ALTURA_FAIXA,
                        stroke=0, fill=1)

    def _curva(self):
        canvas = self.canv
        base = self._y_grafico()
        passo = self._passo()

        # area sob a curva, colorida pelo nivel de cada trecho
        for posicao, valor in enumerate(self.demanda):
            x = self._x(posicao)
            altura = self._y_valor(valor) - base
            canvas.setFillColor(colors.HexColor(self._cor_nivel(valor)))
            canvas.setFillAlpha(0.75)
            canvas.rect(x, base, passo + 0.2, max(altura, 0.2), stroke=0, fill=1)
        canvas.setFillAlpha(1)

        # linha da curva
        canvas.setStrokeColor(colors.HexColor("#333333"))
        canvas.setLineWidth(0.5)
        caminho = canvas.beginPath()
        for posicao, valor in enumerate(self.demanda):
            x = self._x(posicao) + passo / 2
            y = self._y_valor(valor)
            caminho.moveTo(x, y) if posicao == 0 else caminho.lineTo(x, y)
        canvas.drawPath(caminho)

        # capacidade
        if self.capacidade_fora_da_escala:
            y_cap = base + self.ALTURA_GRAFICO
            rotulo = f"{self.capacidade:.0f}% ↑"
        else:
            y_cap = self._y_valor(self.capacidade)
            rotulo = f"{self.capacidade:.0f}%"

        canvas.setStrokeColor(colors.HexColor(VERMELHO))
        canvas.setLineWidth(0.7)
        canvas.setDash(3, 2)
        canvas.line(self.MARGEM_ESQ, y_cap, self.width, y_cap)
        canvas.setDash()
        canvas.setFont("Helvetica", 6)
        canvas.setFillColor(colors.HexColor(VERMELHO))
        canvas.drawRightString(self.MARGEM_ESQ - 2 * mm, y_cap - 0.8 * mm, rotulo)

        # eixo vertical — so as pontas, para nao colidir com o rotulo da
        # capacidade, que fica entre elas
        canvas.setFont("Helvetica", 5.5)
        canvas.setFillColor(colors.HexColor(CINZA))
        # Com a capacidade fora da escala o rotulo dela ocupa o topo; desenhar
        # o teto ali tambem faria os dois se sobreporem.
        marcas = [0.0] if self.capacidade_fora_da_escala else [0.0, self.teto]
        for valor in marcas:
            canvas.drawRightString(self.MARGEM_ESQ - 2 * mm,
                                   self._y_valor(valor) + 0.4 * mm,
                                   f"{valor:.0f}%")


def _legenda_linha_do_tempo(estilos, capacidade, postos, inicio, fim, ritmo,
                            n_veiculos, tem_gargalos=True, pausas=None):
    parado = sum(p2 - p1 for p1, p2 in _pausas_em_segundos(
        inicio, fim, PAUSAS_PADRAO if pausas is None else pausas)) // 60
    sobre_gargalos = (" Faixas rosadas marcam os trechos de gargalo."
                      if tem_gargalos else "")
    aviso = ""
    if abs(ritmo - SEGUNDOS_DISPONIVEIS_POR_UNIDADE) > 30:
        aviso = (f' <font color="{AMBAR}">O cálculo de carga assume '
                 f'{SEGUNDOS_DISPONIVEIS_POR_UNIDADE:.0f} s por veículo; o '
                 f'turno informado dá {ritmo:.0f} s.</font>')
    return Paragraph(
        f'<font size="7" color="{CINZA}">Turno {inicio}–{fim} menos {parado} min de pausa, {n_veiculos} '
        f'veículos, {ritmo / 60:.1f} min por veículo. Demanda de trabalho por '
        f'veículo (média móvel de 5). A linha tracejada é a capacidade dos '
        f'{postos} posto(s) ativo(s) = {capacidade:.0f}%.{sobre_gargalos}'
        f'{aviso}</font>',
        estilos["pequeno"])


def _legenda_faixa(estilos):
    return Paragraph(
        f'<font size="7" color="{CINZA}">Linha do tempo da sequência — cada '
        f'bloco é um veículo, na ordem de produção. '
        f'<font color="{VERMELHO}">■</font> alta &nbsp; '
        f'<font color="{AMBAR}">■</font> média &nbsp; '
        f'<font color="{NEUTRO}">■</font> sem gargalo. '
        f'Os números à esquerda marcam a posição inicial de cada linha.</font>',
        estilos["pequeno"])


def _cor_carga(percentual: float) -> str:
    if percentual >= 100:
        return VERMELHO
    if percentual >= 90:
        return AMBAR
    return VERDE


def _num(valor: float, casas: int = 2) -> str:
    return f"{valor:.{casas}f}".replace(".", ",")


def _estilos():
    base = getSampleStyleSheet()
    return {
        "titulo": ParagraphStyle(
            "titulo", parent=base["Title"], fontSize=17, spaceAfter=2,
            alignment=0, textColor=colors.HexColor(NAVY)),
        "sub": ParagraphStyle(
            "sub", parent=base["Normal"], fontSize=9,
            textColor=colors.HexColor(CINZA), spaceAfter=14),
        "secao": ParagraphStyle(
            "secao", parent=base["Heading2"], fontSize=11, spaceBefore=14,
            spaceAfter=6, textColor=colors.HexColor(NAVY)),
        "corpo": ParagraphStyle("corpo", parent=base["Normal"], fontSize=8.5),
        "pequeno": ParagraphStyle(
            "pequeno", parent=base["Normal"], fontSize=7.5,
            textColor=colors.HexColor(CINZA)),
        "rodape": ParagraphStyle(
            "rodape", parent=base["Normal"], fontSize=7.5,
            textColor=colors.HexColor(CINZA), alignment=TA_RIGHT),
    }


def _estilo_tabela(cabecalho_escuro: bool = True) -> TableStyle:
    comandos = [
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor(LINHA)),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]
    if cabecalho_escuro:
        comandos += [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(CABECALHO)),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ]
    return TableStyle(comandos)


# ---------------------------------------------------------------------------
# Secoes
# ---------------------------------------------------------------------------
def _cabecalho(historia, estilos, resultado):
    historia.append(Paragraph("Relatório de análise do MIX", estilos["titulo"]))
    historia.append(Paragraph(
        f"POPID {resultado.popid} &nbsp;·&nbsp; linhas "
        f"{resultado.primeira_linha}–{resultado.ultima_linha} &nbsp;·&nbsp; "
        f"{resultado.total_linhas} veículos &nbsp;·&nbsp; gerado em "
        f"{datetime.now():%d/%m/%Y %H:%M}",
        estilos["sub"]))

    if getattr(resultado, "faltas", 0):
        if resultado.postos_ativos < resultado.postos_originais:
            situacao = (f"tarefas redistribuídas em {resultado.postos_ativos} "
                        f"de {resultado.postos_originais} variantes")
        else:
            situacao = f"distribuição normal, {resultado.postos_originais} variantes"
        historia.append(Paragraph(
            f"Absenteísmo: {resultado.faltas} pessoa(s) — {situacao}.",
            estilos["corpo"]))
        historia.append(Spacer(1, 6))

    for alerta in getattr(resultado, "alertas", []):
        cor = VERMELHO if alerta["nivel"] == "critico" else AMBAR
        historia.append(Paragraph(
            f'<font color="{cor}"><b>{alerta["texto"]}</b></font> '
            f'<font size="7.5" color="{CINZA}">{alerta["detalhe"]}</font>',
            estilos["corpo"]))
    if getattr(resultado, "alertas", []):
        historia.append(Spacer(1, 8))

    if resultado.avisos:
        for aviso in resultado.avisos:
            historia.append(Paragraph(f"⚠ {aviso}", estilos["pequeno"]))
        historia.append(Spacer(1, 8))


def _resumo_variantes(historia, estilos, resultado):
    historia.append(Paragraph("Carga por variante", estilos["secao"]))

    dados = [["Variante", "Tarefas", "Absenteísmo", "Total"]]
    estilo = _estilo_tabela()
    for i, variante in enumerate(resultado.variantes.values(), 1):
        absenteismo = (f"{_num(variante.absenteismo.percentual)}%"
                       if variante.absenteismo else "—")
        dados.append([variante.nome, f"{_num(variante.subtotal)}%",
                      absenteismo, f"{_num(variante.total)}%"])
        estilo.add("TEXTCOLOR", (3, i), (3, i),
                   colors.HexColor(_cor_carga(variante.total)))
        estilo.add("FONTNAME", (3, i), (3, i), "Helvetica-Bold")

    tabela = Table(dados, colWidths=[70 * mm, 32 * mm, 32 * mm, 32 * mm])
    tabela.setStyle(estilo)
    historia.append(tabela)


def _tarefas_por_variante(historia, estilos, resultado):
    for variante in resultado.variantes.values():
        bloco = [Paragraph(
            f"{variante.nome} — {_num(variante.total)}%", estilos["secao"])]

        dados = [["Tarefa", "Qtde", "Tempo (s)", "Carga"]]
        estilo = _estilo_tabela()
        for i, tarefa in enumerate(variante.linhas(), 1):
            dados.append([tarefa.nome, str(tarefa.contagem),
                          _num(tarefa.tempo, 0), f"{_num(tarefa.percentual)}%"])
            estilo.add("TEXTCOLOR", (3, i), (3, i),
                       colors.HexColor(_cor_carga(tarefa.percentual)))

        tabela = Table(dados, colWidths=[80 * mm, 22 * mm, 26 * mm, 26 * mm])
        tabela.setStyle(estilo)
        bloco.append(tabela)
        historia.append(KeepTogether(bloco))


def _mix(historia, estilos, resultado):
    historia.append(Paragraph("CPM MIX", estilos["secao"]))

    itens = [(resultado.rotulos_mix.get(chave, chave), valor)
             for chave, valor in resultado.mix.items()]

    # tres pares de colunas lado a lado, para nao gastar paginas
    colunas = 3
    por_coluna = -(-len(itens) // colunas)
    linhas = []
    for i in range(por_coluna):
        linha = []
        for c in range(colunas):
            indice = c * por_coluna + i
            if indice < len(itens):
                linha += [itens[indice][0], str(itens[indice][1])]
            else:
                linha += ["", ""]
        linhas.append(linha)

    tabela = Table(linhas, colWidths=[38 * mm, 14 * mm] * colunas)
    tabela.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LINEBELOW", (0, 0), (-1, -2), 0.3, colors.HexColor(LINHA)),
        ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#333333")),
    ] + [("FONTNAME", (c * 2 + 1, 0), (c * 2 + 1, -1), "Helvetica-Bold")
         for c in range(colunas)]))
    historia.append(tabela)


def _gargalos(historia, estilos, dados, membros_de, resultado_ref=None,
              inicio_turno=INICIO_TURNO_PADRAO, fim_turno=FIM_TURNO_PADRAO,
              pausas=None):
    historia.append(Paragraph("Demanda ao longo do turno", estilos["secao"]))

    resumo = dados.get("resumo", {})

    # Sem gargalo a linha do tempo continua valendo: ela mostra a demanda ao
    # longo do turno, e "nenhum gargalo" nao quer dizer "carga tranquila".
    if resumo.get("total"):
        historia.append(Paragraph(
            f"{resumo['total']} faixa(s) — {resumo['alta']} de severidade alta e "
            f"{resumo['media']} média. {resumo['unidades_afetadas']} de "
            f"{dados['meta']['n_unidades']} veículos envolvidos.",
            estilos["corpo"]))
    else:
        historia.append(Paragraph(
            "Nenhuma regra de gargalo disparou neste intervalo.",
            estilos["corpo"]))
    historia.append(Spacer(1, 8))

    unidades = dados.get("unidades", [])
    if unidades and getattr(resultado_ref, "demanda", None):
        capacidade = max(resultado_ref.postos_ativos, 1) * 100.0
        historia.append(KeepTogether([
            LinhaDoTempo(unidades, resultado_ref.demanda, capacidade,
                         dados.get("gargalos", []), inicio_turno, fim_turno,
                         pausas),
            Spacer(1, 2),
            _legenda_linha_do_tempo(
                estilos, capacidade, max(resultado_ref.postos_ativos, 1),
                inicio_turno, fim_turno,
                ritmo_do_turno(len(unidades), inicio_turno, fim_turno, pausas),
                len(unidades), bool(resumo.get("total")), pausas),
            Spacer(1, 10),
        ]))

    if unidades and resumo.get("total"):
        historia.append(KeepTogether([
            FaixaAndon(unidades),
            Spacer(1, 3),
            _legenda_faixa(estilos),
            Spacer(1, 10),
        ]))

    if not resumo.get("total"):
        return

    linhas = [["Sev.", "Regra", "Posições", "Veículos (POPID / SEQ)", "O que fazer"]]
    estilo = _estilo_tabela()
    for i, gargalo in enumerate(dados["gargalos"], 1):
        membros = membros_de(gargalo)
        texto_membros = ", ".join(
            f"{m['popid']}/{m['seq']}" if m["seq"] else m["popid"]
            for m in membros)
        linhas.append([
            gargalo["severidade"],
            Paragraph(gargalo["regra_nome"], estilos["corpo"]),
            f"{gargalo['a'] + 1}–{gargalo['b'] + 1}",
            Paragraph(texto_membros, estilos["pequeno"]),
            Paragraph(gargalo["detalhe"], estilos["pequeno"]),
        ])
        cor = VERMELHO if gargalo["severidade"] == "ALTA" else AMBAR
        estilo.add("TEXTCOLOR", (0, i), (0, i), colors.HexColor(cor))
        estilo.add("FONTNAME", (0, i), (0, i), "Helvetica-Bold")

    tabela = Table(linhas, colWidths=[14 * mm, 52 * mm, 18 * mm, 62 * mm, 111 * mm])
    tabela.setStyle(estilo)
    historia.append(tabela)


def _rodape(canvas, documento):
    canvas.saveState()
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.HexColor(CINZA))
    largura, _ = landscape(A4)
    canvas.drawRightString(largura - 15 * mm, 10 * mm,
                           f"NEXUS-6 · página {documento.page}")
    canvas.setFillColor(colors.HexColor(SCANIA))
    canvas.rect(15 * mm, 12 * mm, largura - 30 * mm, 0.6, stroke=0, fill=1)
    canvas.restoreState()


# ---------------------------------------------------------------------------
# Entrada
# ---------------------------------------------------------------------------
def gerar(resultado, gargalos_dados=None, membros_de=None,
          inicio_turno=INICIO_TURNO_PADRAO,
          fim_turno=FIM_TURNO_PADRAO, pausas=None) -> bytes:
    """Devolve o PDF em bytes, pronto para enviar como download."""
    if not DISPONIVEL:
        raise RuntimeError(
            "reportlab não está instalado. Rode: pip install reportlab"
        )

    buffer = io.BytesIO()
    documento = SimpleDocTemplate(
        buffer, pagesize=landscape(A4),
        leftMargin=15 * mm, rightMargin=15 * mm,
        topMargin=14 * mm, bottomMargin=16 * mm,
        title=f"MIX POPID {resultado.popid}", author="NEXUS-6",
    )
    estilos = _estilos()
    historia: list = []

    _cabecalho(historia, estilos, resultado)
    _resumo_variantes(historia, estilos, resultado)
    _tarefas_por_variante(historia, estilos, resultado)

    historia.append(PageBreak())
    _mix(historia, estilos, resultado)

    if gargalos_dados and membros_de:
        _gargalos(historia, estilos, gargalos_dados, membros_de,
                  resultado, inicio_turno, fim_turno, pausas)

    documento.build(historia, onFirstPage=_rodape, onLaterPages=_rodape)
    return buffer.getvalue()


def nome_arquivo(resultado) -> str:
    return (f"MIX_POPID_{resultado.popid}_"
            f"{datetime.now():%Y%m%d_%H%M}.pdf")

# -*- coding: utf-8 -*-
"""
mix_formatter.py — geração da planilha MIX no formato do MIX_F4.xlsx.

Mantém a ideia que já existe em `formatar_excel_mix_logic`: uma LISTA DE REGRAS
declarativas em vez de `if` espalhado. O que mudou:

* A tabela REGRAS agora consegue expressar tudo o que o MIX_F4 usa: pintar uma
  coluna, pintar a linha inteira, trocar só a cor da fonte, e comparar contra
  uma coluna diferente de "Variantes" (MODELO, MU).

* Foi removido o segundo laço de formatação — aquele que percorria TODAS as
  células no fim e fazia `if "8X2" in cell_value_str: cell.fill = azul_claro`.
  Ele rodava depois de `regras_formatacao` e sobrescrevia as regras, além de
  pintar qualquer célula (inclusive POPID e a própria Variantes) que contivesse
  o texto. Hoje isso virou uma regra normal, com coluna de origem e alvo
  explícitos.

* A planilha é escrita num Workbook novo. O código atual fazia
  `load_workbook(original)` e apagava o valor de cada célula, herdando estilos,
  linhas ocultas e colunas fantasma da planilha de origem (é por isso que o
  arquivo sai com colunas J..BE com largura definida e sem conteúdo).

* Layout (fontes por coluna, larguras, altura de linha, bordas, código de
  barras na coluna A, área de impressão) vive num único dicionário COLUNAS.

Regras conferidas contra o MIX_F4.xlsx — reproduzem 85/85 linhas:
  MU ≠ BR01 → MU vermelho | CHIFRÃO → fonte vermelha | ANJO → fonte azul
  GÁS → amarelo | TANQUE A → laranja | 8X2/8X4 → MODELO azul claro
  EURO 3 → linha inteira verde clara

Duas coisas no MIX_F4 que NÃO saem de regra (ver `PENDENTE` no fim do arquivo):
a linha 58 inteira azul e o CHASSI azul das linhas 56 e 58.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from collections import Counter
import re
from typing import Callable, Sequence

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from mix_processor_fibras import anexar_mix_fibras

# ==========================================================================
# Paleta
# ==========================================================================

# ARGB de 8 dígitos: sem o "FF" de opacidade o Excel grava 00xxxxxx e a cor
# sai transparente/preta em algumas versões.
AZUL = "FF0070C0"
AZUL_CLARO = "FFCDEAF7"
AZUL_ESCURO = "FF163E64"
VERDE = "FF00B050"
VERDE_CLARO = "FFCCFFCC"
VERMELHO = "FFFF0000"
AMARELO = "FFFFFF00"
LARANJA = "FFFFA500"
CINZA_CABECALHO = "FFC0C0C0"

FONTE_CODIGO_BARRAS = "Free 3 of 9 Extended"  # precisa estar instalada na máquina
FONTE_PADRAO = "Arial"

ALTURA_CABECALHO = 30.0
ALTURA_LINHA = 50.1


# ==========================================================================
# Layout das colunas — ordem, origem, largura e tipografia
# ==========================================================================


@dataclass(frozen=True)
class Coluna:
    titulo: str
    origem: str | Callable[[pd.Series], object]  # nome da coluna no df, ou função
    largura: float
    tamanho: int
    fonte: str = FONTE_PADRAO
    formato: str = "General"
    negrito: bool = True

    def valor(self, linha: pd.Series):
        if callable(self.origem):
            return self.origem(linha)
        return linha.get(self.origem)


def _codigo_barras_popid(linha: pd.Series) -> str:
    """Coluna A: o POPID entre asteriscos, que é o que a fonte Code 39 lê."""
    popid = linha.get("POPID")
    if pd.isna(popid):
        return ""
    if isinstance(popid, float) and popid.is_integer():
        popid = int(popid)
    return f"*{popid}*"


COLUNAS: Sequence[Coluna] = (
    Coluna("COD POPID",      _codigo_barras_popid, 32.00, 50, FONTE_CODIGO_BARRAS, negrito=False),
    Coluna("POPID",          "POPID",              20.14, 28, formato="0"),
    Coluna("CHASSI",         "CHASSI",             14.29, 18, formato="0"),
    Coluna("MODELO",         "MODELO",             22.00, 12),
    Coluna("CHASSIS HEIGHT", "CHASSIS HEIGHT",     20.57,  8),
    Coluna("SUSP TRASEIRA",  "SUSP TRASEIRA",       7.00,  8),
    Coluna("MU",             "MU",                  4.71,  8),
    Coluna("COR",            "COR",                 5.29,  8),
    Coluna("Variantes",      "Variantes",          46.00, 26),
)

COLUNAS_PADRAO = tuple(c.titulo for c in COLUNAS)


def _chave_coluna(nome) -> str:
    return " ".join(str(nome).strip().upper().split())


def _colunas_escolhidas(colunas_desejadas=None) -> Sequence["Coluna"]:
    """
    Filtra o layout pelas colunas pedidas, mantendo a ordem original.

    O layout desta funcao e fixo (largura, fonte, formato por coluna), entao
    configurar e escolher quais dessas colunas entram no arquivo final.
    """
    if not colunas_desejadas:
        return COLUNAS
    pedidas = {_chave_coluna(c) for c in colunas_desejadas}
    escolhidas = tuple(c for c in COLUNAS if _chave_coluna(c.titulo) in pedidas)
    return escolhidas or COLUNAS


LINHA_INTEIRA = "*"  # alvo especial


# ==========================================================================
# Regras de formatação condicional
# ==========================================================================


@dataclass(frozen=True)
class Regra:
    nome: str
    origem: str                 # coluna onde procurar o valor
    alvo: str                   # coluna a formatar, ou LINHA_INTEIRA
    comparacao: str             # contem | contem_algum | exato | inicia_com | diferente
    valor: str | tuple[str, ...]
    fill: str | None = None
    cor_fonte: str | None = None
    remover_do_texto: bool = False   # tira o marcador da coluna Variantes no fim
    prioridade: int = 0              # maior = aplicada depois, pinta por cima
    ativa: bool = True

    def casa(self, texto: str) -> bool:
        texto = (texto or "").strip().upper()
        if self.comparacao == "contem":
            return str(self.valor).upper() in texto
        if self.comparacao == "contem_algum":
            return any(str(v).upper() in texto for v in self.valor)
        if self.comparacao == "exato":
            return texto == str(self.valor).upper()
        if self.comparacao == "inicia_com":
            return texto.startswith(str(self.valor).upper())
        if self.comparacao == "diferente":
            return bool(texto) and texto != str(self.valor).upper()
        raise ValueError(f"Comparação desconhecida: {self.comparacao}")


REGRAS: Sequence[Regra] = (
    # ---- coluna MODELO ----
    Regra("SMAC", "Variantes", "MODELO", "contem", "SMAC",
          fill=VERDE, remover_do_texto=True),
    Regra("8x2 / 8x4", "MODELO", "MODELO", "contem_algum", ("8X2", "8X4"),
          fill=AZUL_CLARO),

    # ---- coluna CHASSI ----
    Regra("Chifrão porca", "Variantes", "CHASSI", "contem", "CHIFRÃO PORCA", fill=AZUL),

    # ---- coluna CHASSIS HEIGHT ----
    Regra("Hidden", "Variantes", "CHASSIS HEIGHT", "contem", "HIDDEN",
          fill=AZUL_ESCURO, remover_do_texto=True),

    # ---- coluna MU ----
    # No MIX_F4 o vermelho do MU bate exatamente com "mercado != BR01" (85/85).
    Regra("Mercado externo", "MU", "MU", "diferente", "BR01", fill=VERMELHO),
    # A regra antiga (P-C na Variantes) pinta a mesma célula com a mesma cor;
    # mantida porque o marcador P-C é apagado do texto no fim e eu não consegui
    # verificar se ela é redundante com a de cima ou não.
    Regra("Para-choque", "Variantes", "MU", "contem", "P-C",
          fill=VERMELHO, remover_do_texto=True),

    # ---- coluna Variantes ----
    Regra("Gás", "Variantes", "Variantes", "contem", "GÁS", fill=AMARELO),
    Regra("Tanque A", "Variantes", "Variantes", "contem", "TANQUE A", fill=LARANJA),
    Regra("Chifrão", "Variantes", "Variantes", "contem", "CHIFRÃO", cor_fonte=VERMELHO),
    Regra("Anjo", "Variantes", "Variantes", "contem", "ANJO", cor_fonte=AZUL),

    # ---- linha inteira: prioridade alta, cobre as regras de coluna ----
    # No MIX_F4 as linhas Euro 3 saem verdes por inteiro, inclusive o MU, que
    # pela regra de mercado externo (EAK16 != BR01) seria vermelho.
    Regra("Euro 3", "Variantes", LINHA_INTEIRA, "contem", "EURO 3",
          fill=VERDE_CLARO, prioridade=100),
)

# Termos removidos da coluna Variantes depois que as regras já rodaram.
# Os marcadores das regras com remover_do_texto=True entram aqui automaticamente.
LIXO_VARIANTES = (
    "N/A", "NULL", "undefined", "nan", "None", "null",
    "Não se aplica", "SEM INFO", "VALOR PADRÃO", "---",
)


# ==========================================================================
# Escrita da planilha
# ==========================================================================

_BORDA_PADRAO = Border(*(Side(style="thin"),) * 4)
_BORDA_CODIGO = Border(left=Side(style="medium"), right=Side(style="medium"),
                       top=Side(style="thin"), bottom=Side(style="thin"))
_BORDA_CABECALHO = Border(left=Side(style="hair"), right=Side(style="hair"),
                          top=Side(style="hair"), bottom=Side(style="hair"))
_CENTRO = Alignment(horizontal="center", vertical="center")


def _fill(cor: str) -> PatternFill:
    return PatternFill(start_color=cor, end_color=cor, fill_type="solid")


def _pintar(celula, cor: str) -> None:
    celula.fill = _fill(cor)


def _cor_fonte(celula, cor: str) -> None:
    """Troca só a cor, preservando nome/tamanho/negrito já aplicados."""
    f = celula.font
    celula.font = Font(name=f.name, size=f.size, bold=f.bold, italic=f.italic, color=cor)


def _limpar_variantes(texto: object, termos_removidos: Sequence[str]) -> str:
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return ""
    remover = {t.strip().upper() for t in (*LIXO_VARIANTES, *termos_removidos)}
    partes = [p.strip() for p in re.split(r"\s*\|\s*|\s+-\s+", str(texto))]
    return " | ".join(p for p in partes if p and p.upper() not in remover)


def escrever_planilha(
    df: pd.DataFrame,
    caminho_saida: str,
    data_referencia: date | None = None,
    nome_aba: str = "MIX",
    colunas_desejadas=None,
) -> str:
    """Escreve o df já preparado no layout do MIX_F4."""
    colunas_layout = _colunas_escolhidas(colunas_desejadas)
    data_referencia = data_referencia or date.today()
    sufixo = data_referencia.strftime("%d-%m")

    wb = Workbook()
    ws = wb.active
    ws.title = nome_aba

    # ---------------- cabeçalho ----------------
    for i, coluna in enumerate(colunas_layout, start=1):
        titulo = f"{coluna.titulo} {sufixo}" if coluna.titulo == "Variantes" else coluna.titulo
        celula = ws.cell(row=1, column=i, value=titulo)
        celula.font = Font(name=FONTE_PADRAO, size=coluna.tamanho,
                           bold=(coluna.titulo == "Variantes"))
        celula.fill = _fill(CINZA_CABECALHO)
        celula.alignment = _CENTRO
        celula.border = _BORDA_CABECALHO
        ws.column_dimensions[get_column_letter(i)].width = coluna.largura
    ws.row_dimensions[1].height = ALTURA_CABECALHO

    # ---------------- dados ----------------
    termos_removidos = [
        str(r.valor) for r in REGRAS if r.ativa and r.remover_do_texto and isinstance(r.valor, str)
    ]
    indice_coluna = {c.titulo: i for i, c in enumerate(colunas_layout, start=1)}

    for n, (_, linha) in enumerate(df.iterrows()):
        r = n + 2
        ws.row_dimensions[r].height = ALTURA_LINHA

        for i, coluna in enumerate(colunas_layout, start=1):
            valor = coluna.valor(linha)
            if isinstance(valor, float) and pd.isna(valor):
                valor = None
            celula = ws.cell(row=r, column=i, value=valor)
            celula.font = Font(name=coluna.fonte, size=coluna.tamanho, bold=coluna.negrito)
            celula.alignment = _CENTRO
            celula.border = _BORDA_CODIGO if i == 1 else _BORDA_PADRAO
            celula.number_format = coluna.formato

        _aplicar_regras(ws, r, linha, indice_coluna)

        # limpeza do texto só depois que as regras já leram os marcadores
        col_var = indice_coluna["Variantes"]
        ws.cell(row=r, column=col_var).value = _limpar_variantes(
            linha.get("Variantes"), termos_removidos
        )

    # ---------------- impressão ----------------
    ultima = get_column_letter(max(len(colunas_layout), 1))
    ws.print_area = f"A1:{ultima}{ws.max_row}"
    ws.print_title_rows = "1:1"
    ws.freeze_panes = "A2"
    ws.page_margins.left = ws.page_margins.right = 0.118
    ws.page_margins.top = ws.page_margins.bottom = 0.157

    wb.save(caminho_saida)
    wb.close()
    return caminho_saida


def _aplicar_regras(ws, r: int, linha: pd.Series, indice_coluna: dict[str, int]) -> None:
    for regra in sorted(REGRAS, key=lambda x: x.prioridade):
        if not regra.ativa:
            continue
        origem = linha.get(regra.origem)
        texto = "" if origem is None or (isinstance(origem, float) and pd.isna(origem)) else str(origem)
        if not regra.casa(texto):
            continue

        if regra.alvo == LINHA_INTEIRA:
            alvos = range(1, len(colunas_layout) + 1)
        else:
            if regra.alvo not in indice_coluna:
                continue
            alvos = [indice_coluna[regra.alvo]]

        for c in alvos:
            celula = ws.cell(row=r, column=c)
            if regra.fill:
                _pintar(celula, regra.fill)
            if regra.cor_fonte:
                _cor_fonte(celula, regra.cor_fonte)


def _contar_variantes(df: pd.DataFrame) -> Counter:
    """Conta quantos veículos possuem cada variante, sem duplicar na mesma linha."""
    contagem = Counter()
    lixo = {str(valor).strip().upper() for valor in LIXO_VARIANTES}

    for texto in df.get("Variantes", pd.Series(dtype="string")).dropna():
        variantes_da_linha = {
            parte.strip().upper()
            for parte in re.split(r"\s*\|\s*|\s+-\s+", str(texto))
            if parte.strip() and parte.strip().upper() not in lixo
        }
        contagem.update(variantes_da_linha)

    return contagem


def _formatar_horario(inicio_turno: str, minutos_decorridos: float) -> str:
    inicio = datetime.strptime(inicio_turno, "%H:%M")
    minutos_inicio = inicio.hour * 60 + inicio.minute
    total = minutos_inicio + int(round(minutos_decorridos))
    dias, minutos_no_dia = divmod(total, 24 * 60)
    horas, minutos = divmod(minutos_no_dia, 60)
    sufixo = f" (+{dias} dia)" if dias == 1 else (f" (+{dias} dias)" if dias > 1 else "")
    return f"{horas:02d}:{minutos:02d}{sufixo}"


def _criar_analise_cpm(
    df: pd.DataFrame,
    takt_minutos: float,
    bloco_horas: float,
    producao: int,
    popid_corte: str,
    inicio_turno: str,
) -> dict:
    """Divide os POPIDs já tratados em blocos e conta as variantes por bloco."""
    if takt_minutos <= 0 or bloco_horas <= 0 or producao <= 0:
        raise ValueError("TAKT, Bloco e Produção devem ser maiores que zero.")

    capacidade_bloco = int((bloco_horas * 60) // takt_minutos)
    if capacidade_bloco < 1:
        raise ValueError(
            "O tempo do bloco deve comportar pelo menos um POPID conforme o TAKT."
        )

    try:
        datetime.strptime(inicio_turno, "%H:%M")
    except ValueError as erro:
        raise ValueError("O início do turno deve estar no formato HH:MM.") from erro

    popid_normalizado = _norm_popid(pd.Series([popid_corte])).iloc[0]
    popids_df = _norm_popid(df["POPID"])
    posicoes_corte = [
        posicao
        for posicao, corresponde in enumerate(
            popids_df.eq(popid_normalizado).fillna(False)
        )
        if corresponde
    ]

    if not posicoes_corte:
        raise ValueError(
            f"O POPID de corte {popid_corte} não foi encontrado "
            "nas duas planilhas MIX após o cruzamento."
        )

    indice_corte = posicoes_corte[0]
    dados_disponiveis = df.iloc[indice_corte:].reset_index(drop=True)

    if len(dados_disponiveis) < producao:
        raise ValueError(
            f"A produção informada é {producao}, mas as duas planilhas MIX "
            f"possuem somente {len(dados_disponiveis)} POPIDs a partir "
            f"do corte {popid_corte}."
        )

    dados = dados_disponiveis.head(producao).copy()
    termos_removidos = [
        str(regra.valor)
        for regra in REGRAS
        if regra.ativa
        and regra.remover_do_texto
        and isinstance(regra.valor, str)
    ]

    # Esta é a mesma limpeza aplicada à coluna Variantes na planilha formatada.
    dados["Variantes"] = dados["Variantes"].apply(
        lambda texto: _limpar_variantes(texto, termos_removidos)
    )

    quantidade_blocos = (producao + capacidade_bloco - 1) // capacidade_bloco
    linhas = []

    for indice_bloco in range(quantidade_blocos):
        inicio = indice_bloco * capacidade_bloco
        fim = min(inicio + capacidade_bloco, producao)
        trecho = dados.iloc[inicio:fim]
        contagem = _contar_variantes(trecho)

        if not contagem:
            contagem["SEM VARIANTE"] = 0

        primeiro_popid = trecho.iloc[0].get("POPID")
        ultimo_popid = trecho.iloc[-1].get("POPID")

        for variante, quantidade in sorted(
            contagem.items(), key=lambda item: (-item[1], item[0])
        ):
            linhas.append(
                {
                    "bloco": indice_bloco + 1,
                    "inicio": _formatar_horario(
                        inicio_turno,
                        inicio * takt_minutos,
                    ),
                    "fim": _formatar_horario(
                        inicio_turno,
                        fim * takt_minutos,
                    ),
                    "posicao_inicial": inicio + 1,
                    "posicao_final": fim,
                    "popids_bloco": len(trecho),
                    "primeiro_popid": primeiro_popid,
                    "ultimo_popid": ultimo_popid,
                    "variante": variante,
                    "quantidade": quantidade,
                }
            )

    return {
        "takt_minutos": takt_minutos,
        "bloco_horas": bloco_horas,
        "capacidade_bloco": capacidade_bloco,
        "producao": producao,
        "quantidade_blocos": quantidade_blocos,
        "popid_corte": popid_corte,
        "inicio_turno": inicio_turno,
        "posicao_corte": indice_corte + 1,
        "popids_disponiveis": len(dados_disponiveis),
        "linhas": linhas,
    }


def _anexar_analise_cpm(caminho_arquivo: str, analise: dict) -> None:
    """Anexa a análise ao arquivo somente após a aba MIX estar formatada."""
    wb = load_workbook(caminho_arquivo)

    if "ANALISE CPM" in wb.sheetnames:
        del wb["ANALISE CPM"]

    ws = wb.create_sheet("ANALISE CPM")
    parametros = [
        ("TAKT (minutos por POPID)", analise["takt_minutos"]),
        ("Bloco de análise (horas)", analise["bloco_horas"]),
        ("Capacidade de POPIDs por bloco", analise["capacidade_bloco"]),
        ("Produção analisada", analise["producao"]),
        ("Quantidade de blocos", analise["quantidade_blocos"]),
        ("POPID de corte", analise["popid_corte"]),
        ("Posição do corte na concatenação", analise["posicao_corte"]),
        ("Início do turno", analise["inicio_turno"]),
        ("POPIDs disponíveis após o corte", analise["popids_disponiveis"]),
    ]

    ws.append(["PARÂMETRO", "VALOR"])
    for nome, valor in parametros:
        ws.append([nome, valor])

    linha_cabecalho = ws.max_row + 2
    cabecalhos = [
        "BLOCO",
        "INÍCIO",
        "FIM",
        "POSIÇÃO INICIAL",
        "POSIÇÃO FINAL",
        "POPIDs NO BLOCO",
        "PRIMEIRO POPID",
        "ÚLTIMO POPID",
        "VARIANTE",
        "QUANTIDADE",
    ]

    for coluna, titulo in enumerate(cabecalhos, start=1):
        ws.cell(row=linha_cabecalho, column=coluna, value=titulo)

    for item in analise["linhas"]:
        ws.append(
            [
                item["bloco"],
                item["inicio"],
                item["fim"],
                item["posicao_inicial"],
                item["posicao_final"],
                item["popids_bloco"],
                item["primeiro_popid"],
                item["ultimo_popid"],
                item["variante"],
                item["quantidade"],
            ]
        )

    for numero_linha in (1, linha_cabecalho):
        for celula in ws[numero_linha]:
            celula.font = Font(name=FONTE_PADRAO, bold=True)
            celula.fill = _fill(CINZA_CABECALHO)
            celula.alignment = _CENTRO
            celula.border = _BORDA_PADRAO

    for linha in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for celula in linha:
            celula.alignment = _CENTRO
            celula.border = _BORDA_PADRAO

    larguras = [10, 10, 10, 18, 16, 18, 18, 18, 35, 15]
    for indice, largura in enumerate(larguras, start=1):
        ws.column_dimensions[get_column_letter(indice)].width = largura

    ws.freeze_panes = f"A{linha_cabecalho + 1}"
    wb.save(caminho_arquivo)
    wb.close()


# ==========================================================================
# MIX DE FIBRAS — segunda aba do workbook, sem cruzamento com a MONA
# ==========================================================================

COLUNAS_FIBRAS = (
    ("COD CHASSIS", 28.43, 40, FONTE_CODIGO_BARRAS),
    ("POP ID", 21.00, 28, FONTE_PADRAO),
    ("CHASSI", 8.00, 10, FONTE_PADRAO),
    ("MODELO", 12.43, 8, FONTE_PADRAO),
    ("MU", 5.00, 8, FONTE_PADRAO),
    ("COR", 6.57, 8, FONTE_PADRAO),
    ("DEF", 4.86, 8, FONTE_PADRAO),
    ("LATERAL", 8.14, 8, FONTE_PADRAO),
    ("GRELHA", 13.86, 6, FONTE_PADRAO),
)


def _normalizar_nome_coluna(nome: object) -> str:
    return " ".join(str(nome or "").strip().upper().split())


def _preparar_mix_fibras(caminho_mix: str) -> pd.DataFrame:
    # O gerenciador de contexto é indispensável no Windows: sem ele, o pandas
    # mantém o arquivo aberto e a TemporaryDirectory não consegue apagá-lo.
    with pd.ExcelFile(caminho_mix) as arquivo:
        if len(arquivo.sheet_names) < 2:
            raise ValueError(
                "O arquivo da 1ª planilha MIX precisa possuir uma segunda aba "
                "com o MIX de fibras."
            )
        df_origem = pd.read_excel(arquivo, sheet_name=1)
    df_origem.columns = [_normalizar_nome_coluna(c) for c in df_origem.columns]

    aliases = {
        "POP ID": ("POP ID", "POPID"),
        "CHASSI": ("CHASSI", "COD CHASSIS"),
        "MODELO": ("MODELO",),
        "MU": ("MU",),
        "COR": ("COR",),
        "DEF": ("DEF",),
        "LATERAL": ("LATERAL",),
        "GRELHA": ("GRELHA",),
    }

    faltantes = [
        destino
        for destino, opcoes in aliases.items()
        if not any(opcao in df_origem.columns for opcao in opcoes)
    ]
    if faltantes:
        raise ValueError(
            "A segunda aba (MIX de fibras) não contém as colunas: "
            + ", ".join(faltantes)
            + "."
        )

    dados = {}
    for destino, opcoes in aliases.items():
        origem = next(opcao for opcao in opcoes if opcao in df_origem.columns)
        dados[destino] = df_origem[origem]

    df = pd.DataFrame(dados)
    mascara_hide = df.astype(str).apply(
        lambda linha: linha.str.contains("Hide", case=False, na=False).any(),
        axis=1,
    )
    df = df[~mascara_hide].copy()
    df = df.dropna(subset=["POP ID", "CHASSI"], how="all").reset_index(drop=True)
    return df


def _codigo_barras_chassi(valor: object) -> str:
    if pd.isna(valor):
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    texto = str(valor).strip()
    if texto.startswith("*") and texto.endswith("*"):
        return texto
    return f"*{texto}*"


def _anexar_mix_fibras(
    caminho_mix_origem: str,
    caminho_arquivo_formatado: str,
    data_referencia: date | None = None,
) -> None:
    """Formata diretamente a segunda aba, sem usar qualquer dado da MONA."""
    df = _preparar_mix_fibras(caminho_mix_origem)
    data_referencia = data_referencia or date.today()
    nome_aba = f"MIX DE FIBRAS {data_referencia.strftime('%d-%m-%Y')}"

    wb = load_workbook(caminho_arquivo_formatado)
    if nome_aba in wb.sheetnames:
        del wb[nome_aba]
    ws = wb.create_sheet(nome_aba)

    for coluna, (titulo, largura, tamanho, fonte) in enumerate(
        COLUNAS_FIBRAS,
        start=1,
    ):
        celula = ws.cell(row=1, column=coluna, value=titulo)
        celula.font = Font(
            name=fonte,
            size=tamanho if titulo != "COD CHASSIS" else 25,
            bold=(titulo != "COD CHASSIS"),
            color="FFFFFFFF" if titulo == "GRELHA" else "FF000000",
        )
        celula.fill = _fill(AZUL if titulo == "GRELHA" else CINZA_CABECALHO)
        celula.alignment = _CENTRO
        celula.border = _BORDA_CODIGO if coluna == 1 else _BORDA_PADRAO
        ws.column_dimensions[get_column_letter(coluna)].width = largura

    ws.row_dimensions[1].height = 31.5

    for indice, linha in df.iterrows():
        numero_linha = indice + 2
        ws.row_dimensions[numero_linha].height = 39.95
        valores = [
            _codigo_barras_chassi(linha.get("CHASSI")),
            linha.get("POP ID"),
            linha.get("CHASSI"),
            linha.get("MODELO"),
            linha.get("MU"),
            linha.get("COR"),
            linha.get("DEF"),
            linha.get("LATERAL"),
            linha.get("GRELHA"),
        ]

        for coluna, valor in enumerate(valores, start=1):
            if isinstance(valor, float) and pd.isna(valor):
                valor = None
            celula = ws.cell(row=numero_linha, column=coluna, value=valor)
            celula.alignment = _CENTRO
            celula.border = _BORDA_CODIGO if coluna == 1 else _BORDA_PADRAO

            if coluna == 1:
                celula.font = Font(
                    name=FONTE_CODIGO_BARRAS,
                    size=40,
                    bold=False,
                )
            elif coluna == 2:
                celula.font = Font(name=FONTE_PADRAO, size=28, bold=True)
                celula.number_format = "0"
            else:
                tamanho = 6 if coluna == 9 else (10 if coluna == 3 else 8)
                celula.font = Font(
                    name=FONTE_PADRAO,
                    size=tamanho,
                    bold=(coluna == 9),
                )

        valor_def = str(linha.get("DEF") or "").strip().upper()
        if valor_def == "TETO":
            ws.cell(numero_linha, 7).fill = _fill(AMARELO)

        valor_grelha = str(linha.get("GRELHA") or "").strip().upper()
        estilos_grelha = {
            "KD": (VERDE_CLARO, "FF000000"),
            "CLIMATIZADOR": (AMARELO, "FF000000"),
            "PÓS TRANSPORTE": (VERDE, "FFFFFFFF"),
            "TVL": (AZUL, "FFFFFFFF"),
            "REALOCADO": ("FFDAEEF3", VERDE),
        }
        if valor_grelha in estilos_grelha:
            cor_fundo, cor_fonte = estilos_grelha[valor_grelha]
            celula_grelha = ws.cell(numero_linha, 9)
            celula_grelha.fill = _fill(cor_fundo)
            celula_grelha.font = Font(
                name=FONTE_PADRAO,
                size=6,
                bold=True,
                color=cor_fonte,
            )

    ultima_linha = max(ws.max_row, 1)
    ws.print_area = f"A1:I{ultima_linha}"
    ws.print_title_rows = "1:1"
    ws.freeze_panes = "A2"
    ws.page_setup.orientation = "portrait"
    ws.page_margins.left = ws.page_margins.right = 0.118
    ws.page_margins.top = ws.page_margins.bottom = 0.157
    wb.save(caminho_arquivo_formatado)
    wb.close()


# ==========================================================================
# Pipeline de preparação (mesma lógica de Manual_E_Automatica4, organizada)
# ==========================================================================

COLUNAS_DESCARTADAS = (
    "SCOB", "SEQ", "MIX", "SEQ.", "REPLAN", "ENTRADA LINHA REAL", "EOV", "COMPRIM",
    "OPCIONAIS", "RHD", "DEF", "LATERAL", "SORDER", "FFU", "TANQUE", "LNG",
    "BEHIND CAB", "DW", "FOU", "RETARDER", "ABS", "HIGH LINE", "CAB MOD",
    "COD POPID", "COD CHASSIS", "AUX CAB COOLER", "LEVEL", "SESSAM7",
    "POS. TANQUE COMB", "POSIÇÃO BATERIA", "SCR", "RPX", "BARRA 2°E DIANT",
    "BARRA TRASEIRA", "ETIQUETA CO²", "SCR HIDDEN", "PARACHOQUE TRASEIRO",
    "ETIQUETA ECE",
)

COLUNAS_VARIANTE = ("GÁS", "CAMERA DE RÉ", "SMAC")

# Renomes aplicados aos valores vindos da Mona antes de entrar em Variantes.
RENOMES_MONA = {
    "PARA-LAMA": {"COM": "PL"},
    "TANQUE A": {"COM": "TANQUE A"},
    "PARA-CHOQUE": {"COM": "P-C"},
    "CHIFRAO": {"SIMPLES": "CHIFRÃO SIMPLES", "PORCA": "CHIFRÃO PORCA"},
    "ANJO CHAPA": {"SIMPLES": "ANJO SIMPLES", "MAIOR": "ANJO MAIOR"},
    "P.L ONE PIECE": {"P.L ONE PIECE": "1PC"},
}


def _norm_popid(serie: pd.Series) -> pd.Series:
    s = serie.astype("string").fillna("").str.strip()
    s = s.str.replace(r"[,\s]", "", regex=True)
    s = s.str.replace(r"(?<=\d)\.0+$", "", regex=True)
    return s.replace("", pd.NA)


def _texto_item_mona(nome_item: str, valor: object) -> str | None:
    """Converte uma célula MONA no texto que será enviado para Variantes."""
    if pd.isna(valor):
        return None

    nome = " ".join(str(nome_item).strip().upper().split())
    texto = str(valor).strip()
    if not texto or texto.lower() == "nan":
        return None

    texto_upper = texto.upper()
    renome = RENOMES_MONA.get(nome, {}).get(texto_upper)
    if renome:
        return renome

    # Quando a célula apenas marca a presença do item, usa o nome descritivo
    # da linha 2, e não o texto genérico COM/SIM/X.
    if texto_upper in {"COM", "SIM", "X", "OK", "TRUE"}:
        return str(nome_item).strip()

    # Nos demais casos, o próprio valor contém a variante desejada:
    # ONE PIECE, SMACK AR, HIDRÁULICA, MCC etc.
    return texto


def _ler_e_agrupar_mona(
    caminho_mona: str,
    padroes_relevantes: Sequence[tuple[str, str]] | None = None,
    padroes_excluidos: Sequence[str] | None = None,
) -> pd.DataFrame:
    """
    Lê o formato real da MONA:
      linha 1 = cabeçalhos técnicos (POPID, POS_10, POS_20...)
      linha 2 = nomes dos itens (SMACK, PARA LAMA, DIREÇÃO...)
      linha 3 em diante = dados.

    Dois jeitos de filtrar as colunas POS_ (use um OU outro, nunca os dois):

    padroes_relevantes: LISTA PERMITIDA — uma lista de (regex, texto_final).
    Só entram no cruzamento as colunas cujo nome do item bata com algum
    desses regex; o texto gravado é sempre o `texto_final` combinado, não o
    valor bruto da célula. Usado quando a MONA tem poucas variantes NOVAS
    (que ainda não existem no mix geral) — ex.: Filtro MIX (FA2) e Análise,
    via cpm_core.tags_mona_ativas().

    padroes_excluidos: LISTA DE EXCLUSÃO — uma lista simples de regex. Entram
    no cruzamento TODAS as colunas POS_, EXCETO as que baterem com algum
    desses regex — cada uma mantém seu próprio texto (via
    `_texto_item_mona`), não um texto combinado. Usado quando a maioria das
    variantes da MONA é nova e só uma ou duas já existem no mix geral — ex.:
    FA1.2, onde "GÁS" e "6X6/4X4" já vêm de outras colunas do mix e o resto
    (8X2 EBS, TQ K1, ABS, SCR V8...) precisa entrar em Variantes.

    Sem nenhum dos dois (uso da Função 4), processa todas as colunas POS_,
    como sempre.
    """
    bruto = pd.read_excel(caminho_mona, header=None, dtype=object)

    if bruto.shape[0] < 3:
        raise ValueError("A planilha MONA não possui linhas de dados.")

    cabecalho_tecnico = [
        "" if pd.isna(valor) else str(valor).strip().upper()
        for valor in bruto.iloc[0].tolist()
    ]

    try:
        indice_popid = cabecalho_tecnico.index("POPID")
    except ValueError as erro:
        raise ValueError(
            "A planilha MONA não contém a coluna POPID na primeira linha."
        ) from erro

    colunas_itens = []
    for indice, cabecalho in enumerate(cabecalho_tecnico):
        nome_item = bruto.iloc[1, indice]
        if not (
            cabecalho.startswith("POS_")
            and pd.notna(nome_item)
            and str(nome_item).strip()
        ):
            continue
        nome_item = str(nome_item).strip()
        nome_upper = nome_item.upper()

        texto_final = None
        if padroes_relevantes is not None:
            for padrao, texto in padroes_relevantes:
                if re.search(padrao, nome_upper, flags=re.IGNORECASE):
                    texto_final = texto
                    break
            if texto_final is None:
                continue  # coluna ja coberta pelo mix geral, ignora
        elif padroes_excluidos is not None:
            excluida = any(
                re.search(padrao, nome_upper, flags=re.IGNORECASE)
                for padrao in padroes_excluidos
            )
            if excluida:
                continue  # coluna ja coberta pelo mix geral, ignora

        colunas_itens.append((indice, nome_item, texto_final))

    if not colunas_itens:
        if padroes_relevantes is not None:
            textos = ", ".join(sorted({t for _, t in padroes_relevantes}))
            raise ValueError(
                f"Nenhuma das colunas relevantes ({textos}) foi encontrada "
                "nesta planilha MONA."
            )
        raise ValueError(
            "Nenhum item MONA foi encontrado na segunda linha das colunas POS_."
        )

    registros = []
    for numero_linha in range(2, bruto.shape[0]):
        popid = bruto.iat[numero_linha, indice_popid]
        if pd.isna(popid) or not str(popid).strip():
            continue

        variantes = []
        for indice_coluna, nome_item, texto_final in colunas_itens:
            valor = bruto.iat[numero_linha, indice_coluna]
            if texto_final is not None:
                # Modo filtrado: o texto gravado e sempre o canonico da tag,
                # nao o valor bruto da celula — so importa se ela indica
                # presenca do item.
                if pd.isna(valor):
                    continue
                texto_bruto = str(valor).strip()
                if not texto_bruto or texto_bruto.lower() == "nan":
                    continue
                texto = texto_final
            else:
                texto = _texto_item_mona(nome_item, valor)
            if texto and texto not in variantes:
                variantes.append(texto)

        registros.append(
            {
                "POPID": popid,
                "linha_concat": " | ".join(variantes) if variantes else pd.NA,
            }
        )

    df_mona = pd.DataFrame(registros)
    if df_mona.empty:
        raise ValueError("A planilha MONA não contém POPIDs para cruzamento.")

    df_mona["POPID_norm"] = _norm_popid(df_mona["POPID"])

    def juntar_sem_duplicar(valores) -> str:
        encontrados = []
        for valor in valores:
            if pd.isna(valor):
                continue
            # O separador entre itens e sempre " | " (com espacos); dividir
            # por um "|" solto quebraria tags como "658|660" ou "649|650",
            # que usam a barra como parte do proprio texto.
            for parte in str(valor).split(" | "):
                parte = parte.strip()
                if parte and parte not in encontrados:
                    encontrados.append(parte)
        return " | ".join(encontrados)

    return (
        df_mona.dropna(subset=["POPID_norm", "linha_concat"])
        .groupby("POPID_norm", sort=False)["linha_concat"]
        .apply(juntar_sem_duplicar)
        .reset_index()
    )


def _detectar_formato_mona(caminho_mona: str) -> str:
    """
    "pos" quando a planilha segue o formato técnico da MONA usado na Função 4
    (POPID na 1a linha, colunas POS_ com o nome do item na 2a linha);
    "lista" quando é só uma lista de POPIDs com cabeçalho normal — caso de
    listas exportadas do MONA fora do padrão da Função 4, como uma relação
    de chassis de uma restrição pontual.
    """
    bruto = pd.read_excel(caminho_mona, header=None, nrows=2, dtype=object)
    if bruto.shape[0] < 2:
        return "lista"
    cabecalho = ["" if pd.isna(v) else str(v).strip().upper()
                for v in bruto.iloc[0].tolist()]
    tem_pos = any(c.startswith("POS_") for c in cabecalho)
    return "pos" if (tem_pos and "POPID" in cabecalho) else "lista"


def _ler_lista_popid_simples(caminho: str) -> set[str]:
    """
    Lê uma planilha qualquer que tenha uma coluna POPID (cabeçalho normal,
    na primeira linha) e devolve o conjunto de POPIDs normalizados nela —
    sem exigir o formato técnico de itens em colunas POS_.
    """
    df = pd.read_excel(caminho)
    coluna_popid = next(
        (c for c in df.columns if str(c).strip().upper() == "POPID"), None)
    if coluna_popid is None:
        raise ValueError("A planilha não contém uma coluna POPID.")
    return set(_norm_popid(df[coluna_popid]).dropna())


def cruzar_com_mona(
    df: pd.DataFrame,
    caminho_mona: str,
    coluna_saida: str = "MONA",
    texto_lista_simples: str | None = None,
    itens_excluidos: Sequence[str] | None = None,
    setor: str | None = None,
) -> pd.DataFrame:
    """
    Cruza qualquer DataFrame que já tenha a coluna POPID com a planilha MONA.

    Reaproveita a mesma leitura/normalização da Função 4 (`_ler_e_agrupar_mona`
    + `_norm_popid`), mas devolve o cruzamento pronto para qualquer outra
    função anexar como preferir — em `coluna_saida` (nova coluna) ou juntando
    o texto a uma coluna já existente (ex.: Variantes, OPCIONAIS), como fazem
    o FA1.2 e o FA2.

    Quando a planilha não segue o formato técnico da MONA (só uma lista de
    POPIDs, sem as colunas POS_), `texto_lista_simples` é o texto fixo
    escrito para cada POPID encontrado nela — ex.: "650/649" numa lista de
    chassis de uma restrição pontual.

    Quando a planilha SEGUE o formato POS_ (POPID + itens), há dois modos —
    escolha um dos dois, nunca os dois juntos:

    - Padrão (itens_excluidos=None): só entram no cruzamento as colunas que
      batem com alguma tag ativa (editável em /tags-mona) — as poucas
      variantes novas que ainda não existem no mix geral da fábrica. O resto
      é ignorado. Se além disso `texto_lista_simples` for informado, ele
      também vira um padrão avulso: qualquer coluna cujo nome do item
      contenha esse texto entra no cruzamento também, marcada com esse mesmo
      texto — pra cobrir um valor novo que ainda não virou tag fixa.

    - `itens_excluidos` informado: o inverso — entram TODAS as colunas,
      exceto as que baterem com algum desses regex. Cada coluna mantém seu
      próprio texto (não um texto combinado). Use quando a maioria das
      variantes da MONA é nova e só uma ou duas já existem no mix geral —
      caso do FA1.2 (ver ITENS_MONA_FA12_JA_NO_MIX em mix_processor_1_2.py).

    setor: só importa no modo padrão (`itens_excluidos=None`) — decide de
      qual setor vêm as tags ativas (`cpm_core.tags_mona_ativas`). Sem isso,
      cai no padrão da FA2, igual sempre foi.
    """
    if "POPID" not in df.columns:
        raise ValueError(
            "A planilha não contém a coluna POPID para cruzar com a MONA.")

    df = df.copy()
    df["POPID_norm"] = _norm_popid(df["POPID"])

    if _detectar_formato_mona(caminho_mona) == "pos" and itens_excluidos is not None:
        agrupado = _ler_e_agrupar_mona(
            caminho_mona, padroes_excluidos=itens_excluidos
        ).rename(columns={"linha_concat": coluna_saida})
        df = df.merge(agrupado, on="POPID_norm", how="left", sort=False)
    elif _detectar_formato_mona(caminho_mona) == "pos":
        from cpm_core import tags_mona_ativas
        padroes_relevantes = [
            (padrao, tag.texto)
            for tag in tags_mona_ativas(setor) for padrao in tag.padroes
        ]
        if texto_lista_simples:
            padroes_relevantes.append(
                (re.escape(texto_lista_simples), texto_lista_simples))
        agrupado = _ler_e_agrupar_mona(
            caminho_mona, padroes_relevantes=padroes_relevantes
        ).rename(columns={"linha_concat": coluna_saida})
        df = df.merge(agrupado, on="POPID_norm", how="left", sort=False)
    else:
        if not texto_lista_simples:
            raise ValueError(
                "Esta planilha MONA não está no formato da Função 4 (POPID "
                "+ colunas POS_ com os itens) — é só uma lista de POPIDs. "
                "Informe o texto a escrever para os POPIDs encontrados nela.")
        popids_lista = _ler_lista_popid_simples(caminho_mona)
        df[coluna_saida] = df["POPID_norm"].apply(
            lambda v: texto_lista_simples if v in popids_lista else pd.NA)

    return df.drop(columns=["POPID_norm"], errors="ignore")


def preparar_dataframe(
    caminho_mix: str,
    caminho_mona: str,
    caminho_mix_adicional: str | None = None,
) -> pd.DataFrame:
    df_mix = pd.read_excel(caminho_mix)

    if "POPID" not in df_mix.columns:
        raise ValueError("A planilha MIX não contém a coluna POPID.")

    partes = [df_mix]
    if caminho_mix_adicional:
        df_mix_adicional = pd.read_excel(caminho_mix_adicional)
        if "POPID" not in df_mix_adicional.columns:
            raise ValueError("A segunda planilha MIX não contém a coluna POPID.")
        partes.append(df_mix_adicional)

    # Na análise CPM, as linhas dos dois arquivos MIX são concatenadas antes
    # do cruzamento com a MONA.
    df = pd.concat(partes, ignore_index=True, sort=False)

    # linhas marcadas como "Hide" na origem
    mascara_hide = df.astype(str).apply(
        lambda linha: linha.str.contains("Hide", case=False, na=False).any(), axis=1
    )
    df = df[~mascara_hide].copy()

    # GÁS + CAMERA DE RÉ + SMAC → Variantes
    presentes = [c for c in COLUNAS_VARIANTE if c in df.columns]
    if presentes:
        df["Variantes"] = (
            df[presentes]
            .astype(str)
            .apply(lambda l: " - ".join(v.strip() for v in l if v.strip() and v.strip().lower() != "nan"), axis=1)
        )
    else:
        df["Variantes"] = ""

    # MONA: somente colunas POS_ com nome de item na segunda linha entram
    # em Variantes. CUOBJ, PLAN_START, TYPEDES e demais metadados são ignorados.
    df["POPID_norm"] = _norm_popid(df["POPID"])
    agrupado = _ler_e_agrupar_mona(caminho_mona)
    df = df.merge(agrupado, on="POPID_norm", how="left", sort=False)

    def juntar_variantes(linha: pd.Series) -> str:
        partes = []
        for valor in (linha.get("Variantes"), linha.get("linha_concat")):
            if pd.notna(valor) and str(valor).strip() and str(valor).strip().lower() != "nan":
                partes.append(str(valor).strip())
        return " | ".join(partes)

    df["Variantes"] = df.apply(juntar_variantes, axis=1)
    df = df.drop(columns=["POPID_norm", "linha_concat"], errors="ignore")

    return df.drop(columns=[*COLUNAS_DESCARTADAS, *COLUNAS_VARIANTE], errors="ignore")


# ==========================================================================
# Entrada compatível com o código atual
# ==========================================================================


def formatar_excel_mix_logic_funcao_quatro(
    caminho_mix,
    caminho_mona,
    caminho_arquivo_novo,
    colunas_desejadas=None,
    caminho_mix_adicional=None,
    takt_minutos=None,
    bloco_horas=None,
    producao=None,
    popid_corte=None,
    inicio_turno=None,
):
    """Formata o MIX e, opcionalmente, anexa a análise CPM dos dois MIX."""
    try:
        # A aba MIX formatada usa somente o primeiro arquivo.
        df_formatado = preparar_dataframe(caminho_mix, caminho_mona)
        escrever_planilha(df_formatado, caminho_arquivo_novo,
                          colunas_desejadas=colunas_desejadas)

        # A segunda aba do primeiro workbook é o MIX de fibras. Ela é
        # formatada diretamente e não participa do cruzamento com a MONA.
        anexar_mix_fibras(caminho_mix, caminho_arquivo_novo)

        if caminho_mix_adicional:
            # 1. Concatena os MIX na ordem correta.
            # 2. Cruza o conjunto com a MONA.
            # 3. Aplica o mesmo tratamento usado na formatação.
            # 4. Divide a produção em blocos e conta as variantes.
            df_analise_cpm = preparar_dataframe(
                caminho_mix,
                caminho_mona,
                caminho_mix_adicional,
            )
            analise_cpm = _criar_analise_cpm(
                df_analise_cpm,
                float(takt_minutos),
                float(bloco_horas),
                int(producao),
                str(popid_corte),
                str(inicio_turno),
            )
            _anexar_analise_cpm(caminho_arquivo_novo, analise_cpm)

        return True, "Arquivo formatado e salvo com sucesso!"
    except Exception as e:
        return False, f"Erro ao formatar o arquivo: {e}"


# ==========================================================================
# PENDENTE — coisas do MIX_F4.xlsx que não saíram de nenhuma regra
# ==========================================================================
#
# 1. Linha 58 inteira (A..H) pintada de azul 0070C0, com a coluna I amarela.
#    Nenhuma coluna dessa linha a distingue das outras. Parece marcação manual.
#
# 2. CHASSI azul nas linhas 56 e 58. As outras 9 células de CHASSI azuis são
#    exatamente as que têm CHIFRÃO PORCA na Variantes; essas duas não têm.
#

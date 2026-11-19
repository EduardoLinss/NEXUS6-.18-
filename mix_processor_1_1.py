# -*- coding: utf-8 -*-
"""
mix_processor_1_1.py - MIX da função 1.1: MIX principal + MONA.

Independente da FA1.2: leitura, colunas, regras de VARIANTES e layout são
todos daqui. O que é da 1.1:

  * COLUNAS_PADRAO   — colunas da planilha final, NESTA ordem
  * formatacao: VARIANTES preenchida pinta a celula; 8X2/8X4 a linha
  * FONTES_VARIANTES — o que do MIX principal entra em VARIANTES
  * a MONA (obrigatória): tudo o que ela registra para o POPID entra em
    VARIANTES depois do que veio do MIX, separado por " | "

Linhas ocultas do MIX principal ficam de fora (mesmo critério das outras
funções: linha oculta não faz parte da sequência do dia).
"""
from __future__ import annotations

import os
from dataclasses import dataclass

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from cpm_core import carregar_excel_sem_linhas_ocultas

COLUNA_VARIANTES = "VARIANTES"

# Ordem da planilha final.
COLUNAS_PADRAO = ("SEQ", COLUNA_VARIANTES, "POPID", "MODELO", "MU", "RETARDER")

# A coluna SEQ da planilha final e a sequencia de producao: no MIX principal
# ela se chama "SEQ." (com ponto) — a mesma que a Analise usa. O MIX tambem tem
# uma coluna "SEQ" (sem ponto) com outra numeracao, que NAO e a da linha.
# Sem "SEQ." no arquivo, cai para "SEQ".
ORIGEM_SEQ = ("SEQ.", "SEQ")

# Colunas que o MIX principal PRECISA ter. As outras (RETARDER), se faltarem
# num dia, saem vazias em vez de impedir a planilha.
COLUNAS_OBRIGATORIAS = ("SEQ", "POPID", "MODELO", "MU")

LARGURAS = {"SEQ": 6, COLUNA_VARIANTES: 45, "POPID": 10, "MODELO": 24, "MU": 8,
            "RETARDER": 12}
FONTE = "Arial"

# Formatacao:
#   * VARIANTES preenchida       -> so a celula de VARIANTES fica preta
#   * MODELO com 8X2 ou 8X4      -> a linha inteira fica preta
# Fundo preto sempre com fonte branca, senao o texto some.
PRETO = "FF000000"
MODELOS_LINHA_INTEIRA = ("8X2", "8X4")


@dataclass(frozen=True)
class Fonte:
    """
    Um valor do MIX principal que vai para a coluna VARIANTES.

    tipo:
      "valor"     — se a coluna estiver preenchida, entra o valor da celula
      "contem"    — se a celula contiver `valor`, entra `rotulo` (ou `valor`)
      "maior_que" — se a celula for um numero maior que `valor`, entra `rotulo`
    """
    coluna: str
    tipo: str = "valor"
    valor: object = None
    rotulo: str = ""


# Ordem = ordem em que aparecem em VARIANTES, separados por " - ".
# Para filtrar um item novo, acrescente uma linha aqui. Os exemplos:
#   Fonte("CAMERA DE RÉ")                                  -> valor da celula
#   Fonte("OPCIONAIS", "contem", "ABAFADOR")               -> "ABAFADOR"
#   Fonte("COMPRIM", "maior_que", 8100, rotulo="LONGO")    -> "LONGO"
FONTES_VARIANTES: tuple[Fonte, ...] = (
    Fonte("OPCIONAIS", "contem", "WTA"),
    Fonte("RHD"),
    Fonte("LNG"),
    Fonte("AUX CAB COOLER"),
    Fonte("DUPLO CIRCUITO"),
    Fonte("COOLING COIL"),
)

# Itens da MONA que NAO entram em VARIANTES (regex sobre o nome do item, linha
# 2 da MONA). Vazio: tudo o que a MONA registra entra.
ITENS_MONA_EXCLUIDOS: tuple[str, ...] = ()


def _texto(valor) -> str:
    if valor is None or (not isinstance(valor, str) and pd.isna(valor)):
        return ""
    return str(valor).strip()


def _chave(nome) -> str:
    return " ".join(str(nome).strip().upper().split())


def variantes_da_linha(linha, fontes=FONTES_VARIANTES) -> list[str]:
    """Aplica FONTES_VARIANTES a uma linha do MIX principal."""
    saida = []
    for fonte in fontes:
        if fonte.coluna not in linha.index:
            continue
        texto = _texto(linha[fonte.coluna])
        if not texto:
            continue
        if fonte.tipo == "valor":
            saida.append(texto)
        elif fonte.tipo == "contem":
            if str(fonte.valor).upper() in texto.upper():
                saida.append(fonte.rotulo or str(fonte.valor))
        elif fonte.tipo == "maior_que":
            try:
                if float(texto.replace(",", ".")) > float(fonte.valor):
                    saida.append(fonte.rotulo or texto)
            except ValueError:
                pass
        else:
            raise ValueError(f"Fonte de variante com tipo desconhecido: {fonte.tipo}")
    return saida


def _colunas_escolhidas(colunas_desejadas) -> list[str]:
    """Recorte pedido na tela, sempre na ordem de COLUNAS_PADRAO."""
    if not colunas_desejadas:
        return list(COLUNAS_PADRAO)
    pedidas = {_chave(c) for c in colunas_desejadas}
    escolhidas = [c for c in COLUNAS_PADRAO if _chave(c) in pedidas]
    return escolhidas or list(COLUNAS_PADRAO)


def _numero_se_inteiro(valor):
    """POPID/SEQ voltam como número (o Excel não reclama de 'número como texto')."""
    texto = _texto(valor)
    if texto.endswith(".0"):
        texto = texto[:-2]
    return int(texto) if texto.isdigit() else (texto or None)


def montar_dataframe(caminho_mix: str, caminho_mona: str,
                     texto_mona: str | None = None) -> pd.DataFrame:
    df = carregar_excel_sem_linhas_ocultas(caminho_mix)
    if df is None or df.empty:
        raise ValueError("Não foi possível ler linhas da planilha MIX principal.")
    if "POPID" not in df.columns:
        raise ValueError("A planilha MIX principal não tem a coluna POPID.")

    df = df[df["POPID"].map(_texto) != ""].copy()
    origem = next((c for c in ORIGEM_SEQ if c in df.columns), None)
    if origem is not None:
        df["SEQ"] = df[origem]
    df[COLUNA_VARIANTES] = [" - ".join(variantes_da_linha(linha))
                            for _, linha in df.iterrows()]

    from mix_processor_4 import cruzar_com_mona
    df = cruzar_com_mona(df, caminho_mona, coluna_saida="_MONA",
                         texto_lista_simples=texto_mona or None,
                         itens_excluidos=ITENS_MONA_EXCLUIDOS)
    df[COLUNA_VARIANTES] = [
        " | ".join(p for p in (_texto(v), _texto(m)) if p)
        for v, m in zip(df[COLUNA_VARIANTES], df["_MONA"])
    ]
    return df.drop(columns=["_MONA"])


def pintar_linha_inteira(linha) -> bool:
    """MODELO com 8X2 ou 8X4 pinta a linha toda de preto."""
    modelo = _texto(linha.get("MODELO")).upper()
    return any(m in modelo for m in MODELOS_LINHA_INTEIRA)


def escrever(df: pd.DataFrame, colunas: list[str], caminho_saida: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "MIX 1.1"

    fino = Side(style="thin")
    borda = Border(left=fino, right=fino, top=fino, bottom=fino)
    centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cinza = PatternFill("solid", start_color="FFC0C0C0", end_color="FFC0C0C0")

    for c, nome in enumerate(colunas, start=1):
        cel = ws.cell(row=1, column=c, value=nome)
        cel.font = Font(name=FONTE, bold=True, size=10)
        cel.fill, cel.border, cel.alignment = cinza, borda, centro
        ws.column_dimensions[cel.column_letter].width = LARGURAS.get(nome, 14)

    preto = PatternFill("solid", start_color=PRETO, end_color=PRETO)
    fonte_normal = Font(name=FONTE, bold=True, size=10)
    fonte_branca = Font(name=FONTE, bold=True, size=10, color="FFFFFFFF")

    for r, (_, linha) in enumerate(df.iterrows(), start=2):
        linha_inteira = pintar_linha_inteira(linha)
        variantes_preenchida = bool(_texto(linha.get(COLUNA_VARIANTES)))
        for c, nome in enumerate(colunas, start=1):
            valor = linha[nome] if nome in linha.index else None
            if nome in ("SEQ", "POPID"):
                valor = _numero_se_inteiro(valor)
            else:
                valor = _texto(valor) or None
            cel = ws.cell(row=r, column=c, value=valor)
            cel.font = fonte_normal
            cel.border, cel.alignment = borda, centro
            if linha_inteira or (nome == COLUNA_VARIANTES and variantes_preenchida):
                cel.fill, cel.font = preto, fonte_branca

    ws.freeze_panes = "A2"
    wb.save(caminho_saida)


def formatar_excel_mix_logic_funcao_um_um(caminho_mix, caminho_mona,
                                          texto_mona, caminho_arquivo_novo,
                                          colunas_desejadas=None):
    """
    Contrato do `processadores.CATALOGO`: arquivos (MIX, MONA), depois os
    parâmetros (texto para MONA em lista simples, pode ser None) e por último
    o caminho de saída. Devolve (ok, mensagem).
    """
    if not caminho_mona or not os.path.isfile(caminho_mona):
        return False, "A função 1.1 precisa da planilha MONA."
    try:
        df = montar_dataframe(caminho_mix, caminho_mona, texto_mona)
        colunas = _colunas_escolhidas(colunas_desejadas)
        faltando = [c for c in colunas
                    if c in COLUNAS_OBRIGATORIAS and c not in df.columns]
        if faltando:
            return False, ("A planilha MIX principal não tem a(s) coluna(s): "
                           + ", ".join(faltando))
        escrever(df, colunas, caminho_arquivo_novo)
    except Exception as erro:  # noqa: BLE001
        return False, f"Erro ao gerar o MIX 1.1: {erro}"
    return True, f"MIX 1.1 gerado com sucesso: {len(df)} veículo(s)."

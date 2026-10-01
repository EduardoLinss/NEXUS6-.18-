"""
processadores.py - Registro das funcoes que transformam planilha em planilha.

Cada processador se descreve: titulo, quais arquivos pede, quais parametros
precisa e qual funcao chamar. Uma tela unica (`processar.html`) e montada a
partir dessa descricao, entao acrescentar uma funcao e acrescentar uma entrada
aqui — nao criar rota nem template.

As funcoes seguem o mesmo contrato: recebem os caminhos de entrada, depois os
parametros, e por ultimo o caminho de saida; devolvem (ok, mensagem).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable

from filtro_mix import COLUNAS_DESEJADAS_PADRAO, formatar_excel_mix_logic
from mix_processor_1_2 import COLUNAS_PADRAO as COLUNAS_FA12
from mix_processor_1_2 import formatar_excel_mix_logic_funcao_um_dois
from mix_processor_4 import COLUNAS_PADRAO as COLUNAS_F4
from mix_processor_4 import formatar_excel_mix_logic_funcao_quatro
from mix_processor_fibras import COLUNAS_PADRAO as COLUNAS_FIBRAS
from mix_processor_fibras import formatar_excel_mix_fibras_logic

EXTENSOES = (".xlsx", ".xlsm", ".xls")


@dataclass(frozen=True)
class CampoArquivo:
    nome: str
    rotulo: str
    opcional: bool = False
    ajuda: str = ""
    # A lista de colunas configurada na tela e aplicada a este arquivo antes
    # de a funcao rodar. A MONA, por exemplo, fica de fora: tem formato proprio.
    filtra_colunas: bool = True


@dataclass(frozen=True)
class CampoValor:
    nome: str
    rotulo: str
    tipo: str = "text"          # text | number | time
    placeholder: str = ""
    ajuda: str = ""
    padrao: str = ""
    passo: str = ""
    minimo: str = ""
    # Quando preenchido, o campo so e exigido se este arquivo tiver sido
    # enviado. E o caso dos parametros de CPM da Funcao 4.
    exigido_com: str = ""


@dataclass(frozen=True)
class Processador:
    id: str
    titulo: str
    descricao: str
    funcao: Callable
    prefixo_saida: str
    arquivos: tuple[CampoArquivo, ...]
    valores: tuple[CampoValor, ...] = ()
    observacao: str = ""
    # Colunas que a funcao sabe produzir. A lista salva pelo usuario e um
    # recorte desta: o layout (largura, fonte, cor) e definido por coluna,
    # entao nao da para pedir coluna que a funcao nao conhece.
    colunas_padrao: tuple[str, ...] = ()
    # Mostra na tela a caixa "Cruzar com MONA": um segundo arquivo opcional,
    # cruzado por POPID (mesma leitura da Função 4). A própria Função 4 fica
    # de fora — ela já pede a MONA como arquivo obrigatório.
    aceita_mona: bool = False

    @property
    def usa_lista_colunas(self) -> bool:
        return bool(self.colunas_padrao)

    # Captura do Outlook so faz sentido quando a funcao pede um arquivo so.
    @property
    def aceita_outlook(self) -> bool:
        return len([c for c in self.arquivos if not c.opcional]) == 1

    def campos_obrigatorios(self) -> list[str]:
        return [c.rotulo for c in self.arquivos if not c.opcional]


CATALOGO: dict[str, Processador] = {
    "fa2": Processador(
        id="fa2",
        titulo="FA2",
        descricao="Formata o MIX da FA2: filtra colunas, aplica cores e larguras.",
        funcao=formatar_excel_mix_logic,
        prefixo_saida="FA2_Processado",
        arquivos=(CampoArquivo("arquivo", "Planilha MIX"),),
        colunas_padrao=tuple(COLUNAS_DESEJADAS_PADRAO),
        aceita_mona=True,
    ),
    "fa12": Processador(
        id="fa12",
        titulo="FA1.2",
        descricao="Formata o MIX da FA1.2 com as variantes e cores da função 1.2.",
        funcao=formatar_excel_mix_logic_funcao_um_dois,
        prefixo_saida="FA12_Processado",
        arquivos=(CampoArquivo("arquivo", "Planilha MIX"),),
        colunas_padrao=tuple(COLUNAS_FA12),
        aceita_mona=True,
    ),
    "fibras": Processador(
        id="fibras",
        titulo="Fibras",
        descricao="Gera o MIX DE FIBRAS a partir do MIX geral, com código de "
                  "barras do chassi e a grelha.",
        funcao=formatar_excel_mix_fibras_logic,
        prefixo_saida="MIX_Fibras",
        arquivos=(CampoArquivo("arquivo", "MIX geral"),),
        colunas_padrao=tuple(COLUNAS_FIBRAS),
        observacao="A grelha sai do MIX: LEVEL = KD vira KD e AUX CAB COOLER "
                   "preenchido vira CLIMATIZADOR.",
    ),
    "f4": Processador(
        id="f4",
        titulo="Função 4",
        descricao="Cruza MIX e MONA, formata a planilha e anexa o MIX de fibras.",
        funcao=formatar_excel_mix_logic_funcao_quatro,
        prefixo_saida="F4_Processado",
        arquivos=(
            CampoArquivo("arquivo_mix", "Planilha MIX"),
            CampoArquivo("arquivo_mona", "Planilha MONA",
                         filtra_colunas=False,
                         ajuda="Formato próprio (cabeçalho técnico na linha 1, "
                               "itens na linha 2) — não passa pelo filtro de "
                               "colunas."),
        ),
        colunas_padrao=tuple(COLUNAS_F4),
    ),
}


def obter(processador_id: str) -> Processador | None:
    return CATALOGO.get(processador_id)


def listar() -> list[Processador]:
    return list(CATALOGO.values())

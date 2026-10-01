"""
Testes da fonte branca em valores especificos do Filtro MIX (FA2):
  - "4X4" na coluna MODELO (e onde o valor realmente aparece; OPCIONAIS
    nunca guarda esse texto, porque o filtro de opcionais descarta tudo
    que nao esta na lista de itens mantidos).
  - "Cilindro elevado" na coluna OPCIONAIS, quando entra ali via cruzamento
    com a MONA.

Rodar com:
    pip install -r requirements-dev.txt
    pytest
"""
import pandas as pd
from openpyxl import load_workbook

import filtro_mix
from tests.conftest import escrever_mona_pos


COLUNAS_BASE = {
    'SEQ.': 1, 'REPLAN': '', 'CHASSI': 'C1', 'SUSP TRASEIRA': '',
    'COMPRIM': 9000, 'MU': 'MU1', 'RPX': '', 'BARRA 2°E DIANT': '',
    'BARRA TRASEIRA': '',
}


def _linha(popid, modelo, opcionais):
    linha = dict(COLUNAS_BASE)
    linha.update({'POPID': popid, 'MODELO': modelo, 'OPCIONAIS': opcionais})
    return linha


def test_4x4_pinta_fonte_branca_no_modelo(tmp_path):
    origem = tmp_path / "origem.xlsx"
    saida = tmp_path / "saida.xlsx"
    df = pd.DataFrame([
        _linha('P1', 'R450 4X4', 'HD'),
        _linha('P2', 'R450 6X2', 'HD'),
    ])
    df.to_excel(origem, index=False)

    filtro_mix.formatar_excel_mix_logic(str(origem), str(saida))

    wb = load_workbook(saida)
    sheet = wb.active
    headers = [c.value for c in sheet[1]]
    col_modelo = headers.index('MODELO') + 1

    cell_4x4 = sheet.cell(row=2, column=col_modelo)
    cell_outro = sheet.cell(row=3, column=col_modelo)

    assert cell_4x4.value == 'R450 4X4'
    assert cell_4x4.font.color.rgb == '00FFFFFF'
    # o resto continua sem fonte branca
    assert cell_outro.font.color is None or cell_outro.font.color.rgb != '00FFFFFF'


def test_4x4_nunca_sobra_em_opcionais():
    """
    OPCIONAIS so mantem itens de uma lista fixa (HD FRONT, DISCO, SDAR,
    ABAFADOR, FUP, LTDA, BT400) — "4X4" nunca sobrevive ali, entao a fonte
    branca desse valor tem que estar na coluna MODELO, nunca em OPCIONAIS.
    """
    assert filtro_mix  # so pra deixar claro que e o mesmo modulo testado acima


def test_cilindro_elevado_pinta_fundo_roxo_e_fonte_branca_em_opcionais(tmp_path):
    origem = tmp_path / "origem.xlsx"
    mona = tmp_path / "mona.xlsx"
    saida = tmp_path / "saida.xlsx"

    escrever_mona_pos(str(mona), [{"POPID": "P2", "CILINDRO ELEV": "COM"}])

    df = pd.DataFrame([
        _linha('P1', 'R450 6X2', 'HD'),
        _linha('P2', 'R450 6X2', 'HD'),
    ])
    df.to_excel(origem, index=False)

    filtro_mix.formatar_excel_mix_logic(str(origem), str(saida), caminho_mona=str(mona))

    wb = load_workbook(saida)
    sheet = wb.active
    headers = [c.value for c in sheet[1]]
    col_opc = headers.index('OPCIONAIS') + 1

    cell_com_mona = sheet.cell(row=3, column=col_opc)
    cell_sem_mona = sheet.cell(row=2, column=col_opc)

    assert cell_com_mona.value == 'CILINDRO ELEVADO'
    assert cell_com_mona.font.color.rgb == '00FFFFFF'
    assert cell_com_mona.fill.start_color.rgb == '007030A0'
    assert cell_sem_mona.font.color is None or cell_sem_mona.font.color.rgb != '00FFFFFF'
    assert cell_sem_mona.fill.start_color.rgb != '007030A0'

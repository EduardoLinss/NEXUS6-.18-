"""
Testes do dropdown de colunas na tela de tarefas de variantes
(app.colunas_e_valores).

O pedido: mesmo sem nenhuma planilha carregada ainda, o campo "coluna" das
condicoes de uma tarefa deve oferecer as colunas padrao da funcao (a mesma
lista usada no construtor de regras de gargalos), em vez de ficar so no
texto livre. Quando ha planilha carregada, as colunas e valores reais dela
entram tambem, sem perder as colunas padrao que não estiverem na planilha.

Rodar com:
    pip install -r requirements-dev.txt
    pytest
"""
import pandas as pd

import app as appmod
import gargalos


def test_sem_planilha_usa_colunas_padrao_da_funcao():
    mapa = appmod.colunas_e_valores("", "fa2")
    assert set(mapa.keys()) == set(gargalos.COLUNAS_REGRA_FA2)
    # sem planilha carregada, nao ha valores sugeridos ainda
    assert all(valores == [] for valores in mapa.values())


def test_sem_planilha_fa12_usa_lista_propria():
    mapa = appmod.colunas_e_valores("", "1-2")
    assert set(mapa.keys()) == set(gargalos.COLUNAS_REGRA_1_2)
    assert "SORDER" in mapa
    assert "TPM" in mapa


def test_setor_desconhecido_cai_no_padrao_fa2():
    mapa = appmod.colunas_e_valores("", None)
    assert set(mapa.keys()) == set(gargalos.COLUNAS_REGRA_FA2)


def test_com_planilha_mescla_colunas_reais_e_padrao(tmp_path, monkeypatch):
    token = "token-teste-tarefas"
    df = pd.DataFrame([
        {"MODELO": "R450 8X4", "SORDER": "S1", "COLUNA NOVA": "X"},
        {"MODELO": "R450 6X2", "SORDER": "S2", "COLUNA NOVA": "Y"},
    ])

    def _fake_recuperar_dataframe(tok):
        return df if tok == token else None

    monkeypatch.setattr(appmod, "recuperar_dataframe", _fake_recuperar_dataframe)

    mapa = appmod.colunas_e_valores(token, "1-2")

    # colunas padrao da FA1.2 continuam la, mesmo as que nao estao na planilha
    assert "ENTRADA LINHA REAL" in mapa
    assert mapa["ENTRADA LINHA REAL"] == []

    # colunas reais da planilha entram com os valores mais comuns
    assert set(mapa["MODELO"]) == {"R450 8X4", "R450 6X2"}
    # coluna da planilha que nao estava na lista padrao tambem aparece
    assert "COLUNA NOVA" in mapa
    assert set(mapa["COLUNA NOVA"]) == {"X", "Y"}

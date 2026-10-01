"""
Feature: a tela de Análise (/) mostra qual planilha MONA está carregada
para a planilha atual — antes disso, o nome do arquivo da MONA nunca era
guardado, só o texto usado no cruzamento (para o contador do CPM MIX).

Rodar com:
    pip install -r requirements-dev.txt
    pytest
"""
import io

import pandas as pd
import pytest

import app as app_mod


@pytest.fixture
def cliente(tmp_path, monkeypatch):
    monkeypatch.setattr(app_mod, "PASTA_SESSAO", str(tmp_path))
    app_mod.app.config["TESTING"] = True
    with app_mod.app.test_client() as cliente:
        yield cliente


def _planilha_basica() -> bytes:
    df = pd.DataFrame({
        "POPID": ["P1", "P2"],
        "MODELO": ["6X2", "6X2"],
        "GÁS": ["", ""],
        "OPCIONAIS": ["", ""],
        "COMPRIM": [7000, 7000],
        "SUSP TRASEIRA": ["", ""],
        "BARRA TRASEIRA": ["", ""],
        "RPX": ["", ""],
        "MU": ["", ""],
    })
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer.read()


def _planilha_mona() -> bytes:
    df = pd.DataFrame({"POPID": ["P1"], "SDAR": ["COM"]})
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer.read()


def test_nome_da_mona_aparece_apos_o_cruzamento(cliente):
    conteudo1 = _planilha_basica()
    conteudo2 = _planilha_basica()
    conteudo_mona = _planilha_mona()

    resposta = cliente.post("/analise", data={
        "planilhas": [
            (io.BytesIO(conteudo1), "planilha1.xlsx"),
            (io.BytesIO(conteudo2), "planilha2.xlsx"),
        ],
        "usar_mona": "1",
        "arquivo_mona_extra": (io.BytesIO(conteudo_mona), "MONA_SETEMBRO.xlsx"),
        "texto_mona_fixo": "SDAR",
    }, content_type="multipart/form-data")

    assert resposta.status_code == 200
    assert b"MONA carregada" in resposta.data
    assert b"MONA_SETEMBRO.xlsx" in resposta.data


def test_sem_mona_nao_mostra_marcador(cliente):
    conteudo1 = _planilha_basica()
    conteudo2 = _planilha_basica()

    resposta = cliente.post("/analise", data={
        "planilhas": [
            (io.BytesIO(conteudo1), "planilha1.xlsx"),
            (io.BytesIO(conteudo2), "planilha2.xlsx"),
        ],
    }, content_type="multipart/form-data")

    assert resposta.status_code == 200
    assert b"MONA carregada" not in resposta.data

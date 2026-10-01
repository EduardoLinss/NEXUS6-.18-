"""
Testes do cruzamento com a MONA: cpm_core.TAGS_MONA, mix_processor_4.py
(cruzar_com_mona / _ler_e_agrupar_mona) e cpm_core.analisar_mix.

Rodar com:
    pip install -r requirements-dev.txt
    pytest
"""
import pandas as pd
import pytest

import cpm_core as core
from mix_processor_4 import cruzar_com_mona, _detectar_formato_mona
from tests.conftest import (
    escrever_mona_pos, escrever_mona_lista_simples, escrever_mix_fa2,
)


# ---------------------------------------------------------------------------
# Deteccao de formato
# ---------------------------------------------------------------------------

def test_detecta_formato_pos(tmp_path):
    caminho = tmp_path / "mona_pos.xlsx"
    escrever_mona_pos(str(caminho), [{"POPID": "1001", "SMOLA 658 / 660": "COM"}])
    assert _detectar_formato_mona(str(caminho)) == "pos"


def test_detecta_formato_lista_simples(tmp_path):
    caminho = tmp_path / "mona_lista.xlsx"
    escrever_mona_lista_simples(str(caminho), ["1001", "1002"])
    assert _detectar_formato_mona(str(caminho)) == "lista"


# ---------------------------------------------------------------------------
# cruzar_com_mona — formato tecnico (POS_)
# ---------------------------------------------------------------------------

def test_pos_so_cruza_tags_conhecidas(tmp_path):
    """Colunas que ja existem no mix geral (fora de TAGS_MONA) sao ignoradas."""
    caminho = tmp_path / "mona.xlsx"
    escrever_mona_pos(str(caminho), [
        {"POPID": "1001", "SMOLA 658 / 660": "COM", "8x2": "COM"},
        {"POPID": "1002", "8x2": "COM"},
    ])
    df = pd.DataFrame({"POPID": ["1001", "1002"], "OPCIONAIS": ["", ""]})

    resultado = cruzar_com_mona(df, str(caminho), coluna_saida="_MONA_TEXTO")

    assert resultado.loc[resultado["POPID"] == "1001", "_MONA_TEXTO"].iloc[0] == "MOLA 658|660"
    assert pd.isna(resultado.loc[resultado["POPID"] == "1002", "_MONA_TEXTO"].iloc[0])


def test_pos_reconhece_abreviacao_cilindro_elevado(tmp_path):
    caminho = tmp_path / "mona.xlsx"
    escrever_mona_pos(str(caminho), [{"POPID": "1001", "CILINDRO ELEV.": "COM"}])
    df = pd.DataFrame({"POPID": ["1001"], "OPCIONAIS": [""]})

    resultado = cruzar_com_mona(df, str(caminho), coluna_saida="_MONA_TEXTO")

    assert resultado["_MONA_TEXTO"].iloc[0] == "CILINDRO ELEVADO"


def test_pos_multiplas_tags_no_mesmo_popid(tmp_path):
    caminho = tmp_path / "mona.xlsx"
    escrever_mona_pos(str(caminho), [
        {"POPID": "1001", "SMOLA 658 / 660": "COM", "649|650 RESTRICAO": "COM"},
    ])
    df = pd.DataFrame({"POPID": ["1001"], "OPCIONAIS": [""]})

    resultado = cruzar_com_mona(df, str(caminho), coluna_saida="_MONA_TEXTO")

    texto = resultado["_MONA_TEXTO"].iloc[0]
    assert "658|660" in texto
    assert "649|650" in texto


def test_pos_texto_avulso_soma_as_tags_fixas(tmp_path):
    """O campo de texto livre tambem funciona no formato tecnico, alem das
    tags fixas — sem apagar o que ja e reconhecido automaticamente."""
    caminho = tmp_path / "mona.xlsx"
    escrever_mona_pos(str(caminho), [
        {"POPID": "1001", "5ª RODA": "COM"},
        {"POPID": "1002", "SMOLA 658 / 660": "COM"},
    ])
    df = pd.DataFrame({"POPID": ["1001", "1002"], "OPCIONAIS": ["", ""]})

    resultado = cruzar_com_mona(df, str(caminho), coluna_saida="_MONA_TEXTO",
                                texto_lista_simples="5ª RODA")

    assert resultado.loc[resultado["POPID"] == "1001", "_MONA_TEXTO"].iloc[0] == "5ª RODA"
    assert resultado.loc[resultado["POPID"] == "1002", "_MONA_TEXTO"].iloc[0] == "MOLA 658|660"


def test_pos_sem_nenhuma_tag_relevante_da_erro_claro(tmp_path):
    caminho = tmp_path / "mona.xlsx"
    escrever_mona_pos(str(caminho), [{"POPID": "1001", "8x2": "COM"}])
    df = pd.DataFrame({"POPID": ["1001"], "OPCIONAIS": [""]})

    with pytest.raises(ValueError, match="Nenhuma das colunas relevantes"):
        cruzar_com_mona(df, str(caminho), coluna_saida="_MONA_TEXTO")


def test_pos_modo_exclusao_ignora_so_o_que_ja_existe(tmp_path):
    """
    Ao contrario da lista permitida (usada pelo FA2), o modo de exclusao
    (usado pelo FA1.2) deixa passar a MAIORIA das colunas e so tira o que
    ja existe no mix geral.
    """
    caminho = tmp_path / "mona.xlsx"
    escrever_mona_pos(str(caminho), [
        {"POPID": "1001", "GÁS": "COM", "6X6/4X4": "COM", "ABS": "COM"},
    ])
    df = pd.DataFrame({"POPID": ["1001"], "OPCIONAIS": [""]})

    resultado = cruzar_com_mona(df, str(caminho), coluna_saida="_MONA_TEXTO",
                                itens_excluidos=(r"\bG[ÁA]S\b", r"6X6\D*4X4"))

    assert resultado["_MONA_TEXTO"].iloc[0] == "ABS"


def test_pos_nao_corrompe_tag_com_pipe_em_popid_duplicado(tmp_path):
    """
    Regressao: o mesmo POPID em duas linhas da MONA passava pela funcao de
    'juntar sem duplicar', que quebrava "658|660" em "658" e "660" porque
    usava um "|" solto como separador. O separador tem que ser " | " (com
    espacos), nunca o caractere sozinho.
    """
    caminho = tmp_path / "mona.xlsx"
    # Duas linhas com o MESMO POPID — forca o agrupamento interno.
    escrever_mona_pos(str(caminho), [
        {"POPID": "1001", "SMOLA 658 / 660": "COM"},
        {"POPID": "1001", "SMOLA 658 / 660": "COM"},
    ])
    df = pd.DataFrame({"POPID": ["1001"], "OPCIONAIS": [""]})

    resultado = cruzar_com_mona(df, str(caminho), coluna_saida="_MONA_TEXTO")

    assert resultado["_MONA_TEXTO"].iloc[0] == "MOLA 658|660"


# ---------------------------------------------------------------------------
# cruzar_com_mona — lista simples
# ---------------------------------------------------------------------------

def test_lista_simples_marca_popids_encontrados(tmp_path):
    caminho = tmp_path / "mona_lista.xlsx"
    escrever_mona_lista_simples(str(caminho), ["1001", "1002"])
    df = pd.DataFrame({"POPID": ["1001", "1002", "1003"], "OPCIONAIS": ["", "", ""]})

    resultado = cruzar_com_mona(df, str(caminho), coluna_saida="_MONA_TEXTO",
                                texto_lista_simples="650/649")

    assert resultado.loc[resultado["POPID"] == "1001", "_MONA_TEXTO"].iloc[0] == "650/649"
    assert resultado.loc[resultado["POPID"] == "1002", "_MONA_TEXTO"].iloc[0] == "650/649"
    assert pd.isna(resultado.loc[resultado["POPID"] == "1003", "_MONA_TEXTO"].iloc[0])


def test_lista_simples_sem_texto_da_erro_claro(tmp_path):
    caminho = tmp_path / "mona_lista.xlsx"
    escrever_mona_lista_simples(str(caminho), ["1001"])
    df = pd.DataFrame({"POPID": ["1001"], "OPCIONAIS": [""]})

    with pytest.raises(ValueError, match="lista de POPIDs"):
        cruzar_com_mona(df, str(caminho), coluna_saida="_MONA_TEXTO")


# ---------------------------------------------------------------------------
# Filtro MIX (FA2) — cruzamento ponta a ponta, junta em OPCIONAIS
# ---------------------------------------------------------------------------

def test_fa2_junta_mona_em_opcionais(tmp_path):
    import filtro_mix as fm

    caminho_mix = tmp_path / "mix.xlsx"
    caminho_mona = tmp_path / "mona.xlsx"
    caminho_saida = tmp_path / "saida.xlsx"

    escrever_mix_fa2(str(caminho_mix), [
        {"POPID": "1001", "OPCIONAIS": "HD FRONT"},
        {"POPID": "1002", "OPCIONAIS": ""},
    ])
    escrever_mona_lista_simples(str(caminho_mona), ["1001"])

    ok, msg = fm.formatar_excel_mix_logic(
        str(caminho_mix), str(caminho_saida),
        caminho_mona=str(caminho_mona), texto_mona="650/649",
    )
    assert ok, msg

    df = pd.read_excel(caminho_saida)
    # FA2 normaliza OPCIONAIS (filtrar_opcionais): "HD FRONT" vira "HD".
    assert df.loc[df["POPID"] == 1001, "OPCIONAIS"].iloc[0] == "HD | 650/649"
    linha_1002 = df.loc[df["POPID"] == 1002, "OPCIONAIS"].iloc[0]
    assert pd.isna(linha_1002) or linha_1002 == ""


# ---------------------------------------------------------------------------
# analisar_mix — contadores fixos (TAGS_MONA) e dinamico (contador_mona)
# ---------------------------------------------------------------------------

def _df_mix_basico(opcionais):
    n = len(opcionais)
    return pd.DataFrame({
        "POPID": [f"P{i}" for i in range(n)],
        "MODELO": ["6X4"] * n,
        "OPCIONAIS": opcionais,
        "SUSP TRASEIRA": ["T"] * n,
        "RPX": [""] * n,
        "MU": ["BR01"] * n,
        "COMPRIM": [7000] * n,
        "BARRA TRASEIRA": [""] * n,
    })


def _chave_mix(texto):
    import re
    return "MONA_" + re.sub(r"\W+", "_", texto).strip("_").upper()


def test_tags_fixas_contam_automaticamente():
    df = _df_mix_basico(["MOLA 658|660", "MOLA 649|650 | CILINDRO ELEVADO", "DISCO"])
    resultado = core.analisar_mix(df, "P0", len(df), takt=60)

    assert resultado.mix[_chave_mix("MOLA 658|660")] == 1
    assert resultado.mix[_chave_mix("MOLA 649|650")] == 1
    assert resultado.mix[_chave_mix("CILINDRO ELEVADO")] == 1


def test_contador_dinamico_para_texto_avulso():
    df = _df_mix_basico(["5ª RODA", "DISCO"])
    resultado = core.analisar_mix(df, "P0", len(df), takt=60, contador_mona="5ª RODA")

    assert resultado.mix[_chave_mix("5ª RODA")] == 1


def test_contador_dinamico_nao_duplica_tag_fixa():
    """Se o texto digitado bater com uma tag fixa, nao deve virar um segundo
    contador — CATALOGO_MIX ja tem um pra ela."""
    df = _df_mix_basico(["658|660"])
    resultado = core.analisar_mix(df, "P0", len(df), takt=60, contador_mona="658|660")

    chave = _chave_mix("658|660")
    assert list(resultado.mix.keys()).count(chave) == 1
    assert resultado.mix[chave] == 1


def test_contador_so_conta_dentro_da_janela_analisada():
    """POPIDs fora do intervalo POPID+Producao nao devem ser contados."""
    df = _df_mix_basico(["DISCO", "MOLA 658|660", "DISCO", "DISCO"])
    # Janela: comeca em P0, so 2 linhas (P0 e P1) — "MOLA 658|660" esta na P1, entra.
    resultado = core.analisar_mix(df, "P0", 2, takt=60)
    assert resultado.mix[_chave_mix("MOLA 658|660")] == 1

    # Janela comecando em P2 (so P2 e P3) — "MOLA 658|660" (P1) fica de fora.
    resultado2 = core.analisar_mix(df, "P2", 2, takt=60)
    assert resultado2.mix[_chave_mix("MOLA 658|660")] == 0

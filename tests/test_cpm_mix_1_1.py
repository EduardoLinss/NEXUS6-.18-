"""
CPM MIX do setor 1-1: mesma estrutura das outras funcoes (modelos e Veiculo
longo); os itens proprios sao os do MIX da 1.1 — WTA (em OPCIONAIS), RHD, LNG,
AUX CAB COOLER, DUPLO CIRCUITO e COOLING COIL.
"""
import pandas as pd
import pytest

import cpm_core as core
import gargalos
import tags_mona
from mix_processor_1_1 import FONTES_VARIANTES


@pytest.fixture(autouse=True)
def isolar_tags(tmp_path, monkeypatch):
    monkeypatch.setattr(tags_mona, "caminho_dados", lambda nome: str(tmp_path / nome))
    yield


ITENS_1_1 = {"WTA", "RHD", "LNG", "AUX_CAB_COOLER", "DUPLO_CIRCUITO", "COOLING_COIL"}
ESTRUTURA = {"10x4", "8x4", "8x2", "6x6", "6x4", "6x2", "4x4", "4x2", "VEICULO_LONGO"}


def test_1_1_tem_catalogo_proprio():
    chaves = {c.chave for c in core.catalogo_mix_do_setor("1-1")}
    assert chaves == ESTRUTURA | ITENS_1_1
    # nada da FA2 nem da FA1.2 fora a estrutura comum
    assert "GAS" not in chaves and "N_MODULE" not in chaves and "ABAFADOR" not in chaves


def test_outros_setores_nao_mudaram():
    assert core.catalogo_mix_do_setor("fa2") == core.CATALOGO_MIX_FA2
    assert core.catalogo_mix_do_setor("1-2") == core.CATALOGO_MIX_1_2


def test_itens_do_cpm_sao_as_mesmas_colunas_do_filtro_mix_1_1():
    colunas_cpm = set()
    for c in core.catalogo_mix_do_setor("1-1"):
        if c.chave in ITENS_1_1:
            colunas_cpm.update(c.criterio.colunas())
    assert colunas_cpm == {f.coluna for f in FONTES_VARIANTES}


def test_1_1_nao_herda_as_tags_mona_da_fa2():
    assert core.tags_mona_padrao_do_setor("1-1") == ()
    assert core.tags_mona_ativas("1-1") == ()
    assert core.tags_mona_padrao_do_setor("fa2") == core.TAGS_MONA_PADRAO


def test_contagem_no_analisar_mix():
    df = pd.DataFrame([
        {"POPID": "1", "MODELO": "R 460 A6X2", "OPCIONAIS": "EBS/COOLING/WTA",
         "COOLING COIL": "COOLING COIL", "RHD": None, "COMPRIM": 7125},
        {"POPID": "2", "MODELO": "P 320 B8X2", "OPCIONAIS": "EBS", "RHD": "RHD",
         "AUX CAB COOLER": "AUX CAB", "COMPRIM": 9000},
        {"POPID": "3", "MODELO": "G 560 A8X4", "OPCIONAIS": "", "LNG": "LNG",
         "DUPLO CIRCUITO": "DUPLO", "COMPRIM": 7000},
    ])
    r = core.analisar_mix(df, "1", "3", setor="1-1")
    assert r.mix["WTA"] == 1
    assert r.mix["COOLING_COIL"] == 1
    assert r.mix["RHD"] == 1          # celula vazia (None) nao conta
    assert r.mix["AUX_CAB_COOLER"] == 1
    assert r.mix["LNG"] == 1
    assert r.mix["DUPLO_CIRCUITO"] == 1
    assert r.mix["8x2"] == 1 and r.mix["8x4"] == 1 and r.mix["6x2"] == 1
    assert r.mix["VEICULO_LONGO"] == 1
    assert not any(k.startswith("MONA_") for k in r.mix)


def test_preenchido_ignora_vazio_e_espaco():
    df = pd.DataFrame({"RHD": ["RHD", "", "  ", None, float("nan")]})
    ctx = core.ContextoDados(df)
    assert ctx.contar(core.Preenchido("RHD")) == 1


def test_construtor_de_regras_do_1_1_sugere_o_cabecalho_completo():
    colunas = gargalos.colunas_regra_do_setor("1-1")
    for coluna in ("RETARDER", "COOLING COIL", "DUPLO CIRCUITO", "AUX CAB COOLER"):
        assert coluna in colunas

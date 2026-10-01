"""
Testes do CPM MIX por setor: cada função (FA2, FA1.2) tem seu próprio
catálogo de itens contados (`catalogo_mix_do_setor`) e sua própria lista de
tags da MONA (`tags_mona_padrao_do_setor` / tags_mona.py), independentes uma
da outra — mesmo padrão já usado para as colunas do construtor de regras
(COLUNAS_REGRA_FA2 / COLUNAS_REGRA_1_2).

Rodar com:
    pip install -r requirements-dev.txt
    pytest
"""
import pandas as pd
import pytest

import catalogo as catalogo_mod
import cpm_core as core
import tags_mona
from cpm_core import Contem, Tarefa, Variante
from mix_processor_4 import cruzar_com_mona
from tests.conftest import escrever_mona_pos


@pytest.fixture(autouse=True)
def isolar_tags(tmp_path, monkeypatch):
    monkeypatch.setattr(tags_mona, "caminho_dados", lambda nome: str(tmp_path / nome))
    yield


@pytest.fixture(autouse=True)
def isolar_catalogo(tmp_path, monkeypatch):
    """Cada teste grava o catalogo de tarefas (Variantes) num arquivo proprio,
    isolado por setor, nunca no dados/ do projeto."""
    monkeypatch.setattr(catalogo_mod, "caminho_dados", lambda nome: str(tmp_path / nome))
    yield


# ---------------------------------------------------------------------------
# catalogo_mix_do_setor
# ---------------------------------------------------------------------------

def test_fa2_mantem_o_catalogo_historico():
    assert core.catalogo_mix_do_setor("fa2") == core.CATALOGO_MIX_FA2
    assert core.catalogo_mix_do_setor(None) == core.CATALOGO_MIX_FA2
    assert core.CATALOGO_MIX == core.CATALOGO_MIX_FA2  # compat


def test_1_2_tem_catalogo_proprio_com_gas_e_n_module():
    chaves = {c.chave for c in core.catalogo_mix_do_setor("1-2")}
    assert "GAS" in chaves
    assert "N_MODULE" in chaves
    assert "VEICULO_LONGO" in chaves
    assert "6x2" in chaves and "8x4" in chaves  # modelos, igual FA2
    # itens que na FA1.2 só entram via MONA não aparecem aqui direto:
    assert "SDAR" not in chaves
    assert "LTDA" not in chaves


def test_setor_desconhecido_cai_no_padrao_fa2():
    assert core.catalogo_mix_do_setor("outro-setor-qualquer") == core.CATALOGO_MIX_FA2


# ---------------------------------------------------------------------------
# tags_mona_padrao_do_setor
# ---------------------------------------------------------------------------

def test_tags_padrao_1_2_tem_os_itens_da_mona():
    textos = {t.texto for t in core.tags_mona_padrao_do_setor("1-2")}
    assert textos == {
        "LTDA", "CALHA INVERTIDA", "RADAR LATERAL", "SDAR", "DIR. ELETRICA",
        "BATERIA TRASEIRA", "K1 20L", "CHICOTE FOU", 'SOQUETE TANK "D',
    }


def test_tags_padrao_fa2_nao_muda():
    assert core.tags_mona_padrao_do_setor("fa2") == core.TAGS_MONA_PADRAO
    assert core.tags_mona_padrao_do_setor(None) == core.TAGS_MONA_PADRAO


# ---------------------------------------------------------------------------
# persistencia por setor (tags_mona.json / tags_mona.py)
# ---------------------------------------------------------------------------

def test_editar_tags_de_um_setor_nao_afeta_o_outro():
    tags_mona.adicionar("Item só da FA1.2", "", setor="1-2")

    textos_1_2 = [t.texto for t in tags_mona.carregar("1-2")]
    textos_fa2 = [t.texto for t in tags_mona.carregar("fa2")]

    assert "Item só da FA1.2" in textos_1_2
    assert "Item só da FA1.2" not in textos_fa2


def test_tags_mona_ativas_usa_a_lista_do_setor_certo():
    tags_mona.adicionar("Extra FA2", "", setor="fa2")
    tags_mona.adicionar("Extra 1-2", "", setor="1-2")

    ativas_fa2 = [t.texto for t in core.tags_mona_ativas("fa2")]
    ativas_1_2 = [t.texto for t in core.tags_mona_ativas("1-2")]

    assert "Extra FA2" in ativas_fa2 and "Extra 1-2" not in ativas_fa2
    assert "Extra 1-2" in ativas_1_2 and "Extra FA2" not in ativas_1_2


# ---------------------------------------------------------------------------
# analisar_mix() contando por setor
# ---------------------------------------------------------------------------

def _df_basico(popids, modelo="6X2", gas="", opcionais=""):
    return pd.DataFrame({
        "POPID": popids,
        "MODELO": [modelo] * len(popids),
        "GÁS": [gas] * len(popids),
        "OPCIONAIS": [opcionais] * len(popids),
        "COMPRIM": [7000] * len(popids),
        "SUSP TRASEIRA": [""] * len(popids),
        "BARRA TRASEIRA": [""] * len(popids),
        "RPX": [""] * len(popids),
        "MU": [""] * len(popids),
    })


def test_gas_so_conta_na_fa1_2():
    df = _df_basico(["P1", "P2"], gas="GÁS")
    resultado_1_2 = core.analisar_mix(df, "P1", 2, setor="1-2")
    resultado_fa2 = core.analisar_mix(df, "P1", 2, setor="fa2")

    assert resultado_1_2.mix.get("GAS") == 2
    assert "GAS" not in resultado_fa2.mix


def test_n_module_conta_na_fa1_2_via_opcionais():
    df = _df_basico(["P1", "P2"], opcionais="N_MODULE")
    resultado = core.analisar_mix(df, "P1", 2, setor="1-2")
    assert resultado.mix.get("N_MODULE") == 2


def test_mona_crossing_traz_tags_da_fa1_2_para_o_cpm_mix(tmp_path):
    """
    O mesmo cruzamento que o FA2 usa (cruzar_com_mona, modo padrão) tem que
    trazer as tags da FA1.2 quando `setor="1-2"` é passado — sem isso, so as
    tags da FA2 (MOLA/CILINDRO) seriam reconhecidas, mesmo analisando dados
    da FA1.2.
    """
    caminho_mona = tmp_path / "mona.xlsx"
    escrever_mona_pos(str(caminho_mona), [{"POPID": "P2", "CALHA INVERTIDA": "COM"}])

    df = _df_basico(["P1", "P2"])
    cruzado = cruzar_com_mona(df, str(caminho_mona), coluna_saida="_MONA_TEXTO",
                              setor="1-2")
    assert cruzado.loc[cruzado["POPID"] == "P2", "_MONA_TEXTO"].iloc[0] == "CALHA INVERTIDA"
    assert pd.isna(cruzado.loc[cruzado["POPID"] == "P1", "_MONA_TEXTO"].iloc[0])

    # junta o texto cruzado em OPCIONAIS, como o app.py faz antes de analisar
    cruzado["OPCIONAIS"] = cruzado["_MONA_TEXTO"].fillna("")
    resultado = core.analisar_mix(cruzado, "P1", 2, setor="1-2")
    assert resultado.mix.get("MONA_CALHA_INVERTIDA") == 1


# ---------------------------------------------------------------------------
# catalogo_ativo() / analisar_mix() — tarefas de variantes por setor
# ---------------------------------------------------------------------------

def _catalogo_de_teste(nome_tarefa: str) -> tuple[Variante, ...]:
    return (
        Variante("V1", "Variante 1", (
            Tarefa(nome_tarefa, 111, Contem("MODELO", "6X2"), f"{nome_tarefa} v1"),
        )),
    )


def _df_para_tarefa():
    return pd.DataFrame({
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


def test_catalogo_ativo_usa_o_catalogo_do_setor_certo():
    catalogo_mod.salvar(_catalogo_de_teste("Tarefa só da FA2"), setor="fa2")
    catalogo_mod.salvar(_catalogo_de_teste("Tarefa só da FA1.2"), setor="1-2")

    nomes_fa2 = {t.nome for v in core.catalogo_ativo("fa2") for t in v.tarefas}
    nomes_1_2 = {t.nome for v in core.catalogo_ativo("1-2") for t in v.tarefas}

    assert "Tarefa só da FA2" in nomes_fa2 and "Tarefa só da FA1.2" not in nomes_fa2
    assert "Tarefa só da FA1.2" in nomes_1_2 and "Tarefa só da FA2" not in nomes_1_2


def test_analisar_mix_nao_puxa_as_tarefas_da_fa2_quando_setor_e_1_2():
    """
    Regressão: selecionar a FA1.2 na tela de Análise tinha voltado a mostrar
    as tarefas da FA2 — `analisar_mix(..., setor="1-2")` chamava
    `catalogo_ativo()` sem repassar o setor, então o catálogo de
    Variantes/Tarefas sempre caía no da FA2, mesmo com o CPM MIX e as tags
    da MONA já corretos para a FA1.2.
    """
    catalogo_mod.salvar(_catalogo_de_teste("Tarefa só da FA2"), setor="fa2")
    catalogo_mod.salvar(_catalogo_de_teste("Tarefa só da FA1.2"), setor="1-2")

    df = _df_para_tarefa()
    resultado_1_2 = core.analisar_mix(df, "P1", 2, setor="1-2")
    resultado_fa2 = core.analisar_mix(df, "P1", 2, setor="fa2")

    nomes_1_2 = {t.nome for t in resultado_1_2.variantes["V1"].tarefas}
    nomes_fa2 = {t.nome for t in resultado_fa2.variantes["V1"].tarefas}

    assert nomes_1_2 == {"Tarefa só da FA1.2"}
    assert nomes_fa2 == {"Tarefa só da FA2"}

"""
Regressão: os links "baixar .json" / "baixar .tsv" da tela /tarefas sempre
baixavam o catálogo da FA2, não importa qual setor estivesse selecionado —
a rota `/tarefas/exportar.<formato>` chamava `catalogo_mod.exportar_*()` sem
setor nenhum, mesmo a pré-visualização na própria tela já usando o setor
certo. `/gargalos/regras/exportar.<formato>` já fazia isso certo; essa rota
de tarefas tinha ficado pra trás.

Rodar com:
    pip install -r requirements-dev.txt
    pytest
"""
import pytest

import app as app_mod
import catalogo as catalogo_mod
from cpm_core import Contem, Tarefa, Variante


@pytest.fixture(autouse=True)
def isolar_catalogo(tmp_path, monkeypatch):
    monkeypatch.setattr(catalogo_mod, "caminho_dados", lambda nome: str(tmp_path / nome))
    yield


@pytest.fixture
def cliente():
    app_mod.app.config["TESTING"] = True
    with app_mod.app.test_client() as cliente:
        yield cliente


def _catalogo_de_teste(nome_tarefa: str) -> tuple[Variante, ...]:
    return (
        Variante("V1", "Variante 1", (
            Tarefa(nome_tarefa, 60, Contem("OPCIONAIS", "X"), f"{nome_tarefa} v1"),
        )),
    )


def test_baixar_json_respeita_o_setor_selecionado(cliente):
    catalogo_mod.salvar(_catalogo_de_teste("Tarefa só da FA2"), setor="fa2")
    catalogo_mod.salvar(_catalogo_de_teste("Tarefa só da FA1.2"), setor="1-2")

    resposta_1_2 = cliente.get("/tarefas/exportar.json?setor=1-2")
    resposta_fa2 = cliente.get("/tarefas/exportar.json?setor=fa2")

    assert b"Tarefa s\xc3\xb3 da FA1.2" in resposta_1_2.data
    assert b"Tarefa s\xc3\xb3 da FA2" not in resposta_1_2.data

    assert b"Tarefa s\xc3\xb3 da FA2" in resposta_fa2.data
    assert b"Tarefa s\xc3\xb3 da FA1.2" not in resposta_fa2.data


def test_baixar_tsv_respeita_o_setor_selecionado(cliente):
    catalogo_mod.salvar(_catalogo_de_teste("Tarefa só da FA2"), setor="fa2")
    catalogo_mod.salvar(_catalogo_de_teste("Tarefa só da FA1.2"), setor="1-2")

    resposta = cliente.get("/tarefas/exportar.tsv?setor=1-2")

    assert b"Tarefa s\xc3\xb3 da FA1.2" in resposta.data
    assert b"Tarefa s\xc3\xb3 da FA2" not in resposta.data

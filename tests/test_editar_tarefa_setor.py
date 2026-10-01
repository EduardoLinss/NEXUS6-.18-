"""
Regressão: em /tarefas, clicar para editar uma tarefa da FA1.2 abria o
formulário em branco — `encontrar_tarefa(editar_id)` era chamado sem
`setor`, então a busca sempre caía no catálogo da FA2 (padrão), não achava
o id da FA1.2 e devolvia (None, None), fazendo a tela cair no formulário
vazio em vez de preenchido.

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


def _catalogo_de_teste(tarefa_id: str, nome_tarefa: str) -> tuple[Variante, ...]:
    return (
        Variante("V1", "Variante 1", (
            Tarefa(nome_tarefa, 60, Contem("OPCIONAIS", "X"), f"{nome_tarefa} legado",
                   id=tarefa_id),
        )),
    )


def test_editar_tarefa_da_fa1_2_preenche_o_formulario(cliente):
    catalogo_mod.salvar(_catalogo_de_teste("id-fa2", "Tarefa só da FA2"), setor="fa2")
    catalogo_mod.salvar(_catalogo_de_teste("id-1-2", "Tarefa só da FA1.2"), setor="1-2")

    resposta = cliente.get("/tarefas?editar=id-1-2&setor=1-2")

    assert resposta.status_code == 200
    # o nome também aparece na lista lateral de tarefas da variante, mesmo
    # com o formulário em branco — por isso o teste checa o campo do
    # formulário (value="...") e não só a presença do texto na página.
    assert b'name="nome" value="Tarefa s\xc3\xb3 da FA1.2"' in resposta.data


def test_editar_tarefa_da_fa2_continua_preenchendo(cliente):
    catalogo_mod.salvar(_catalogo_de_teste("id-fa2", "Tarefa só da FA2"), setor="fa2")
    catalogo_mod.salvar(_catalogo_de_teste("id-1-2", "Tarefa só da FA1.2"), setor="1-2")

    resposta = cliente.get("/tarefas?editar=id-fa2&setor=fa2")

    assert resposta.status_code == 200
    assert b'name="nome" value="Tarefa s\xc3\xb3 da FA2"' in resposta.data

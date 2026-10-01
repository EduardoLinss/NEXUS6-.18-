"""Testes do cadastro de tags da MONA (tags_mona.py)."""
import pytest

import tags_mona
import cpm_core as core


@pytest.fixture(autouse=True)
def isolar_arquivo(tmp_path, monkeypatch):
    """
    Cada teste grava num tags_mona.json proprio, nunca no do projeto.
    `arquivo_do_setor` (o arquivo por setor e o fallback antigo) e
    `copia_de_seguranca` (os backups) usam `caminho_dados` — isolando essa
    funcao, isola tudo de uma vez.
    """
    monkeypatch.setattr(tags_mona, "caminho_dados",
                        lambda nome: str(tmp_path / nome))
    yield


def test_sem_arquivo_usa_padrao():
    tags = tags_mona.carregar()
    assert tags == core.TAGS_MONA_PADRAO


def test_adicionar_grava_e_aparece_no_carregar():
    tags_mona.adicionar("649|650 extra", r"649\D*650\D*extra")
    textos = [t.texto for t in tags_mona.carregar()]
    assert "649|650 extra" in textos


def test_adicionar_sem_padrao_usa_texto_exato_escapado():
    tags_mona.adicionar("Item Novo", "")
    tag = next(t for t in tags_mona.carregar() if t.texto == "Item Novo")
    assert len(tag.padroes) == 1
    import re
    assert re.search(tag.padroes[0], "algo com Item Novo no meio", re.IGNORECASE)


def test_adicionar_duplicado_falha():
    tags_mona.adicionar("Repetido", "")
    with pytest.raises(tags_mona.ErroTagsMona, match="Já existe"):
        tags_mona.adicionar("Repetido", "")


def test_adicionar_texto_vazio_falha():
    with pytest.raises(tags_mona.ErroTagsMona):
        tags_mona.adicionar("   ", "")


def test_adicionar_regex_invalido_falha():
    with pytest.raises(tags_mona.ErroTagsMona, match="Padrão inválido"):
        tags_mona.adicionar("Item", "(")


def test_remover_tira_da_lista():
    tags_mona.adicionar("Temporario", "")
    tags_mona.remover("Temporario")
    textos = [t.texto for t in tags_mona.carregar()]
    assert "Temporario" not in textos


def test_restaurar_padrao_volta_as_tres_originais():
    tags_mona.adicionar("Extra", "")
    tags_mona.restaurar_padrao()
    assert tags_mona.carregar() == core.TAGS_MONA_PADRAO


def test_tags_mona_ativas_reflete_o_arquivo_editado():
    """cpm_core.tags_mona_ativas() precisa refletir a edicao na hora, sem
    precisar reiniciar nada — é exatamente o ponto de ter isso num JSON."""
    tags_mona.adicionar("Dinamico", "")
    ativas = [t.texto for t in core.tags_mona_ativas()]
    assert "Dinamico" in ativas

"""
O link de download do Filtro MIX precisa sobreviver a um reinicio do
servidor: o indice nome -> caminho fica em processados.json, nao em memoria.
"""
import importlib
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


@pytest.fixture
def appmod(tmp_path, monkeypatch):
    monkeypatch.setenv("MENUMIX_DADOS", str(tmp_path))
    import armazenamento
    importlib.reload(armazenamento)
    import app
    app = importlib.reload(app)
    app.app.config["TESTING"] = True
    yield app
    monkeypatch.delenv("MENUMIX_DADOS")
    importlib.reload(armazenamento)
    importlib.reload(app)


def test_download_continua_valido_apos_reinicio(appmod, tmp_path):
    gerado = tmp_path / "saida" / "MIX_FA2_20261118_101010.xlsx"
    gerado.parent.mkdir()
    gerado.write_bytes(b"conteudo")
    appmod.registrar_processado(gerado.name, str(gerado))

    # "reinicio": o modulo e recarregado, nada em memoria sobrevive
    appmod = importlib.reload(appmod)
    resposta = appmod.app.test_client().get(f"/processados/{gerado.name}")

    assert resposta.status_code == 200
    assert resposta.data == b"conteudo"


def test_nome_fora_do_indice_nao_vira_caminho(appmod, tmp_path):
    (tmp_path / "segredo.txt").write_text("x")
    cliente = appmod.app.test_client()

    assert appmod.caminho_processado("segredo.txt") is None
    assert appmod.caminho_processado("../segredo.txt") is None
    assert cliente.get("/processados/segredo.txt").status_code == 302


def test_indice_descarta_arquivos_apagados_e_limita_tamanho(appmod, tmp_path, monkeypatch):
    monkeypatch.setattr(appmod, "MAX_PROCESSADOS", 3)
    apagado = tmp_path / "apagado.xlsx"
    apagado.write_text("x")
    appmod.registrar_processado(apagado.name, str(apagado))
    apagado.unlink()

    for i in range(5):
        arq = tmp_path / f"f{i}.xlsx"
        arq.write_text("x")
        appmod.registrar_processado(arq.name, str(arq))

    assert list(appmod._ler_processados()) == ["f2.xlsx", "f3.xlsx", "f4.xlsx"]


def test_indice_corrompido_nao_derruba_o_download(appmod):
    with open(appmod.ARQUIVO_PROCESSADOS, "w", encoding="utf-8") as saida:
        saida.write("{nao e json")
    assert appmod.caminho_processado("qualquer.xlsx") is None

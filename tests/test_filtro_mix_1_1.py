"""
Função 1.1 do Filtro MIX: MIX principal + MONA (obrigatória), cruzados por
POPID. Planilha final: SEQ, VARIANTES, POPID, MODELO, MU — nessa ordem.
"""
import importlib
import io
import os
import sys

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from conftest import escrever_mona_lista_simples, escrever_mona_pos  # noqa: E402

import processadores  # noqa: E402
from mix_processor_1_1 import (  # noqa: E402
    COLUNAS_PADRAO, FONTES_VARIANTES, Fonte,
    formatar_excel_mix_logic_funcao_um_um, variantes_da_linha,
)

CABECALHO_MIX = ["SEQ", "MIX", "SEQ.", "POPID", "CHASSI", "MODELO", "COMPRIM",
                 "OPCIONAIS", "MU", "RHD", "GÁS", "LNG", "AUX CAB COOLER",
                 "DUPLO CIRCUITO", "COOLING COIL", "RETARDER"]


def escrever_mix(caminho, linhas, ocultas=()):
    """MIX principal no layout real (SEQ e SEQ. sao colunas diferentes)."""
    wb = Workbook()
    ws = wb.active
    ws.append(CABECALHO_MIX)
    for i, linha in enumerate(linhas, start=1):
        base = {"SEQ": 100 + i, "SEQ.": i, "CHASSI": 4000 + i,
                "MODELO": "R 460 A6X2 E6", "COMPRIM": 7125, "MU": "BR01",
                "OPCIONAIS": "COMMON RAIL/EBS"}
        base.update(linha)
        ws.append([base.get(c) for c in CABECALHO_MIX])
    for indice in ocultas:  # indice da linha de dados, a partir de 1
        ws.row_dimensions[indice + 1].hidden = True
    wb.save(caminho)


def ler(caminho):
    ws = load_workbook(caminho).active
    return [list(r) for r in ws.iter_rows(values_only=True)]


def gerar(tmp_path, linhas_mix, linhas_mona, texto=None, ocultas=(), colunas=None):
    mix, mona, saida = tmp_path / "mix.xlsx", tmp_path / "mona.xlsx", tmp_path / "out.xlsx"
    escrever_mix(mix, linhas_mix, ocultas)
    if linhas_mona and isinstance(linhas_mona[0], dict):
        escrever_mona_pos(mona, linhas_mona)
    else:
        escrever_mona_lista_simples(mona, linhas_mona)
    ok, msg = formatar_excel_mix_logic_funcao_um_um(
        str(mix), str(mona), texto, str(saida), colunas_desejadas=colunas)
    assert ok, msg
    return ler(saida)


def test_registrada_no_seletor_com_mix_e_mona_obrigatorios():
    p = processadores.obter("fa11")
    assert p is not None and p.titulo == "1.1"
    assert [c.nome for c in p.arquivos] == ["arquivo_mix", "arquivo_mona"]
    assert not any(c.opcional for c in p.arquivos)
    assert not p.aceita_outlook
    assert p.colunas_padrao == ("SEQ", "VARIANTES", "POPID", "MODELO", "MU", "RETARDER")
    ids = [x.id for x in processadores.listar()]
    assert ids.index("fa11") == ids.index("fa12") + 1


def test_colunas_na_ordem_e_valores_certos(tmp_path):
    linhas = gerar(tmp_path, [
        {"POPID": 1001, "SEQ": 58, "SEQ.": 52, "MODELO": "P 320 B4X2", "MU": "BR02",
         "RETARDER": "RETARDER"},
        {"POPID": 1002, "SEQ": 59, "SEQ.": 53, "MODELO": "R 460 A6X2", "MU": "BR01"},
    ], [{"POPID": "9999", "ITEM": "X"}])
    assert linhas[0] == ["SEQ", "VARIANTES", "POPID", "MODELO", "MU", "RETARDER"]
    # SEQ da planilha = coluna "SEQ." do MIX (sequencia da linha), nao "SEQ"
    assert linhas[1] == [52, None, 1001, "P 320 B4X2", "BR02", "RETARDER"]
    assert linhas[2] == [53, None, 1002, "R 460 A6X2", "BR01", None]


def test_variantes_do_mix_e_da_mona_numa_coluna(tmp_path):
    linhas = gerar(tmp_path, [
        {"POPID": 1001, "OPCIONAIS": "EBS/COOLING/WTA", "RHD": "RHD",
         "COOLING COIL": "COOLING COIL", "GÁS": "GÁS", "COMPRIM": 9000},
        {"POPID": 1002, "AUX CAB COOLER": "AUX CAB"},
    ], [{"POPID": "1001", "TOMADA 7 PINOS": "X"}])
    variantes = {l[2]: l[1] for l in linhas[1:]}
    # GÁS e COMPRIM longo eram da FA1.2: nao entram na 1.1
    assert variantes[1001] == "WTA - RHD - COOLING COIL | TOMADA 7 PINOS"
    assert variantes[1002] == "AUX CAB"


def test_mona_lista_simples_usa_o_texto_informado(tmp_path):
    linhas = gerar(tmp_path, [{"POPID": 1001}, {"POPID": 1002, "LNG": "LNG"}],
                   ["1002"], texto="650/649")
    variantes = {l[2]: l[1] for l in linhas[1:]}
    assert variantes[1001] is None
    assert variantes[1002] == "LNG | 650/649"


def test_linhas_ocultas_ficam_de_fora(tmp_path):
    linhas = gerar(tmp_path, [{"POPID": 1}, {"POPID": 2, "RHD": "RHD"}, {"POPID": 3}],
                   [{"POPID": "9", "ITEM": "X"}], ocultas=(2,))
    assert [l[2] for l in linhas[1:]] == [1, 3]


def test_recorte_de_colunas_mantem_a_ordem(tmp_path):
    linhas = gerar(tmp_path, [{"POPID": 1}], [{"POPID": "9", "ITEM": "X"}],
                   colunas=["MU", "POPID", "SEQ"])
    assert linhas[0] == ["SEQ", "POPID", "MU"]


def test_sem_mona_recusa(tmp_path):
    mix = tmp_path / "mix.xlsx"
    escrever_mix(mix, [{"POPID": 1}])
    ok, msg = formatar_excel_mix_logic_funcao_um_um(str(mix), None, None, str(tmp_path / "o.xlsx"))
    assert not ok and "MONA" in msg


def test_regras_de_variantes_da_1_1():
    linha = pd.Series({
        "OPCIONAIS": "COMMON RAIL/EBS/COOLING/WTA", "RHD": "RHD", "LNG": "LNG",
        "AUX CAB COOLER": "AUX CAB", "DUPLO CIRCUITO": "DUPLO",
        "COOLING COIL": "COOLING COIL", "GÁS": "GÁS", "ABS": "ABS",
    })
    assert variantes_da_linha(linha) == [
        "WTA", "RHD", "LNG", "AUX CAB", "DUPLO", "COOLING COIL"]
    assert variantes_da_linha(pd.Series({"OPCIONAIS": "EBS", "RHD": None})) == []


def test_tipos_de_fonte():
    linha = pd.Series({"CAMERA DE RÉ": "CBU", "OPCIONAIS": "EBS/abafador",
                       "COMPRIM": "8200", "LNG": None})
    fontes = (Fonte("CAMERA DE RÉ"), Fonte("LNG"),
              Fonte("OPCIONAIS", "contem", "ABAFADOR"),
              Fonte("COMPRIM", "maior_que", 8100, rotulo="LONGO"),
              Fonte("COLUNA QUE NAO EXISTE"))
    assert variantes_da_linha(linha, fontes) == ["CBU", "ABAFADOR", "LONGO"]
    assert all(f.tipo in ("valor", "contem", "maior_que") for f in FONTES_VARIANTES)


@pytest.fixture
def cliente(tmp_path, monkeypatch):
    monkeypatch.setenv("MENUMIX_DADOS", str(tmp_path))
    import armazenamento
    import filtro_mix
    importlib.reload(armazenamento)
    importlib.reload(filtro_mix)
    import app
    app = importlib.reload(app)
    yield app.app.test_client()
    monkeypatch.delenv("MENUMIX_DADOS")
    importlib.reload(armazenamento)
    importlib.reload(filtro_mix)
    importlib.reload(app)


def test_tela_mostra_os_dois_campos(cliente):
    html = cliente.get("/filtro-mix?funcao=fa11").get_data(as_text=True)
    assert 'href="/filtro-mix?funcao=fa11"' in html
    assert 'name="arquivo_mix"' in html and 'name="arquivo_mona"' in html
    assert 'name="texto_mona"' in html
    for coluna in COLUNAS_PADRAO:
        assert coluna in html


def test_envio_manual_gera_arquivo_para_download(cliente, tmp_path):
    mix, mona = tmp_path / "mix.xlsx", tmp_path / "mona.xlsx"
    escrever_mix(mix, [{"POPID": 1001}])
    escrever_mona_pos(mona, [{"POPID": "1001", "TOMADA 7 PINOS": "X"}])
    resposta = cliente.post("/filtro-mix", data={
        "acao": "manual", "funcao": "fa11", "texto_mona": "",
        "arquivo_mix": (io.BytesIO(mix.read_bytes()), "mix.xlsx"),
        "arquivo_mona": (io.BytesIO(mona.read_bytes()), "mona.xlsx"),
    }, content_type="multipart/form-data")
    html = resposta.get_data(as_text=True)
    assert "Informe:" not in html
    assert "/processados/FA11_Processado_" in html


def _cores(caminho):
    """{POPID: [preta? por coluna]} e se a fonte da celula preta e branca."""
    ws = load_workbook(caminho).active
    saida = {}
    for r in range(2, ws.max_row + 1):
        celulas = [ws.cell(r, c) for c in range(1, ws.max_column + 1)]
        pretas = [c.fill.fgColor.rgb == "FF000000" for c in celulas]
        for c, p in zip(celulas, pretas):
            if p:
                assert c.font.color is not None and c.font.color.rgb == "FFFFFFFF"
        saida[ws.cell(r, 3).value] = pretas
    return saida


def test_formatacao_variantes_e_modelos_8x2_8x4(tmp_path):
    mix, mona, saida = tmp_path / "mix.xlsx", tmp_path / "mona.xlsx", tmp_path / "out.xlsx"
    escrever_mix(mix, [
        {"POPID": 1, "COOLING COIL": "COOLING COIL", "MODELO": "R 550 A6X4 E5"},
        {"POPID": 2, "MODELO": "P 320 B8X2 E6"},
        {"POPID": 3, "MODELO": "G 560 A8X4 E6", "COOLING COIL": "COOLING COIL"},
        {"POPID": 4, "MODELO": "R 500 A6X2 E6"},
    ])
    escrever_mona_pos(mona, [{"POPID": "9", "ITEM": "X"}])
    ok, msg = formatar_excel_mix_logic_funcao_um_um(str(mix), str(mona), None, str(saida))
    assert ok, msg
    cores = _cores(saida)
    #            SEQ    VARIANTES POPID  MODELO MU     RETARDER
    assert cores[1] == [False, True, False, False, False, False]  # so VARIANTES
    assert cores[2] == [True] * 6                                 # 8X2: linha toda
    assert cores[3] == [True] * 6                                 # 8X4: linha toda
    assert cores[4] == [False] * 6                                # nada
    cabecalho = load_workbook(saida).active.cell(1, 1)
    assert cabecalho.fill.fgColor.rgb != "FF000000"


def test_mix_sem_retarder_ainda_gera(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["SEQ", "POPID", "MODELO", "MU"])
    ws.append([1, 1001, "R 460 A6X2", "BR01"])
    mix, mona, saida = tmp_path / "mix.xlsx", tmp_path / "mona.xlsx", tmp_path / "out.xlsx"
    wb.save(mix)
    escrever_mona_pos(mona, [{"POPID": "9", "ITEM": "X"}])
    ok, msg = formatar_excel_mix_logic_funcao_um_um(str(mix), str(mona), None, str(saida))
    assert ok, msg
    assert ler(saida)[1] == [1, None, 1001, "R 460 A6X2", "BR01", None]

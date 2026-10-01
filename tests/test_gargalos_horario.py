"""
Testes de gargalos.anotar_horarios: o horario estimado que aparece na regua
embaixo da faixa de gargalos (numero + horario por posicao).

Rodar com:
    pip install -r requirements-dev.txt
    pytest
"""
import gargalos
from setores import Setor


def _setor(**kwargs):
    base = dict(
        id="fa2", nome="FA2", takt_segundos=300,
        inicio_turno="07:05", fim_turno="16:05",
        pausas=[["09:00", "09:04"], ["10:55", "11:55"], ["14:00", "14:10"]],
    )
    base.update(kwargs)
    return Setor(**base)


def test_horario_da_primeira_posicao_bate_com_inicio_do_turno():
    dados = {"unidades": [{"pos": 0, "popid": "P1"}]}
    gargalos.anotar_horarios(dados, _setor(), offset=0)
    assert dados["unidades"][0]["horario"] == "07:05"


def test_offset_desloca_o_horario_pro_meio_do_dia():
    # 10 carros ja passaram (offset) + a posicao 0 do recorte = 10 carros
    # de takt (300s = 5min) depois do inicio do turno.
    dados = {"unidades": [{"pos": 0, "popid": "P1"}]}
    gargalos.anotar_horarios(dados, _setor(), offset=10)
    assert dados["unidades"][0]["horario"] == "07:55"


def test_horarios_crescem_com_a_posicao_dentro_do_recorte():
    dados = {"unidades": [{"pos": i, "popid": f"P{i}"} for i in range(3)]}
    gargalos.anotar_horarios(dados, _setor(), offset=0)
    horarios = [u["horario"] for u in dados["unidades"]]
    assert horarios == ["07:05", "07:10", "07:15"]


def test_chamada_sem_offset_comeca_no_inicio_do_turno():
    """
    E assim que app.py chama de verdade (sem offset): a posicao 0 da faixa
    tem que ser o inicio do turno configurado, nunca um horario deslocado
    pelo numero da linha dentro da planilha inteira — foi exatamente isso
    (usar `resultado.primeira_linha - 1`, um numero de linha do arquivo,
    como offset) que fazia a regua mostrar horarios de madrugada com um
    turno configurado de manha.
    """
    dados = {"unidades": [{"pos": i, "popid": f"P{i}"} for i in range(3)]}
    gargalos.anotar_horarios(dados, _setor())
    horarios = [u["horario"] for u in dados["unidades"]]
    assert horarios == ["07:05", "07:10", "07:15"]

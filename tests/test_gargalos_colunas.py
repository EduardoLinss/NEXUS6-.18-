"""
Testes do seletor de colunas do construtor de regras de gargalos.

Cobre o pedido de deixar cada funcao (setor) com sua propria lista de
colunas sugeridas no construtor, sem perder o acesso a qualquer coluna
original da planilha do mix nas condicoes das regras — mesmo as que nao
aparecem na lista sugerida de um setor.

Rodar com:
    pip install -r requirements-dev.txt
    pytest
"""
import pandas as pd

import gargalos


# ---------------------------------------------------------------------------
# Lista sugerida por setor
# ---------------------------------------------------------------------------

def test_colunas_fa2_continua_como_antes():
    assert gargalos.colunas_regra_do_setor("fa2") == gargalos.COLUNAS_REGRA_FA2
    assert "OPCIONAIS" in gargalos.COLUNAS_REGRA_FA2
    assert "RPX" in gargalos.COLUNAS_REGRA_FA2


def test_colunas_fa12_tem_lista_propria_e_independente():
    colunas_1_2 = gargalos.colunas_regra_do_setor("1-2")
    assert colunas_1_2 == gargalos.COLUNAS_REGRA_1_2
    assert colunas_1_2 is not gargalos.COLUNAS_REGRA_FA2
    # Colunas originais da planilha que a FA1.2 precisa e a FA2 nao expunha.
    assert "SORDER" in colunas_1_2
    assert "ENTRADA LINHA REAL" in colunas_1_2
    # Cobre o cabecalho completo da planilha real (todas as colunas do mix).
    assert "TPM" in colunas_1_2
    assert "GÁS" in colunas_1_2
    assert "SCR HIDDEN" in colunas_1_2
    assert len(colunas_1_2) == 56


def test_setor_desconhecido_cai_no_padrao_fa2():
    assert gargalos.colunas_regra_do_setor("outro-setor-qualquer") == gargalos.COLUNAS_REGRA_FA2
    assert gargalos.colunas_regra_do_setor(None) == gargalos.COLUNAS_REGRA_FA2


# ---------------------------------------------------------------------------
# unidades_de_dataframe expoe qualquer coluna original, nao so as fixas
# ---------------------------------------------------------------------------

def test_unidade_expoe_colunas_originais_alem_das_fixas():
    df = pd.DataFrame([
        {
            "SEQ.": "1", "POPID": "ABC123", "CHASSI": "9999999",
            "MODELO": "R450 8X4", "SUSP TRASEIRA": "Z", "COMPRIM": "9000",
            "MU": "MU01", "SORDER": "S001", "ENTRADA LINHA REAL": "TVL",
            "OPCIONAIS": "HD",
        },
    ])
    unidades = gargalos.unidades_de_dataframe(df)
    assert len(unidades) == 1
    unidade = unidades[0]

    # chaves fixas continuam funcionando (regras antigas nao quebram)
    assert unidade["MODELO"] == "R450 8X4"
    assert unidade["COMPRIM"] == "9000"

    # colunas novas (usadas pela FA1.2) ficam acessiveis pelo nome exato,
    # prontas para qualquer regra usar como condicao
    assert unidade["SORDER"] == "S001"
    assert unidade["ENTRADA LINHA REAL"] == "TVL"
    assert unidade["SEQ."] == "1"
    assert unidade["POPID"] == "ABC123"


def test_regra_unidade_critica_pode_usar_coluna_nao_listada_no_fa2():
    """
    Uma regra pode usar SORDER como condicao mesmo essa coluna nao estando
    na lista sugerida da FA2 — a lista so guia o autocompletar, nao
    restringe o que o motor consegue avaliar.
    """
    df = pd.DataFrame([
        {"MODELO": "R450 8X4", "SORDER": "PRIORIDADE"},
        {"MODELO": "R450 6X2", "SORDER": "NORMAL"},
    ])
    unidades = gargalos.unidades_de_dataframe(df)
    config = {
        "regras_unidade": [],
        "regras_padrao": [{
            "id": "SORDER_PRIORIDADE",
            "nome": "SORDER prioridade",
            "tipo": "unidade_critica",
            "severidade": "ALTA",
            "predicado": {
                "logica": "E",
                "condicoes": [{"coluna": "SORDER", "operador": "igual",
                              "valor": "PRIORIDADE"}],
            },
        }],
    }
    import analisador_cpm as motor
    resultado = motor.analisar_unidades(unidades, config)
    assert resultado["meta"]["n_gargalos"] == 1

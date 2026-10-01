"""
gargalos.py - Ponte entre a analise de MIX e o motor de gargalos.

O motor (`analisador_cpm.py`) veio pronto e nao foi alterado: ele recebe uma
lista de "unidades" e devolve as faixas de gargalo. O que este modulo faz:

  * monta as unidades a partir do MESMO recorte de DataFrame que a analise de
    MIX usa (POPID digitado + produção do dia), em vez de reler o arquivo;
  * cuida do CRUD de `regras.json` (o construtor de regras da tela).

`regras_fabrica.json` guarda as regras originais e nunca e escrito — e dele
que sai o "restaurar padrao".
"""

from __future__ import annotations

import json
import os
import logging
import re
import unicodedata

import pandas as pd

import analisador_cpm as motor
import cpm_core

from armazenamento import caminho_codigo, caminho_dados
from setores import SETOR_PADRAO

logger = logging.getLogger(__name__)

PASTA_APP = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_FABRICA = caminho_codigo("regras_fabrica.json")  # somente leitura


def arquivo_do_setor(setor: str | None = None) -> str:
    """
    Regras de um setor: dados/<setor>/regras.json.

    O arquivo antigo (dados/regras.json) vira as regras do setor padrao
    enquanto o novo nao existir, para nao perder o que ja foi criado.
    """
    setor_id = setor or SETOR_PADRAO
    novo = caminho_dados(os.path.join(setor_id, "regras.json"))
    if not os.path.exists(novo) and setor_id == SETOR_PADRAO:
        antigo = caminho_dados("regras.json")
        if os.path.exists(antigo):
            return antigo
    return novo


# Mantido para compatibilidade: aponta para o setor padrao.
ARQUIVO_REGRAS = arquivo_do_setor()

OPERADORES = {
    "contem": "contém",
    "nao_contem": "não contém",
    "igual": "é igual a",
    "diferente": "é diferente de",
    "maior_que": "é maior que",
    "maior_igual": "é maior ou igual a",
    "menor_que": "é menor que",
    "menor_igual": "é menor ou igual a",
    "vazio": "está vazio",
    "nao_vazio": "não está vazio",
}
OPERADORES_SEM_VALOR = ("vazio", "nao_vazio")
# Comparam numero, nao texto — e o caso de COMPRIM > 7400.
OPERADORES_NUMERICOS = ("maior_que", "maior_igual", "menor_que", "menor_igual")

TIPOS = {
    "unidade_critica": "o produto sozinho já é gargalo (variante de alta complexidade)",
    "contagem_janela": "N ocorrências dentro de uma janela de X carros",
    "sequencia_consecutiva": "K ou mais carros consecutivos",
    "coocorrencia_janela": "dois grupos coincidindo dentro de uma janela",
}

SEVERIDADES = ("ALTA", "MEDIA")

MAX_CONDICOES = 3
MAX_GRUPOS = 2

# Colunas que o motor enxerga em cada unidade (as chaves em MAIUSCULA de
# `carrega`, mais as de exibicao).
#
# Cada funcao tem seu proprio seletor de colunas no construtor de regras —
# igual ao padrao ja usado em COLUNAS_PADRAO (mix_processor_1_2.py) vs
# COLUNAS_DESEJADAS_PADRAO (filtro_mix.py): listas independentes, nao uma
# so compartilhada. `unidades_de_dataframe` ja deixa QUALQUER coluna da
# planilha original disponivel para o predicado (pelo nome exato dela); as
# listas abaixo so controlam o que aparece como sugestao no seletor de cada
# tela — nao limitam o que a regra pode de fato usar.
COLUNAS_REGRA_FA2 = [
    "MODELO", "OPCIONAIS", "RPX", "BARRA TRASEIRA", "SUSP TRASEIRA",
    "COMPRIM", "MU", "CHASSI",
]

# FA1.2: todas as colunas originais da planilha geral do mix (a mesma que a
# FA2 recebe por Outlook/upload na tela de Analise) — a FA1.2 nao fica mais
# restrita ao subconjunto pensado para a FA2. Lista completa passada pelo
# usuario (cabecalho real da planilha de mix).
COLUNAS_REGRA_1_2 = [
    "SCOB", "SEQ", "MIX", "SEQ.", "REPLAN", "ENTRADA LINHA REAL", "POPID",
    "CHASSI", "MODELO", "CHASSIS HEIGHT", "SUSP TRASEIRA", "EOV", "COMPRIM",
    "OPCIONAIS", "MU", "RHD", "COR", "DEF", "LATERAL", "SORDER", "FFU",
    "RPX", "TANQUE", "GÁS", "LNG", "BEHIND CAB", "DW", "FOU", "RETARDER",
    "ABS", "HIGH LINE", "CAB MOD", "COD POPID", "COD CHASSIS",
    "BARRA 2°E DIANT", "BARRA TRASEIRA", "SMAC", "AUX CAB COOLER", "LEVEL",
    "SESSAM7", "POSIÇÃO DO TANQUE", "CAMERA DE RÉ", "POSIÇÃO BATERIA",
    "SCR", "ETIQUETA ECE", "ETIQUETA CO²", "SCR HIDDEN",
    "PARACHOQUE TRASEIRO", "DUTO DE AR", "TETO TEMPORARIO", "LITRAGEM",
    "DUPLO CIRCUITO", "COOLING COIL", "CABINE 14", "SEM 5 RODA", "TPM",
]

# Mantido para compatibilidade com quem ainda importa o nome antigo.
COLUNAS_REGRA = COLUNAS_REGRA_FA2


def colunas_regra_do_setor(setor: str | None = None) -> list[str]:
    """Lista de colunas sugeridas no seletor do construtor de regras, por setor."""
    setor_id = setor or SETOR_PADRAO
    if setor_id == "1-2":
        return COLUNAS_REGRA_1_2
    return COLUNAS_REGRA_FA2


class ErroRegra(ValueError):
    """Erro de validacao mostrado na tela do construtor de regras."""


# ---------------------------------------------------------------------------
# Unidades a partir do DataFrame ja carregado
# ---------------------------------------------------------------------------
def _texto(valor) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    if isinstance(valor, float) and float(valor).is_integer():
        return str(int(valor))
    return str(valor).strip()


def _coluna(linha, *nomes) -> str:
    for nome in nomes:
        if nome in linha.index:
            return _texto(linha[nome])
    return ""


def unidades_de_dataframe(df: pd.DataFrame) -> list[dict]:
    """
    Converte o recorte analisado nas unidades que o motor espera.

    Mesmo mapeamento de `analisador_cpm.carrega`, so que partindo de um
    DataFrame ja em memoria — assim os gargalos saem exatamente do mesmo
    intervalo de POPID/produção que a carga por variante.
    """
    ja_formatado = "SEQ." in df.columns and "BARRA TRASEIRA" in df.columns

    unidades = []
    for _, linha in df.iterrows():
        modelo = _coluna(linha, "MODELO")
        if not modelo:
            continue

        opcionais = _coluna(linha, "OPCIONAIS")
        opc_fmt = opcionais if ja_formatado else motor._filtrar_opcionais(opcionais)

        unidade = {
            "seq": _coluna(linha, "SEQ.", "SEQ", "SEQUENCIA"),
            "popid": _coluna(linha, "POPID"),
            "chassi": _coluna(linha, "CHASSI"),
            "modelo": modelo,
            "comprim": _coluna(linha, "COMPRIM"),
            "susp": _coluna(linha, "SUSP TRASEIRA"),
            "mu": _coluna(linha, "MU"),
            "rpx": _coluna(linha, "RPX"),
            "opc_fmt": opc_fmt,
            "barra2e": _coluna(linha, "BARRA 2°E DIANT"),
            "barra_tras": _coluna(linha, "BARRA TRASEIRA"),
            # chaves que os predicados das regras usam
            "MODELO": modelo,
            "OPCIONAIS": opcionais,
            "RPX": _coluna(linha, "RPX"),
            "BARRA TRASEIRA": _coluna(linha, "BARRA TRASEIRA"),
            "SUSP TRASEIRA": _coluna(linha, "SUSP TRASEIRA"),
            "COMPRIM": _coluna(linha, "COMPRIM"),
            "MU": _coluna(linha, "MU"),
            "CHASSI": _coluna(linha, "CHASSI"),
        }

        # Qualquer outra coluna da planilha original (SORDER, ENTRADA LINHA
        # REAL, SEQ., REPLAN, POPID, etc.) tambem fica disponivel para o
        # predicado das regras, com a chave exatamente como aparece na
        # planilha — assim cada funcao pode montar regras usando as
        # colunas que fazem sentido para ela, sem depender de uma lista
        # fixa aqui no motor.
        for coluna in df.columns:
            nome = str(coluna).strip()
            if nome and nome not in unidade:
                unidade[nome] = _texto(linha[coluna])

        unidades.append(unidade)

    return unidades


def analisar(df: pd.DataFrame, setor: str | None = None) -> dict:
    """Roda o motor sobre o recorte e devolve o dicionario de analise."""
    unidades = unidades_de_dataframe(df)
    if not unidades:
        return {"meta": {"n_unidades": 0, "n_gargalos": 0},
                "regras": [], "unidades": [], "gargalos": [], "resumo": {}}

    dados = motor.analisar_unidades(unidades, carregar_config(setor))
    dados["resumo"] = _resumir(dados)
    return dados


def _resumir(dados: dict) -> dict:
    gargalos = dados.get("gargalos", [])
    por_regra: dict[str, int] = {}
    for gargalo in gargalos:
        por_regra[gargalo["regra_nome"]] = por_regra.get(gargalo["regra_nome"], 0) + 1

    return {
        "total": len(gargalos),
        "alta": sum(1 for g in gargalos if g["severidade"] == "ALTA"),
        "media": sum(1 for g in gargalos if g["severidade"] == "MEDIA"),
        "unidades_afetadas": sum(
            1 for u in dados.get("unidades", []) if u.get("heat") != "NENHUM"
        ),
        "por_regra": por_regra,
    }


def descrever_membros(dados: dict, gargalo: dict) -> list[dict]:
    """POPID e SEQ das unidades que compoem o gargalo."""
    por_pos = {u["pos"]: u for u in dados.get("unidades", [])}
    membros = []
    for posicao in gargalo.get("membros", []):
        unidade = por_pos.get(posicao)
        if unidade:
            membros.append({
                "popid": unidade.get("popid", ""),
                "seq": unidade.get("seq", ""),
                "modelo": unidade.get("modelo", ""),
                "pos": posicao + 1,
            })
    return membros


def anotar_horarios(dados: dict, setor, offset: int = 0) -> dict:
    """
    Preenche 'horario' (HH:MM estimado) em cada unidade da faixa, usando o
    turno e o takt do setor.

    A posicao 0 da faixa sempre corresponde ao inicio do turno configurado
    (`offset=0`, o padrao) — cada faixa mostra um recorte novo da producao,
    entao o horario e sempre relativo ao comeco desse recorte, nunca ao
    numero da linha dentro da planilha inteira. `resultado.primeira_linha`
    NAO deve ser usado como offset aqui: ele conta linhas desde o topo do
    arquivo carregado (que pode ter varios turnos/dias empilhados), nao
    veiculos desde o inicio do turno — usa-lo faz o horario derrapar para
    muito depois do `fim_turno` e virar a meia-noite. `offset` so existe
    para quem realmente sabe quantos veiculos deste MESMO turno ja ficaram
    para tras antes do recorte.
    """
    for unidade in dados.get("unidades", []):
        unidade["horario"] = cpm_core.horario_da_posicao(
            offset + unidade.get("pos", 0),
            setor.inicio_turno, setor.takt_segundos, setor.fim_turno,
            setor.pausas_tuplas,
        )
    return dados


# ---------------------------------------------------------------------------
# regras.json
# ---------------------------------------------------------------------------
def carregar_config(setor: str | None = None) -> dict:
    for caminho in (arquivo_do_setor(setor), ARQUIVO_FABRICA):
        try:
            with open(caminho, encoding="utf-8") as arquivo:
                config = json.load(arquivo)
            config.setdefault("regras_unidade", [])
            config.setdefault("regras_padrao", [])
            return config
        except FileNotFoundError:
            continue
        except json.JSONDecodeError as erro:
            logger.warning("%s inválido (%s)", os.path.basename(caminho), erro)
    return {"config": {}, "regras_unidade": [], "regras_padrao": []}


def salvar_config(config: dict, setor: str | None = None) -> dict:
    caminho = arquivo_do_setor(setor)
    os.makedirs(os.path.dirname(caminho), exist_ok=True)
    temporario = caminho + ".tmp"
    with open(temporario, "w", encoding="utf-8") as arquivo:
        json.dump(config, arquivo, ensure_ascii=False, indent=1)
    os.replace(temporario, caminho)
    return config


def listar_regras(setor: str | None = None) -> list[dict]:
    return carregar_config(setor).get("regras_padrao", [])


def encontrar_regra(regra_id: str, setor: str | None = None) -> dict | None:
    for regra in listar_regras(setor):
        if regra["id"] == regra_id:
            return regra
    return None


def exportar_json(setor: str | None = None) -> str:
    """Mesmo conteudo de regras.json — serve de backup e de importacao."""
    return json.dumps(carregar_config(setor), ensure_ascii=False, indent=1)


def exportar_tsv(setor: str | None = None) -> str:
    """Regras de um setor em colunas, para abrir no Excel (nao reimporta)."""
    linhas = ["Regra\tId\tTipo\tSeveridade\tCondicoes"]
    for regra in listar_regras(setor):
        linhas.append("\t".join([
            regra.get("nome", ""), regra.get("id", ""),
            TIPOS.get(regra.get("tipo", ""), regra.get("tipo", "")),
            regra.get("severidade", ""), descrever_regra(regra),
        ]))
    return "\n".join(linhas)


def exportar_texto(setor: str | None = None) -> str:
    partes = []
    for regra in listar_regras(setor):
        partes.append(f"[{regra.get('severidade', '')}] {regra.get('nome', '')}")
        partes.append(f"    {descrever_regra(regra)}")
    return "\n".join(partes) or "(sem regras)"


def exportar_tsv_todos(setores: list) -> str:
    """Regras de todos os setores num arquivo so, com o setor na primeira coluna."""
    linhas = ["Setor\tCodigo do setor\tRegra\tId\tTipo\tSeveridade\tCondicoes"]
    for setor in setores:
        regras = listar_regras(setor.id)
        if not regras:
            linhas.append(f"{setor.nome}\t{setor.id}\t(sem regras)\t\t\t\t")
            continue
        for regra in regras:
            linhas.append("\t".join([
                setor.nome, setor.id, regra.get("nome", ""), regra.get("id", ""),
                TIPOS.get(regra.get("tipo", ""), regra.get("tipo", "")),
                regra.get("severidade", ""), descrever_regra(regra),
            ]))
    return "\n".join(linhas)


def _erro_de_formato(conteudo: str, erro) -> str:
    """
    Mensagem util quando o arquivo nao e JSON.

    O caso comum e enviar o .tsv, que existe para abrir no Excel e nao para
    reimportar. Dizer so "JSON inválido" nao ajuda ninguem a se achar.
    """
    inicio = (conteudo or "").lstrip()[:400]
    primeira = inicio.split("\n", 1)[0]

    if "\t" in primeira:
        return ("Esse é o arquivo .tsv, feito para abrir no Excel. Para "
                "importar, use o .json — é ele que guarda as regras inteiras.")
    if primeira.lower().startswith(("setor", "variante", "=== ")):
        return ("Esse arquivo é a exportação em texto/planilha, que serve para "
                "leitura. Para importar, use o .json.")
    if not inicio:
        return "O arquivo está vazio."
    return (f"Arquivo não é um JSON válido ({erro}). Use um arquivo gerado "
            f"pelo botão 'baixar .json'.")


def exportar_json_todos(setores: list) -> str:
    """Backup das regras de todos os setores, um bloco por setor."""
    conteudo = {"setores": {s.id: carregar_config(s.id) for s in setores}}
    return json.dumps(conteudo, ensure_ascii=False, indent=1)


def exportar_texto_todos(setores: list) -> str:
    partes = []
    for setor in setores:
        partes.append(f"=== {setor.nome} ({setor.id}) ===")
        regras = listar_regras(setor.id)
        if not regras:
            partes.append("  (sem regras)")
        for regra in regras:
            partes.append(f"  [{regra.get('severidade', '')}] {regra.get('nome', '')}")
            partes.append(f"      {descrever_regra(regra)}")
        partes.append("")
    return "\n".join(partes)


def _copia_de_seguranca(setor: str | None = None) -> str | None:
    """Guarda as regras atuais antes de sobrescrever."""
    caminho = arquivo_do_setor(setor)
    if not os.path.exists(caminho):
        return None
    from datetime import datetime
    setor_id = setor or SETOR_PADRAO
    nome = f"regras_backup_{setor_id}_{datetime.now():%Y%m%d_%H%M%S}.json"
    destino = caminho_dados(nome)
    with open(caminho, encoding="utf-8") as origem, \
            open(destino, "w", encoding="utf-8") as copia:
        copia.write(origem.read())
    return nome


def importar(conteudo: str, setor: str | None = None) -> tuple[list[dict], str | None]:
    """
    Substitui as regras pelo conteudo de um .json exportado.

    Valida antes de gravar e guarda uma copia do que estava no ar.
    """
    try:
        dados = json.loads(conteudo)
    except json.JSONDecodeError as erro:
        raise ErroRegra(_erro_de_formato(conteudo, erro))

    # Aceita tambem o arquivo da fabrica inteira ({"setores": {id: config}}),
    # pegando o bloco do setor em edicao.
    if isinstance(dados, dict) and "setores" in dados and "regras_padrao" not in dados:
        blocos = dados.get("setores") or {}
        alvo = setor or SETOR_PADRAO
        if alvo not in blocos:
            disponiveis = ", ".join(sorted(blocos)) or "nenhum"
            raise ErroRegra(
                f"Este arquivo tem as regras da fábrica, mas não do setor "
                f"'{alvo}'. Setores no arquivo: {disponiveis}."
            )
        dados = blocos[alvo]

    if not isinstance(dados, dict) or "regras_padrao" not in dados:
        raise ErroRegra(
            "Formato não reconhecido: falta a chave 'regras_padrao'. Use um "
            "arquivo gerado pelo botão 'baixar .json' desta tela."
        )

    regras = dados.get("regras_padrao") or []
    if not isinstance(regras, list) or not regras:
        raise ErroRegra("O arquivo não tem nenhuma regra.")

    # valida cada regra montando o criterio, como o motor faria
    for posicao, regra in enumerate(regras, 1):
        rotulo = regra.get("nome") or f"regra {posicao}"
        if regra.get("tipo") not in TIPOS:
            raise ErroRegra(f"{rotulo}: tipo desconhecido ({regra.get('tipo')}).")
        if not regra.get("id"):
            raise ErroRegra(f"{rotulo}: falta o campo 'id'.")
        if regra["tipo"] == "coocorrencia_janela":
            if len(regra.get("grupos") or []) < 2:
                raise ErroRegra(f"{rotulo}: coocorrência precisa de dois grupos.")
        elif not (regra.get("predicado") or {}).get("condicoes"):
            raise ErroRegra(f"{rotulo}: sem condições.")

    config = carregar_config(setor)
    config["regras_padrao"] = regras
    if "config" in dados:
        config["config"] = dados["config"]
    if "regras_unidade" in dados:
        config["regras_unidade"] = dados["regras_unidade"]

    backup = _copia_de_seguranca(setor)
    salvar_config(config, setor)
    return regras, backup


def restaurar_padrao(setor: str | None = None) -> dict:
    with open(ARQUIVO_FABRICA, encoding="utf-8") as arquivo:
        return salvar_config(json.load(arquivo), setor)


def remover_regra(regra_id: str, setor: str | None = None) -> dict:
    config = carregar_config(setor)
    antes = len(config["regras_padrao"])
    config["regras_padrao"] = [r for r in config["regras_padrao"]
                               if r["id"] != regra_id]
    if len(config["regras_padrao"]) == antes:
        raise ErroRegra("Regra não encontrada.")
    return salvar_config(config, setor)


def _id_a_partir_do_nome(nome: str, existentes: set[str]) -> str:
    base = unicodedata.normalize("NFD", nome.upper())
    base = "".join(c for c in base if unicodedata.category(c) != "Mn")
    base = re.sub(r"[^A-Z0-9]+", "_", base).strip("_") or "REGRA"
    base = base[:40]

    candidato, n = base, 2
    while candidato in existentes:
        candidato = f"{base}_{n}"
        n += 1
    return candidato


def _predicado_de_condicoes(condicoes: list[dict], logica: str,
                            rotulo: str = "condição") -> dict:
    limpas = []
    for condicao in condicoes:
        coluna = (condicao.get("coluna") or "").strip()
        operador = (condicao.get("operador") or "contem").strip()
        valor = (condicao.get("valor") or "").strip()

        if not coluna:
            continue
        if operador not in OPERADORES:
            raise ErroRegra(f"Operador desconhecido: {operador}")
        if operador in OPERADORES_NUMERICOS:
            try:
                float(valor.replace(",", "."))
            except ValueError:
                raise ErroRegra(
                    f"'{valor}' não é um número — a {rotulo} em '{coluna}' "
                    f"{OPERADORES[operador]} precisa de um valor numérico."
                )
        if operador not in OPERADORES_SEM_VALOR and not valor:
            raise ErroRegra(
                f"A {rotulo} em '{coluna}' precisa de um valor "
                f"(ou use 'está vazio' / 'não está vazio')."
            )
        limpas.append({"coluna": coluna, "operador": operador, "valor": valor})

    if not limpas:
        raise ErroRegra(f"Informe pelo menos uma {rotulo} (coluna e valor).")

    logica = (logica or "E").upper()
    if logica not in ("E", "OU"):
        raise ErroRegra("A lógica deve ser E ou OU.")
    return {"logica": logica, "condicoes": limpas}


def _inteiro(valor, rotulo: str, minimo: int = 1) -> int:
    try:
        numero = int(str(valor).strip())
    except (TypeError, ValueError):
        raise ErroRegra(f"{rotulo} deve ser um número inteiro.")
    if numero < minimo:
        raise ErroRegra(f"{rotulo} deve ser {minimo} ou mais.")
    return numero


def salvar_regra(nome: str, tipo: str, severidade: str, parametros: dict,
                 regra_id: str | None = None, setor: str | None = None) -> dict:
    """Cria ou atualiza uma regra de gargalo. `parametros` vem do formulário."""
    nome = (nome or "").strip()
    if not nome:
        raise ErroRegra("O nome da regra é obrigatório.")
    if tipo not in TIPOS:
        raise ErroRegra(f"Tipo de regra desconhecido: {tipo}")
    severidade = (severidade or "ALTA").upper()
    if severidade not in SEVERIDADES:
        raise ErroRegra(f"Severidade desconhecida: {severidade}")

    regra: dict = {"id": regra_id or "", "nome": nome, "tipo": tipo,
                   "severidade": severidade}

    if tipo == "unidade_critica":
        # Sem janela nem minimo: cada unidade que bate ja e o gargalo.
        regra["predicado"] = _predicado_de_condicoes(
            parametros.get("condicoes", []), parametros.get("logica", "E"))

    elif tipo == "contagem_janela":
        regra["janela"] = _inteiro(parametros.get("janela"), "A janela", 2)
        regra["min_ocorrencias"] = _inteiro(
            parametros.get("min_ocorrencias"), "O mínimo de ocorrências", 2)
        if regra["min_ocorrencias"] > regra["janela"]:
            raise ErroRegra(
                f"O mínimo de ocorrências ({regra['min_ocorrencias']}) não pode "
                f"ser maior que a janela ({regra['janela']})."
            )
        regra["predicado"] = _predicado_de_condicoes(
            parametros.get("condicoes", []), parametros.get("logica", "E"))

    elif tipo == "sequencia_consecutiva":
        regra["min_seq"] = _inteiro(parametros.get("min_seq"),
                                    "O mínimo de consecutivos", 2)
        regra["predicado"] = _predicado_de_condicoes(
            parametros.get("condicoes", []), parametros.get("logica", "E"))

    else:  # coocorrencia_janela
        regra["janela"] = _inteiro(parametros.get("janela"), "A janela", 2)
        grupos = []
        for indice, bruto in enumerate(parametros.get("grupos", []), 1):
            condicoes = bruto.get("condicoes", [])
            if not any((c.get("coluna") or "").strip() for c in condicoes):
                continue
            grupos.append(_predicado_de_condicoes(
                condicoes, bruto.get("logica", "E"), f"condição do grupo {indice}"))
        if len(grupos) < 2:
            raise ErroRegra(
                "A coocorrência precisa de dois grupos preenchidos — é o "
                "encontro deles na janela que caracteriza o gargalo."
            )
        regra["grupos"] = grupos

    config = carregar_config(setor)
    existentes = {r["id"] for r in config["regras_padrao"]}
    if not regra["id"]:
        regra["id"] = _id_a_partir_do_nome(nome, existentes)

    for posicao, atual in enumerate(config["regras_padrao"]):
        if atual["id"] == regra["id"]:
            config["regras_padrao"][posicao] = regra
            break
    else:
        config["regras_padrao"].append(regra)

    salvar_config(config, setor)
    return regra


# ---------------------------------------------------------------------------
# Descricao legivel (usada na listagem e na exportacao)
# ---------------------------------------------------------------------------
def descrever_predicado(predicado: dict) -> str:
    juncao = " ou " if predicado.get("logica", "E").upper() == "OU" else " e "
    partes = []
    for condicao in predicado.get("condicoes", []):
        rotulo = OPERADORES.get(condicao["operador"], condicao["operador"])
        if condicao["operador"] in OPERADORES_SEM_VALOR:
            partes.append(f"{condicao['coluna']} {rotulo}")
        else:
            partes.append(f"{condicao['coluna']} {rotulo} {condicao['valor']}")
    return juncao.join(partes)


def descrever_regra(regra: dict) -> str:
    if regra["tipo"] == "unidade_critica":
        return f"cada carro onde {descrever_predicado(regra['predicado'])}"
    if regra["tipo"] == "contagem_janela":
        return (f"{regra['min_ocorrencias']}+ carros em {regra['janela']} "
                f"onde {descrever_predicado(regra['predicado'])}")
    if regra["tipo"] == "sequencia_consecutiva":
        return (f"{regra['min_seq']}+ consecutivos onde "
                f"{descrever_predicado(regra['predicado'])}")
    grupos = " · com · ".join(descrever_predicado(g) for g in regra.get("grupos", []))
    return f"em {regra['janela']} carros: {grupos}"

"""
catalogo.py - Catalogo de tarefas por variante, gravado em `catalogo.json`.

Tira o catalogo de dentro do codigo. Sem esse arquivo, `cpm_core` continua
usando o catalogo embutido (`CATALOGO_PADRAO`); na primeira gravacao o JSON
nasce com esse mesmo conteudo, entao nada se perde.

Formato:

    {
      "variantes": [
        {
          "id": "V1",
          "nome": "Variante 1",
          "tarefas": [
            {
              "id": "a1b2c3d4",
              "nome": "Haste traseira",
              "tempo": 200,
              "chave_legado": "HASTE TRASEIRA v1",
              "condicoes": [
                {"coluna": "SUSP TRASEIRA", "operador": "contem", "valor": "Z"}
              ]
            }
          ]
        }
      ]
    }

Todas as condicoes de uma tarefa sao combinadas com E (todas precisam valer),
que e como as tarefas originais ja funcionavam.
"""

from __future__ import annotations

import json
import os
import uuid
from typing import Iterable, Sequence

from cpm_core import (
    CATALOGO_PADRAO, Contem, Criterio, E, EmJanela, EmSequencia, MaiorQue,
    MenorQue, NaoContem, Tarefa, Variante,
)

from armazenamento import caminho_dados
from setores import SETOR_PADRAO, obter as obter_setor

PASTA = os.path.dirname(os.path.abspath(__file__))


def arquivo_do_setor(setor: str | None = None) -> str:
    """
    Catalogo de um setor: dados/<setor>/catalogo.json.

    O arquivo antigo (dados/catalogo.json, de quando havia um setor so) e
    adotado como catalogo do setor padrao enquanto o novo nao existir — assim
    quem ja tinha tarefas criadas nao perde nada ao atualizar.
    """
    setor_id = setor or SETOR_PADRAO
    novo = caminho_dados(os.path.join(setor_id, "catalogo.json"))
    if not os.path.exists(novo) and setor_id == SETOR_PADRAO:
        antigo = caminho_dados("catalogo.json")
        if os.path.exists(antigo):
            return antigo
    return novo


# Mantido para compatibilidade: aponta para o setor padrao.
ARQUIVO = arquivo_do_setor()

OPERADORES = {
    "contem": "contém o texto",
    "nao_contem": "NÃO contém o texto",
    "maior_que": "é maior que",
    "menor_que": "é menor que",
}

OPERADORES_NUMERICOS = ("maior_que", "menor_que")

MODOS_SEQUENCIA = {
    "primeiro": "a variante atua 1 vez por grupo",
    "todos": "conta todos os veículos do grupo",
    "excedente": "conta só os veículos que passam do mínimo",
}

# Versao curta, para caber nas exportacoes
OPERADORES_CURTOS = {
    "contem": "contém",
    "nao_contem": "não contém",
    "maior_que": ">",
    "menor_que": "<",
}


class ErroCatalogo(ValueError):
    pass


# ---------------------------------------------------------------------------
# Conversao entre criterios e dicionarios
# ---------------------------------------------------------------------------
def sequencia_de_criterio(criterio: Criterio) -> dict | None:
    """
    Devolve o agrupamento da tarefa, ou None se ela conta veiculo a veiculo.

    Sem 'janela' -> veiculos colados (EmSequencia).
    Com 'janela' -> N dentro de uma janela de X (EmJanela).
    """
    if isinstance(criterio, EmSequencia):
        return {"minimo": criterio.minimo, "modo": criterio.modo, "janela": None}
    if isinstance(criterio, EmJanela):
        return {"minimo": criterio.minimo, "modo": criterio.modo,
                "janela": criterio.janela}
    return None


def condicoes_de_criterio(criterio: Criterio) -> list[dict]:
    """Achata um criterio em uma lista de condicoes simples."""
    if isinstance(criterio, (EmSequencia, EmJanela)):
        return condicoes_de_criterio(criterio.criterio)
    if isinstance(criterio, E):
        condicoes = []
        for parte in criterio.criterios:
            condicoes.extend(condicoes_de_criterio(parte))
        return condicoes
    if isinstance(criterio, Contem):
        return [{"coluna": criterio.coluna, "operador": "contem", "valor": criterio.valor}]
    if isinstance(criterio, NaoContem):
        return [{"coluna": criterio.coluna, "operador": "nao_contem",
                 "valor": criterio.valor}]
    if isinstance(criterio, MaiorQue):
        return [{"coluna": criterio.coluna, "operador": "maior_que",
                 "valor": ("%g" % criterio.limite)}]
    if isinstance(criterio, MenorQue):
        return [{"coluna": criterio.coluna, "operador": "menor_que",
                 "valor": ("%g" % criterio.limite)}]
    raise ErroCatalogo(f"Critério não suportado: {type(criterio).__name__}")


def criterio_de_condicoes(condicoes: Sequence[dict],
                         sequencia: dict | None = None) -> Criterio:
    partes = []
    for condicao in condicoes:
        coluna = (condicao.get("coluna") or "").strip()
        operador = (condicao.get("operador") or "contem").strip()
        valor = str(condicao.get("valor") or "").strip()

        if not coluna or not valor:
            continue
        if operador == "contem":
            partes.append(Contem(coluna, valor))
        elif operador == "nao_contem":
            partes.append(NaoContem(coluna, valor))
        elif operador in OPERADORES_NUMERICOS:
            try:
                limite = float(valor.replace(",", "."))
            except ValueError:
                raise ErroCatalogo(
                    f"'{valor}' não é um número válido para a condição "
                    f"'{coluna} {OPERADORES[operador]}'."
                )
            classe = MaiorQue if operador == "maior_que" else MenorQue
            partes.append(classe(coluna, limite))
        else:
            raise ErroCatalogo(f"Operador desconhecido: {operador}")

    if not partes:
        raise ErroCatalogo("Informe pelo menos uma condição (coluna e valor).")

    criterio = partes[0] if len(partes) == 1 else E(*partes)

    if sequencia:
        try:
            minimo = int(sequencia.get("minimo") or 0)
        except (TypeError, ValueError):
            raise ErroCatalogo("O mínimo da sequência deve ser um número inteiro.")
        janela_bruta = sequencia.get("janela")
        try:
            janela = int(janela_bruta) if str(janela_bruta or "").strip() else 0
        except (TypeError, ValueError):
            raise ErroCatalogo("A janela deve ser um número inteiro de veículos.")

        if minimo > 1:
            modo = sequencia.get("modo") or "primeiro"
            if modo not in MODOS_SEQUENCIA:
                raise ErroCatalogo(f"Modo de agrupamento desconhecido: {modo}")
            if janela:
                if janela < minimo:
                    raise ErroCatalogo(
                        f"A janela ({janela}) não pode ser menor que o mínimo "
                        f"({minimo})."
                    )
                criterio = EmJanela(criterio, janela=janela, minimo=minimo, modo=modo)
            else:
                criterio = EmSequencia(criterio, minimo=minimo, modo=modo)
        elif minimo == 1:
            raise ErroCatalogo(
                "Sequência de 1 é o mesmo que nenhuma sequência — deixe o campo "
                "vazio ou use 2 ou mais."
            )

    return criterio


# ---------------------------------------------------------------------------
# Leitura e gravacao
# ---------------------------------------------------------------------------
def novo_id() -> str:
    return uuid.uuid4().hex[:8]


def _id_da_tarefa(tarefa: Tarefa, variante_id: str) -> str:
    if tarefa.id:
        return tarefa.id
    base = tarefa.chave_legado or f"{variante_id}-{tarefa.nome}"
    return uuid.uuid5(uuid.NAMESPACE_OID, base).hex[:8]


def _tarefa_para_dict(tarefa: Tarefa, variante_id: str) -> dict:
    bruto = {
        "id": _id_da_tarefa(tarefa, variante_id),
        "nome": tarefa.nome,
        "tempo": tarefa.tempo,
        "chave_legado": tarefa.chave_legado,
        "condicoes": condicoes_de_criterio(tarefa.criterio),
    }
    sequencia = sequencia_de_criterio(tarefa.criterio)
    if sequencia:
        bruto["sequencia"] = sequencia
    return bruto


def para_dict(variantes: Iterable[Variante]) -> dict:
    saida = {"variantes": []}
    for variante in variantes:
        saida["variantes"].append({
            "id": variante.id,
            "nome": variante.nome,
            "tarefas": [_tarefa_para_dict(tarefa, variante.id)
                        for tarefa in variante.tarefas],
        })
    return saida


def de_dict(dados: dict) -> tuple[Variante, ...]:
    variantes = []
    for bruta in dados.get("variantes", []):
        tarefas = []
        for bruta_tarefa in bruta.get("tarefas", []):
            tarefas.append(Tarefa(
                nome=bruta_tarefa["nome"],
                tempo=float(bruta_tarefa["tempo"]),
                criterio=criterio_de_condicoes(
                    bruta_tarefa.get("condicoes", []),
                    bruta_tarefa.get("sequencia"),
                ),
                chave_legado=bruta_tarefa.get("chave_legado", ""),
                id=bruta_tarefa.get("id") or novo_id(),
            ))
        variantes.append(Variante(
            id=bruta["id"], nome=bruta.get("nome", bruta["id"]), tarefas=tuple(tarefas)
        ))
    return tuple(variantes)


def carregar(setor: str | None = None) -> tuple[Variante, ...]:
    """Le o catalogo do setor; se nao existir, devolve o catalogo embutido."""
    caminho = arquivo_do_setor(setor)
    if not os.path.exists(caminho):
        return de_dict(para_dict(CATALOGO_PADRAO))  # normaliza os ids
    with open(caminho, encoding="utf-8") as arquivo:
        return de_dict(json.load(arquivo))


def salvar(variantes: Iterable[Variante], setor: str | None = None) -> None:
    """Grava de forma atomica, para nao corromper o arquivo se algo falhar."""
    caminho = arquivo_do_setor(setor)
    os.makedirs(os.path.dirname(caminho), exist_ok=True)
    temporario = caminho + ".tmp"
    with open(temporario, "w", encoding="utf-8") as arquivo:
        json.dump(para_dict(variantes), arquivo, ensure_ascii=False, indent=2)
    os.replace(temporario, caminho)


def _copia_de_seguranca(setor: str | None = None) -> str | None:
    """Guarda o catalogo atual antes de sobrescrever. Devolve o nome do arquivo."""
    caminho = arquivo_do_setor(setor)
    if not os.path.exists(caminho):
        return None
    from datetime import datetime
    setor_id = setor or SETOR_PADRAO
    nome = f"catalogo_backup_{setor_id}_{datetime.now():%Y%m%d_%H%M%S}.json"
    destino = caminho_dados(nome)
    with open(caminho, encoding="utf-8") as origem, \
            open(destino, "w", encoding="utf-8") as copia:
        copia.write(origem.read())
    return nome


def importar(conteudo: str, setor: str | None = None) -> tuple[tuple[Variante, ...], str | None]:
    """
    Substitui o catalogo pelo conteudo de um arquivo .json exportado.

    O arquivo e validado ANTES de gravar: se estiver quebrado, nada muda. Uma
    copia do catalogo atual fica na pasta de dados, para o import poder ser
    desfeito.
    """
    try:
        dados = json.loads(conteudo)
    except json.JSONDecodeError as erro:
        # O engano comum e enviar o .tsv, que existe para abrir no Excel.
        primeira = (conteudo or "").lstrip().split("\n", 1)[0]
        if "\t" in primeira:
            raise ErroCatalogo(
                "Esse é o arquivo .tsv, feito para abrir no Excel. Para "
                "importar, use o .json — é ele que guarda o catálogo inteiro."
            )
        if not (conteudo or "").strip():
            raise ErroCatalogo("O arquivo está vazio.")
        raise ErroCatalogo(
            f"Arquivo não é um JSON válido ({erro}). Use um arquivo gerado "
            f"pelo botão 'baixar .json'."
        )

    if not isinstance(dados, dict) or "variantes" not in dados:
        raise ErroCatalogo(
            "Formato não reconhecido: falta a chave 'variantes'. Use um arquivo "
            "gerado pelo botão 'baixar .json' desta tela."
        )

    try:
        variantes = de_dict(dados)
    except ErroCatalogo:
        raise
    except Exception as erro:  # noqa: BLE001
        raise ErroCatalogo(f"Não consegui interpretar o arquivo: {erro}")

    if not variantes:
        raise ErroCatalogo("O arquivo não tem nenhuma variante.")

    backup = _copia_de_seguranca(setor)
    salvar(variantes, setor)
    return variantes, backup


def restaurar_padrao(setor: str | None = None) -> tuple[Variante, ...]:
    variantes = de_dict(para_dict(CATALOGO_PADRAO))
    salvar(variantes, setor)
    return variantes


# ---------------------------------------------------------------------------
# Operacoes de edicao
# ---------------------------------------------------------------------------
def _substituir_tarefas(variante: Variante, tarefas: Sequence[Tarefa]) -> Variante:
    return Variante(id=variante.id, nome=variante.nome, tarefas=tuple(tarefas))


def salvar_tarefa(variante_id: str, nome: str, tempo, condicoes: Sequence[dict],
                  tarefa_id: str | None = None,
                  sequencia: dict | None = None,
                  setor: str | None = None) -> tuple[Variante, ...]:
    """Cria uma tarefa nova ou atualiza a existente (quando tarefa_id vem preenchido)."""
    nome = (nome or "").strip()
    if not nome:
        raise ErroCatalogo("O nome da tarefa é obrigatório.")

    try:
        tempo = float(str(tempo).replace(",", "."))
    except (TypeError, ValueError):
        raise ErroCatalogo("O tempo deve ser um número em segundos.")
    if tempo <= 0:
        raise ErroCatalogo("O tempo deve ser maior que zero.")

    criterio = criterio_de_condicoes(condicoes, sequencia)

    variantes = list(carregar(setor))
    indice = next((i for i, v in enumerate(variantes) if v.id == variante_id), None)
    if indice is None:
        raise ErroCatalogo(f"Variante '{variante_id}' não encontrada.")

    alvo = variantes[indice]
    tarefas = list(alvo.tarefas)

    if tarefa_id:
        posicao = next((i for i, t in enumerate(tarefas) if t.id == tarefa_id), None)
        if posicao is None:
            # A tarefa mudou de variante: remove da antiga e adiciona na nova.
            variantes = _remover(variantes, tarefa_id)
            alvo = next(v for v in variantes if v.id == variante_id)
            indice = variantes.index(alvo)
            tarefas = list(alvo.tarefas)
            tarefas.append(Tarefa(nome, tempo, criterio,
                                  _chave_legado(nome, variante_id), tarefa_id))
        else:
            antiga = tarefas[posicao]
            tarefas[posicao] = Tarefa(nome, tempo, criterio,
                                      antiga.chave_legado or _chave_legado(nome, variante_id),
                                      antiga.id)
    else:
        tarefas.append(Tarefa(nome, tempo, criterio,
                              _chave_legado(nome, variante_id), novo_id()))

    variantes[indice] = _substituir_tarefas(alvo, tarefas)
    salvar(variantes, setor)
    return tuple(variantes)


def _chave_legado(nome: str, variante_id: str) -> str:
    return f"{nome.lower()} {variante_id.lower()}"


def _remover(variantes: list[Variante], tarefa_id: str) -> list[Variante]:
    return [
        _substituir_tarefas(v, [t for t in v.tarefas if t.id != tarefa_id])
        for v in variantes
    ]


def remover_tarefa(tarefa_id: str, setor: str | None = None) -> tuple[Variante, ...]:
    variantes = list(carregar(setor))
    if not any(t.id == tarefa_id for v in variantes for t in v.tarefas):
        raise ErroCatalogo("Tarefa não encontrada.")
    variantes = _remover(variantes, tarefa_id)
    salvar(variantes, setor)
    return tuple(variantes)


def adicionar_variante(variante_id: str, nome: str, setor: str | None = None) -> tuple[Variante, ...]:
    variante_id = (variante_id or "").strip().upper()
    nome = (nome or "").strip() or variante_id
    if not variante_id:
        raise ErroCatalogo("Informe o código da variante (ex.: V5).")

    variantes = list(carregar(setor))
    if any(v.id == variante_id for v in variantes):
        raise ErroCatalogo(f"Já existe uma variante '{variante_id}'.")

    variantes.append(Variante(id=variante_id, nome=nome, tarefas=()))
    salvar(variantes, setor)
    return tuple(variantes)


def remover_variante(variante_id: str, setor: str | None = None) -> tuple[Variante, ...]:
    variantes = [v for v in carregar(setor) if v.id != variante_id]
    salvar(variantes, setor)
    return tuple(variantes)


def descrever_condicoes(criterio: Criterio) -> str:
    """'MODELO contém 6X2 e OPCIONAIS contém DISCO'"""
    texto = " e ".join(
        f"{c['coluna']} {OPERADORES_CURTOS[c['operador']]} {c['valor']}"
        for c in condicoes_de_criterio(criterio)
    )
    grupo = sequencia_de_criterio(criterio)
    if grupo:
        rotulo = {"primeiro": "1x por grupo",
                  "todos": "todos do grupo",
                  "excedente": "excedente do grupo"}[grupo["modo"]]
        onde = (f"{grupo['minimo']} em cada {grupo['janela']}"
                if grupo.get("janela")
                else f"{grupo['minimo']}+ seguidos")
        texto += f"  [{onde} → {rotulo}]"
    return texto


def exportar_tsv(variantes: Iterable[Variante] | None = None,
                 setor: str | None = None) -> str:
    """Regras em colunas separadas por TAB, para colar no Excel."""
    variantes = variantes if variantes is not None else carregar(setor)
    linhas = ["Variante\tCodigo\tTarefa\tTempo (s)\tCondicoes"]
    for variante in variantes:
        if not variante.tarefas:
            linhas.append(f"{variante.nome}\t{variante.id}\t(sem tarefas)\t\t")
            continue
        for tarefa in variante.tarefas:
            linhas.append("\t".join([
                variante.nome, variante.id, tarefa.nome,
                ("%g" % tarefa.tempo), descrever_condicoes(tarefa.criterio),
            ]))
    return "\n".join(linhas)


def exportar_texto(variantes: Iterable[Variante] | None = None,
                   setor: str | None = None) -> str:
    """Regras em texto legivel, para e-mail ou documentacao."""
    variantes = variantes if variantes is not None else carregar(setor)
    linhas = ["REGRAS DAS VARIANTES", ""]
    for variante in variantes:
        linhas.append(f"{variante.id} - {variante.nome}")
        if not variante.tarefas:
            linhas.append("  (sem tarefas)")
        for tarefa in variante.tarefas:
            linhas.append(f"  {tarefa.nome} ({'%g' % tarefa.tempo}s)")
            linhas.append(f"      quando {descrever_condicoes(tarefa.criterio)}")
        linhas.append("")
    return "\n".join(linhas)


def exportar_json(variantes: Iterable[Variante] | None = None,
                  setor: str | None = None) -> str:
    """Mesmo conteudo do catalogo.json — serve de backup e de importacao."""
    variantes = variantes if variantes is not None else carregar(setor)
    return json.dumps(para_dict(variantes), ensure_ascii=False, indent=2)


def encontrar_tarefa(tarefa_id: str, setor: str | None = None):
    """Devolve (variante, tarefa) ou (None, None)."""
    for variante in carregar(setor):
        for tarefa in variante.tarefas:
            if tarefa.id == tarefa_id:
                return variante, tarefa
    return None, None

"""
tags_mona.py - Tags especiais da MONA (variantes que ainda nao existem no mix
geral da fabrica), gravadas em `dados/<setor>/tags_mona.json` — uma lista por
setor, mesmo padrao de `gargalos.arquivo_do_setor` para regras.json.

Tira a lista de dentro do codigo. Sem esse arquivo, `cpm_core` continua
usando a lista embutida do setor (`tags_mona_padrao_do_setor`); na primeira
gravacao pela tela o JSON nasce com esse mesmo conteudo, entao nada se perde.

Formato:

    {
      "tags": [
        {"texto": "658|660", "padroes": ["658\\\\D*660"]},
        {"texto": "649|650", "padroes": ["649\\\\D*650"]},
        {"texto": "Cilindro elevado", "padroes": ["CILINDRO\\\\s*ELEV"]}
      ]
    }

`texto`   : o que entra em OPCIONAIS/Variantes quando o POPID tem essa tag, e
            o que o contador do CPM MIX procura.
`padroes` : expressoes regex (uma ou mais) que reconhecem essa tag pelo nome
            do item numa planilha MONA no formato tecnico da Funcao 4.
"""

from __future__ import annotations

import json
import logging
import os
import re
from datetime import datetime

from armazenamento import caminho_dados
from cpm_core import TagMona, tags_mona_padrao_do_setor
from setores import SETOR_PADRAO

logger = logging.getLogger(__name__)


class ErroTagsMona(ValueError):
    pass


def arquivo_do_setor(setor: str | None = None) -> str:
    """
    Tags de um setor: dados/<setor>/tags_mona.json.

    O arquivo antigo (dados/tags_mona.json) vira as tags do setor padrao
    (FA2) enquanto o novo nao existir, para nao perder o que ja foi criado
    antes de existir mais de um setor.
    """
    setor_id = setor or SETOR_PADRAO
    novo = caminho_dados(os.path.join(setor_id, "tags_mona.json"))
    if not os.path.exists(novo) and setor_id == SETOR_PADRAO:
        antigo = caminho_dados("tags_mona.json")
        if os.path.exists(antigo):
            return antigo
    return novo


def _validar_regex(padrao: str) -> str:
    try:
        re.compile(padrao, re.IGNORECASE)
    except re.error as erro:
        raise ErroTagsMona(f"Padrão inválido '{padrao}': {erro}") from erro
    return padrao


def para_dict(tags: "list[TagMona] | tuple[TagMona, ...]") -> dict:
    return {
        "tags": [
            {"texto": tag.texto, "padroes": list(tag.padroes)}
            for tag in tags
        ]
    }


def de_dict(dados: dict) -> tuple[TagMona, ...]:
    tags = []
    for item in dados.get("tags", []):
        texto = str(item.get("texto") or "").strip()
        if not texto:
            continue
        padroes = tuple(
            str(p).strip() for p in (item.get("padroes") or []) if str(p).strip()
        )
        tags.append(TagMona(texto, padroes))
    return tuple(tags)


def carregar(setor: str | None = None) -> tuple[TagMona, ...]:
    """Lê a lista gravada do setor; se o arquivo não existir, devolve a
    embutida daquele setor."""
    arquivo = arquivo_do_setor(setor)
    if not os.path.exists(arquivo):
        return tags_mona_padrao_do_setor(setor)
    with open(arquivo, encoding="utf-8") as f:
        return de_dict(json.load(f))


def salvar(tags: "list[TagMona] | tuple[TagMona, ...]", setor: str | None = None) -> None:
    """Valida e grava de forma atômica, para não corromper o arquivo."""
    vistos = set()
    for tag in tags:
        if not tag.texto.strip():
            raise ErroTagsMona("O texto da tag não pode ficar em branco.")
        chave = tag.texto.strip().lower()
        if chave in vistos:
            raise ErroTagsMona(f"O texto '{tag.texto}' está duplicado.")
        vistos.add(chave)
        for padrao in tag.padroes:
            _validar_regex(padrao)

    arquivo = arquivo_do_setor(setor)
    os.makedirs(os.path.dirname(arquivo), exist_ok=True)
    temporario = arquivo + ".tmp"
    with open(temporario, "w", encoding="utf-8") as f:
        json.dump(para_dict(tags), f, ensure_ascii=False, indent=2)
    os.replace(temporario, arquivo)


def copia_de_seguranca(setor: str | None = None) -> str | None:
    """Guarda a lista atual antes de sobrescrever. Devolve o nome do arquivo."""
    arquivo = arquivo_do_setor(setor)
    if not os.path.exists(arquivo):
        return None
    setor_id = setor or SETOR_PADRAO
    nome = f"tags_mona_{setor_id}_backup_{datetime.now():%Y%m%d_%H%M%S}.json"
    destino = caminho_dados(nome)
    with open(arquivo, encoding="utf-8") as origem, \
            open(destino, "w", encoding="utf-8") as copia:
        copia.write(origem.read())
    return nome


def restaurar_padrao(setor: str | None = None) -> tuple[TagMona, ...]:
    copia_de_seguranca(setor)
    padrao = tags_mona_padrao_do_setor(setor)
    salvar(padrao, setor)
    return padrao


def adicionar(texto: str, padroes_texto: str, setor: str | None = None) -> tuple[TagMona, ...]:
    """
    padroes_texto: um padrão por linha (como vem de um <textarea> na tela).
    Sem nenhum padrão informado, usa o próprio texto (escapado) como padrão —
    reconhece a tag numa coluna da MONA cujo nome seja exatamente esse texto.
    """
    texto = texto.strip()
    if not texto:
        raise ErroTagsMona("Informe o texto da tag.")
    padroes = tuple(
        linha.strip() for linha in (padroes_texto or "").splitlines()
        if linha.strip()
    ) or (re.escape(texto),)

    atuais = list(carregar(setor))
    if any(t.texto.strip().lower() == texto.lower() for t in atuais):
        raise ErroTagsMona(f"Já existe uma tag com o texto '{texto}'.")

    atuais.append(TagMona(texto, padroes))
    copia_de_seguranca(setor)
    salvar(atuais, setor)
    return tuple(atuais)


def remover(texto: str, setor: str | None = None) -> tuple[TagMona, ...]:
    atuais = [t for t in carregar(setor) if t.texto != texto]
    copia_de_seguranca(setor)
    salvar(atuais, setor)
    return tuple(atuais)

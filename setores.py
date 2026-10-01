"""
setores.py - Registro dos setores e a configuracao de cada um.

Ate aqui o sistema atendia um setor so (FA2), e os parametros dele estavam
escritos como constantes de modulo no `cpm_core`: takt, turno, pausas, tabela
de postos por falta. Pareciam parametros do sistema, mas sao do FA2.

Aqui eles viram configuracao por setor. Cada setor tem:

  * parametros de calculo  — takt, turno, pausas, postos, limiar de gargalos
  * dados proprios         — dados/<id>/catalogo.json e dados/<id>/regras.json

Assim dois setores com takt e turno diferentes convivem sem um interferir no
outro. `setores.json` fica na pasta de dados (volume, em deploy) e pode ser
editado; `setores_fabrica.json` acompanha o codigo e e a base do "restaurar".
"""

from __future__ import annotations

import json
import os
import logging
import re
from dataclasses import asdict, dataclass, field

from armazenamento import caminho_codigo, caminho_dados

logger = logging.getLogger(__name__)

ARQUIVO = caminho_dados("setores.json")
ARQUIVO_FABRICA = caminho_codigo("setores_fabrica.json")

SETOR_PADRAO = "fa2"

# Valores que valem para a fabrica e servem de ponto de partida para cada setor
# novo. Ficam no mesmo setores.json, num bloco separado.
PADROES_FABRICA = {
    "producao": "88",
    "takt_segundos": 300.0,
    "inicio_turno": "07:05",
    "fim_turno": "16:05",
    "pausas": [["09:00", "09:04"], ["10:55", "11:55"], ["14:00", "14:10"]],
    "postos": 4,
    "postos_por_faltas": {"0": None, "1": None, "2": 2},
    "gargalos_varios": 2,
}

class ErroSetor(ValueError):
    """Erro de validacao mostrado na tela de setores."""


@dataclass
class Setor:
    id: str
    nome: str
    descricao: str = ""

    # --- parametros de calculo ---
    takt_segundos: float = 300.0        # tempo de ciclo usado na carga
    inicio_turno: str = "07:05"
    fim_turno: str = "16:05"
    pausas: list = field(default_factory=list)   # [["09:00", "09:04"], ...]
    postos: int = 4                     # variantes/postos do setor

    # absenteismo: quantos postos seguem ativos com N ausentes.
    # chave "0" e "1" com valor null = todos os postos.
    postos_por_faltas: dict = field(default_factory=lambda: {"0": None, "1": None, "2": 2})
    gargalos_varios: int = 2            # a partir de quantos gargalos e "varios"

    @property
    def pausas_tuplas(self) -> tuple[tuple[str, str], ...]:
        return tuple((p[0], p[1]) for p in self.pausas if len(p) >= 2)

    @property
    def tabela_postos(self) -> dict[int, int | None]:
        """A chave vem do JSON como texto; o calculo usa inteiro."""
        tabela = {}
        for chave, valor in (self.postos_por_faltas or {}).items():
            try:
                tabela[int(chave)] = None if valor is None else int(valor)
            except (TypeError, ValueError):
                continue
        return tabela

    def pasta(self, arquivo: str) -> str:
        """Caminho de um arquivo de dados deste setor."""
        return caminho_dados(os.path.join(self.id, arquivo))


# ---------------------------------------------------------------------------
# Persistencia
# ---------------------------------------------------------------------------
def _de_dict(bruto: dict) -> Setor:
    campos = {c: bruto[c] for c in Setor.__dataclass_fields__ if c in bruto}
    if "id" not in campos or not campos["id"]:
        raise ErroSetor("Setor sem 'id'.")
    if "nome" not in campos or not campos["nome"]:
        campos["nome"] = campos["id"].upper()
    return Setor(**campos)


def _ler_arquivo() -> dict:
    for caminho in (ARQUIVO, ARQUIVO_FABRICA):
        try:
            with open(caminho, encoding="utf-8") as arquivo:
                return json.load(arquivo)
        except FileNotFoundError:
            continue
        except json.JSONDecodeError as erro:
            logger.warning("%s inválido (%s)", os.path.basename(caminho), erro)
    return {}


def carregar_padroes() -> dict:
    padroes = dict(PADROES_FABRICA)
    padroes.update(_ler_arquivo().get("padroes") or {})
    return padroes


def salvar_padroes(novos: dict) -> dict:
    """Valida e grava os padrões da fábrica, sem tocar nos setores."""
    padroes = carregar_padroes()

    if "producao" in novos:
        producao = (novos["producao"] or "").strip()
        if producao:
            try:
                if int(float(producao.replace(",", "."))) <= 0:
                    raise ValueError
            except ValueError:
                raise ErroSetor("A produção padrão deve ser um número maior que zero.")
        padroes["producao"] = producao

    if "takt_segundos" in novos:
        try:
            takt = float(str(novos["takt_segundos"]).replace(",", "."))
        except ValueError:
            raise ErroSetor("O takt padrão deve ser um número de segundos.")
        if takt <= 0:
            raise ErroSetor("O takt padrão deve ser maior que zero.")
        padroes["takt_segundos"] = takt

    if "inicio_turno" in novos:
        padroes["inicio_turno"] = _horario(novos["inicio_turno"], "Início do turno")
    if "fim_turno" in novos:
        padroes["fim_turno"] = _horario(novos["fim_turno"], "Fim do turno")
    if "pausas" in novos:
        padroes["pausas"] = pausas_de_texto(novos["pausas"])
    if "postos" in novos:
        try:
            postos = int(novos["postos"])
        except (TypeError, ValueError):
            raise ErroSetor("O número padrão de postos deve ser inteiro.")
        if postos < 1:
            raise ErroSetor("O número padrão de postos deve ser 1 ou mais.")
        padroes["postos"] = postos
    if "postos_por_faltas" in novos:
        padroes["postos_por_faltas"] = _tabela_de_texto(
            novos["postos_por_faltas"], padroes.get("postos", 4))
    if "gargalos_varios" in novos:
        try:
            padroes["gargalos_varios"] = max(1, int(novos["gargalos_varios"]))
        except (TypeError, ValueError):
            raise ErroSetor("O limiar de gargalos deve ser inteiro.")

    dados = _ler_arquivo()
    dados["padroes"] = padroes
    dados.setdefault("setores", [asdict(s) for s in carregar()])

    temporario = ARQUIVO + ".tmp"
    with open(temporario, "w", encoding="utf-8") as arquivo:
        json.dump(dados, arquivo, ensure_ascii=False, indent=2)
    os.replace(temporario, ARQUIVO)
    return padroes


def setor_a_partir_dos_padroes() -> Setor:
    """Setor em branco ja preenchido com os padrões — o formulário de criação."""
    padroes = carregar_padroes()
    return Setor(
        id="", nome="", descricao="",
        takt_segundos=padroes["takt_segundos"],
        inicio_turno=padroes["inicio_turno"], fim_turno=padroes["fim_turno"],
        pausas=[list(p) for p in padroes["pausas"]],
        postos=padroes["postos"],
        postos_por_faltas=dict(padroes["postos_por_faltas"]),
        gargalos_varios=padroes["gargalos_varios"],
    )


def carregar() -> list[Setor]:
    for caminho in (ARQUIVO, ARQUIVO_FABRICA):
        try:
            with open(caminho, encoding="utf-8") as arquivo:
                dados = json.load(arquivo)
            return [_de_dict(s) for s in dados.get("setores", [])]
        except FileNotFoundError:
            continue
        except (json.JSONDecodeError, ErroSetor, TypeError) as erro:
            logger.warning("%s inválido (%s)", os.path.basename(caminho), erro)
    return [Setor(id=SETOR_PADRAO, nome="FA2")]


def salvar(setores: list[Setor]) -> list[Setor]:
    if not setores:
        raise ErroSetor("É preciso haver pelo menos um setor.")
    conteudo = _ler_arquivo()
    conteudo["setores"] = [asdict(s) for s in setores]
    padroes = carregar_padroes()
    if padroes != PADROES_FABRICA:
        conteudo["padroes"] = padroes    # gravar setores nao apaga os padroes
    temporario = ARQUIVO + ".tmp"
    with open(temporario, "w", encoding="utf-8") as arquivo:
        json.dump(conteudo, arquivo, ensure_ascii=False, indent=2)
    os.replace(temporario, ARQUIVO)
    return setores


def obter(setor_id: str | None) -> Setor:
    """
    Devolve o setor pedido. Sem id, ou com id desconhecido, cai no primeiro —
    assim um link antigo sem `?setor=` continua funcionando.
    """
    setores = carregar()
    if setor_id:
        for setor in setores:
            if setor.id == setor_id:
                return setor
    for setor in setores:
        if setor.id == SETOR_PADRAO:
            return setor
    return setores[0]


def existe(setor_id: str) -> bool:
    return any(s.id == setor_id for s in carregar())


def restaurar_padrao() -> list[Setor]:
    with open(ARQUIVO_FABRICA, encoding="utf-8") as arquivo:
        return salvar([_de_dict(s) for s in json.load(arquivo).get("setores", [])])


# ---------------------------------------------------------------------------
# CRUD usado pela tela
# ---------------------------------------------------------------------------
def _id_valido(bruto: str) -> str:
    limpo = re.sub(r"[^a-z0-9_-]+", "-", (bruto or "").strip().lower()).strip("-")
    if not limpo:
        raise ErroSetor("Informe um identificador (ex.: fa2, fa12, fibras).")
    return limpo[:24]


def _horario(valor: str, rotulo: str) -> str:
    if not re.fullmatch(r"\d{1,2}:\d{2}", (valor or "").strip()):
        raise ErroSetor(f"{rotulo} deve estar no formato HH:MM.")
    hora, minuto = (int(p) for p in valor.split(":"))
    if hora > 23 or minuto > 59:
        raise ErroSetor(f"{rotulo} não é um horário válido.")
    return f"{hora:02d}:{minuto:02d}"


def pausas_de_texto(texto: str) -> list[list[str]]:
    """
    Lê '09:00-09:04' por linha. Formato curto de propósito: pausa muda pouco e
    uma tabela na tela seria mais trabalho do que ajuda.
    """
    pausas = []
    for linha in (texto or "").replace(";", "\n").splitlines():
        linha = linha.strip()
        if not linha:
            continue
        partes = re.split(r"\s*(?:-|até|as|às)\s*", linha, maxsplit=1)
        if len(partes) != 2:
            raise ErroSetor(f"Pausa inválida: '{linha}'. Use 09:00-09:04.")
        pausas.append([_horario(partes[0], "Início da pausa"),
                       _horario(partes[1], "Fim da pausa")])
    return pausas


def pausas_para_texto(pausas) -> str:
    return "\n".join(f"{p[0]}-{p[1]}" for p in pausas or [] if len(p) >= 2)


def salvar_setor(dados: dict, setor_id: str | None = None) -> Setor:
    """Cria ou atualiza um setor a partir do formulário."""
    nome = (dados.get("nome") or "").strip()
    if not nome:
        raise ErroSetor("O nome do setor é obrigatório.")

    novo_id = _id_valido(dados.get("id") or setor_id or nome)

    try:
        takt = float(str(dados.get("takt_segundos") or 0).replace(",", "."))
    except ValueError:
        raise ErroSetor("O takt deve ser um número de segundos.")
    if takt <= 0:
        raise ErroSetor("O takt deve ser maior que zero.")

    try:
        postos = int(dados.get("postos") or 0)
    except (TypeError, ValueError):
        raise ErroSetor("O número de postos deve ser inteiro.")
    if postos < 1:
        raise ErroSetor("O setor precisa de pelo menos um posto.")

    try:
        varios = int(dados.get("gargalos_varios") or 2)
    except (TypeError, ValueError):
        raise ErroSetor("O limiar de gargalos deve ser inteiro.")

    setor = Setor(
        id=novo_id,
        nome=nome,
        descricao=(dados.get("descricao") or "").strip(),
        takt_segundos=takt,
        inicio_turno=_horario(dados.get("inicio_turno"), "Início do turno"),
        fim_turno=_horario(dados.get("fim_turno"), "Fim do turno"),
        pausas=pausas_de_texto(dados.get("pausas", "")),
        postos=postos,
        postos_por_faltas=_tabela_de_texto(dados.get("postos_por_faltas", ""), postos),
        gargalos_varios=max(1, varios),
    )

    setores = carregar()
    for posicao, atual in enumerate(setores):
        if atual.id == (setor_id or novo_id):
            setores[posicao] = setor
            break
    else:
        if any(s.id == novo_id for s in setores):
            raise ErroSetor(f"Já existe um setor com o identificador '{novo_id}'.")
        setores.append(setor)

    salvar(setores)
    return setor


def _tabela_de_texto(texto: str, postos: int) -> dict:
    """
    Lê 'faltas=postos' por linha; 'todos' mantém o setor completo.
    Vazio devolve a regra atual do FA2.
    """
    if not (texto or "").strip():
        return {"0": None, "1": None, "2": min(2, postos)}

    tabela: dict[str, int | None] = {}
    for linha in texto.replace(";", "\n").splitlines():
        linha = linha.strip()
        if not linha:
            continue
        if "=" not in linha:
            raise ErroSetor(f"Regra de absenteísmo inválida: '{linha}'. Use 2=2.")
        faltas, valor = (p.strip() for p in linha.split("=", 1))
        try:
            chave = str(int(faltas))
        except ValueError:
            raise ErroSetor(f"'{faltas}' não é um número de faltas.")
        if valor.lower() in ("todos", "all", "*"):
            tabela[chave] = None
        else:
            try:
                tabela[chave] = max(1, min(int(valor), postos))
            except ValueError:
                raise ErroSetor(f"'{valor}' não é um número de postos.")
    return tabela


def tabela_para_texto(tabela: dict) -> str:
    linhas = []
    for chave in sorted(tabela or {}, key=lambda c: int(c)):
        valor = tabela[chave]
        linhas.append(f"{chave}={'todos' if valor is None else valor}")
    return "\n".join(linhas)


def remover(setor_id: str) -> list[Setor]:
    setores = [s for s in carregar() if s.id != setor_id]
    if len(setores) == len(carregar()):
        raise ErroSetor("Setor não encontrado.")
    if not setores:
        raise ErroSetor("Não dá para remover o último setor.")
    return salvar(setores)

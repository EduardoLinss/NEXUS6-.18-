"""
visao_fabrica.py - Consolida a analise de varios setores em uma visao so.

A tela de Analise de MIX responde "como esta o FA2". Esta responde "como esta a
linha": mesma planilha, um POPID de corte por setor, e o resultado lado a lado.

Cada setor calcula com os proprios parametros (takt, postos, pausas, regras de
gargalo), entao os numeros nao sao comparaveis por acaso — sao comparaveis
porque cada um esta certo dentro do seu setor.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field

import cpm_core
import gargalos as gargalos_mod

logger = logging.getLogger(__name__)

# Altura util do mini-grafico, em unidades do viewBox do SVG.
ALTURA_SVG = 40.0
LARGURA_SVG = 100.0


@dataclass
class LinhaSetor:
    """O que a tela mostra de um setor."""

    setor: object
    popid: str = ""
    producao: str = ""
    faltas: str = ""
    erro: str | None = None

    resultado: object | None = None
    gargalos: dict | None = None

    # dados ja prontos para o SVG
    pontos: str = ""
    y_capacidade: float = 0.0
    blocos: list = field(default_factory=list)

    # -- numeros de topo --
    @property
    def ok(self) -> bool:
        return self.resultado is not None and self.erro is None

    @property
    def carga_maxima(self) -> float:
        if not self.resultado or not self.resultado.variantes:
            return 0.0
        return max(v.total for v in self.resultado.variantes.values())

    @property
    def carga_media(self) -> float:
        if not self.resultado or not self.resultado.variantes:
            return 0.0
        totais = [v.total for v in self.resultado.variantes.values()]
        return sum(totais) / len(totais)

    @property
    def capacidade(self) -> float:
        return max(getattr(self.resultado, "postos_ativos", 0), 1) * 100.0

    @property
    def n_gargalos(self) -> int:
        return (self.gargalos or {}).get("resumo", {}).get("total", 0)

    @property
    def n_alta(self) -> int:
        return (self.gargalos or {}).get("resumo", {}).get("alta", 0)

    @property
    def veiculos_afetados(self) -> int:
        return (self.gargalos or {}).get("resumo", {}).get("unidades_afetadas", 0)

    @property
    def alertas(self) -> list:
        return getattr(self.resultado, "alertas", []) or []

    @property
    def situacao(self) -> str:
        """
        Estado do setor em uma palavra, para ordenar e colorir.

        Considera as duas coisas que travam a linha: carga acima da capacidade
        e concentracao de gargalos. Um setor pode estar folgado na media e
        ainda assim ser critico num trecho — por isso a carga usada e a maxima.
        """
        if not self.ok:
            return "sem-dados"
        if self.alertas or self.carga_maxima >= 100:
            return "critico"
        if self.n_alta or self.carga_maxima >= 90:
            return "atencao"
        return "ok"


ORDEM_SITUACAO = {"critico": 0, "atencao": 1, "ok": 2, "sem-dados": 3}


def _pontos_do_grafico(demanda: list[float], teto: float) -> str:
    """Converte a curva em coordenadas do SVG (y invertido)."""
    if not demanda or teto <= 0:
        return ""
    passo = LARGURA_SVG / max(len(demanda) - 1, 1)
    partes = []
    for i, valor in enumerate(demanda):
        y = ALTURA_SVG - min(valor / teto, 1.0) * ALTURA_SVG
        partes.append(f"{i * passo:.2f},{y:.2f}")
    return " ".join(partes)


def analisar_setor(df, setor, popid: str, producao: str, faltas: str) -> LinhaSetor:
    """Roda a analise de um setor e prepara o que a tela precisa."""
    linha = LinhaSetor(setor=setor, popid=(popid or "").strip(),
                       producao=(producao or "").strip(),
                       faltas=(faltas or "").strip())

    if not linha.popid or not linha.producao:
        return linha

    try:
        n_faltas = int(float((linha.faltas or "0").replace(",", ".")))
    except ValueError:
        linha.erro = "Absenteísmo deve ser um número."
        return linha

    try:
        resultado = cpm_core.analisar_mix(df, linha.popid, linha.producao,
                                          takt=setor.takt_segundos)
    except ValueError as erro:
        linha.erro = str(erro)
        return linha

    recorte = df.iloc[resultado.primeira_linha - 1:resultado.ultima_linha]
    try:
        dados = gargalos_mod.analisar(recorte, setor.id)
    except Exception as erro:  # noqa: BLE001
        logger.warning("gargalos do setor %s: %s", setor.id, erro)
        dados = {"resumo": {}, "gargalos": [], "unidades": []}

    cpm_core.aplicar_absenteismo(
        resultado, n_faltas, dados.get("resumo", {}).get("total", 0),
        setor.tabela_postos, setor.gargalos_varios)

    linha.resultado = resultado
    linha.gargalos = dados

    demanda = cpm_core.suavizar(resultado.demanda, 5)
    capacidade = max(resultado.postos_ativos, 1) * 100.0
    teto = max(capacidade * 1.25, max(demanda or [0]) * 1.1, 1)
    linha.pontos = _pontos_do_grafico(demanda, teto)
    linha.y_capacidade = ALTURA_SVG - min(capacidade / teto, 1.0) * ALTURA_SVG
    linha.blocos = [u.get("heat", "NENHUM") for u in dados.get("unidades", [])]

    return linha


def consolidar(linhas: list[LinhaSetor]) -> dict:
    """Resumo da linha inteira, para o topo da tela."""
    ativos = [l for l in linhas if l.ok]
    todos_gargalos = []
    for linha in ativos:
        for gargalo in (linha.gargalos or {}).get("gargalos", []):
            todos_gargalos.append({
                "setor": linha.setor,
                "severidade": gargalo["severidade"],
                "regra": gargalo["regra_nome"],
                "a": gargalo["a"] + 1,
                "b": gargalo["b"] + 1,
                "detalhe": gargalo["detalhe"],
                "membros": gargalos_mod.descrever_membros(linha.gargalos, gargalo),
            })

    # Alta primeiro, depois pela posicao: quem trava mais cedo no turno aparece
    # antes, que e a ordem em que da para agir.
    todos_gargalos.sort(key=lambda g: (0 if g["severidade"] == "ALTA" else 1, g["a"]))

    return {
        "setores_analisados": len(ativos),
        "setores_criticos": sum(1 for l in ativos if l.situacao == "critico"),
        "setores_atencao": sum(1 for l in ativos if l.situacao == "atencao"),
        "gargalos": todos_gargalos,
        "total_gargalos": len(todos_gargalos),
        "gargalos_alta": sum(1 for g in todos_gargalos if g["severidade"] == "ALTA"),
        "veiculos": sum(l.resultado.total_linhas for l in ativos),
        "pior_carga": max((l.carga_maxima for l in ativos), default=0.0),
    }

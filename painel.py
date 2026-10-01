"""
painel.py - Visao geral da fabrica: todos os setores lado a lado.

As telas existentes olham um setor por vez. Aqui a mesma planilha alimenta
todos os setores cadastrados, cada um com o seu POPID de corte, e o resultado
vem consolidado: carga, gargalos e a curva de demanda de cada um.
"""

from __future__ import annotations

import logging
import cpm_core
import gargalos as gargalos_mod

logger = logging.getLogger(__name__)


LIMITE_ATENCAO = 90.0
LIMITE_CRITICO = 100.0

# Cores das linhas do grafico geral. Sao 8 porque a fabrica tem 7 setores
# previstos e ainda sobra uma; acima disso o ciclo se repete.
CORES_SETORES = ("#7FD4FF", "#FF5FA2", "#6FB6A4", "#D99A4E",
                 "#A78BFA", "#F0705F", "#9FB4C7", "#E8D45A")

# O 1o turno usa o horario cadastrado em cada setor. O 2o e o mesmo para a
# fabrica toda e atravessa a meia-noite, por isso vem escrito aqui.
TURNOS = {
    "1": {"id": "1", "nome": "1º turno",
          "descricao": "usa o horário cadastrado em cada setor",
          "inicio": None, "fim": None, "takt_obrigatorio": False},
    "2": {"id": "2", "nome": "2º turno",
          "descricao": "17:30 às 02:08, atravessa a meia-noite",
          "inicio": "17:30", "fim": "02:08", "takt_obrigatorio": True},
}
TURNO_PADRAO = "1"


def obter_turno(turno_id: str | None) -> dict:
    return TURNOS.get((turno_id or "").strip() or TURNO_PADRAO, TURNOS[TURNO_PADRAO])


def horario_do_setor(setor, turno: dict) -> tuple[str, str]:
    """Início e fim que valem para este setor neste turno."""
    if turno.get("inicio") and turno.get("fim"):
        return turno["inicio"], turno["fim"]
    return setor.inicio_turno, setor.fim_turno


def janelas_de_texto(texto: str) -> list[tuple[str, str, int]] | None:
    """
    Le 'faltas por horário', uma janela por linha: '10:00-12:00=2'.

    Devolve None quando o texto e o formato antigo (um numero so, sem ':') —
    nesse caso quem chama trata como o absenteismo fixo do turno inteiro, como
    sempre foi. So quando ha pelo menos uma linha com horario e que a analise
    passa a variar a capacidade ao longo do dia.
    """
    linhas = [l.strip() for l in (texto or "").replace(";", "\n").splitlines()
             if l.strip()]
    if not linhas or not any(":" in l for l in linhas):
        return None

    janelas = []
    for linha in linhas:
        if "=" not in linha or "-" not in linha.split("=")[0]:
            continue
        periodo, valor = linha.split("=", 1)
        inicio, fim = periodo.split("-", 1)
        try:
            faltas = int(valor.strip())
        except ValueError:
            continue
        janelas.append((inicio.strip(), fim.strip(), max(0, faltas)))
    return janelas or None


def _faixa(percentual: float) -> str:
    if percentual >= LIMITE_CRITICO:
        return "critico"
    if percentual >= LIMITE_ATENCAO:
        return "atencao"
    return "ok"


def _curva_svg(demanda: list[float], capacidade: float,
               largura: int = 260, altura: int = 54) -> dict:
    """
    Pontos de uma mini curva de demanda, ja em coordenadas do SVG.

    Feito no servidor de proposito: sem biblioteca de grafico, sem JavaScript,
    o que importa numa rede onde instalar coisa nova e caro.
    """
    if not demanda:
        return {"linha": "", "area": "", "teto": capacidade or 1}

    suave = cpm_core.suavizar(demanda, 5)
    teto = max(max(suave), capacidade, 1) * 1.1
    passo = largura / max(len(suave) - 1, 1)

    pontos = []
    for i, valor in enumerate(suave):
        x = i * passo
        y = altura - (valor / teto) * altura
        pontos.append(f"{x:.1f},{y:.1f}")

    linha = " ".join(pontos)
    area = f"0,{altura} " + linha + f" {largura},{altura}"
    y_cap = altura - (capacidade / teto) * altura if teto else 0

    return {"linha": linha, "area": area, "teto": teto,
            "y_capacidade": max(0.0, min(y_cap, altura)),
            "largura": largura, "altura": altura}


def demanda_por_variante(linha: dict) -> list[dict]:
    """
    Quanto trabalho cada VARIANTE recebe, veiculo a veiculo.

    A curva do setor soma tudo; aqui a soma e separada por posto. E o que
    responde "qual variante esta no aperto", que a soma esconde: dois setores
    com o mesmo total podem ter um distribuido e outro concentrado num posto.

    O percentual e sobre UM posto (100% = posto saturado), entao as linhas de
    variantes de setores diferentes sao comparaveis entre si.
    """
    resultado = linha.get("resultado")
    if not resultado:
        return []

    total = len(resultado.demanda)
    takt = linha.get("takt") or resultado and linha["setor"].takt_segundos

    series = []
    for variante in resultado.variantes.values():
        tempos = [0.0] * total
        for tarefa in variante.tarefas:
            for item in tarefa.itens:
                indice = item["posicao"] - 1
                if 0 <= indice < total:
                    tempos[indice] += tarefa.tempo
        series.append({
            "variante": variante,
            "valores": [cpm_core.carga_percentual(1, t, 1, takt) for t in tempos],
        })
    return series


def _relogios(linha: dict, quantidade: int) -> list[float]:
    """
    Horario de relogio de cada posicao.

    No 2o turno valem 17:30-02:08. As pausas cadastradas no setor sao de
    manha, ficam fora dessa janela e sao descartadas sozinhas — se o 2o turno
    tiver pausas proprias, elas precisam ser cadastradas.
    """
    inicio = linha.get("inicio_turno") or linha["setor"].inicio_turno
    fim = linha.get("fim_turno") or linha["setor"].fim_turno
    pausas = linha["setor"].pausas_tuplas or None

    ritmo = cpm_core.ritmo_do_turno(quantidade, inicio, fim, pausas)
    return [cpm_core.relogio_apos(i * ritmo, inicio, fim, pausas)
            for i in range(quantidade)]


def curva_geral(linhas: list[dict], altura: int = 260,
                largura_por_hora: int = 190) -> dict:
    """
    Uma linha por VARIANTE de cada setor, no eixo do relogio.

    Duas escolhas que mudam a leitura:

    * o eixo horizontal e a hora do dia, nao a posicao na sequencia. Setores
      comecam em POPIDs diferentes e tem turnos diferentes; so o relogio
      permite dizer "as 10h todos apertam ao mesmo tempo".
    * o eixo vertical e % de UM posto. 100% e o posto saturado, para qualquer
      variante de qualquer setor.
    """
    series = []
    faixas = []          # uma trilha de gargalos por setor
    for indice, linha in enumerate(l for l in linhas if l.get("resultado")):
        cor_setor = CORES_SETORES[indice % len(CORES_SETORES)]
        por_variante = demanda_por_variante(linha)
        if not por_variante:
            continue

        relogios = _relogios(linha, len(linha["resultado"].demanda))
        faixas.append({
            "setor": linha["setor"],
            "cor": cor_setor,
            "gargalos": [
                {"a": g["a"], "b": g["b"],
                 "inicio": relogios[min(g["a"], len(relogios) - 1)],
                 "fim": relogios[min(g["b"], len(relogios) - 1)],
                 "severidade": g["severidade"],
                 "regra": g["regra_nome"]}
                for g in (linha.get("gargalos") or {}).get("gargalos", [])
            ],
        })
        # Com horario de falta, cada veiculo divide pela capacidade DAQUELE
        # momento — e o que faz o pico crescer visivelmente durante a janela
        # de menos gente, mesmo que a demanda bruta nao tenha mudado.
        serie_capacidade = linha.get("capacidade_serie")
        capacidade = max(linha.get("capacidade") or 100.0, 1.0)
        if serie_capacidade:
            brutos = [d / max(c, 1.0) * 100.0
                     for d, c in zip(linha["resultado"].demanda, serie_capacidade)]
        else:
            brutos = [v / capacidade * 100.0 for v in linha["resultado"].demanda]
        valores = cpm_core.suavizar(brutos, 5)
        # Qual variante puxa o pico esta no cartao do setor; aqui o que
        # interessa e QUANDO a funcao aperta ao longo do dia.
        pior = max(por_variante,
                   key=lambda item: max(item["valores"], default=0.0))
        series.append({
            "setor": linha["setor"],
            "rotulo": linha["setor"].nome,
            "cor": cor_setor,
            "valores": valores,
            "relogios": relogios,
            "pico": max(valores) if valores else 0.0,
            "media": sum(valores) / len(valores) if valores else 0.0,
            "capacidade": capacidade,
            "variante_critica": pior["variante"].nome if por_variante else "",
        })

    if not series:
        return {"series": [], "vazio": True}

    inicio = min(s["relogios"][0] for s in series)
    fim = max(s["relogios"][-1] for s in series)
    duracao = max(fim - inicio, 1)

    teto = max(max(s["pico"] for s in series) * 1.12, 120.0)
    margem_esq = 40
    horas = duracao / 3600.0
    largura = int(margem_esq + max(horas * largura_por_hora, 600))
    util = largura - margem_esq - 10

    def x_de(relogio):
        return margem_esq + (relogio - inicio) / duracao * util

    for serie in series:
        pontos = []
        for valor, relogio in zip(serie["valores"], serie["relogios"]):
            y = altura - (valor / teto) * altura
            pontos.append(f"{x_de(relogio):.1f},{y:.1f}")
        serie["pontos"] = " ".join(pontos)

        # Hora do pico: e a informacao que a tela responde de imediato.
        indice = serie["valores"].index(serie["pico"])
        serie["hora_pico"] = cpm_core._hhmm(serie["relogios"][indice])
        serie["x_pico"] = x_de(serie["relogios"][indice])
        serie["y_pico"] = altura - (serie["pico"] / teto) * altura

    # Trilhas de gargalo: uma faixa por setor, abaixo do grafico. Desenhar por
    # cima das curvas escondia justamente o que se quer comparar.
    altura_trilha = 11
    espaco_trilha = 3
    topo_trilhas = altura + 20
    faixas = [f for f in faixas if f["gargalos"]]
    for ordem, faixa in enumerate(faixas):
        faixa["y"] = topo_trilhas + ordem * (altura_trilha + espaco_trilha)
        for gargalo in faixa["gargalos"]:
            x1 = x_de(gargalo["inicio"])
            x2 = x_de(gargalo["fim"])
            gargalo["x"] = x1
            # gargalo de um veiculo so ficaria invisivel; 3px e o minimo legivel
            gargalo["largura"] = max(x2 - x1, 3.0)
            gargalo["opacidade"] = 0.95 if gargalo["severidade"] == "ALTA" else 0.45

    altura_total = (topo_trilhas + len(faixas) * (altura_trilha + espaco_trilha)
                    + 16)

    def y_de(valor):
        return altura - (valor / teto) * altura

    marcas = [v for v in (0, 50, 100, 150, 200, 250, 300) if v <= teto]

    # marcas de hora cheia dentro do intervalo
    horas_marcadas = []
    marca = (int(inicio // 3600) + 1) * 3600
    while marca <= fim:
        horas_marcadas.append({"rotulo": cpm_core._hhmm(marca), "x": x_de(marca)})
        marca += 3600

    return {
        "series": series, "vazio": False,
        "largura": largura, "altura": altura, "margem_esq": margem_esq,
        "marcas": [{"valor": v, "y": y_de(v)} for v in marcas],
        "horas": horas_marcadas,
        "faixas": faixas, "altura_trilha": altura_trilha,
        "topo_trilhas": topo_trilhas, "altura_total": altura_total,
        "inicio": cpm_core._hhmm(inicio), "fim": cpm_core._hhmm(fim),
        "total_gargalos": sum(len(f["gargalos"]) for f in faixas),
    }


def analisar_setor(df, setor, popid, producao, faltas=0, takt=None,
                   turno=None, janelas_falta=None) -> dict:
    """Roda a analise completa de um setor. Devolve sempre um dicionario."""
    turno = turno or TURNOS[TURNO_PADRAO]
    inicio_turno, fim_turno = horario_do_setor(setor, turno)
    base = {
        "setor": setor, "popid": str(popid or "").strip(),
        "producao": str(producao or "").strip(),
        "faltas": faltas, "erro": None, "resultado": None,
        "gargalos": None, "curva": None, "alertas": [],
        "turno": turno, "inicio_turno": inicio_turno, "fim_turno": fim_turno,
        "takt": takt or setor.takt_segundos,
    }

    if not base["popid"] or not base["producao"]:
        base["erro"] = "Informe POPID e produção."
        return base

    try:
        resultado = cpm_core.analisar_mix(
            df, base["popid"], base["producao"],
            takt=takt or setor.takt_segundos)
    except ValueError as erro:
        base["erro"] = str(erro)
        return base

    recorte = df.iloc[resultado.primeira_linha - 1:resultado.ultima_linha]
    try:
        dados = gargalos_mod.analisar(recorte, setor.id)
    except Exception as erro:  # noqa: BLE001
        logger.warning("gargalos de %s: %s", setor.id, erro)
        dados = {"resumo": {"total": 0, "alta": 0, "media": 0,
                            "unidades_afetadas": 0},
                 "gargalos": [], "unidades": [], "meta": {"n_unidades": 0}}

    # Absenteismo fixo do turno inteiro (comportamento de sempre) — usado
    # para os alertas e o cartao do setor: o pior momento do dia e o que
    # decide se e preciso "recurso", entao ele manda no numero exibido.
    faltas_pico = faltas
    capacidade_serie = None
    faltas_serie = None

    if janelas_falta:
        quantidade = len(resultado.demanda)
        relogios = _relogios(base, quantidade)
        janelas_abs = []
        for inicio_j, fim_j, faltas_j in janelas_falta:
            recortada = cpm_core._pausas_em_segundos(
                inicio_turno, fim_turno, [(inicio_j, fim_j)])
            if recortada:
                p1, p2 = recortada[0]
                janelas_abs.append((p1, p2, faltas_j))

        faltas_serie = []
        for relogio in relogios:
            no_momento = [fj for p1, p2, fj in janelas_abs if p1 <= relogio <= p2]
            faltas_serie.append(max(no_momento) if no_momento else 0)

        # resultado.postos_originais so e preenchido DENTRO de
        # aplicar_absenteismo, que ainda nao rodou aqui — o total de
        # variantes atuais e o mesmo numero, so que disponivel mais cedo.
        total_postos = len(resultado.variantes)
        capacidade_serie = [
            max(cpm_core.postos_ativos_por_faltas(
                f, total_postos, setor.tabela_postos), 1) * 100.0
            for f in faltas_serie
        ]
        faltas_pico = max(faltas_serie) if faltas_serie else 0
        base["relogios"] = relogios

    cpm_core.aplicar_absenteismo(
        resultado, faltas_pico, dados["resumo"].get("total", 0),
        setor.tabela_postos, setor.gargalos_varios)

    capacidade = max(resultado.postos_ativos, 1) * 100.0
    base.update({
        "resultado": resultado,
        "gargalos": dados,
        "curva": _curva_svg(resultado.demanda, capacidade),
        "capacidade": capacidade,
        "capacidade_serie": capacidade_serie,   # None = capacidade fixa
        "faltas_serie": faltas_serie,
        "janelas_falta": janelas_falta,
        "alertas": resultado.alertas,
        "membros_de": lambda g, d=dados: gargalos_mod.descrever_membros(d, g),
    })
    return base


def momentos_criticos(linhas: list[dict]) -> list[dict]:
    """
    O cruzamento que o Eduardo pediu: onde a fabrica aperta MAIS, juntando
    montagem complexa (gargalo) com menos gente (janela de falta).

    So entra aqui o setor que tem horario de falta configurado — sem isso nao
    ha "momento" de falta para cruzar, so um numero unico do turno inteiro.
    Um gargalo que cai dentro de uma janela reduzida diferente da capacidade
    plena do setor e um "momento critico": a mesma sequencia que o setor
    aguentaria num dia normal, aqui pega sem gente suficiente.
    """
    momentos = []
    for linha in linhas:
        serie_capacidade = linha.get("capacidade_serie")
        resultado = linha.get("resultado")
        if not serie_capacidade or not resultado:
            continue

        relogios = linha.get("relogios") or _relogios(linha, len(resultado.demanda))
        cheia = resultado.postos_originais * 100.0
        faltas_serie = linha.get("faltas_serie") or []
        gargalos_lista = (linha.get("gargalos") or {}).get("gargalos", [])

        for gargalo in gargalos_lista:
            a = max(0, min(gargalo["a"], len(serie_capacidade) - 1))
            b = max(0, min(gargalo["b"], len(serie_capacidade) - 1))
            pior_capacidade = min(serie_capacidade[a:b + 1])
            if pior_capacidade >= cheia:
                continue  # gargalo aconteceu com o setor completo — nao e o cruzamento

            faltas_no_trecho = max(faltas_serie[a:b + 1]) if faltas_serie else 0
            pico_no_trecho = max(resultado.demanda[a:b + 1])
            momentos.append({
                "setor": linha["setor"],
                "severidade": gargalo["severidade"],
                "regra": gargalo["regra_nome"],
                "detalhe": gargalo["detalhe"],
                "a": a + 1, "b": b + 1,
                "hora": (f"{cpm_core._hhmm(relogios[a])}–"
                        f"{cpm_core._hhmm(relogios[b])}"),
                "faltas": faltas_no_trecho,
                "postos_no_trecho": pior_capacidade / 100.0,
                "postos_cheios": resultado.postos_originais,
                "pico_local": pico_no_trecho,
                "sobrecarga": pico_no_trecho / pior_capacidade * 100.0,
            })

    # pior sobrecarga primeiro — e o "momento mais complicado do dia" que se
    # quer ver logo no topo.
    momentos.sort(key=lambda m: m["sobrecarga"], reverse=True)
    return momentos


def consolidar(linhas: list[dict]) -> dict:
    """Numeros da fábrica inteira e a lista de gargalos de todos os setores."""
    validos = [l for l in linhas if l.get("resultado")]

    gargalos = []
    for linha in validos:
        for gargalo in linha["gargalos"].get("gargalos", []):
            gargalos.append({
                "setor": linha["setor"],
                "severidade": gargalo["severidade"],
                "regra": gargalo["regra_nome"],
                "a": gargalo["a"] + 1,
                "b": gargalo["b"] + 1,
                "detalhe": gargalo["detalhe"],
                "membros": linha["membros_de"](gargalo),
            })

    # Alta primeiro; dentro da severidade, na ordem da sequencia.
    gargalos.sort(key=lambda g: (0 if g["severidade"] == "ALTA" else 1, g["a"]))

    postos_ativos = sum(l["resultado"].postos_ativos for l in validos)
    postos_totais = sum(l["resultado"].postos_originais for l in validos)

    return {
        "setores_analisados": len(validos),
        "setores_com_erro": len(linhas) - len(validos),
        "gargalos": gargalos,
        "total_gargalos": len(gargalos),
        "alta": sum(1 for g in gargalos if g["severidade"] == "ALTA"),
        "media": sum(1 for g in gargalos if g["severidade"] == "MEDIA"),
        "postos_ativos": postos_ativos,
        "postos_totais": postos_totais,
        "faltas": sum(int(l.get("faltas") or 0) for l in validos),
        "alertas": [{"setor": l["setor"], **a}
                    for l in validos for a in l["alertas"]],
        "variantes_criticas": sum(
            1 for l in validos
            for v in l["resultado"].variantes.values()
            if v.total >= LIMITE_CRITICO),
    }


def barras_de_carga(linhas: list[dict]) -> list[dict]:
    """Uma barra por variante, agrupada por setor, para o gráfico da tela."""
    barras = []
    for linha in linhas:
        resultado = linha.get("resultado")
        if not resultado:
            continue
        for variante in resultado.variantes.values():
            barras.append({
                "setor": linha["setor"].nome,
                "variante": variante.nome,
                "total": variante.total,
                "faixa": _faixa(variante.total),
                # 150% ocupa a largura toda; acima disso a barra satura.
                "largura": min(variante.total / 150.0 * 100.0, 100.0),
            })
    return barras

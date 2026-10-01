"""
app.py - Versao web do MENUMIX/CPM.

Reaproveita `cpm_core.py` sem nenhuma alteracao: toda a regra de negocio
continua no mesmo lugar, muda so a camada de apresentacao (HTML em vez de
Tkinter).

Como rodar:
    pip install flask pandas openpyxl
    python app.py
    -> abre http://127.0.0.1:5000

Rodando na propria maquina Windows, o restante do app (Outlook via win32com)
continua funcionando normalmente e pode alimentar esta mesma tela.
"""

from __future__ import annotations

import argparse
import logging
import os
import pickle
import secrets
import socket
import tempfile
import threading
import time
import traceback
import uuid
from datetime import datetime

import pandas as pd
from flask import (
    Flask, Response, abort, g, redirect, render_template, request, send_file,
    url_for,
)

import catalogo as catalogo_mod
from catalogo import ErroCatalogo
import armazenamento
import filtro_mix as filtro_mix_mod
import gargalos as gargalos_mod
import painel as painel_mod
import relatorio_fabrica as pdf_fabrica
import processadores as processadores_mod
import setores as setores_mod
import tags_mona as tags_mona_mod
import relatorio_pdf as pdf_mod
import cpm_core
from cpm_core import (
    FIM_TURNO_PADRAO, INICIO_TURNO_PADRAO, analisar_mix, aplicar_absenteismo, concatenar_planilhas,
    concatenar_planilhas_em_memoria, exportar_texto, exportar_tsv,
    ordenar_por_data,
)
from outlook_downloader import COM_DISPONIVEL, OutlookDownloader, OutlookIndisponivel

# Logs estruturados em vez de print(): aparecem com nivel e hora nos logs do
# Railway (ou no terminal, local), e dao pra filtrar por WARNING/ERROR sem
# precisar pedir print de tela pra quem esta usando o sistema.
logging.basicConfig(
    level=os.environ.get("MENUMIX_LOG_LEVEL", "INFO").upper(),
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
logger = logging.getLogger("menumix")

app = Flask(__name__)

# Sem isso, um upload gigante (ou um ataque simples) consome memoria e disco
# sem limite. 25 MB cobre folgado uma planilha real de MIX.
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

PASTA_SESSAO = armazenamento.pasta_sessao()

DIAS_PADRAO = 7
MAX_PLANILHAS = 3

baixador = OutlookDownloader()

# O Outlook nao gosta de varias chamadas COM ao mesmo tempo. Com mais de uma
# pessoa usando a tela, as buscas entram em fila.
trava_outlook = threading.Lock()


# ---------------------------------------------------------------------------
# Login (HTTP Basic) — so entra em vigor se MENUMIX_USUARIO/MENUMIX_SENHA
# estiverem definidas. Sem essas variaveis (uso local, na rede interna), o
# app funciona sem pedir senha, como sempre.
# ---------------------------------------------------------------------------
USUARIO_ESPERADO = os.environ.get("MENUMIX_USUARIO", "").strip()
SENHA_ESPERADA = os.environ.get("MENUMIX_SENHA", "")


@app.before_request
def exigir_login():
    if not USUARIO_ESPERADO or not SENHA_ESPERADA:
        return None  # login desligado: nenhuma das duas variaveis foi definida
    auth = request.authorization
    autenticado = (
        auth is not None
        and secrets.compare_digest(auth.username or "", USUARIO_ESPERADO)
        and secrets.compare_digest(auth.password or "", SENHA_ESPERADA)
    )
    if not autenticado:
        return Response(
            "Login necessario.", 401,
            {"WWW-Authenticate": 'Basic realm="NEXUS-6"'},
        )
    return None

FERRAMENTAS = [
    {
        "rota": "painel",
        "nome": "Painel da fábrica",
        "descricao": "Todos os setores numa tela: carga, gargalos e a curva de "
                     "demanda de cada um.",
    },
    {
        "rota": "analise",
        "nome": "Análise de MIX",
        "descricao": "Carga de trabalho por variante a partir do POPID, "
                     "com CPM MIX e exportação.",
    },
    {
        "rota": "filtro_mix",
        "nome": "Filtro MIX",
        "descricao": "FA2, FA1.2, Fibras e Função 4 — formata o MIX a partir "
                     "do Outlook ou de um arquivo enviado.",
    },
    {
        "rota": "gargalos_tela",
        "nome": "Gargalos da sequência",
        "descricao": "Onde a sequência de produção concentra carga, "
                     "a partir do POPID e da produção do dia.",
    },
    {
        "rota": "tarefas",
        "nome": "Tarefas das variantes",
        "descricao": "Cadastro das regras que alimentam o cálculo de carga.",
    },
    {
        "rota": "tela_setores",
        "nome": "Setores",
        "descricao": "Takt, turno, pausas e postos de cada setor. Cada um tem "
                     "catálogo e regras próprios.",
    },
    # A ajuda nao entra como card: o botao AJUDA, no canto superior direito
    # do menu, leva ao mesmo lugar. Para voltar a mostrar o card, basta
    # devolver aqui uma entrada com "rota": "ajuda".
]


# ---------------------------------------------------------------------------
# Cache simples do DataFrame concatenado, para nao reenviar as planilhas
# a cada mudanca de POPID.
# ---------------------------------------------------------------------------
COOKIE_SETOR = "menumix_setor"
COOKIE_TOKEN = "menumix_token"
COOKIE_POPID = "menumix_popid"
COOKIE_PRODUCAO = "menumix_producao"
COOKIE_DIAS = 7
HORAS_VALIDADE = 24


def limpar_cache_antigo(horas: int = HORAS_VALIDADE) -> None:
    """Remove os arquivos de sessão mais velhos que `horas`."""
    limite = time.time() - horas * 3600
    for nome in os.listdir(PASTA_SESSAO):
        caminho = os.path.join(PASTA_SESSAO, nome)
        try:
            if os.path.isfile(caminho) and os.path.getmtime(caminho) < limite:
                os.remove(caminho)
        except OSError:
            pass


def setor_atual() -> "setores_mod.Setor":
    """
    Setor em uso: o do formulario/URL, senao o do cookie, senao o primeiro.

    O setor acompanha a navegacao porque tudo o que a tela mostra — catalogo,
    regras, takt, turno — vem dele.
    """
    pedido = (request.values.get("setor")
              or request.cookies.get(COOKIE_SETOR)
              or "")
    setor = setores_mod.obter(pedido)
    g.setor_para_cookie = setor.id
    return setor


def token_atual() -> str:
    """
    Token da planilha carregada.

    Procura no formulário, na URL e, por último, no cookie. O cookie é o que
    faz a planilha sobreviver à troca de tela: antes o token viajava só em
    campo escondido, então qualquer link para outra aba perdia os dados.
    """
    return (request.values.get("token")
            or request.cookies.get(COOKIE_TOKEN)
            or "")


def lembrar_token(token: str) -> None:
    """Marca o token para virar cookie no fim da requisição."""
    if token:
        g.token_para_cookie = token


def lembrar_intervalo(popid: str, producao: str) -> None:
    g.intervalo_para_cookie = (popid or "", producao or "")


@app.after_request
def gravar_cookie(resposta):
    token = getattr(g, "token_para_cookie", None)
    if token and request.cookies.get(COOKIE_TOKEN) != token:
        resposta.set_cookie(COOKIE_TOKEN, token, max_age=COOKIE_DIAS * 86400,
                            samesite="Lax")

    setor = getattr(g, "setor_para_cookie", None)
    if setor and request.cookies.get(COOKIE_SETOR) != setor:
        resposta.set_cookie(COOKIE_SETOR, setor, max_age=COOKIE_DIAS * 86400,
                            samesite="Lax")

    painel_cookie = getattr(g, "painel_para_cookie", None)
    if painel_cookie is not None:
        resposta.set_cookie(COOKIE_PAINEL, painel_cookie,
                            max_age=COOKIE_DIAS * 86400, samesite="Lax")

    intervalo = getattr(g, "intervalo_para_cookie", None)
    if intervalo:
        popid, producao = intervalo
        for nome, valor in ((COOKIE_POPID, popid), (COOKIE_PRODUCAO, producao)):
            if valor and request.cookies.get(nome) != valor:
                resposta.set_cookie(nome, valor, max_age=COOKIE_DIAS * 86400,
                                    samesite="Lax")
    return resposta


def guardar_dataframe(df: pd.DataFrame) -> str:
    limpar_cache_antigo()
    token = uuid.uuid4().hex
    with open(os.path.join(PASTA_SESSAO, f"{token}.pkl"), "wb") as arquivo:
        pickle.dump(df, arquivo)
    lembrar_token(token)
    return token


def recuperar_dataframe(token: str) -> pd.DataFrame | None:
    if not token:
        return None
    caminho = os.path.join(PASTA_SESSAO, f"{token}.pkl")
    if not os.path.exists(caminho):
        return None
    with open(caminho, "rb") as arquivo:
        return pickle.load(arquivo)


def _resolver_texto_mona(form) -> str | None:
    """
    Le o texto a marcar quando a MONA bater um POPID: escolhido no menu de
    tags fixas (tags_mona.carregar(), editável em /tags-mona) ou digitado a
    mao quando a opcao
    for "Outro" — assim, pras tags ja conhecidas, ninguem precisa digitar
    nada, so escolher.
    """
    selecionado = (form.get("texto_mona_fixo") or "").strip()
    if selecionado == "__outro__":
        return (form.get("texto_mona_outro") or "").strip() or None
    return selecionado or None


def _aplicar_mona_em_opcionais(df: pd.DataFrame, request_atual, pasta: str | None = None,
                               setor: str | None = None):
    """
    Le a caixa "Cruzar com MONA" e o arquivo/texto que vierem no formulario e,
    se marcada, devolve o df com o texto da MONA juntado em OPCIONAIS (mesmo
    cruzamento usado no Filtro MIX). Devolve (df, erro, aviso, texto_mona,
    nome_mona): df inalterado e um texto de erro quando algo falta ou a
    planilha nao bate no formato; `aviso` traz quantos POPIDs bateram, pra dar
    visibilidade sem precisar adivinhar se o cruzamento realmente encontrou
    algo; `texto_mona` e o texto usado (lista simples) — repassado pra virar
    contador dinamico no CPM MIX, ja que so ele identifica o que foi cruzado;
    `nome_mona` e o nome do arquivo da planilha MONA enviada, pra mostrar na
    tela qual planilha MONA esta em uso.
    """
    if request_atual.form.get("usar_mona") != "1":
        return df, None, None, None, None

    enviado_mona = request_atual.files.get("arquivo_mona_extra")
    if enviado_mona is None or not enviado_mona.filename:
        return df, "Marque Mona: selecione a planilha MONA para cruzar.", None, None, None
    if not enviado_mona.filename.lower().endswith(processadores_mod.EXTENSOES):
        return df, "Planilha MONA: envie um arquivo .xlsx, .xlsm ou .xls.", None, None, None
    if "OPCIONAIS" not in df.columns:
        return df, "A planilha carregada não tem a coluna OPCIONAIS para cruzar com a MONA.", None, None, None

    pasta_mona = pasta or os.path.join(PASTA_SESSAO, uuid.uuid4().hex)
    os.makedirs(pasta_mona, exist_ok=True)
    caminho_mona = os.path.join(pasta_mona, os.path.basename(enviado_mona.filename))
    enviado_mona.save(caminho_mona)
    texto_mona = _resolver_texto_mona(request_atual.form)

    try:
        from mix_processor_4 import cruzar_com_mona
        df = df.copy()
        df = cruzar_com_mona(df, caminho_mona, coluna_saida="_MONA_TEXTO",
                             texto_lista_simples=texto_mona, setor=setor)
    except Exception as erro:  # noqa: BLE001
        return df, f"Erro ao cruzar com a MONA: {erro}", None, None, None

    encontrados = int(df["_MONA_TEXTO"].notna().sum())
    popids_batidos = (df.loc[df["_MONA_TEXTO"].notna(), "POPID"]
                      .astype(str).head(15).tolist())
    aviso = (f"Cruzamento com MONA: {encontrados} de {len(df)} linha(s) "
            f"da planilha carregada bateram com a MONA.")
    if popids_batidos:
        aviso += " POPIDs: " + ", ".join(popids_batidos)
        if encontrados > len(popids_batidos):
            aviso += "…"

    def _juntar_com_mona(linha):
        partes = []
        for valor in (linha.get("OPCIONAIS"), linha.get("_MONA_TEXTO")):
            texto = str(valor).strip() if pd.notna(valor) else ""
            if texto and texto.lower() != "nan":
                partes.append(texto)
        return " | ".join(partes)

    df["OPCIONAIS"] = df.apply(_juntar_com_mona, axis=1)
    df = df.drop(columns=["_MONA_TEXTO"], errors="ignore")
    return df, None, aviso, texto_mona, enviado_mona.filename


def salvar_uploads(arquivos) -> list[str]:
    """Grava em uma subpasta unica, para o arquivo manter o nome original
    (o nome aparece nos avisos de leitura da planilha)."""
    pasta = os.path.join(PASTA_SESSAO, uuid.uuid4().hex)
    caminhos = []
    for enviado in arquivos:
        if not enviado or not enviado.filename:
            continue
        os.makedirs(pasta, exist_ok=True)
        destino = os.path.join(pasta, os.path.basename(enviado.filename))
        enviado.save(destino)
        caminhos.append(destino)
    return caminhos


def guardar_lista_arquivos(token: str, arquivos: list[dict]) -> None:
    resumo = [{k: v for k, v in a.items() if k != "recebido"} for a in arquivos]
    with open(os.path.join(PASTA_SESSAO, f"{token}.arquivos.pkl"), "wb") as saida:
        pickle.dump(resumo, saida)


def recuperar_lista_arquivos(token: str) -> list[dict]:
    caminho = os.path.join(PASTA_SESSAO, f"{token}.arquivos.pkl")
    if not token or not os.path.exists(caminho):
        return []
    with open(caminho, "rb") as entrada:
        return pickle.load(entrada)


def guardar_texto_mona(token: str, texto: str | None) -> None:
    """Lembra qual texto foi usado no cruzamento com a MONA desta planilha,
    pra reaparecer como contador no CPM MIX mesmo quando a analise seguinte
    reaproveita o cache sem reenviar arquivos."""
    if not token or not texto:
        return
    with open(os.path.join(PASTA_SESSAO, f"{token}.mona.txt"), "w", encoding="utf-8") as saida:
        saida.write(texto)


def recuperar_texto_mona(token: str) -> str | None:
    caminho = os.path.join(PASTA_SESSAO, f"{token}.mona.txt")
    if not token or not os.path.exists(caminho):
        return None
    with open(caminho, "r", encoding="utf-8") as entrada:
        texto = entrada.read().strip()
    return texto or None


def guardar_nome_mona(token: str, nome: str | None) -> None:
    """Lembra o nome do arquivo da planilha MONA usada no cruzamento desta
    planilha, pra mostrar na tela qual MONA esta carregada mesmo depois de
    reaproveitar o cache sem reenviar arquivos."""
    caminho = os.path.join(PASTA_SESSAO, f"{token}.mona_nome.txt")
    if not token:
        return
    if not nome:
        # sem MONA nesta analise: nao deixa um nome antigo enganando a tela
        if os.path.exists(caminho):
            os.remove(caminho)
        return
    with open(caminho, "w", encoding="utf-8") as saida:
        saida.write(nome)


def recuperar_nome_mona(token: str) -> str | None:
    caminho = os.path.join(PASTA_SESSAO, f"{token}.mona_nome.txt")
    if not token or not os.path.exists(caminho):
        return None
    with open(caminho, "r", encoding="utf-8") as entrada:
        nome = entrada.read().strip()
    return nome or None


def buscar_no_outlook(acao: str, contexto: dict) -> list[dict]:
    with trava_outlook:
        return _buscar_no_outlook(acao, contexto)


def _buscar_no_outlook(acao: str, contexto: dict) -> list[dict]:
    if acao == "outlook_recentes":
        try:
            dias = int(contexto["dias"])
        except ValueError:
            raise ValueError("Dias deve ser um número inteiro.")
        return baixador.baixar_recentes(dias=dias, maximo=MAX_PLANILHAS)

    if not contexto["data_inicio"] or not contexto["data_fim"]:
        raise ValueError("Informe as datas inicial e final do período.")

    inicio = datetime.strptime(contexto["data_inicio"], "%Y-%m-%d")
    fim = datetime.strptime(contexto["data_fim"], "%Y-%m-%d").replace(
        hour=23, minute=59, second=59
    )
    if inicio > fim:
        raise ValueError("A data inicial não pode ser maior que a final.")
    return baixador.baixar_por_periodo(inicio, fim, maximo=MAX_PLANILHAS)


def _presenteismo_do_form(form) -> tuple[bool, str, list[dict]]:
    """
    Le o checkbox de presenteismo e os pares hora-inicio/hora-fim gerados
    dinamicamente na tela. Cada par e uma pessoa que esta no turno mas sai por
    um periodo — diferente do absenteismo, que falta o turno inteiro.

    Devolve (ativo, texto_da_quantidade, [{'inicio':..,'fim':..}, ...]) —
    exatamente o que o template precisa para redesenhar a tela apos o POST.
    """
    ativo = form.get("presenteismo") == "sim"
    qtd_texto = (form.get("presenteismo_qtd") or "").strip()
    if not ativo:
        return False, qtd_texto, []

    try:
        quantidade = min(max(int(qtd_texto), 0), 30)
    except ValueError:
        quantidade = 0

    janelas = []
    for i in range(1, quantidade + 1):
        inicio = (form.get(f"presenteismo_inicio_{i}") or "").strip()
        fim = (form.get(f"presenteismo_fim_{i}") or "").strip()
        if inicio and fim:
            janelas.append({"inicio": inicio, "fim": fim})
    return ativo, qtd_texto, janelas


def _presenteismo_por_posicao(janelas: list[dict], relogios: list[float],
                              inicio_turno: str, fim_turno: str) -> list[int]:
    """
    Quantas pessoas em presenteismo estao FORA em cada posicao da sequencia.

    Cada janela e uma pessoa; janelas que se sobrepoem SOMAM (duas pessoas
    fora ao mesmo tempo pesam o dobro) — diferente do horario de falta do
    painel da fabrica, onde uma janela ja declara o total de ausentes.
    """
    absolutas = []
    for janela in janelas:
        recorte = cpm_core._pausas_em_segundos(
            inicio_turno, fim_turno, [(janela["inicio"], janela["fim"])])
        if recorte:
            absolutas.append(recorte[0])

    return [sum(1 for p1, p2 in absolutas if p1 <= relogio <= p2)
            for relogio in relogios]


def _momentos_de_presenteismo(resultado, dados_gargalos, faltas_serie,
                              total_postos, tabela_postos) -> list[dict]:
    """
    O cruzamento: gargalo (montagem complexa numa janela da sequência) que
    caiu num trecho onde o presenteísmo reduziu a capacidade do posto.

    Mesma lógica do painel da fábrica (ver painel.momentos_criticos), só que
    aqui para o setor único desta tela, com a série de faltas vinda do
    presenteísmo em vez de um horário de falta do turno inteiro.
    """
    if not any(faltas_serie):
        return []

    capacidade_serie = [
        max(cpm_core.postos_ativos_por_faltas(f, total_postos, tabela_postos), 1)
        * 100.0
        for f in faltas_serie
    ]
    cheia = total_postos * 100.0

    momentos = []
    for gargalo in (dados_gargalos or {}).get("gargalos", []):
        a = max(0, min(gargalo["a"], len(capacidade_serie) - 1))
        b = max(0, min(gargalo["b"], len(capacidade_serie) - 1))
        pior_capacidade = min(capacidade_serie[a:b + 1])
        if pior_capacidade >= cheia:
            continue  # o gargalo caiu com o posto completo — nao e o cruzamento

        pico_local = max(resultado.demanda[a:b + 1])
        momentos.append({
            "severidade": gargalo["severidade"],
            "regra": gargalo["regra_nome"],
            "detalhe": gargalo["detalhe"],
            "a": a + 1, "b": b + 1,
            "fora_no_trecho": max(faltas_serie[a:b + 1]),
            "postos_no_trecho": pior_capacidade / 100.0,
            "postos_cheios": total_postos,
            "sobrecarga": pico_local / pior_capacidade * 100.0,
        })

    momentos.sort(key=lambda m: m["sobrecarga"], reverse=True)
    return momentos


def executar_analise(df, contexto: dict, setor=None, contador_mona: str | None = None) -> dict:
    setor = setor or setor_atual()
    if not contexto["popid"] or not contexto["producao"]:
        contexto["erro"] = "Informe POPID e produção do dia."
        return contexto

    try:
        takt = float((contexto["takt"] or "0").replace(",", "."))
    except ValueError:
        contexto["erro"] = "O takt deve ser um número de segundos."
        return contexto
    if takt <= 0:
        contexto["erro"] = "O takt deve ser maior que zero."
        return contexto

    try:
        faltas = int(float((contexto["faltas"] or "0").replace(",", ".")))
    except ValueError:
        contexto["erro"] = "O absenteísmo deve ser um número de pessoas."
        return contexto
    if faltas < 0:
        contexto["erro"] = "O absenteísmo não pode ser negativo."
        return contexto

    try:
        resultado = analisar_mix(df, contexto["popid"], contexto["producao"],
                                 takt=takt, contador_mona=contador_mona,
                                 setor=setor.id)
    except ValueError as erro:
        contexto["erro"] = str(erro)
        return contexto

    lembrar_intervalo(contexto["popid"], contexto["producao"])

    # Gargalos primeiro: os alertas de absenteísmo dependem de quantos são.
    recorte = df.iloc[resultado.primeira_linha - 1:resultado.ultima_linha]
    n_gargalos = 0
    try:
        dados = gargalos_mod.analisar(recorte, setor.id)
        gargalos_mod.anotar_horarios(dados, setor)
        n_gargalos = dados["resumo"].get("total", 0)
        contexto["gargalos_dados"] = dados
        contexto["membros_de"] = lambda g: gargalos_mod.descrever_membros(dados, g)
    except Exception as erro:  # noqa: BLE001
        logger.warning("gargalos: falha na análise (%s)", erro)
        contexto["gargalos_dados"] = None

    # Presenteismo: gente que esta no turno mas sai por um periodo. Some ao
    # absenteismo fixo SO durante a janela de cada pessoa — o resto do turno
    # continua com o posto completo. E o que faz o calculo de gargalo
    # enxergar o momento exato em que a montagem complexa bateu com menos
    # gente, em vez de aplicar a falta o dia inteiro.
    faltas_para_alerta = faltas
    if contexto.get("presenteismo_janelas"):
        total_postos = len(resultado.variantes)
        pausas = setor.pausas_tuplas or None
        ritmo = cpm_core.ritmo_do_turno(
            len(resultado.demanda), contexto["inicio_turno"],
            contexto["fim_turno"], pausas)
        relogios = [
            cpm_core.relogio_apos(i * ritmo, contexto["inicio_turno"],
                                  contexto["fim_turno"], pausas)
            for i in range(len(resultado.demanda))
        ]
        presentes_fora = _presenteismo_por_posicao(
            contexto["presenteismo_janelas"], relogios,
            contexto["inicio_turno"], contexto["fim_turno"])
        faltas_serie = [faltas + fora for fora in presentes_fora]

        contexto["presenteismo_momentos"] = _momentos_de_presenteismo(
            resultado, dados, faltas_serie, total_postos, setor.tabela_postos)

        # O pior instante do turno e o que decide se falta recurso — mesma
        # logica usada no painel da fabrica para o horario de falta.
        faltas_para_alerta = max(faltas_serie) if faltas_serie else faltas

    aplicar_absenteismo(resultado, faltas_para_alerta, n_gargalos,
                        setor.tabela_postos, setor.gargalos_varios)

    contexto["resultado"] = resultado
    contexto["export_tsv"] = exportar_tsv(resultado)
    contexto["export_texto"] = exportar_texto(resultado)
    contexto["pdf_disponivel"] = pdf_mod.DISPONIVEL
    contexto["params_pdf"] = {
        "token": contexto.get("token", ""),
        "popid": contexto["popid"],
        "producao": contexto["producao"],
        "absenteismo": contexto["faltas"] or "0",
        "takt": contexto["takt"],
        "inicio_turno": contexto["inicio_turno"],
        "fim_turno": contexto["fim_turno"],
        "setor": setor.id,
    }
    return contexto


# ---------------------------------------------------------------------------
# Rota unica
# ---------------------------------------------------------------------------
@app.route("/")
def menu():
    """Menu inicial: so leva ate as telas."""
    return render_template("menu.html", ferramentas=FERRAMENTAS,
                           setor=setor_atual(), setores=setores_mod.carregar())


# ---------------------------------------------------------------------------
# Setores
# ---------------------------------------------------------------------------
@app.route("/setores", methods=["GET", "POST"])
def tela_setores():
    setor = setor_atual()
    mensagem = erro = None
    editando = None

    if request.method == "POST":
        acao = request.form.get("acao", "")
        try:
            if acao == "salvar":
                salvo = setores_mod.salvar_setor(
                    request.form, request.form.get("setor_id") or None)
                mensagem = f"Setor {salvo.nome} salvo."
            elif acao == "remover":
                setores_mod.remover(request.form.get("setor_id", ""))
                mensagem = "Setor removido."
            elif acao == "salvar_padroes":
                campos = ("producao", "takt_segundos", "inicio_turno",
                          "fim_turno", "pausas", "postos",
                          "postos_por_faltas", "gargalos_varios")
                setores_mod.salvar_padroes(
                    {c: request.form[c] for c in campos if c in request.form})
                mensagem = "Padrões da fábrica salvos."
        except setores_mod.ErroSetor as falha:
            erro = str(falha)
            editando = setores_mod.Setor(
                id=request.form.get("setor_id", ""),
                nome=request.form.get("nome", ""),
                descricao=request.form.get("descricao", ""),
            )

    lista = setores_mod.carregar()
    padroes = setores_mod.carregar_padroes()

    if editando is None:
        pedido = request.args.get("editar", "")
        editando = next((s for s in lista if s.id == pedido), None)
        if editando is None:
            # Setor novo ja nasce com os padrões da fábrica preenchidos.
            editando = setores_mod.setor_a_partir_dos_padroes()

    # numeros de cada setor, para a tela mostrar o que ja existe la dentro
    contagens = {}
    for item in lista:
        try:
            variantes_setor = catalogo_mod.carregar(item.id)
            tarefas = sum(len(v.tarefas) for v in variantes_setor)
            n_variantes = len(variantes_setor)
        except Exception:  # noqa: BLE001
            tarefas, n_variantes = 0, 0
        try:
            regras = len(gargalos_mod.listar_regras(item.id))
        except Exception:  # noqa: BLE001
            regras = 0
        contagens[item.id] = {"tarefas": tarefas, "regras": regras,
                              "variantes": n_variantes}

    return render_template(
        "setores.html", setores=lista, setor=setor, editando=editando,
        mensagem=mensagem, erro=erro,
        pausas_texto=setores_mod.pausas_para_texto(editando.pausas),
        tabela_texto=setores_mod.tabela_para_texto(editando.postos_por_faltas),
        contagens=contagens,
        padroes=padroes,
        padroes_pausas=setores_mod.pausas_para_texto(padroes["pausas"]),
        padroes_tabela=setores_mod.tabela_para_texto(padroes["postos_por_faltas"]),
    )


# ---------------------------------------------------------------------------
# Processadores: uma tela para todas as funcoes de planilha
# ---------------------------------------------------------------------------
_processados: dict[str, str] = {}


def _valor_do_form(campo, formulario) -> object:
    """Converte o campo conforme o tipo declarado, aceitando vírgula decimal."""
    bruto = (formulario.get(campo.nome) or "").strip()
    if not bruto:
        return None
    if campo.tipo != "number":
        return bruto

    texto = bruto.replace("\u00a0", "").replace(" ", "")
    # Aceita 1.234,56 e 1,234.56 além de 88 e 88.5
    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    else:
        texto = texto.replace(",", ".")

    try:
        numero = float(texto)
    except ValueError:
        raise ValueError(f"{campo.rotulo}: '{bruto}' não é um número.")
    if campo.minimo and numero < float(campo.minimo):
        raise ValueError(f"{campo.rotulo} deve ser {campo.minimo} ou mais.")
    return int(numero) if campo.passo == "1" else numero


@app.route("/processados/<nome>")
def baixar_processado(nome):
    caminho = _processados.get(nome)
    if not caminho or not os.path.isfile(caminho):
        return redirect(url_for("menu"))
    return send_file(caminho, as_attachment=True, download_name=nome)


# ---------------------------------------------------------------------------
# Painel: a fabrica inteira numa tela
# ---------------------------------------------------------------------------
COOKIE_PAINEL = "menumix_painel"


def _faltas_do_campo(texto: str) -> tuple[int, list | None]:
    """
    O campo 'Faltas' aceita um numero (turno inteiro) ou um horario por linha
    ('10:00-12:00=2'). Um numero so continua funcionando como sempre — so
    quando ha horario a analise passa a variar a capacidade ao longo do dia.
    """
    janelas = painel_mod.janelas_de_texto(texto)
    if janelas is not None:
        return 0, janelas
    try:
        return int(float((texto or "0").replace(",", "."))), None
    except ValueError:
        return 0, None


def _pedidos_do_painel(setores_lista) -> dict:
    """POPID e produção de cada setor, do formulário ou do último uso."""
    salvos = {}
    bruto = request.cookies.get(COOKIE_PAINEL, "")
    for parte in bruto.split("|"):
        if ":" in parte:
            chave, _, valor = parte.partition(":")
            salvos[chave] = valor

    # No POST vale exatamente o que o formulario mandou: sem isso, apagar um
    # POPID nao tiraria o setor da analise — o valor lembrado voltaria.
    enviado = request.method == "POST"

    def valor(campo, chave):
        if enviado:
            return (request.form.get(campo) or "").strip()
        return salvos.get(chave, "").strip()


    # O turno vem primeiro: producao, takt e POPIDs sao guardados POR TURNO.
    # O 2o turno tem produção e ponto de corte proprios; um valor so faria a
    # troca de turno sobrescrever os numeros do outro.
    # O turno vem tambem pela URL: trocar no seletor navega em GET, e ai os
    # campos sao lidos do que estava guardado PARA AQUELE turno.
    turno_id = ((request.values.get("turno") or "").strip()
                or salvos.get("geral_turno", "").strip()
                or painel_mod.TURNO_PADRAO)

    producao = valor("producao_geral", f"q{turno_id}")
    if not producao and not enviado:
        producao = str(setores_mod.carregar_padroes().get("producao") or "")

    pedidos = {"_geral": {
        "producao": producao,
        "takt": valor("takt_geral", f"t{turno_id}"),
        "turno": turno_id,
    }}
    for setor in setores_lista:
        pedidos[setor.id] = {
            "popid": valor(f"popid_{setor.id}", f"p{turno_id}_{setor.id}"),
            "faltas": valor(f"faltas_{setor.id}", f"f{turno_id}_{setor.id}"),
        }
    return pedidos


def _guardar_pedidos(pedidos: dict) -> None:
    geral = pedidos.get("_geral", {})
    turno_id = geral.get("turno") or painel_mod.TURNO_PADRAO

    # Mantem o que ja estava guardado do OUTRO turno.
    anterior = {}
    for parte in request.cookies.get(COOKIE_PAINEL, "").split("|"):
        if ":" in parte:
            chave, _, valor = parte.partition(":")
            anterior[chave] = valor

    anterior["geral_turno"] = turno_id
    anterior[f"q{turno_id}"] = geral.get("producao", "")
    anterior[f"t{turno_id}"] = geral.get("takt", "")
    for setor_id, dados in pedidos.items():
        if setor_id == "_geral":
            continue
        anterior[f"p{turno_id}_{setor_id}"] = dados["popid"]
        # ';' no lugar de quebra de linha: o cookie e um cabecalho HTTP, uma
        # quebra de linha de verdade dentro dele quebraria o cookie inteiro.
        # O parser de horario ja trata ';' como separador de linha.
        anterior[f"f{turno_id}_{setor_id}"] = (
            dados["faltas"].replace("\r\n", "\n").replace("\n", ";"))

    g.painel_para_cookie = "|".join(f"{c}:{v}" for c, v in anterior.items())


@app.route("/painel", methods=["GET", "POST"])
def painel():
    setores_lista = setores_mod.carregar()
    contexto = {
        "setores": setores_lista,
        "setor": setor_atual(),
        "token": token_atual(),
        "linhas": [],
        "resumo": None,
        "barras": [],
        "erro": None,
        "mensagem": None,
        "info_dados": None,
        "avisos_planilha": [],
        "arquivos_outlook": [],
        "com_disponivel": COM_DISPONIVEL,
        "curva_geral": None,
        "pdf_disponivel": pdf_mod.DISPONIVEL,
        "params_pdf": None,
        "turnos": painel_mod.TURNOS,
        "turno": painel_mod.obter_turno(
            _pedidos_do_painel(setores_lista)["_geral"]["turno"]),
        "pedidos": _pedidos_do_painel(setores_lista),
    }

    df = recuperar_dataframe(contexto["token"])
    acao = request.form.get("acao", "") if request.method == "POST" else ""

    # -- planilhas: upload ou Outlook --
    if acao == "outlook":
        try:
            with trava_outlook:
                arquivos = baixador.baixar_recentes(dias=DIAS_PADRAO,
                                                    maximo=MAX_PLANILHAS)
        except OutlookIndisponivel as erro:
            contexto["erro"] = str(erro)
            arquivos = []
        except Exception as erro:  # noqa: BLE001
            contexto["erro"] = f"Falha ao falar com o Outlook: {erro}"
            arquivos = []

        if arquivos:
            df = concatenar_planilhas([a["caminho"] for a in arquivos])
            contexto["token"] = guardar_dataframe(df)
            guardar_lista_arquivos(contexto["token"], arquivos)
            contexto["arquivos_outlook"] = arquivos
            contexto["mensagem"] = f"{len(arquivos)} planilha(s) baixada(s)."
            contexto["avisos_planilha"] = list(df.attrs.get("avisos", []))
        elif not contexto["erro"]:
            contexto["erro"] = "Nenhuma planilha encontrada no Outlook."

    elif acao == "analisar":
        caminhos = salvar_uploads(request.files.getlist("planilhas"))
        if len(caminhos) >= 2:
            caminhos, avisos = ordenar_por_data(caminhos)
            try:
                df = concatenar_planilhas_em_memoria(caminhos[0], caminhos[1],
                                                     verbose=False)
            except Exception as erro:  # noqa: BLE001
                contexto["erro"] = f"Não foi possível juntar as planilhas: {erro}"
                df = None
            if df is not None:
                contexto["token"] = guardar_dataframe(df)
                contexto["avisos_planilha"] = avisos + list(
                    df.attrs.get("avisos", []))
        elif len(caminhos) == 1:
            contexto["erro"] = "Envie as duas planilhas juntas."

    if df is None:
        if not contexto["erro"]:
            contexto["erro"] = ("Carregue as planilhas para ver o painel — "
                                "pelo Outlook ou enviando os arquivos.")
        return render_template("painel.html", **contexto)

    contexto["info_dados"] = f"{len(df)} linhas carregadas"
    contexto["arquivos_outlook"] = (contexto["arquivos_outlook"]
                                    or recuperar_lista_arquivos(contexto["token"]))
    lembrar_token(contexto["token"])

    pedidos = _pedidos_do_painel(setores_lista)
    contexto["pedidos"] = pedidos
    if request.method == "POST":
        _guardar_pedidos(pedidos)

    geral = pedidos["_geral"]
    turno = painel_mod.obter_turno(geral["turno"])
    contexto["turno"] = turno
    producao_geral = geral["producao"]
    takt_geral = None
    if geral["takt"]:
        try:
            takt_geral = float(geral["takt"].replace(",", "."))
            if takt_geral <= 0:
                raise ValueError
        except ValueError:
            contexto["erro"] = "O takt deve ser um número de segundos maior que zero."
            return render_template("painel.html", **contexto)

    if not any(pedidos[s.id]["popid"] for s in setores_lista):
        return render_template("painel.html", **contexto)

    if not producao_geral:
        contexto["erro"] = "Informe a produção do dia."
        return render_template("painel.html", **contexto)

    # O takt cadastrado no setor e o do 1o turno; no 2o ele precisa ser dito.
    if turno["takt_obrigatorio"] and takt_geral is None:
        contexto["erro"] = (f"Informe o takt para o {turno['nome']} — o valor "
                            f"cadastrado em cada setor vale para o 1º turno.")
        return render_template("painel.html", **contexto)

    linhas = []
    for setor in setores_lista:
        pedido = pedidos[setor.id]
        if not pedido["popid"]:
            continue  # setor sem POPID fica de fora, sem virar erro
        faltas, janelas = _faltas_do_campo(pedido["faltas"])
        linhas.append(painel_mod.analisar_setor(
            df, setor, pedido["popid"], producao_geral, faltas, takt_geral,
            turno, janelas))

    contexto["linhas"] = linhas
    contexto["resumo"] = painel_mod.consolidar(linhas)
    contexto["barras"] = painel_mod.barras_de_carga(linhas)
    contexto["curva_geral"] = painel_mod.curva_geral(linhas)
    contexto["momentos_criticos"] = painel_mod.momentos_criticos(linhas)
    contexto["pdf_disponivel"] = pdf_mod.DISPONIVEL

    # O PDF refaz a analise com os mesmos parametros; o link os carrega.
    parametros = {"turno": turno["id"], "producao": producao_geral,
                  "takt": geral["takt"]}
    for setor in setores_lista:
        pedido = pedidos[setor.id]
        if pedido["popid"]:
            parametros[f"p_{setor.id}"] = pedido["popid"]
            parametros[f"f_{setor.id}"] = pedido["faltas"] or "0"
    contexto["params_pdf"] = parametros
    return render_template("painel.html", **contexto)


@app.route("/painel/relatorio.pdf")
def relatorio_fabrica_pdf():
    """PDF do painel: refaz a analise a partir dos parametros do link."""
    if not pdf_mod.DISPONIVEL:
        return Response(
            "Geração de PDF indisponível: instale com 'pip install reportlab'.",
            status=503, mimetype="text/plain; charset=utf-8")

    df = recuperar_dataframe(token_atual())
    if df is None:
        return redirect(url_for("painel"))

    turno = painel_mod.obter_turno(request.args.get("turno"))
    producao = (request.args.get("producao") or "").strip()
    if not producao:
        return redirect(url_for("painel"))

    try:
        takt = float((request.args.get("takt") or "").replace(",", "."))
        takt = takt if takt > 0 else None
    except ValueError:
        takt = None
    if turno["takt_obrigatorio"] and takt is None:
        return redirect(url_for("painel"))

    linhas = []
    for setor in setores_mod.carregar():
        popid = (request.args.get(f"p_{setor.id}") or "").strip()
        if not popid:
            continue
        faltas, janelas = _faltas_do_campo(request.args.get(f"f_{setor.id}") or "0")
        linhas.append(painel_mod.analisar_setor(
            df, setor, popid, producao, faltas, takt, turno, janelas))

    if not linhas:
        return redirect(url_for("painel"))

    contexto = {"turno": turno, "producao": producao}
    conteudo = pdf_fabrica.gerar(
        linhas, painel_mod.consolidar(linhas), painel_mod.curva_geral(linhas),
        contexto, painel_mod.momentos_criticos(linhas))

    return Response(conteudo, mimetype="application/pdf", headers={
        "Content-Disposition":
            f'attachment; filename="{pdf_fabrica.nome_arquivo(contexto)}"'})


@app.route("/ajuda")
def ajuda():
    """Tutorial das telas. Pagina estatica — nao depende de dados carregados."""
    return render_template("ajuda.html")


@app.route("/tags-mona", methods=["GET", "POST"])
def tags_mona_tela():
    """
    Cadastro das tags fixas da MONA (variantes que ainda não existem no mix
    geral da fábrica) — mesma lista usada pelo Filtro MIX e pela Análise,
    tanto para contar no CPM MIX quanto para filtrar as colunas relevantes
    de uma planilha MONA no formato técnico da Função 4.
    """
    setor = setor_atual()
    mensagem = erro = None

    if request.method == "POST":
        acao = request.form.get("acao", "")
        try:
            if acao == "adicionar":
                tags_mona_mod.adicionar(
                    request.form.get("texto", ""),
                    request.form.get("padroes", ""),
                    setor.id,
                )
                mensagem = "Tag adicionada."
            elif acao == "remover":
                tags_mona_mod.remover(request.form.get("texto", ""), setor.id)
                mensagem = "Tag removida."
            elif acao == "restaurar":
                tags_mona_mod.restaurar_padrao(setor.id)
                mensagem = "Lista restaurada para o padrão."
        except tags_mona_mod.ErroTagsMona as falha:
            erro = str(falha)

    return render_template(
        "tags_mona.html",
        tags=tags_mona_mod.carregar(setor.id),
        mensagem=mensagem, erro=erro,
        setor=setor, setores=setores_mod.carregar(),
    )


@app.route("/analise", methods=["GET", "POST"])
def analise():
    setor = setor_atual()
    contexto = {
        "setor": setor,
        "setores": setores_mod.carregar(),
        "token": "",
        "popid": "",
        "producao": "",
        "faltas": "",
        "presenteismo_ativo": False,
        "presenteismo_qtd": "",
        "presenteismo_janelas": [],
        "presenteismo_momentos": [],
        "takt": f"{setor.takt_segundos:g}",
        "inicio_turno": setor.inicio_turno,
        "fim_turno": setor.fim_turno,
        "resultado": None,
        "erro": None,
        "info_dados": None,

        "arquivos_outlook": [],
        "mensagem": None,
        "com_disponivel": COM_DISPONIVEL,
        "dias": str(DIAS_PADRAO),
        "avisos_planilha": [],
        "export_tsv": "",
        "export_texto": "",
        "gargalos_dados": None,
        "membros_de": None,
        "pdf_disponivel": pdf_mod.DISPONIVEL,
        "params_pdf": None,
        "data_inicio": "",
        "data_fim": "",
        "tags_mona_fixas": [tag.texto for tag in cpm_core.tags_mona_ativas(setor.id)],
        "mona_arquivo": None,
    }

    if request.method != "POST":
        # Voltando de outra aba: reaproveita a planilha que ficou no cache.
        token = token_atual()
        df_cache = recuperar_dataframe(token)
        if df_cache is not None:
            contexto["token"] = token
            contexto["info_dados"] = f"{len(df_cache)} linhas carregadas"
            contexto["arquivos_outlook"] = recuperar_lista_arquivos(token)
            contexto["popid"] = request.cookies.get(COOKIE_POPID, "")
            contexto["producao"] = request.cookies.get(COOKIE_PRODUCAO, "")
            contexto["mona_arquivo"] = recuperar_nome_mona(token)
            lembrar_token(token)
        return render_template("index.html", **contexto)

    contexto["popid"] = request.form.get("popid", "").strip()
    contexto["producao"] = request.form.get("producao", "").strip()
    contexto["faltas"] = request.form.get("absenteismo", "").strip()
    (contexto["presenteismo_ativo"], contexto["presenteismo_qtd"],
     contexto["presenteismo_janelas"]) = _presenteismo_do_form(request.form)
    contexto["takt"] = (request.form.get("takt")
                        or f"{setor.takt_segundos:g}").strip()
    contexto["inicio_turno"] = (request.form.get("inicio_turno")
                                or setor.inicio_turno).strip()
    contexto["fim_turno"] = (request.form.get("fim_turno")
                             or setor.fim_turno).strip()
    token = token_atual()
    acao = request.form.get("acao", "analisar")
    contexto["dias"] = request.form.get("dias", str(DIAS_PADRAO)).strip() or str(DIAS_PADRAO)
    contexto["data_inicio"] = request.form.get("data_inicio", "").strip()
    contexto["data_fim"] = request.form.get("data_fim", "").strip()

    # 1a. Captura automatica no Outlook
    if acao in ("outlook_recentes", "outlook_periodo"):
        try:
            arquivos = buscar_no_outlook(acao, contexto)
        except (OutlookIndisponivel, ValueError) as erro:
            contexto["erro"] = str(erro)
            return render_template("index.html", **contexto)
        except Exception as erro:
            contexto["erro"] = f"Falha ao falar com o Outlook: {erro}"
            return render_template("index.html", **contexto)

        if not arquivos:
            contexto["erro"] = "Nenhuma planilha encontrada no Outlook com esse filtro."
            return render_template("index.html", **contexto)

        df = concatenar_planilhas([a["caminho"] for a in arquivos])
        df, erro_mona, aviso_mona, texto_mona, nome_mona = _aplicar_mona_em_opcionais(df, request, setor=setor.id)
        if erro_mona:
            contexto["erro"] = erro_mona
            return render_template("index.html", **contexto)
        contexto["token"] = token = guardar_dataframe(df)
        contexto["arquivos_outlook"] = arquivos
        contexto["info_dados"] = f"{len(df)} linhas de {len(arquivos)} planilha(s)"
        contexto["avisos_planilha"] = list(df.attrs.get("avisos", []))
        contexto["mona_arquivo"] = nome_mona
        if aviso_mona:
            contexto["avisos_planilha"].append(aviso_mona)
        contexto["mensagem"] = (
            f"{len(arquivos)} planilha(s) baixada(s) do Outlook."
        )
        guardar_lista_arquivos(token, arquivos)
        guardar_texto_mona(token, texto_mona)
        guardar_nome_mona(token, nome_mona)

        # So analisa em seguida se POPID e producao ja estiverem preenchidos
        if not contexto["popid"] or not contexto["producao"]:
            return render_template("index.html", **contexto)
        return render_template("index.html",
                               **executar_analise(df, contexto, setor,
                                                  contador_mona=texto_mona))

    # 1b. Planilhas: usa as novas se vierem, senao reaproveita as ja carregadas
    enviados = request.files.getlist("planilhas")
    caminhos = salvar_uploads(enviados)
    texto_mona = None

    if len(caminhos) >= 2:
        # O navegador entrega os arquivos na ordem em que o sistema listou.
        # A concatenacao espera a planilha mais antiga primeiro.
        caminhos, avisos_ordem = ordenar_por_data(caminhos)
        contexto["avisos_planilha"] = list(avisos_ordem)
        try:
            df = concatenar_planilhas_em_memoria(caminhos[0], caminhos[1], verbose=False)
        except Exception as erro:  # noqa: BLE001
            contexto["erro"] = f"Não foi possível juntar as planilhas: {erro}"
            return render_template("index.html", **contexto)
        if df is None:
            contexto["erro"] = "Não foi possível ler as planilhas enviadas."
            return render_template("index.html", **contexto)
        df, erro_mona, aviso_mona, texto_mona, nome_mona = _aplicar_mona_em_opcionais(df, request, setor=setor.id)
        if erro_mona:
            contexto["erro"] = erro_mona
            return render_template("index.html", **contexto)
        if aviso_mona:
            contexto["avisos_planilha"].append(aviso_mona)
        token = guardar_dataframe(df)
        guardar_texto_mona(token, texto_mona)
        guardar_nome_mona(token, nome_mona)
    elif len(caminhos) == 1:
        contexto["erro"] = "Envie as duas planilhas juntas."
        return render_template("index.html", **contexto)
    else:
        df = recuperar_dataframe(token)
        if df is None:
            contexto["erro"] = "Selecione as duas planilhas para começar."
            return render_template("index.html", **contexto)
        texto_mona = recuperar_texto_mona(token)
        nome_mona = recuperar_nome_mona(token)

    contexto["token"] = token
    lembrar_token(token)
    contexto["info_dados"] = f"{len(df)} linhas carregadas"
    contexto["avisos_planilha"] = (list(contexto.get("avisos_planilha") or [])
                                   + list(df.attrs.get("avisos", [])))
    contexto["arquivos_outlook"] = recuperar_lista_arquivos(token)
    contexto["mona_arquivo"] = nome_mona

    # 2. Analise
    return render_template("index.html",
                           **executar_analise(df, contexto, setor,
                                              contador_mona=texto_mona))


# ---------------------------------------------------------------------------
# Tela de tarefas (adicionar / editar / excluir)
# ---------------------------------------------------------------------------
MAX_CONDICOES = 3


def colunas_e_valores(token: str, setor: str | None = None) -> dict[str, list[str]]:
    """
    Colunas para o seletor de condicoes (dropdown) e, quando ha planilha
    carregada, alguns valores de cada uma para o formulario sugerir em vez
    de exigir digitacao.

    Sempre inclui as colunas padrao da funcao (mesma lista do construtor de
    regras de gargalos, `gargalos_mod.colunas_regra_do_setor`) — assim o
    dropdown aparece mesmo sem nenhuma planilha carregada ainda. Quando ha
    planilha, as colunas dela entram tambem (e ganham os valores mais
    comuns), incluindo as que nao estao na lista padrao.
    """
    mapa: dict[str, list[str]] = {
        coluna: [] for coluna in gargalos_mod.colunas_regra_do_setor(setor)
    }

    df = recuperar_dataframe(token)
    if df is None:
        return mapa

    for coluna in df.columns:
        try:
            valores = (
                df[coluna].dropna().astype(str).str.strip().str.upper()
                .replace("", pd.NA).dropna()
                .value_counts().head(30).index.tolist()
            )
        except Exception:
            valores = []
        mapa[str(coluna)] = valores
    return mapa


def sequencia_do_form(formulario) -> dict | None:
    minimo = (formulario.get("sequencia_min") or "").strip()
    if not minimo:
        return None
    return {"minimo": minimo,
            "modo": formulario.get("sequencia_modo") or "primeiro",
            "janela": (formulario.get("sequencia_janela") or "").strip()}


def condicoes_do_form(formulario) -> list[dict]:
    condicoes = []
    for i in range(1, MAX_CONDICOES + 1):
        coluna = formulario.get(f"coluna_{i}", "").strip()
        valor = formulario.get(f"valor_{i}", "").strip()
        if not coluna or not valor:
            continue
        condicoes.append({
            "coluna": coluna,
            "operador": formulario.get(f"operador_{i}", "contem"),
            "valor": valor,
        })
    return condicoes


# ---------------------------------------------------------------------------
# Filtro MIX (antigo filtroMix.py, agora sem Tkinter)
# ---------------------------------------------------------------------------
def _opcoes_de_colunas(processador, funcao_id, config) -> list[str]:
    """
    Caixas a mostrar: as colunas que a funcao conhece, mais qualquer coluna
    ja salva que nao esteja nelas — assim uma configuracao antiga nao some da
    tela sem aviso.
    """
    opcoes = list(processador.colunas_padrao)
    salvas = filtro_mix_mod.colunas_da_funcao(
        funcao_id, config, processador.colunas_padrao)
    opcoes += [c for c in salvas if c not in opcoes]
    return opcoes


@app.route("/filtro-mix", methods=["GET", "POST"])
def filtro_mix():
    """
    Uma tela para todas as funcoes de MIX.

    O seletor do topo troca a funcao; o resto da tela se adapta ao que ela
    declara em `processadores.py`: quais arquivos pede, quais parametros, e se
    tem lista de colunas para configurar.
    """
    funcao_id = (request.values.get("funcao") or "fa2").strip()
    processador = processadores_mod.obter(funcao_id)
    if processador is None:
        processador = processadores_mod.obter("fa2")
        funcao_id = "fa2"

    config = filtro_mix_mod.carregar_config()
    contexto = {
        "config": config,
        "processador": processador,
        "processadores": processadores_mod.listar(),
        "colunas_funcao": filtro_mix_mod.colunas_da_funcao(
            funcao_id, config, processador.colunas_padrao),
        "colunas_disponiveis": _opcoes_de_colunas(processador, funcao_id, config),
        "enviados": {c.nome: "" for c in processador.valores},
        "com_disponivel": filtro_mix_mod.COM_DISPONIVEL,
        "mensagem": None,
        "erro": None,
        "info_email": None,
        "arquivo_gerado": None,
        "diagnostico": None,
        "relatorio_colunas": None,
    }

    def responder():
        return render_template("filtro_mix.html", **contexto)

    if request.method != "POST":
        return responder()

    acao = request.form.get("acao", "")
    contexto["enviados"] = {c.nome: (request.form.get(c.nome) or "")
                            for c in processador.valores}

    # -- salvar configuracoes --
    if acao == "salvar_config":
        campos = ("assunto", "remetente", "dias_recentes",
                  "palavras_chave", "pasta_saida")
        novos = {c: request.form[c] for c in campos if c in request.form}
        novos["funcao"] = funcao_id

        # As colunas vem das caixas marcadas. A ordem e a do layout da funcao,
        # nao a da marcacao — quem define largura e formato e a funcao.
        if "colunas_enviadas" in request.form:
            marcadas = set(request.form.getlist("colunas"))
            ordem = list(processador.colunas_padrao) + [
                c for c in filtro_mix_mod.colunas_da_funcao(
                    funcao_id, config, processador.colunas_padrao)
                if c not in processador.colunas_padrao
            ]
            novos["colunas_desejadas"] = [c for c in ordem if c in marcadas]
        try:
            contexto["config"] = filtro_mix_mod.salvar_config(novos)
            contexto["colunas_funcao"] = filtro_mix_mod.colunas_da_funcao(
                funcao_id, contexto["config"], processador.colunas_padrao)
            contexto["colunas_disponiveis"] = _opcoes_de_colunas(
                processador, funcao_id, contexto["config"])
            contexto["mensagem"] = "Configurações salvas."
        except (ValueError, OSError) as erro:
            contexto["erro"] = str(erro)
        return responder()

    # -- diagnostico --
    if acao == "diagnostico":
        try:
            contexto["diagnostico"] = filtro_mix_mod.diagnosticar(config)
        except Exception as erro:  # noqa: BLE001
            contexto["erro"] = f"Falha no diagnóstico: {erro}"
        return responder()

    # -- automatico: baixa do Outlook e processa --
    if acao == "automatico":
        if not processador.aceita_outlook:
            contexto["erro"] = (f"{processador.titulo} precisa de mais de um "
                                f"arquivo — use o envio manual.")
            return responder()

        caminho_mona = None
        if processador.aceita_mona and request.form.get("usar_mona") == "1":
            enviado_mona = request.files.get("arquivo_mona_extra")
            if enviado_mona is None or not enviado_mona.filename:
                contexto["erro"] = ("Marque Mona: selecione a planilha MONA "
                                    "para cruzar.")
                return responder()
            if not enviado_mona.filename.lower().endswith(processadores_mod.EXTENSOES):
                contexto["erro"] = ("Planilha MONA: envie um arquivo .xlsx, "
                                    ".xlsm ou .xls.")
                return responder()
            pasta_mona = os.path.join(PASTA_SESSAO, uuid.uuid4().hex)
            os.makedirs(pasta_mona, exist_ok=True)
            caminho_mona = os.path.join(
                pasta_mona, os.path.basename(enviado_mona.filename))
            enviado_mona.save(caminho_mona)
        texto_mona = (request.form.get("texto_mona") or "").strip() or None

        try:
            with trava_outlook:
                info, mensagem = filtro_mix_mod.baixar_ultimo_anexo_excel(config)
        except OutlookIndisponivel as erro:
            contexto["erro"] = str(erro)
            return responder()
        except Exception as erro:  # noqa: BLE001
            contexto["erro"] = f"Falha ao falar com o Outlook: {erro}"
            return responder()

        if not info:
            contexto["erro"] = mensagem
            return responder()

        contexto["info_email"] = info
        _executar(processador, [info["caminho_arquivo"]], [], contexto,
                  caminho_mona=caminho_mona, texto_mona=texto_mona)
        try:
            os.remove(info["caminho_arquivo"])
        except OSError:
            pass
        return responder()

    # -- manual: arquivos enviados pela tela --
    if acao == "manual":
        pasta = os.path.join(PASTA_SESSAO, uuid.uuid4().hex)
        caminhos, enviados_por_campo = [], {}

        for campo in processador.arquivos:
            enviado = request.files.get(campo.nome)
            if enviado is None or not enviado.filename:
                if campo.opcional:
                    caminhos.append(None)
                    enviados_por_campo[campo.nome] = False
                    continue
                contexto["erro"] = f"Selecione o arquivo: {campo.rotulo}."
                return responder()

            if not enviado.filename.lower().endswith(processadores_mod.EXTENSOES):
                contexto["erro"] = (f"{campo.rotulo}: envie um arquivo .xlsx, "
                                    f".xlsm ou .xls.")
                return responder()

            os.makedirs(pasta, exist_ok=True)
            destino = os.path.join(pasta, os.path.basename(enviado.filename))
            enviado.save(destino)
            caminhos.append(destino)
            enviados_por_campo[campo.nome] = True

        valores = []
        for campo in processador.valores:
            # Parametro atrelado a um arquivo opcional so vale se ele veio.
            if campo.exigido_com and not enviados_por_campo.get(campo.exigido_com):
                valores.append(None)
                continue
            try:
                valor = _valor_do_form(campo, request.form)
            except ValueError as erro:
                contexto["erro"] = str(erro)
                return responder()
            if valor is None:
                contexto["erro"] = f"Informe: {campo.rotulo}."
                return responder()
            valores.append(valor)

        caminho_mona = None
        if processador.aceita_mona and request.form.get("usar_mona") == "1":
            enviado_mona = request.files.get("arquivo_mona_extra")
            if enviado_mona is None or not enviado_mona.filename:
                contexto["erro"] = ("Marque Mona: selecione a planilha MONA "
                                    "para cruzar.")
                return responder()
            if not enviado_mona.filename.lower().endswith(processadores_mod.EXTENSOES):
                contexto["erro"] = ("Planilha MONA: envie um arquivo .xlsx, "
                                    ".xlsm ou .xls.")
                return responder()
            os.makedirs(pasta, exist_ok=True)
            caminho_mona = os.path.join(
                pasta, os.path.basename(enviado_mona.filename))
            enviado_mona.save(caminho_mona)

        texto_mona = (request.form.get("texto_mona") or "").strip() or None

        _executar(processador, caminhos, valores, contexto, pasta,
                  caminho_mona=caminho_mona, texto_mona=texto_mona)
        return responder()

    return responder()


def _executar(processador, caminhos, valores, contexto, pasta=None,
              caminho_mona=None, texto_mona=None) -> None:
    """Roda a funcao escolhida e guarda o arquivo gerado para download."""
    limpar_cache_antigo()
    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_saida = f"{processador.prefixo_saida}_{carimbo}.xlsx"

    destino_local = filtro_mix_mod.carregar_config().get("pasta_saida")
    try:
        filtro_mix_mod.criar_pasta_se_nao_existir(destino_local)
        saida = os.path.join(destino_local, nome_saida)
    except OSError:
        saida = os.path.join(pasta or PASTA_SESSAO, nome_saida)
        os.makedirs(os.path.dirname(saida), exist_ok=True)

    colunas = filtro_mix_mod.colunas_da_funcao(
        processador.id, None, processador.colunas_padrao)

    kwargs = {"colunas_desejadas": colunas}
    if processador.aceita_mona:
        kwargs["caminho_mona"] = caminho_mona
        kwargs["texto_mona"] = texto_mona

    try:
        ok, mensagem = processador.funcao(*caminhos, *valores, saida, **kwargs)
    except TypeError as erro:
        contexto["erro"] = (f"A função {processador.titulo} não aceitou os "
                            f"parâmetros informados: {erro}")
        return
    except Exception as erro:  # noqa: BLE001
        contexto["erro"] = f"Falha ao processar: {erro}"
        logger.exception("filtro-mix/%s: falha ao processar", processador.id)
        return

    if ok:
        contexto["mensagem"] = mensagem
        contexto["arquivo_gerado"] = nome_saida
        _processados[nome_saida] = saida
        if processador.usa_lista_colunas:
            contexto["relatorio_colunas"] = filtro_mix_mod.relatorio_colunas()
    else:
        contexto["erro"] = mensagem


# ---------------------------------------------------------------------------
# Gargalos: tela dedicada e construtor de regras
# ---------------------------------------------------------------------------
def condicoes_do_form_gargalo(formulario, prefixo: str = "") -> list[dict]:
    condicoes = []
    for i in range(1, gargalos_mod.MAX_CONDICOES + 1):
        condicoes.append({
            "coluna": formulario.get(f"{prefixo}coluna_{i}", ""),
            "operador": formulario.get(f"{prefixo}operador_{i}", "contem"),
            "valor": formulario.get(f"{prefixo}valor_{i}", ""),
        })
    return condicoes


def parametros_do_form(formulario) -> dict:
    return {
        "janela": formulario.get("janela", ""),
        "min_ocorrencias": formulario.get("min_ocorrencias", ""),
        "min_seq": formulario.get("min_seq", ""),
        "logica": formulario.get("logica", "E"),
        "condicoes": condicoes_do_form_gargalo(formulario),
        "grupos": [
            {"logica": formulario.get(f"g{n}_logica", "E"),
             "condicoes": condicoes_do_form_gargalo(formulario, f"g{n}_")}
            for n in range(1, gargalos_mod.MAX_GRUPOS + 1)
        ],
    }


@app.route("/relatorio.pdf")
def relatorio_pdf_rota():
    """Relatorio completo: variantes, tarefas, CPM MIX e gargalos."""
    if not pdf_mod.DISPONIVEL:
        return Response(
            "Geração de PDF indisponível: instale a biblioteca com "
            "'pip install reportlab' e reinicie o app.",
            status=503, mimetype="text/plain; charset=utf-8")

    setor = setor_atual()
    df = recuperar_dataframe(token_atual())
    if df is None:
        return redirect(url_for("analise"))

    try:
        faltas = int(float((request.args.get("absenteismo") or "0").replace(",", ".")))
    except ValueError:
        faltas = 0

    try:
        takt = float((request.args.get("takt") or "").replace(",", "."))
    except ValueError:
        takt = 0.0
    if takt <= 0:
        takt = setor.takt_segundos

    try:
        resultado = analisar_mix(df, request.args.get("popid", ""),
                                 request.args.get("producao", ""), takt=takt,
                                 setor=setor.id)
    except ValueError as erro:
        return Response(str(erro), status=400,
                        mimetype="text/plain; charset=utf-8")

    recorte = df.iloc[resultado.primeira_linha - 1:resultado.ultima_linha]
    try:
        dados = gargalos_mod.analisar(recorte, setor.id)
        gargalos_mod.anotar_horarios(dados, setor)
        membros = lambda g: gargalos_mod.descrever_membros(dados, g)  # noqa: E731
        n_gargalos = dados["resumo"].get("total", 0)
    except Exception as erro:  # noqa: BLE001
        logger.warning("pdf: gargalos fora do relatório (%s)", erro)
        dados, membros, n_gargalos = None, None, 0

    aplicar_absenteismo(resultado, faltas, n_gargalos,
                        setor.tabela_postos, setor.gargalos_varios)

    conteudo = pdf_mod.gerar(
        resultado, dados, membros,
        inicio_turno=request.args.get("inicio_turno") or setor.inicio_turno,
        fim_turno=request.args.get("fim_turno") or setor.fim_turno,
        pausas=setor.pausas_tuplas or None)
    return Response(conteudo, mimetype="application/pdf", headers={
        "Content-Disposition":
            f'attachment; filename="{pdf_mod.nome_arquivo(resultado)}"'})


@app.route("/gargalos", methods=["GET", "POST"])
def gargalos_tela():
    """Só os gargalos, a partir do mesmo POPID e produção do dia."""
    setor = setor_atual()
    contexto = {
        "setor": setor,
        "setores": setores_mod.carregar(),
        "token": token_atual(),
        "popid": (request.values.get("popid")
                  or request.cookies.get(COOKIE_POPID, "")).strip(),
        "producao": (request.values.get("producao")
                     or request.cookies.get(COOKIE_PRODUCAO, "")).strip(),
        "erro": None,
        "info_dados": None,
        "gargalos_dados": None,
        "membros_de": None,
        "resultado": None,
        "pdf_disponivel": pdf_mod.DISPONIVEL,
        "params_pdf": None,
    }

    df = recuperar_dataframe(contexto["token"])
    if df is None:
        contexto["erro"] = ("Nenhuma planilha carregada. Faça a análise de MIX "
                            "primeiro — os gargalos usam os mesmos dados.")
        return render_template("gargalos.html", **contexto)

    contexto["info_dados"] = f"{len(df)} linhas carregadas"
    if not contexto["popid"] or not contexto["producao"]:
        return render_template("gargalos.html", **contexto)

    try:
        resultado = analisar_mix(df, contexto["popid"], contexto["producao"],
                                 takt=setor.takt_segundos, setor=setor.id)
    except ValueError as erro:
        contexto["erro"] = str(erro)
        return render_template("gargalos.html", **contexto)

    contexto["resultado"] = resultado
    recorte = df.iloc[resultado.primeira_linha - 1:resultado.ultima_linha]
    dados = gargalos_mod.analisar(recorte, setor.id)
    gargalos_mod.anotar_horarios(dados, setor)
    contexto["gargalos_dados"] = dados
    contexto["membros_de"] = lambda g: gargalos_mod.descrever_membros(dados, g)
    contexto["pdf_disponivel"] = pdf_mod.DISPONIVEL
    contexto["params_pdf"] = {"token": contexto["token"],
                              "popid": contexto["popid"],
                              "producao": contexto["producao"],
                              "setor": setor.id}
    return render_template("gargalos.html", **contexto)


@app.route("/gargalos/regras", methods=["GET", "POST"])
def gargalos_regras():
    setor = setor_atual()
    mensagem = erro = None
    editando = None

    if request.method == "POST":
        acao = request.form.get("acao", "")
        try:
            if acao == "salvar":
                gargalos_mod.salvar_regra(
                    nome=request.form.get("nome", ""),
                    tipo=request.form.get("tipo", ""),
                    severidade=request.form.get("severidade", "ALTA"),
                    parametros=parametros_do_form(request.form),
                    regra_id=request.form.get("regra_id") or None,
                    setor=setor.id,
                )
                mensagem = "Regra salva."
            elif acao == "remover":
                gargalos_mod.remover_regra(request.form.get("regra_id", ""), setor.id)
                mensagem = "Regra removida."
            elif acao == "importar":
                enviados = request.files.getlist("arquivo")
                if not enviados or not enviados[0].filename:
                    raise gargalos_mod.ErroRegra(
                        "Selecione o arquivo .json exportado.")
                conteudo = enviados[0].read().decode("utf-8-sig", errors="replace")
                importadas, backup = gargalos_mod.importar(conteudo, setor.id)
                mensagem = f"Importado: {len(importadas)} regra(s)."
                if backup:
                    mensagem += f" As regras anteriores ficaram salvas como {backup}."
            elif acao == "restaurar":
                gargalos_mod.restaurar_padrao(setor.id)
                mensagem = "Regras restauradas para o padrão."
        except gargalos_mod.ErroRegra as falha:
            erro = str(falha)
            editando = {
                "id": request.form.get("regra_id", ""),
                "nome": request.form.get("nome", ""),
                "tipo": request.form.get("tipo", "contagem_janela"),
                "severidade": request.form.get("severidade", "ALTA"),
                "params": parametros_do_form(request.form),
            }

    regras = gargalos_mod.listar_regras(setor.id)

    if editando is None:
        regra = gargalos_mod.encontrar_regra(request.args.get("editar", ""), setor.id)
        if regra:
            editando = _regra_para_form(regra)
        else:
            editando = {"id": "", "nome": "", "tipo": "contagem_janela",
                        "severidade": "ALTA", "params": _params_vazios()}

    return render_template(
        "gargalos_regras.html",
        regras=regras, editando=editando, mensagem=mensagem, erro=erro,
        setor=setor, setores=setores_mod.carregar(),
        tipos=gargalos_mod.TIPOS, operadores=gargalos_mod.OPERADORES,
        severidades=gargalos_mod.SEVERIDADES,
        colunas=gargalos_mod.colunas_regra_do_setor(setor.id),
        max_condicoes=gargalos_mod.MAX_CONDICOES,
        max_grupos=gargalos_mod.MAX_GRUPOS,
        descrever=gargalos_mod.descrever_regra,
    )


def _params_vazios() -> dict:
    vazia = [{"coluna": "", "operador": "contem", "valor": ""}
             for _ in range(gargalos_mod.MAX_CONDICOES)]
    return {
        "janela": "", "min_ocorrencias": "", "min_seq": "", "logica": "E",
        "condicoes": list(vazia),
        "grupos": [{"logica": "E", "condicoes": list(vazia)}
                   for _ in range(gargalos_mod.MAX_GRUPOS)],
    }


def _completar(condicoes: list[dict]) -> list[dict]:
    cheias = list(condicoes)
    while len(cheias) < gargalos_mod.MAX_CONDICOES:
        cheias.append({"coluna": "", "operador": "contem", "valor": ""})
    return cheias[:gargalos_mod.MAX_CONDICOES]


def _regra_para_form(regra: dict) -> dict:
    params = _params_vazios()
    params["janela"] = regra.get("janela", "")
    params["min_ocorrencias"] = regra.get("min_ocorrencias", "")
    params["min_seq"] = regra.get("min_seq", "")

    predicado = regra.get("predicado")
    if predicado:
        params["logica"] = predicado.get("logica", "E")
        params["condicoes"] = _completar(predicado.get("condicoes", []))

    grupos = regra.get("grupos", [])
    for indice in range(gargalos_mod.MAX_GRUPOS):
        if indice < len(grupos):
            params["grupos"][indice] = {
                "logica": grupos[indice].get("logica", "E"),
                "condicoes": _completar(grupos[indice].get("condicoes", [])),
            }

    return {"id": regra["id"], "nome": regra["nome"], "tipo": regra["tipo"],
            "severidade": regra.get("severidade", "ALTA"), "params": params}


@app.route("/gargalos/regras/exportar.<formato>")
def exportar_regras_setor(formato):
    """Regras do setor em edicao. O .json e o unico que o import aceita."""
    setor = setor_atual()
    geradores = {
        "json": (gargalos_mod.exportar_json, "application/json"),
        "tsv": (gargalos_mod.exportar_tsv, "text/tab-separated-values"),
        "txt": (gargalos_mod.exportar_texto, "text/plain"),
    }
    if formato not in geradores:
        return redirect(url_for("gargalos_regras"))

    gerar, tipo = geradores[formato]
    carimbo = datetime.now().strftime("%Y%m%d_%H%M")
    return Response(
        gerar(setor.id),
        mimetype=f"{tipo}; charset=utf-8",
        headers={"Content-Disposition":
                 f'attachment; filename="regras_{setor.id}_{carimbo}.{formato}"'},
    )


@app.route("/gargalos/regras/todos.<formato>")
def exportar_regras_todos(formato):
    """Regras de gargalo de TODOS os setores, num arquivo so."""
    lista = setores_mod.carregar()
    geradores = {
        "json": (gargalos_mod.exportar_json_todos, "application/json"),
        "tsv": (gargalos_mod.exportar_tsv_todos, "text/tab-separated-values"),
        "txt": (gargalos_mod.exportar_texto_todos, "text/plain"),
    }
    if formato not in geradores:
        return redirect(url_for("gargalos_regras"))

    gerar, tipo = geradores[formato]
    carimbo = datetime.now().strftime("%Y%m%d_%H%M")
    return Response(
        gerar(lista),
        mimetype=f"{tipo}; charset=utf-8",
        headers={"Content-Disposition":
                 f'attachment; filename="regras_gargalo_fabrica_{carimbo}.{formato}"'},
    )


@app.route("/tarefas", methods=["GET", "POST"])
def tarefas():
    setor = setor_atual()
    token = token_atual()
    erro = mensagem = None
    editando = None

    if request.method == "POST":
        acao = request.form.get("acao", "")
        try:
            if acao == "salvar_tarefa":
                catalogo_mod.salvar_tarefa(
                    variante_id=request.form.get("variante", ""),
                    nome=request.form.get("nome", ""),
                    tempo=request.form.get("tempo", ""),
                    condicoes=condicoes_do_form(request.form),
                    tarefa_id=request.form.get("tarefa_id") or None,
                    sequencia=sequencia_do_form(request.form),
                    setor=setor.id,
                )
                mensagem = "Tarefa salva."
            elif acao == "remover_tarefa":
                catalogo_mod.remover_tarefa(request.form.get("tarefa_id", ""), setor.id)
                mensagem = "Tarefa excluída."
            elif acao == "nova_variante":
                catalogo_mod.adicionar_variante(
                    request.form.get("variante_id", ""), request.form.get("variante_nome", "")
                )
                mensagem = "Variante criada."
            elif acao == "remover_variante":
                catalogo_mod.remover_variante(request.form.get("variante_id", ""), setor.id)
                mensagem = "Variante excluída."
            elif acao == "importar":
                enviados = request.files.getlist("arquivo")
                if not enviados or not enviados[0].filename:
                    raise ErroCatalogo("Selecione o arquivo .json exportado.")
                conteudo = enviados[0].read().decode("utf-8-sig", errors="replace")
                variantes, backup = catalogo_mod.importar(conteudo, setor.id)
                mensagem = (
                    f"Importado: {len(variantes)} variante(s), "
                    f"{sum(len(v.tarefas) for v in variantes)} tarefa(s)."
                )
                if backup:
                    mensagem += f" O catálogo anterior ficou salvo como {backup}."
            elif acao == "restaurar":
                catalogo_mod.restaurar_padrao(setor.id)
                mensagem = "Catálogo restaurado para o padrão de fábrica."
        except ErroCatalogo as falha:
            erro = str(falha)
        except Exception as falha:  # noqa: BLE001
            erro = f"Falha ao gravar o catálogo: {falha}"

        if erro:
            # devolve o formulario preenchido para o usuario corrigir
            editando = {
                "id": request.form.get("tarefa_id", ""),
                "variante": request.form.get("variante", ""),
                "nome": request.form.get("nome", ""),
                "tempo": request.form.get("tempo", ""),
                "condicoes": condicoes_do_form(request.form),
                "sequencia": sequencia_do_form(request.form),
            }

    variantes = catalogo_mod.carregar(setor.id)

    if editando is None:
        editar_id = request.args.get("editar", "")
        if editar_id:
            variante, tarefa = catalogo_mod.encontrar_tarefa(editar_id, setor.id)
            if tarefa:
                editando = {
                    "id": tarefa.id,
                    "variante": variante.id,
                    "nome": tarefa.nome,
                    "tempo": ("%g" % tarefa.tempo),
                    "condicoes": catalogo_mod.condicoes_de_criterio(tarefa.criterio),
                    "sequencia": catalogo_mod.sequencia_de_criterio(tarefa.criterio),
                }
        elif request.args.get("nova"):
            editando = {"id": "", "variante": request.args["nova"],
                        "nome": "", "tempo": "", "condicoes": [],
                        "sequencia": None}

    if editando is None:
        editando = {"id": "", "variante": variantes[0].id if variantes else "",
                    "nome": "", "tempo": "", "condicoes": [],
                    "sequencia": None}

    while len(editando["condicoes"]) < MAX_CONDICOES:
        editando["condicoes"].append({"coluna": "", "operador": "contem", "valor": ""})

    return render_template(
        "tarefas.html",
        variantes=variantes,
        editando=editando,
        erro=erro,
        mensagem=mensagem,
        token=token,
        planilha_carregada=recuperar_dataframe(token) is not None,
        colunas=colunas_e_valores(token, setor.id),
        operadores=catalogo_mod.OPERADORES,
        max_condicoes=MAX_CONDICOES,
        modos_sequencia=catalogo_mod.MODOS_SEQUENCIA,
        sequencia_de_criterio=catalogo_mod.sequencia_de_criterio,
        export_regras_tsv=catalogo_mod.exportar_tsv(variantes),
        export_regras_texto=catalogo_mod.exportar_texto(variantes),
        setor=setor, setores=setores_mod.carregar(),
        condicoes_de_criterio=catalogo_mod.condicoes_de_criterio,
    )


@app.route("/tarefas/exportar.<formato>")
def exportar_regras(formato):
    """Baixa as regras: json (backup/importacao), tsv (Excel) ou txt."""
    setor = setor_atual()
    geradores = {
        "json": (catalogo_mod.exportar_json, "application/json"),
        "tsv": (catalogo_mod.exportar_tsv, "text/tab-separated-values"),
        "txt": (catalogo_mod.exportar_texto, "text/plain"),
    }
    if formato not in geradores:
        return redirect(url_for("tarefas"))

    gerar, tipo = geradores[formato]
    carimbo = datetime.now().strftime("%Y%m%d_%H%M")
    conteudo = gerar(setor=setor.id)
    return Response(
        conteudo,
        mimetype=f"{tipo}; charset=utf-8",
        headers={"Content-Disposition":
                 f'attachment; filename="regras_variantes_{carimbo}.{formato}"'},
    )


@app.template_filter("itens_texto")
def itens_texto(itens) -> str:
    """POPID/SEQ em TAB, pronto para colar no Excel."""
    if not itens:
        return ""
    tem_seq = any(item.get("seq") for item in itens)
    linhas = ["POPID\tSEQ" if tem_seq else "POPID\tPosição"]
    for item in itens:
        linhas.append(f"{item['popid']}\t{item['seq'] if tem_seq else item['posicao']}")
    return "\n".join(linhas)


@app.template_filter("faixa")
def faixa(percentual: float) -> str:
    """Classe CSS conforme a carga."""
    if percentual >= 100:
        return "critico"
    if percentual >= 90:
        return "atencao"
    return "ok"


def ip_da_rede() -> str:
    """IP da maquina na rede local, para passar o endereco aos colegas."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as teste:
            teste.connect(("8.8.8.8", 80))
            return teste.getsockname()[0]
    except Exception:
        return socket.gethostbyname(socket.gethostname())


# Em servidor (Railway), quem sobe o app e o gunicorn: `gunicorn app:app`.
# O bloco abaixo so roda no uso local, pela linha de comando.
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="NEXUS-6 web")
    parser.add_argument("--rede", action="store_true",
                        help="aceita acesso de outras maquinas da rede local")
    parser.add_argument("--porta", type=int, default=5082)
    parser.add_argument("--debug", action="store_true",
                        help="NAO use com --rede: o debugger do Flask permite "
                             "executar codigo remotamente")
    args = parser.parse_args()

    host = "0.0.0.0" if args.rede else "127.0.0.1"

    if args.rede:
        if args.debug:
            raise SystemExit(
                "Recusado: --debug junto com --rede expõe um console remoto. "
                "Rode sem --debug."
            )
        print(f"\n  Acesso pelos colegas:  http://{ip_da_rede()}:{args.porta}\n")
        try:
            from waitress import serve
            serve(app, host=host, port=args.porta, threads=8)
        except ImportError:
            print("  (dica: 'pip install waitress' para um servidor mais robusto)\n")
            app.run(host=host, port=args.porta, debug=False, threaded=True)
    else:
        # use_reloader liga o recarregamento automatico ao salvar um arquivo,
        # sem abrir o console do debugger. Sem isso era preciso reiniciar o
        # servidor na mao a cada alteracao — e a tela continuava a antiga.
        app.run(host=host, port=args.porta, debug=args.debug,
                use_reloader=True)

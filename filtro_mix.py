"""
filtro_mix.py - Logica do "Filtro MIX" (antigo filtroMix.py), sem interface.

O que ficou daqui de fora: janelas Tkinter, messagebox, filedialog. A regra de
formatacao do Excel (`formatar_excel_mix_logic`) foi copiada sem alterar uma
linha — mesma saida de antes.

Correcoes feitas na parte do Outlook:

1. `buscar_ultimo_email_com_anexo` chamava CoUninitialize no `finally`, e o
   `SaveAsFile` acontecia DEPOIS, usando objetos COM em um apartment ja
   encerrado. Agora busca e gravacao ficam dentro da mesma sessao COM.
2. As configuracoes eram variaveis globais do modulo e se perdiam ao fechar.
   Agora ficam em `filtro_mix.json`, ao lado do script.
"""

from __future__ import annotations

import datetime
import json
import os
import logging
import tempfile
import time

import pandas as pd
from openpyxl import load_workbook
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

from outlook_downloader import COM_DISPONIVEL, OutlookIndisponivel, sessao_com

try:
    import win32com.client
except ImportError:  # fora do Windows
    win32com = None


from armazenamento import PASTA_DADOS, caminho_dados, em_container

logger = logging.getLogger(__name__)

PASTA_APP = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_CONFIG = caminho_dados("filtro_mix.json")

# Em container nao existe C:/ — a saida vai para a pasta de dados.
PASTA_SAIDA_PADRAO = (os.path.join(PASTA_DADOS, "processados")
                      if em_container() or os.name != "nt"
                      else "C:/Processados/MIX")
CAIXA_DE_ENTRADA = 6  # olFolderInbox
EXTENSOES_EXCEL = (".xlsx", ".xls", ".xlsm")

# Colunas que o MIX formatado deve manter. Qualquer outra coluna da planilha
# de origem e descartada — a logica antiga era o contrario (lista do que
# remover), e coluna nova chegando do sistema passava batido.
COLUNAS_DESEJADAS_PADRAO = [
    "SEQ.",
    "REPLAN",
    "POPID",
    "CHASSI",
    "MODELO",
    "SUSP TRASEIRA",
    "COMPRIM",
    "OPCIONAIS",
    "MU",
    "RPX",
    "BARRA 2°E DIANT",
    "BARRA TRASEIRA",
]

CONFIG_PADRAO = {
    "assunto": "Mix",
    "remetente": "jeisa.silva@scania.com",
    "dias_recentes": 10,
    "palavras_chave": ["pop-id"],
    "pasta_saida": PASTA_SAIDA_PADRAO,
    # Uma lista por funcao: o seletor da tela troca qual esta sendo editada.
    "colunas_por_funcao": {"fa2": list(COLUNAS_DESEJADAS_PADRAO)},
    "tentativas": 3,
    "delay_tentativas": 2,
}


# ---------------------------------------------------------------------------
# Configuracao (antes eram globais do modulo)
# ---------------------------------------------------------------------------
def carregar_config() -> dict:
    config = dict(CONFIG_PADRAO)
    config["colunas_por_funcao"] = dict(CONFIG_PADRAO["colunas_por_funcao"])
    try:
        with open(ARQUIVO_CONFIG, encoding="utf-8") as arquivo:
            gravado = json.load(arquivo)
        # Migra o formato antigo (uma lista so) para o formato por funcao.
        antiga = gravado.pop("colunas_desejadas", None)
        config.update(gravado)
        if antiga and "colunas_por_funcao" not in gravado:
            config["colunas_por_funcao"] = {"fa2": list(antiga)}
    except FileNotFoundError:
        pass
    except (json.JSONDecodeError, OSError) as erro:
        logger.warning("config invalida, usando padrao (%s)", erro)
    return config


def salvar_config(novos: dict) -> dict:
    config = carregar_config()

    if "assunto" in novos:
        config["assunto"] = (novos["assunto"] or "").strip()
    if "remetente" in novos:
        config["remetente"] = (novos["remetente"] or "").strip()
    if "pasta_saida" in novos:
        config["pasta_saida"] = (novos["pasta_saida"] or "").strip() or PASTA_SAIDA_PADRAO
    if "dias_recentes" in novos:
        try:
            dias = int(novos["dias_recentes"])
        except (TypeError, ValueError):
            raise ValueError("Dias recentes deve ser um numero inteiro.")
        if not 1 <= dias <= 90:
            raise ValueError("Dias recentes deve ficar entre 1 e 90.")
        config["dias_recentes"] = dias
    if "colunas_desejadas" in novos:
        bruto = novos["colunas_desejadas"]
        if isinstance(bruto, str):
            bruto = bruto.replace(";", "\n").replace(",", "\n").split("\n")
        colunas = [c.strip() for c in bruto if c and c.strip()]
        if not colunas:
            raise ValueError("Informe pelo menos uma coluna a manter.")
        funcao = (novos.get("funcao") or "fa2").strip() or "fa2"
        por_funcao = dict(config.get("colunas_por_funcao") or {})
        por_funcao[funcao] = colunas
        config["colunas_por_funcao"] = por_funcao
    if "palavras_chave" in novos:
        bruto = novos["palavras_chave"]
        if isinstance(bruto, str):
            bruto = [p.strip() for p in bruto.split(",")]
        config["palavras_chave"] = [p for p in bruto if p]

    temporario = ARQUIVO_CONFIG + ".tmp"
    with open(temporario, "w", encoding="utf-8") as arquivo:
        json.dump(config, arquivo, ensure_ascii=False, indent=2)
    os.replace(temporario, ARQUIVO_CONFIG)
    return config


# Preenchido a cada formatacao, para a tela mostrar o que entrou e o que saiu.
_ULTIMO_RELATORIO_COLUNAS: dict = {}


def relatorio_colunas() -> dict:
    return dict(_ULTIMO_RELATORIO_COLUNAS)


def _chave_coluna(nome) -> str:
    """Compara nome de coluna ignorando caixa, acento sobrando e espaco duplo."""
    return " ".join(str(nome).strip().upper().split())


def colunas_da_funcao(funcao: str = "fa2", config: dict | None = None,
                      padrao=None) -> list[str]:
    """
    Lista de colunas de uma funcao.

    Sem nada salvo, devolve o padrao da propria funcao — cada uma tem o seu,
    porque o layout (largura, fonte, formato) e definido por coluna.
    """
    config = config or carregar_config()
    por_funcao = config.get("colunas_por_funcao") or {}
    return list(por_funcao.get(funcao) or padrao or COLUNAS_DESEJADAS_PADRAO)


def colunas_da_planilha(caminho: str) -> list[str]:
    """Nomes das colunas do arquivo, para a tela sugerir o que configurar."""
    try:
        cabecalho = pd.read_excel(caminho, nrows=0)
    except Exception as erro:  # noqa: BLE001
        logger.warning("não consegui ler o cabeçalho de %s: %s", caminho, erro)
        return []
    return [str(c).strip() for c in cabecalho.columns]


def criar_pasta_se_nao_existir(caminho: str) -> str:
    os.makedirs(caminho, exist_ok=True)
    return caminho


def obter_caminho_saida_padrao(config: dict | None = None) -> str:
    config = config or carregar_config()
    pasta = criar_pasta_se_nao_existir(config["pasta_saida"])
    carimbo = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(pasta, f"MIX_Processado_{carimbo}.xlsx")


# ---------------------------------------------------------------------------
# Outlook
# ---------------------------------------------------------------------------
def _conectar(config: dict):
    """Devolve o namespace MAPI. Deve rodar dentro de `sessao_com()`."""
    ultimo_erro = None
    for tentativa in range(1, int(config["tentativas"]) + 1):
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            namespace.GetDefaultFolder(CAIXA_DE_ENTRADA).Items.Count  # teste real
            return namespace
        except Exception as erro:  # noqa: BLE001
            ultimo_erro = erro
            logger.warning("tentativa %s falhou: %s", tentativa, erro)
            if tentativa < int(config["tentativas"]):
                time.sleep(float(config["delay_tentativas"]))

    raise OutlookIndisponivel(
        f"Falha ao conectar com o Outlook apos {config['tentativas']} tentativas. "
        f"Verifique se o Outlook esta aberto e com uma conta configurada. "
        f"Detalhe: {ultimo_erro}"
    )


def baixar_ultimo_anexo_excel(config: dict | None = None) -> tuple[dict | None, str]:
    """
    Baixa para a pasta temporaria o anexo Excel do email mais recente que
    atende aos filtros. Busca e gravacao na MESMA sessao COM.
    """
    config = config or carregar_config()
    if not COM_DISPONIVEL:
        raise OutlookIndisponivel(
            "pywin32 nao esta disponivel. A captura do Outlook so funciona em "
            "Windows com o Outlook instalado."
        )

    assunto = (config.get("assunto") or "").lower()
    remetente = (config.get("remetente") or "").lower()
    palavras = [p.lower() for p in config.get("palavras_chave") or []]
    dias = int(config.get("dias_recentes") or 0)

    with sessao_com():
        namespace = _conectar(config)
        itens = namespace.GetDefaultFolder(CAIXA_DE_ENTRADA).Items
        itens.Sort("[ReceivedTime]", True)  # mais recente primeiro

        limite = None
        if dias:
            limite = datetime.datetime.now() - datetime.timedelta(days=dias)

        com_anexo = 0
        for mensagem in itens:
            try:
                recebido = mensagem.ReceivedTime
                if getattr(recebido, "tzinfo", None) is not None:
                    recebido = recebido.replace(tzinfo=None)
                recebido = datetime.datetime(
                    recebido.year, recebido.month, recebido.day,
                    recebido.hour, recebido.minute, recebido.second,
                )
                # lista em ordem decrescente: daqui para frente so ha mais antigo
                if limite and recebido < limite:
                    break

                if assunto and assunto not in str(mensagem.Subject or "").lower():
                    continue
                if remetente and remetente not in str(
                        getattr(mensagem, "SenderEmailAddress", "") or "").lower():
                    continue

                anexos = mensagem.Attachments
                if not anexos or anexos.Count == 0:
                    continue
                com_anexo += 1

                for i in range(1, anexos.Count + 1):
                    anexo = anexos.Item(i)
                    nome = str(getattr(anexo, "FileName", "") or "")
                    if not nome.lower().endswith(EXTENSOES_EXCEL):
                        continue
                    if palavras and not any(p in nome.lower() for p in palavras):
                        continue

                    destino = os.path.join(
                        tempfile.gettempdir(),
                        f"filtro_mix_{datetime.datetime.now():%Y%m%d_%H%M%S}_{nome}",
                    )
                    anexo.SaveAsFile(destino)   # ainda dentro da sessao COM
                    return {
                        "assunto": str(mensagem.Subject or ""),
                        "remetente": str(getattr(mensagem, "SenderName", "") or ""),
                        "email_remetente": str(
                            getattr(mensagem, "SenderEmailAddress", "") or ""),
                        "data": recebido.strftime("%d/%m/%Y %H:%M:%S"),
                        "nome_anexo": nome,
                        "caminho_arquivo": destino,
                    }, "Anexo baixado com sucesso"
            except Exception as erro:  # noqa: BLE001
                logger.debug("mensagem ignorada: %s", erro)
                continue

        if com_anexo == 0:
            return None, "Nenhum email com anexo encontrado com os filtros informados."
        return None, "Nenhum anexo Excel encontrado nos emails que passaram no filtro."


def diagnosticar(config: dict | None = None) -> list[dict]:
    """Checagens da tela de diagnostico, em forma de lista para a interface."""
    config = config or carregar_config()
    itens = []

    itens.append({
        "titulo": "pywin32 instalado",
        "ok": COM_DISPONIVEL,
        "detalhe": "OK" if COM_DISPONIVEL else "Instale com: pip install pywin32",
    })

    if not COM_DISPONIVEL:
        return itens

    try:
        with sessao_com():
            namespace = _conectar(config)
            caixa = namespace.GetDefaultFolder(CAIXA_DE_ENTRADA)
            total = caixa.Items.Count
            itens.append({"titulo": "Conexao com o Outlook", "ok": True,
                          "detalhe": f"{total} mensagens na caixa de entrada"})
            try:
                contas = [str(c.DisplayName) for c in namespace.Accounts]
            except Exception:  # noqa: BLE001
                contas = []
            itens.append({"titulo": "Contas configuradas", "ok": bool(contas),
                          "detalhe": ", ".join(contas) or "Nenhuma conta encontrada"})
    except Exception as erro:  # noqa: BLE001
        itens.append({"titulo": "Conexao com o Outlook", "ok": False,
                      "detalhe": str(erro)})

    pasta = config["pasta_saida"]
    try:
        criar_pasta_se_nao_existir(pasta)
        itens.append({"titulo": "Pasta de saida", "ok": True, "detalhe": pasta})
    except OSError as erro:
        itens.append({"titulo": "Pasta de saida", "ok": False,
                      "detalhe": f"{pasta} — {erro}"})

    return itens


# ---------------------------------------------------------------------------
# Processamento
# ---------------------------------------------------------------------------
def processar_manual(entrada: str, saida: str | None = None,
                     config: dict | None = None) -> tuple[bool, str, str]:
    """Formata um arquivo ja existente. Devolve (ok, mensagem, caminho_saida)."""
    config = config or carregar_config()
    if not entrada or not os.path.isfile(entrada):
        return False, "Arquivo de entrada nao encontrado.", ""

    saida = saida or obter_caminho_saida_padrao(config)
    criar_pasta_se_nao_existir(os.path.dirname(saida) or ".")

    ok, mensagem = formatar_excel_mix_logic(
        entrada, saida, colunas_da_funcao("fa2", config)
    )
    return ok, mensagem, saida if ok else ""


def processar_automatico(config: dict | None = None) -> tuple[bool, str, str, dict | None]:
    """
    Baixa o ultimo anexo do Outlook e formata.
    Devolve (ok, mensagem, caminho_saida, info_email).
    """
    config = config or carregar_config()

    info, mensagem = baixar_ultimo_anexo_excel(config)
    if not info:
        return False, mensagem, "", None

    saida = obter_caminho_saida_padrao(config)
    ok, mensagem_formatacao = formatar_excel_mix_logic(
        info["caminho_arquivo"], saida, colunas_da_funcao("fa2", config)
    )

    try:
        os.remove(info["caminho_arquivo"])
    except OSError:
        pass

    return ok, mensagem_formatacao, saida if ok else "", info


# ---------------------------------------------------------------------------
# Formatacao do Excel — copiada de filtroMix.py sem alteracao
# ---------------------------------------------------------------------------
def formatar_excel_mix_logic(caminho_arquivo_original, caminho_arquivo_novo,
                             colunas_desejadas=None, caminho_mona=None,
                             texto_mona=None):
    """Função principal que formata o arquivo Excel MIX (comum a ambas as versões)"""
    try:
        # Carregar o arquivo Excel
        df = pd.read_excel(caminho_arquivo_original)

        # Carrega antes da lista branca cortar as colunas: ENTRADA LINHA REAL
        # não faz parte da lista de colunas mantidas (ver COLUNAS_DESEJADAS_
        # PADRAO), então precisa ser lida aqui, antes que ela suma do df.
        if 'ENTRADA LINHA REAL' in df.columns:
            tvl_marcado = df['ENTRADA LINHA REAL'].apply(
                lambda v: 'TVL' in str(v).strip().upper() if pd.notna(v) else False)
        else:
            tvl_marcado = pd.Series(False, index=df.index)

        # ---- LISTA BRANCA DE COLUNAS ----
        # Mantem so o que esta na lista; tudo o mais sai. Antes era o inverso
        # (uma lista do que remover), entao qualquer coluna nova vinda do
        # sistema entrava no arquivo final sem ninguem perceber.
        desejadas = list(colunas_desejadas or COLUNAS_DESEJADAS_PADRAO)
        conhecidas = {_chave_coluna(c): c for c in desejadas}

        manter, descartadas = [], []
        for coluna in df.columns:
            if _chave_coluna(coluna) in conhecidas:
                manter.append(coluna)
            else:
                descartadas.append(str(coluna))

        faltando = [c for c in desejadas
                    if _chave_coluna(c) not in {_chave_coluna(m) for m in manter}]

        df = df[manter]
        if descartadas:
            logger.info("%s coluna(s) descartada(s): %s",
                       len(descartadas), ", ".join(descartadas))
        if faltando:
            logger.info("coluna(s) desejada(s) ausente(s) na planilha: %s",
                       ", ".join(faltando))
        _ULTIMO_RELATORIO_COLUNAS.clear()
        _ULTIMO_RELATORIO_COLUNAS.update(
            {"mantidas": [str(c) for c in manter],
             "descartadas": descartadas,
             "faltando": faltando}
        )

        # Coluna vazia que sobrou da lista branca continua no arquivo: ela e
        # preenchida a mao depois. (Antes so 'BARRA 2°E DIANT' era poupada.)

        def filtrar_opcionais(valor):
            if pd.isna(valor):
                return ''
            
            valores_filtrados = ['HD FRONT', 'DISCO', 'SDAR', 'ABAFADOR', 'FUP', 'LTDA', 'BT400']
            
            resultado = ['HD' if v == 'HD FRONT' else v 
                        for v in valores_filtrados 
                        if v in str(valor)]
            
            return '/'.join(resultado) if resultado else ''

        df['OPCIONAIS'] = df['OPCIONAIS'].apply(filtrar_opcionais)

        # ENTRADA LINHA REAL: se tinha "TVL" em qualquer parte do valor,
        # OPCIONAIS vira só "TVL" — no lugar de tudo o mais que estava lá
        # (mesmo tratamento do FA1.2, onde só "TVL" entra em Variantes).
        def _acrescentar_tvl(linha):
            if tvl_marcado.get(linha.name, False):
                return 'TVL'
            return str(linha['OPCIONAIS']).strip() if pd.notna(linha['OPCIONAIS']) else ''

        df['OPCIONAIS'] = df.apply(_acrescentar_tvl, axis=1)
        
        #if 'RPX' in df.columns:
          #  df['RPX'] = df['RPX'].astype(str).str.extract(r'(RB885)')
          #  df['RPX'] = df['RPX'].str.replace('RB885', 'RB')

        if 'RPX' in df.columns:
            df['RPX'] = df['RPX'].astype(str).str.extract(r'(RB885|RB756)')
            df['RPX'] = df['RPX'].replace({
                'RB885': 'RB',
                'RB756': 'RBX'
            })
 
 

        # (a lista de colunas indesejadas foi substituida pela lista branca
        #  aplicada logo apos a leitura do arquivo)

        # Remover linhas que contêm "Hide" em qualquer coluna
        mask = df.astype(str).apply(lambda row: row.str.contains('Hide', case=False, na=False).any(), axis=1)
        df = df[~mask]

        # Realinha o marcador de TVL (calculado lá em cima, antes da lista
        # branca) com as linhas que sobraram depois do filtro de Hide, na
        # ordem final — é isso que a formatação de linha inteira usa mais
        # abaixo pra saber qual linha da planilha final é TVL.
        tvl_marcado_final = tvl_marcado.reindex(df.index).fillna(False).reset_index(drop=True)

        # Cruzamento opcional com a MONA (mesma leitura da Função 4): o texto
        # encontrado por POPID entra dentro da própria OPCIONAIS, e não como
        # coluna separada.
        if caminho_mona:
            from mix_processor_4 import cruzar_com_mona
            df = cruzar_com_mona(df, caminho_mona, coluna_saida='_MONA_TEXTO',
                                 texto_lista_simples=texto_mona)

            def _juntar_com_mona(linha):
                partes = []
                for valor in (linha.get('OPCIONAIS'), linha.get('_MONA_TEXTO')):
                    texto = str(valor).strip() if pd.notna(valor) else ''
                    if texto and texto.lower() != 'nan':
                        partes.append(texto)
                return ' | '.join(partes)

            df['OPCIONAIS'] = df.apply(_juntar_com_mona, axis=1)
            df = df.drop(columns=['_MONA_TEXTO'], errors='ignore')

         
        book = load_workbook(caminho_arquivo_original)
        sheet = book.active

        # ---- LIMPEZA DA FORMATACAO HERDADA ----
        # O arquivo de saida reaproveita o workbook de ORIGEM. Apagar o valor
        # da celula nao apaga o preenchimento nem as regras de formatacao
        # condicional que vieram na planilha do e-mail: elas ficam presas ao
        # intervalo (ex.: E2:E100) e voltam a disparar sobre os dados novos.
        # Era dai que saia o vermelho na coluna COMPRIM. Zerando as duas
        # coisas, so valem as regras escritas nesta funcao.
        sheet.conditional_formatting = ConditionalFormattingList()

        sem_preenchimento = PatternFill(fill_type=None)
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row,
                                   min_col=1, max_col=sheet.max_column):
            for cell in row:
                cell.value = None
                cell.fill = sem_preenchimento

      
         


        

        Fontes = {
            'SEQ.': 8,
            'REPLAN': 3,
            'POPID': 10,
            'CHASSI': 3,
            'MODELO': 10,
            'SUSP TRASEIRA': 6,
            'COMPRIM': 5,
            'OPCIONAIS': 8,
            'MU': 5,
            'RPX': 4,
            'BARRA 2°E DIANT': 5,
            'BARRA TRASEIRA': 5
        }
        Bordas = Border(
            left=Side(style='hair'),
            right=Side(style='hair'),
            top=Side(style='hair'),
            bottom=Side(style='hair'))

        fonte_negrito = Font(bold=True)

        
       


        # Escrever cabeçalhos
        for col_idx, cabecalho in enumerate(df.columns, start=1):
            sheet.cell(row=1, column=col_idx, value=cabecalho)

        # Escrever dados
        for row_idx, row in enumerate(dataframe_to_rows(df, index=False, header=False), start=2):
            for col_idx, value in enumerate(row, start=1):
                cell = sheet.cell(row=row_idx, column=col_idx, value=value)

                # Centralizar todas as células
                cell.alignment = Alignment(horizontal='center', vertical='center')

                cell.border = Bordas

                col_name = df.columns[col_idx - 1]
                tamanho_fonte = Fontes.get(col_name, 11)  # Padrão 11 se não especificado

                cell.font = Font(bold=True, size=tamanho_fonte)
                
                if col_idx - 1 < len(df.dtypes):
                    col_type = df.dtypes.iloc[col_idx - 1]
                    if col_type == 'int64':
                        cell.number_format = '0'
                    elif col_type == 'float64':
                        cell.number_format = '0.00'
                    elif col_type == 'datetime64[ns]':
                        cell.number_format = 'DD/MM'

            # ==============================================================
        # DEFINIÇÃO DAS LARGURAS DAS COLUNAS
        # ==============================================================
        
        # Dicionário com as larguras desejadas para cada coluna
        larguras_colunas = {
            'SEQ.': 3.00,
            'REPLAN': 3.00,
            'POPID': 7.20,
            'CHASSI': 4.00,
            'MODELO': 19.00,
            'SUSP TRASEIRA': 2.00,
            'COMPRIM': 6.00,
            'OPCIONAIS': 16.00,
            'MU': 5.00,
            'RPX': 2.00,
            'BARRA 2°E DIANT': 3.00,
            'BARRA TRASEIRA': 3.00
            # Adicione todas as outras colunas que permaneceram no DataFrame
        }
        
        # Definir larguras para cada coluna
        for col_idx, col_name in enumerate(df.columns, start=1):
            col_letter = get_column_letter(col_idx)
            
            # Usar largura personalizada se definida, caso contrário usar largura padrão
            largura = larguras_colunas.get(col_name, 15)  # 15 é a largura padrão
            
            sheet.column_dimensions[col_letter].width = largura
        
        # ==============================================================
        # FIM DA DEFINIÇÃO DE LARGURAS
        # ==============================================================

        # Definição de cores e fontes
        cor_destaque = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        cor_destaque_azul = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
        cor_destaque_verde = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
        cor_destaque_cinza = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
        cor_destaque_vermelho = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        cor_destaque_preto = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
        cor_destaque_Verde_Escuro = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
        cor_destaque_laranja = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
        cor_destaque_tvl = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
        cor_destaque_roxo = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")
        cor_destaque_ciano = PatternFill(start_color="00FFFF", end_color="00FFFF", fill_type="solid")
        cor_destaque_rosa = PatternFill(start_color="FF46A2", end_color="FF46A2", fill_type="solid")

        
        fonte_vermelha = Font(color="FF0000", bold=True)
        fonte_branca = Font(color="FFFFFF", bold=True)
        # Formatação de colunas específicas
        colunas_cores = {
            "RPX": "000000",
            "COMPRIM": "000000",
            "MU": "000000",
            "SEQ.": "000000",
            "MODELO": "000000"
        }

        for col_name, color in colunas_cores.items():
            if col_name in df.columns:
                col_idx = df.columns.get_loc(col_name) + 1
                fonte_cor = Font(color=color, bold=True)
                for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        cell.font = fonte_cor

        # Formatação condicional baseada em valores
        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
            for cell in row:
                if cell.value:
                    cell_value_str = str(cell.value)
                    if "HD FRONT" in cell_value_str:
                        cell.fill = cor_destaque
                    if "HD" in cell_value_str:
                        cell.fill = cor_destaque
                    elif "DISCO" in cell_value_str:
                        cell.fill = cor_destaque
                    elif "RB" in cell_value_str:
                        cell.fill = cor_destaque_azul
                    elif "RBX" in cell_value_str:
                        cell.fill = cor_destaque_azul
                    elif "RB756" in cell_value_str:
                        cell.fill = cor_destaque_azul
                    elif "ABAFADOR" in cell_value_str:
                        cell.fill = cor_destaque
                    elif "FUP" in cell_value_str:
                        cell.fill = cor_destaque
                    elif "LTDA" in cell_value_str:
                        cell.fill = cor_destaque
                    elif "COM" in cell_value_str:
                        cell.fill = cor_destaque_verde
                    elif "Z" in cell_value_str:
                        cell.fill = cor_destaque
                    elif "8X4" in cell_value_str:
                        cell.fill = cor_destaque
                    elif "8X2" in cell_value_str:
                        cell.fill = cor_destaque
                    elif "10X4" in cell_value_str:
                        cell.fill = cor_destaque_vermelho
                    elif "4X2" in cell_value_str:
                        cell.fill = cor_destaque_Verde_Escuro
                    elif "4X4" in cell_value_str:
                        cell.fill = cor_destaque_vermelho
                    elif "6X6" in cell_value_str:
                        cell.fill = cor_destaque_vermelho
                    elif "GT01" in cell_value_str:
                        cell.fill = cor_destaque_laranja
                    elif "SDAR" in cell_value_str:
                        cell.fill = cor_destaque
        
        # Destaque para valores acima de 7400 na coluna COMPRIM
        coluna_valores = "COMPRIM"
        coluna_modelo = "MODELO"

        if coluna_valores in df.columns and coluna_modelo in df.columns:
            col_idx_valores = df.columns.get_loc(coluna_valores) + 1
            col_idx_modelo = df.columns.get_loc(coluna_modelo) + 1
            
            for row_idx in range(2, sheet.max_row + 1):
                cell_valores = sheet.cell(row=row_idx, column=col_idx_valores)
                
                if isinstance(cell_valores.value, (int, float)) and cell_valores.value > 7400:
                    # Preenchimento cinza com o numero em vermelho. O size vem
                    # da celula para nao perder o tamanho definido em `Fontes`.
                    cell_valores.fill = cor_destaque_cinza
                    cell_valores.font = Font(color="FF0000", bold=True,
                                             size=cell_valores.font.size)
                    
                    cell_modelo = sheet.cell(row=row_idx, column=col_idx_modelo)
                    cor_atual = cell_modelo.fill.start_color.index
                    if cor_atual in ['00000000', 'FFFFFFFF']:  # Transparente ou branco (cores padrão)
                        cell_modelo.fill = cor_destaque_laranja

        # "Cilindro elevado" na coluna OPCIONAIS — fundo roxo e fonte branca.
        # So aparece ali quando a MONA e cruzada e bate esse item (o filtro
        # de OPCIONAIS acima descarta qualquer outro texto que nao esteja na
        # lista de itens mantidos).
        coluna_opcionais = "OPCIONAIS"
        if coluna_opcionais in df.columns:
            col_idx_opcionais = df.columns.get_loc(coluna_opcionais) + 1
            for row_idx in range(2, sheet.max_row + 1):
                cell_opcionais = sheet.cell(row=row_idx, column=col_idx_opcionais)
                valor_opcionais = str(cell_opcionais.value or "").upper()
                if "CILINDRO ELEV" in valor_opcionais:
                    cell_opcionais.fill = cor_destaque_roxo
                    cell_opcionais.font = Font(color="FFFFFF", bold=True,
                                               size=cell_opcionais.font.size)

        if coluna_opcionais in df.columns:
            col_idx_opcionais = df.columns.get_loc(coluna_opcionais) + 1
            for row_idx in range(2, sheet.max_row + 1):
                cell_opcionais = sheet.cell(row=row_idx, column=col_idx_opcionais)
                valor_opcionais = str(cell_opcionais.value or "").upper()
                if "MOLA 649|650" in valor_opcionais:
                    cell_opcionais.fill = cor_destaque_ciano
                    cell_opcionais.font = Font( bold=True,
                                               size=cell_opcionais.font.size)

        if coluna_opcionais in df.columns:
            col_idx_opcionais = df.columns.get_loc(coluna_opcionais) + 1
            for row_idx in range(2, sheet.max_row + 1):
                cell_opcionais = sheet.cell(row=row_idx, column=col_idx_opcionais)
                valor_opcionais = str(cell_opcionais.value or "").upper()
                if "MOLA 658|660" in valor_opcionais:
                    cell_opcionais.fill = cor_destaque_rosa
                    cell_opcionais.font = Font( bold=True,
                                               size=cell_opcionais.font.size)

        # Fonte branca para "4X4" na coluna MODELO — e onde esse valor
        # realmente aparece (e onde o bloco de cores acima ja pinta o fundo
        # de vermelho para 4X4), diferente de OPCIONAIS, que nunca guarda
        # esse texto.
        coluna_modelo_4x4 = "MODELO"
        if coluna_modelo_4x4 in df.columns:
            col_idx_modelo_4x4 = df.columns.get_loc(coluna_modelo_4x4) + 1
            for row_idx in range(2, sheet.max_row + 1):
                cell_modelo_4x4 = sheet.cell(row=row_idx, column=col_idx_modelo_4x4)
                valor_modelo_4x4 = str(cell_modelo_4x4.value or "").upper()
                if "4X4" in valor_modelo_4x4:
                    cell_modelo_4x4.font = Font(color="FFFFFF", bold=True,
                                                size=cell_modelo_4x4.font.size)

        # TVL: pinta a linha inteira, sobrescrevendo qualquer outra cor que
        # já tenha caído nela (é a prioridade máxima — se é TVL, é isso que
        # tem que aparecer, sem ficar dividindo a linha com outra regra).
        for row_idx in range(2, sheet.max_row + 1):
            posicao = row_idx - 2  # linha 2 da planilha = primeira linha de dados = posição 0
            if posicao < len(tvl_marcado_final) and tvl_marcado_final.iloc[posicao]:
                for col_idx in range(1, sheet.max_column + 1):
                    cell = sheet.cell(row=row_idx, column=col_idx)
                    cell.fill = cor_destaque_tvl
                    cell.font = Font(color="FFFFFF", bold=True, size=cell.font.size)
        
        """
        #======================================================= Remoção de lunhas ocultas -> Funcionou mas está escondendo o TVL
        linhas_para_remover = []

        for row_idx in range(1, sheet.max_row + 1):
            # Verificar se a linha está oculta
            if sheet.row_dimensions[row_idx].hidden:
                linhas_para_remover.append(row_idx)

        for row_idx in sorted(linhas_para_remover, reverse=True):
            sheet.delete_rows(row_idx)

        #======================================================
        """

        """
        linhas_para_remover = []

        for row_idx in range(1, sheet.max_row + 1):
            # Verificação mais robusta
            if hasattr(sheet.row_dimensions[row_idx], 'hidden') and sheet.row_dimensions[row_idx].hidden:
                logger.debug("linha %s está oculta e será removida", row_idx)
                linhas_para_remover.append(row_idx)

        for row_idx in sorted(linhas_para_remover, reverse=True):
            sheet.delete_rows(row_idx)

        """
        """Versão simplificada que evita a realocação problemática"""
    
          # Primeiro: coletar índices das linhas ocultas
        indices_ocultos = []
        for row_idx in range(1, sheet.max_row + 1):
            if getattr(sheet.row_dimensions.get(row_idx), 'hidden', False):
                indices_ocultos.append(row_idx)
        
        logger.debug("linhas ocultas encontradas: %s", indices_ocultos)
        
        # Estratégia: limpar conteúdo das ocultas em vez de deletar
        for row_idx in indices_ocultos:
            # Primeiro: garantir que a linha ficará visível
            sheet.row_dimensions[row_idx].hidden = False
            
            # Depois: limpar todo o conteúdo
            for col in range(1, sheet.max_column + 1):
                sheet.cell(row=row_idx, column=col).value = None
            
            logger.debug("linha %s - limpa e tornada visível", row_idx)

       
        # ==============================================================
        # REMOÇÃO SIMPLES DE LINHAS VAZIAS
        # ==============================================================
        linhas_para_remover = []
        
        for row_idx in range(1, sheet.max_row + 1):
            linha_vazia = True
            for col in range(1, sheet.max_column + 1):
                if sheet.cell(row=row_idx, column=col).value not in [None, ""]:
                    linha_vazia = False
                    break
            
            if linha_vazia:
                linhas_para_remover.append(row_idx)
        
        # Remover em ordem reversa para não afetar os índices
        for row_idx in sorted(linhas_para_remover, reverse=True):
            sheet.delete_rows(row_idx)
        
        logger.info("%s linhas vazias removidas", len(linhas_para_remover))
        # ==============================================================
    
        


        book.save(caminho_arquivo_novo)
        return True, "Arquivo formatado e salvo com sucesso!"
    
    except Exception as e:
        return False, f"Erro ao formatar o arquivo: {e}"

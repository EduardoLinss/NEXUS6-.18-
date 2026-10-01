"""
outlook_downloader.py - Download automatico das planilhas anexadas no Outlook.

Modulo unico, usado tanto pelo app Tkinter quanto pela versao web. Conversa
com o Outlook instalado na maquina via COM (win32com), entao so funciona em
Windows com o Outlook aberto/configurado.

    from outlook_downloader import OutlookDownloader

    baixador = OutlookDownloader()
    arquivos = baixador.baixar_recentes(dias=7, maximo=3)
    # -> [{'nome':..., 'assunto':..., 'data':..., 'recebido':..., 'caminho':...}, ...]
      (ordem cronologica: mais antigo primeiro)

Correcoes em relacao a versao embutida no TesteCPMporPeriodo2.py:

1. COM inicializado uma unica vez em volta de TODA a operacao. Antes,
   `buscar_emails_com_anexos` chamava CoUninitialize no `finally` e o
   `SaveAsFile` acontecia DEPOIS disso, usando objetos COM em um apartment
   ja finalizado.
2. A caixa de entrada e percorrida do mais novo para o mais antigo e o laco
   para assim que passa da data limite. Antes varria a caixa inteira.
3. Um retorno so (lista de dicionarios) no lugar da tupla de 7 posicoes.
4. Aceita 1, 2 ou 3 arquivos. A tela antiga exigia exatamente 3 e dava erro
   quando so havia 2, apesar de o download suportar menos.
"""

from __future__ import annotations

import os
import logging
import unicodedata
from contextlib import contextmanager
from datetime import datetime, timedelta

logger = logging.getLogger(__name__)

try:
    import pythoncom
    import win32com.client
    COM_DISPONIVEL = True
except ImportError:  # Linux, ou pywin32 nao instalado
    pythoncom = None
    win32com = None
    COM_DISPONIVEL = False


PASTA_PADRAO = os.path.join(os.path.expanduser("~"), "Downloads", "Planilhas_Outlook")
EXTENSOES = (".xlsx", ".xls", ".xlsm")
CAIXA_DE_ENTRADA = 6  # olFolderInbox


class OutlookIndisponivel(RuntimeError):
    pass


def normalizar(texto: str) -> str:
    """minusculas, sem acento e sem separadores: 'POP-ID' -> 'popid'."""
    if not texto:
        return ""
    texto = str(texto).lower()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    for caractere in (" ", "-", "_", ".", "/", "\\"):
        texto = texto.replace(caractere, "")
    return texto


@contextmanager
def sessao_com():
    """Mantem o COM inicializado durante toda a operacao, inclusive o SaveAsFile."""
    if not COM_DISPONIVEL:
        raise OutlookIndisponivel(
            "pywin32 nao esta disponivel. O download automatico do Outlook "
            "so funciona em Windows com o Outlook instalado."
        )
    pythoncom.CoInitialize()
    try:
        yield
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


class OutlookDownloader:
    def __init__(self, pasta_downloads: str = PASTA_PADRAO,
                 filtros_assunto=("mix",), palavras_anexo=("popid",)):
        self.pasta_downloads = pasta_downloads
        self.filtros_assunto = [normalizar(f) for f in filtros_assunto]
        self.palavras_anexo = [normalizar(p) for p in palavras_anexo]

    # -- API principal -------------------------------------------------------
    def baixar_recentes(self, dias: int = 7, maximo: int = 3) -> list[dict]:
        limite = datetime.now().replace(tzinfo=None) - timedelta(days=dias)
        return self._baixar(lambda recebido: recebido >= limite,
                            parar_em=limite, maximo=maximo, prefixo="recentes")

    def baixar_por_periodo(self, data_inicio: datetime, data_fim: datetime,
                           maximo: int = 3) -> list[dict]:
        return self._baixar(lambda r: data_inicio <= r <= data_fim,
                            parar_em=data_inicio, maximo=maximo, prefixo="periodo")

    # -- implementacao -------------------------------------------------------
    def _baixar(self, aceita, parar_em: datetime, maximo: int, prefixo: str) -> list[dict]:
        with sessao_com():
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            itens = namespace.GetDefaultFolder(CAIXA_DE_ENTRADA).Items
            itens.Sort("[ReceivedTime]", True)  # mais novo primeiro

            encontrados = []
            for mensagem in itens:
                if len(encontrados) >= maximo:
                    break
                try:
                    recebido = mensagem.ReceivedTime
                    if getattr(recebido, "tzinfo", None) is not None:
                        recebido = recebido.replace(tzinfo=None)
                    recebido = datetime(recebido.year, recebido.month, recebido.day,
                                        recebido.hour, recebido.minute, recebido.second)

                    # Como a lista esta em ordem decrescente, tudo daqui para
                    # frente e mais antigo que o limite: pode parar.
                    if recebido < parar_em:
                        break
                    if not aceita(recebido):
                        continue

                    assunto = getattr(mensagem, "Subject", "") or ""
                    if not self._assunto_bate(assunto):
                        continue

                    for anexo in self._anexos_validos(mensagem):
                        encontrados.append({
                            "anexo": anexo,
                            "nome": str(anexo.FileName),
                            "assunto": assunto,
                            "recebido": recebido,
                        })
                        if len(encontrados) >= maximo:
                            break
                except Exception as erro:
                    logger.debug("mensagem ignorada: %s", erro)
                    continue

            if not encontrados:
                return []

            encontrados.reverse()  # ordem cronologica: mais antigo primeiro
            return self._salvar(encontrados, prefixo)

    def _assunto_bate(self, assunto: str) -> bool:
        if not self.filtros_assunto:
            return True
        assunto_norm = normalizar(assunto)
        return any(f in assunto_norm for f in self.filtros_assunto)

    def _anexos_validos(self, mensagem):
        anexos = getattr(mensagem, "Attachments", None)
        if not anexos or anexos.Count == 0:
            return
        for i in range(1, anexos.Count + 1):
            anexo = anexos.Item(i)
            nome = str(getattr(anexo, "FileName", "") or "")
            if not nome.lower().endswith(EXTENSOES):
                continue
            if self.palavras_anexo:
                nome_norm = normalizar(nome)
                if not any(p in nome_norm for p in self.palavras_anexo):
                    continue
            yield anexo

    def _salvar(self, encontrados: list[dict], prefixo: str) -> list[dict]:
        # Uma subpasta por lote: os anexos mantem o nome original, que e o que
        # aparece nos avisos de leitura da planilha.
        carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
        pasta = os.path.join(self.pasta_downloads, f"{prefixo}_{carimbo}")
        os.makedirs(pasta, exist_ok=True)

        salvos = []
        for i, item in enumerate(encontrados, 1):
            caminho = os.path.join(pasta, f"{i}_{item['nome']}")
            try:
                item["anexo"].SaveAsFile(caminho)
            except Exception as erro:
                logger.warning("falha ao salvar %s: %s", item['nome'], erro)
                continue
            salvos.append({
                "nome": item["nome"],
                "assunto": item["assunto"],
                "recebido": item["recebido"],
                "data": item["recebido"].strftime("%d/%m/%Y %H:%M"),
                "caminho": caminho,
            })
        return salvos


# ---------------------------------------------------------------------------
# Compatibilidade com a tela Tkinter atual (tupla de 7 posicoes)
# ---------------------------------------------------------------------------
class OutlookAutoDownloader(OutlookDownloader):
    """Mantem a assinatura antiga para nao quebrar TesteCPMporPeriodo2.py."""

    def __init__(self):
        super().__init__()
        self.dias_recentes = 7

    def _para_tupla(self, arquivos, mensagem_ok, mensagem_vazio):
        if not arquivos:
            return (None,) * 6 + (mensagem_vazio,)
        caminhos = [a["caminho"] for a in arquivos] + [None, None, None]
        infos = list(arquivos) + [None, None, None]
        return (caminhos[0], caminhos[1], caminhos[2],
                infos[0], infos[1], infos[2],
                f"{mensagem_ok} ({len(arquivos)} encontrada(s))")

    def baixar_ultimas_duas_planilhas(self):
        try:
            arquivos = self.baixar_recentes(dias=self.dias_recentes, maximo=3)
        except Exception as erro:
            return (None,) * 6 + (str(erro),)
        return self._para_tupla(arquivos, "Planilhas baixadas com sucesso",
                                "Não foi encontrada nenhuma planilha recente no Outlook")

    def buscar_ultimos_tres_arquivos_por_periodo(self, data_inicio, data_fim):
        try:
            arquivos = self.baixar_por_periodo(data_inicio, data_fim, maximo=3)
        except Exception as erro:
            return (None,) * 6 + (str(erro),)
        return self._para_tupla(arquivos, "Planilhas do período baixadas com sucesso",
                                "Nenhum anexo encontrado no período informado.")

"""
armazenamento.py - Onde ficam os arquivos que os usuarios alteram.

Rodando na sua maquina, e a propria pasta do projeto — nada muda.

No Railway (e em qualquer container) o disco e efemero: tudo o que foi escrito
volta ao estado da imagem a cada deploy ou reinicio. As tarefas criadas, as
regras de gargalo e as configuracoes do Filtro MIX sumiriam sozinhas.

Para persistir, monte um Volume no Railway e aponte a variavel de ambiente:

    MENUMIX_DADOS=/dados

Os arquivos de fabrica (`regras_fabrica.json`) continuam vindo da pasta do
codigo — sao somente leitura e nao devem ir para o volume.
"""

from __future__ import annotations

import os

PASTA_CODIGO = os.path.dirname(os.path.abspath(__file__))

# Fora de um container, grava ao lado do codigo (comportamento de sempre).
PASTA_DADOS = os.environ.get("MENUMIX_DADOS", "").strip() or PASTA_CODIGO


def caminho_dados(nome_arquivo: str) -> str:
    """Caminho de um arquivo que os usuarios alteram (JSON de configuracao)."""
    os.makedirs(PASTA_DADOS, exist_ok=True)
    return os.path.join(PASTA_DADOS, nome_arquivo)


def caminho_codigo(nome_arquivo: str) -> str:
    """Caminho de um arquivo somente leitura, versionado junto com o codigo."""
    return os.path.join(PASTA_CODIGO, nome_arquivo)


def pasta_sessao() -> str:
    """Cache das planilhas enviadas. Pode ser efemero — e so cache."""
    caminho = os.path.join(PASTA_DADOS, "sessoes")
    os.makedirs(caminho, exist_ok=True)
    return caminho


def em_container() -> bool:
    """True quando MENUMIX_DADOS foi definido (deploy) ou estamos no Railway."""
    return bool(os.environ.get("MENUMIX_DADOS")
                or os.environ.get("RAILWAY_ENVIRONMENT"))

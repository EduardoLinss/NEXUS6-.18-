# Publicar no Railway

## O problema que precisa ser resolvido primeiro

O disco de um container é **efêmero**. Tudo que o app escreve volta ao estado
da imagem a cada deploy ou reinício — e o Railway reinicia sozinho. Sem
tratamento, as tarefas criadas em `/tarefas`, as regras de `/gargalos/regras` e
as configurações do Filtro MIX desapareceriam sozinhas, sem erro nenhum, o que
é pior do que falhar.

A solução aqui é um **Volume**: um disco de verdade, montado numa pasta, que
sobrevive aos deploys.

## Passo a passo

1. **Suba o código** (GitHub ou `railway up`). O `Procfile` já está pronto.

2. **Crie um Volume** no serviço, com mount path `/dados`.

3. **Defina a variável de ambiente**:

   ```
   MENUMIX_DADOS=/dados
   ```

   É ela que manda o app gravar no volume. Sem ela, ele grava ao lado do
   código — que é o certo na sua máquina e errado no servidor.

4. **Não defina `PORT`** — o Railway define sozinho, e o `Procfile` usa.

O `.gitignore` já exclui `catalogo.json`, `regras.json` e `filtro_mix.json`,
para um deploy não sobrescrever o que os usuários criaram.

## O que fica em cada lugar

| Arquivo | Onde | Muda? |
|---|---|---|
| `catalogo.json` | volume | sim, pela tela de tarefas |
| `regras.json` | volume | sim, pelo construtor de regras |
| `filtro_mix.json` | volume | sim, pelas configurações |
| `sessoes/` | volume | cache das planilhas (pode sumir sem problema) |
| `regras_fabrica.json` | código | não — é a fonte do "restaurar padrão" |

## Um worker, de propósito

```
gunicorn app:app --workers 1 --threads 8
```

Com mais de um worker, cada processo teria sua própria fila do Outlook e sua
própria lista de arquivos gerados pelo Filtro MIX, e os JSON poderiam ser
gravados ao mesmo tempo por dois processos. Threads dão a concorrência
necessária sem esse problema. Para o volume de uso do setor, sobra.

## O que NÃO funciona no Railway

**A captura automática do Outlook.** O `win32com` conversa com o Outlook
instalado na máquina; no servidor não existe nem Windows nem Outlook. A tela
detecta a ausência do `pywin32` e esconde a seção. Restam duas saídas: enviar a
planilha manualmente (funciona igual) ou migrar para a Microsoft Graph API, que
é projeto à parte — registro de app no Azure AD e aprovação da TI.

O mesmo vale para o botão automático do Filtro MIX. O modo manual continua
inteiro.

## Antes de divulgar o endereço

**Não existe login.** No Railway o app ganha uma URL pública: qualquer pessoa
com o link vê o MIX de produção e pode apagar tarefas e regras. Na rede interna
isso era aceitável; publicado, não é.

O mínimo viável é uma senha única via HTTP Basic — dá para fazer em poucas
linhas. Vale resolver isso **antes** do primeiro deploy, não depois.

## Backup

O volume protege contra reinício, não contra alguém clicar em "restaurar
padrão". Use os botões de baixar `.json` das telas de tarefas e de regras de
vez em quando — são backups completos, e restaurar é substituir o arquivo.

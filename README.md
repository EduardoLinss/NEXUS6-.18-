# MENUMIX Web — protótipo

Mesma análise do app Tkinter, servida como página HTML.
`cpm_core.py` é **byte a byte o mesmo arquivo** da versão desktop — só a
camada de apresentação mudou.

## Rodar

```
pip install flask pandas openpyxl
python app.py
```
Abre em http://127.0.0.1:5000

## Estrutura

```
app.py                 ~140 linhas — rotas, upload, cache do DataFrame
templates/index.html   ~200 linhas — formulário + colunas por variante (CSS grid)
cpm_core.py            inalterado
```

As colunas são `grid-template-columns: repeat(auto-fit, minmax(230px, 1fr))`.
Isso é o `PainelVariantes` inteiro do Tkinter (≈200 linhas de Canvas,
scrollbars e `place`) em uma linha de CSS. Acrescentar uma variante V5 não
exige nada: a grade se reorganiza sozinha.

## Cache das planilhas

Ao enviar as duas planilhas, o DataFrame concatenado vai para um pickle em
`%TEMP%/menumix_web/` e o token volta num campo escondido. Trocar o POPID e
reanalisar não reenvia arquivo. Limpe a pasta de vez em quando.

## O que este protótipo ainda NÃO faz

- captura automática do Outlook (`win32com`) — ver observação abaixo
- tela de gargalo (`gargalo3.py`)
- filtroMix
- login / histórico

## Outlook

`win32com.client` conversa com o Outlook **instalado na máquina**. Enquanto o
Flask rodar no seu Windows, isso continua funcionando: basta chamar o
`OutlookAutoDownloader` dentro de uma rota. Se um dia isso virar site em
servidor (Railway), a automação COM não vai junto — nesse cenário as opções
são upload manual ou Microsoft Graph API.

## Captura automática do Outlook

`outlook_downloader.py` é o **mesmo arquivo** usado pelo app desktop.

Na tela: **BAIXAR RECENTES** (últimos N dias) ou **BAIXAR POR PERÍODO** (datas).
Se POPID e produção já estiverem preenchidos, ele baixa, concatena e analisa
numa tacada só. O checkbox "buscar ao abrir a página" fica salvo no navegador
(`localStorage`) e dispara a busca sozinho quando você abre o `localhost:5000`.

Detalhe importante do Flask: cada requisição roda numa thread nova, e COM exige
`CoInitialize` por thread. O `sessao_com()` cuida disso — por isso o download
funciona igual no servidor local e no desktop.

Se `pywin32` não estiver instalado, a seção some e a tela avisa; o envio manual
das planilhas continua disponível.

    pip install pywin32

## Editor de tarefas (`/tarefas`)

O catálogo saiu do código e virou `catalogo.json`. A tela permite:

- **adicionar** tarefa: nome, variante, tempo em segundos e até 3 condições
- **editar** (inclusive mover de variante) e **excluir**
- **criar/excluir variante** (V5, V6…)
- **restaurar padrão** — volta ao catálogo original de fábrica

Condições disponíveis: *contém*, *não contém*, *é maior que*, *é menor que*.
Todas as condições de uma tarefa são combinadas com **E**.

Com uma planilha carregada, o campo de coluna sugere as colunas reais e o
campo de valor sugere os valores que aparecem nelas. Sem planilha, dá para
digitar o nome da coluna.

Se `catalogo.json` não existir, o catálogo embutido em `cpm_core.CATALOGO_PADRAO`
é usado e o JSON nasce dele na primeira gravação — nada se perde. A gravação é
atômica (`.tmp` + `os.replace`).

O app desktop lê o **mesmo** `catalogo.json`: basta manter `catalogo.py` e
`catalogo.json` na pasta. Os campos de absenteísmo das duas telas são gerados
a partir do catálogo, então uma variante nova ganha campo automaticamente.

## Cabeçalho de planilha repetido

Planilha com nome de coluna repetido (`OPCIONAIS` duas vezes) ou célula de
cabeçalho vazia quebrava o `pd.concat` com
`InvalidIndexError: Reindexing only valid with uniquely valued Index objects`.

Na leitura, coluna sem nome vira `COLUNA_<posição>` e nome repetido ganha
sufixo (`OPCIONAIS`, `OPCIONAIS.1`) — o mesmo que o `pd.read_excel` faz. Nada
é descartado, e a tela mostra um aviso amarelo dizendo o que foi renomeado.

As planilhas passaram a ser alinhadas pela **união** das colunas: o que existe
só em um arquivo vira vazio nos outros em vez de ser descartado em silêncio.

## Acesso de outras máquinas (e do celular)

```
python app.py --rede
```

Ele mostra o endereço a passar adiante (`http://192.168.x.x:5000`). Sem
`--rede`, só a sua máquina enxerga.

**No celular, `localhost` não funciona.** `localhost` significa "este
aparelho": digitado no celular, ele procura um servidor dentro do próprio
celular. Use o IP que o `--rede` imprime, com o celular no **mesmo Wi‑Fi** do
PC.

Se não abrir, na ordem: o Windows precisa liberar a porta no firewall (ele
pergunta na primeira execução, e é preciso marcar rede privada); e algumas
redes corporativas e de visitantes isolam os aparelhos entre si, o que impede
qualquer conexão direta — nesse caso não há ajuste no código que resolva.

- `--debug` junto com `--rede` é **recusado**: o debugger do Flask abre um
  console Python remoto para quem alcançar a porta.
- Se `waitress` estiver instalado (`pip install waitress`), ele é usado no
  lugar do servidor de desenvolvimento.
- Sua máquina precisa ficar ligada, e as buscas no Outlook usam **a sua**
  caixa de entrada — os colegas veem as planilhas que chegam para você.
  Chamadas simultâneas ao Outlook entram em fila (`trava_outlook`).
- `catalogo.json` é compartilhado: se duas pessoas editarem tarefas ao mesmo
  tempo, vale a última gravação.

### Layout no celular

Todas as telas têm `<meta viewport>` e um bloco `@media (max-width: 700px)`:
as colunas das variantes viram uma só, os campos ocupam a largura toda, os
botões ganham altura de toque de 44px e a tabela de tarefas vira lista. Campos
de texto usam 16px no celular, senão o iOS dá zoom sozinho ao focar.

Não é um app separado — é a mesma página se ajustando. Só a tela de análise com
muitas variantes exige rolagem vertical, o que é esperado.

## Telas

```
/             menu inicial (só leva às telas)
/analise      Análise de MIX — carga por variante
/filtro-mix   Filtro MIX — Excel formatado a partir do Outlook ou de um arquivo
/tarefas      cadastro das regras das variantes
```

### Filtro MIX

Porte do antigo `filtroMix.py`. A regra de formatação do Excel
(`formatar_excel_mix_logic`) foi copiada **sem alterar uma linha** — a saída é
idêntica à do app Tkinter. O que ficou de fora foi só a camada de janelas.

Os mesmos controles de antes: **automático** (baixa o último anexo do Outlook
que passa nos filtros e formata), **manual** (envia um arquivo e formata),
**configurações** (assunto, remetente, dias, palavras-chave do anexo, pasta de
saída) e **diagnóstico**.

Duas diferenças que a web exige:

- No manual, o arquivo é enviado pelo navegador em vez de escolhido por
  `filedialog` — um colega em outra máquina não enxerga o disco do servidor.
- O resultado é salvo na pasta de saída **da máquina que roda o servidor** e
  também fica disponível por download, pelo mesmo motivo.

As configurações eram variáveis globais e se perdiam ao fechar o app. Agora
ficam em `filtro_mix.json`, ao lado do script.

### Colunas: lista branca

A lógica era uma lista de **colunas indesejadas** (`colunas_indesejadas`, ~45
nomes): qualquer coluna nova vinda do sistema entrava no arquivo final sem
ninguém perceber, e a lista precisava crescer para sempre.

Agora é o contrário: uma lista de **colunas a manter** (`colunas_desejadas`),
editável nas Configurações, uma por linha. Tudo que estiver fora dela é
descartado. A comparação ignora caixa e espaço extra, então `popid`,
`POPID` e ` POPID ` são a mesma coluna.

Depois de cada processamento a tela mostra três listas: mantidas, descartadas
e — em amarelo — as que estão na lista mas **não existem** na planilha, que é
o sinal de que a origem mudou um nome de coluna.

Efeito colateral bom: a regra `COLUNA_PARA_MANTER = 'BARRA 2°E DIANT'`, que
existia só para poupar essa coluna da remoção de colunas vazias, deixou de ser
necessária — qualquer coluna da lista branca sobrevive mesmo vazia.

O `filtroMix.py` original não é mais usado pela versão web.

## Gargalos da sequência

O motor (`analisador_cpm.py`) veio pronto e **não foi alterado**. O que mudou é
de onde vêm os dados: em vez de reler um arquivo, `gargalos.py` monta as
unidades a partir do mesmo recorte de DataFrame que a análise de MIX usa — o
POPID digitado e a produção do dia. Os gargalos e a carga por variante falam
sempre do mesmo intervalo.

**Onde aparece**

- `/analise` — painel abaixo de EXPORTAR RESULTADO, junto com o resto
- `/gargalos` — só os gargalos, com POPID e produção próprios
- `/gargalos/regras` — construtor de regras

**O painel** traz o resumo (quantos de alta e de média, quantos veículos
afetados), uma faixa andon com um bloco por veículo na ordem de produção, e a
tabela com regra, faixa de posições, POPID/SEQ de cada veículo envolvido e o
que fazer.

**Construtor de regras** — os três tipos do motor:

| Tipo | Parâmetros |
|---|---|
| `unidade_critica` | nenhum — o produto sozinho já é gargalo |
| `contagem_janela` | N ocorrências dentro de uma janela de X carros |
| `sequencia_consecutiva` | K ou mais carros colados |
| `coocorrencia_janela` | dois grupos coincidindo numa janela |

Cada tipo aceita até 3 condições combinadas com E ou OU. Operadores: contém,
não contém, é igual, é diferente, **é maior que**, é maior ou igual, **é menor
que**, é menor ou igual, está vazio, não está vazio.

Os quatro de comparação são **numéricos** — comparam número, não texto, que é o
que `COMPRIM > 7400` exige (como texto, "800" seria maior que "7400"). Valor não
numérico é recusado na hora de salvar.

**`unidade_critica`** é para a variante que já é gargalo por si só: cada carro
que atende às condições vira uma faixa de uma posição. Duas seguidas viram uma
faixa única, e o texto de ação muda — "produto de alta complexidade" para uma,
"intercalar produtos simples" para várias em sequência. Ela também pode estar
dentro de uma sequência complicada; as duas regras disparam e aparecem
separadas.
O id sai do nome automaticamente.

As regras ficam em `regras.json`. `regras_fabrica.json` guarda as originais e
nunca é escrito — é dele que sai o "restaurar padrão".

## Relatório em PDF

Botão **RELATÓRIO EM PDF** no bloco de exportação da análise e na tela de
gargalos. Gera um arquivo único, em paisagem, com:

1. cabeçalho — POPID, intervalo de linhas, veículos, data, e os avisos de
   leitura da planilha, se houver
2. resumo de carga por variante (tarefas + absenteísmo + total, com cor por
   faixa)
3. uma tabela por variante, tarefa a tarefa
4. CPM MIX em três colunas
5. demanda ao longo do turno — a **linha do tempo** (eixo de horas, faixa de
   veículos e a curva de demanda contra a capacidade), a faixa da sequência (um bloco por veículo,
   colorido pela severidade, com a posição inicial marcada no começo de cada
   linha) e a tabela com regra, posições, POPID/SEQ dos veículos e o que fazer

O PDF é montado do **mesmo** POPID e produção do dia da tela, incluindo os
absenteísmos digitados: o link carrega esses valores, então o arquivo bate com
o que está na tela.

Precisa de `reportlab`:

```
pip install reportlab
```

Sem a biblioteca, o botão some e a tela explica como instalar — o resto do app
continua funcionando.

## Cache das planilhas entre as telas

O DataFrame concatenado é gravado em `%TEMP%/menumix_web/<token>.pkl` e o
token vai para um **cookie** (`menumix_token`). Antes ele viajava só em campo
escondido do formulário, então qualquer link para outra aba perdia a planilha
e era preciso enviar tudo de novo.

Agora, ao trocar de tela e voltar, a planilha continua carregada — e o POPID e
a produção do dia também voltam preenchidos (cookies `menumix_popid` e
`menumix_producao`).

- Validade: 7 dias no navegador; os arquivos em disco são apagados depois de
  24 horas, na próxima vez que alguém carrega uma planilha.
- É por navegador: cada colega tem o seu token, então dois acessos simultâneos
  não misturam dados.
- Limpar os cookies do navegador tem o efeito de "recomeçar do zero".


## Absenteísmo do setor

O campo passou a ser **um só** — o número de pessoas ausentes no setor — em
vez de um por variante. E o efeito mudou: antes o absenteísmo entrava como um
tempo fixo somado à variante; agora quem cobre a ausência são os postos que
sobraram, e é isso que a tela mostra.

| Ausentes | Postos ativos | Mensagem |
|---|---|---|
| 0 ou 1 | todos | — |
| 2 | 2 variantes | *Necessário recurso*, se houver algum gargalo |
| 3 ou mais | 1 variante | *Linha descoberta para várias tarefas*; e *Necessário recurso* se houver vários gargalos |

É uma tabela, não uma subtração: com 1 ausência a distribuição segue normal.
Os limites ficam em `cpm_core.POSTOS_POR_FALTAS` e `GARGALOS_VARIOS`
("vários" = 2 ou mais).

As tarefas são repartidas entre os postos restantes sempre jogando a próxima
no posto mais leve, o que deixa os totais próximos entre si. Cada tarefa vinda
de outra variante aparece com uma etiqueta (`V3`) ao lado do nome, para você
ver de onde veio. A carga total do setor se conserva — só muda em quantas
colunas ela está dividida.

Os alertas aparecem acima das colunas, e também no PDF e nas duas abas de
exportação.


## Linha do tempo do turno (PDF)

Mostra, ao longo do turno: o eixo de horas, um bloco por veículo colorido pela
severidade, e a curva de **demanda de trabalho** contra a **capacidade**.

A demanda de cada veículo é a soma dos tempos das tarefas que se aplicam a ele,
no mesmo percentual usado na carga por variante. Por isso a média da curva bate
exatamente com a carga total do setor — não é estimativa, é o mesmo cálculo
visto por veículo em vez de somado no fim.

A capacidade é `postos_ativos × 100%`: com absenteísmo, a linha tracejada desce
sozinha de 400% para 300% ou 100%.

O eixo de horas vem dos campos **Início** e **Fim do turno** (padrão
07:05–16:05) e das pausas, em `cpm_core.PAUSAS_PADRAO`:

```python
PAUSAS_PADRAO = (
    ("09:00", "09:04"),
    ("10:55", "11:55"),
    ("14:00", "14:10"),
)
```

O ritmo é derivado: 9 h de turno menos 74 min de pausa = 7 h 46 de produção;
dividido por 88 veículos dá 318 s (5,3 min) cada. As pausas empurram o relógio,
então as horas em volta de uma parada aparecem coladas no eixo — nenhum veículo
passa nesse intervalo. Cada pausa é marcada com uma linha tracejada.

Esse ritmo **não** é o mesmo que `SEGUNDOS_DISPONIVEIS_POR_UNIDADE` (300 s), que
segue mandando na fórmula de carga. São coisas diferentes: um é o relógio, o
outro é o tempo de ciclo do cálculo. Quando os dois divergem mais de 30 s, a
legenda avisa em amarelo. Com as pausas descontadas os valores ficaram a 18 s
um do outro (318 contra 300), o que sugere que o tempo de ciclo do cálculo
está coerente com o turno real.

A curva exibida é média móvel de 5 veículos — a curva crua é serrilhada demais
para ler. As cores seguem a razão demanda/capacidade: verde abaixo de 60%,
amarelo até 85%, laranja até 100%, vermelho acima.

**Presenteísmo fica de fora** — não existe essa informação em nenhuma fonte do
sistema hoje.


### Sem gargalo

A seção passou a se chamar **Demanda ao longo do turno** e o gráfico aparece
sempre — "nenhum gargalo" não quer dizer "carga tranquila", e é justamente a
curva que mostra isso. Sem gargalos, saem só a tabela de faixas e a faixa
andon, que ficariam vazias ou monocromáticas.

Quando o pico de demanda fica bem abaixo da capacidade, a escala segue o pico
em vez da capacidade — senão a curva ficaria esmagada na base. Nesse caso a
linha tracejada vai para o topo com uma seta (`400% ↑`), indicando que está
fora da escala.


## Ordem das planilhas enviadas

O navegador entrega os arquivos na ordem em que o sistema listou, não na ordem
que interessa — por isso o envio manual às vezes vinha invertido e a
concatenação ficava fora de sequência.

Agora a ordem sai da **data no nome do arquivo**, no formato que o MIX usa:

```
mix a partir do pop-id 775301 ( 10_09 ).xlsx
```

A mais antiga vai primeiro. São aceitos `( 10_09 )`, `(10-09)`, `(10/09/2026)`
e variações com ou sem espaço, entre parênteses ou colchetes.

Quando só há dia e mês, um par 31/12 e 01/01 é lido como virada de ano, não
como seis meses de distância.

Se algum arquivo não tiver data no nome, a ordem de envio é mantida e a tela
avisa em amarelo. Quando a ordem é corrigida, o aviso mostra a sequência
usada — vale conferir na primeira vez.

## Tutorial

`/ajuda` — página com o guia de uso, ligada no menu como **Como usar** e no
topo da tela de análise. Nove seções: carregar planilhas, analisar, absenteísmo,
gargalos, cadastro de tarefas, regras de gargalo, exportar/importar/PDF, Filtro
MIX e uma tabela de sintomas e causas prováveis.

É uma página estática (`templates/ajuda.html`) — não depende de dados
carregados, então abre mesmo antes de enviar qualquer planilha. Ao mudar
comportamento do sistema, vale revisar a seção correspondente.

## Setores

Antes o sistema atendia um setor só, e os parâmetros do FA2 — takt, turno,
pausas, tabela de postos por falta, limiar de gargalos — eram constantes de
módulo no `cpm_core`. Pareciam parâmetros do sistema, mas eram do FA2.

Agora são configuração por setor, em `setores.py` e `setores.json`. Tela
`/setores` para cadastrar, e uma barra de seleção no topo de todas as telas.

**Cada setor tem os seus dados:**

```
dados/<setor>/catalogo.json   tarefas das variantes
dados/<setor>/regras.json     regras de gargalo
```

O `setores_fabrica.json` acompanha o código e traz o FA2. Os arquivos antigos
(`dados/catalogo.json` e `dados/regras.json`, de quando havia um setor só) são
adotados como dados do FA2 enquanto os novos não existirem — nada do que já foi
criado se perde na atualização.

**Takt na tela:** o campo *Takt (s/veículo)* na análise vem preenchido com o do
setor e pode ser alterado só para aquela análise — útil para simular um ritmo
diferente sem mexer no cadastro. Quando o valor usado difere do cadastrado, a
tela avisa e oferece o link para alterar no setor. O PDF sai com o mesmo takt
da tela.

**O que o setor controla:** takt, início e fim do turno, pausas, tabela de
absenteísmo (`faltas=postos`) e a partir de quantos gargalos se considera
"vários". O mesmo intervalo analisado em dois setores dá cargas diferentes,
porque o takt é outro.

**Padrões da fábrica.** O botão de engrenagem no canto superior direito de
`/setores` abre um popup com o que quase não muda entre setores: produção do
dia, takt, postos, turno, pausas, tabela de absenteísmo e o limiar de "vários
gargalos".

Eles fazem duas coisas: preenchem o formulário de um **setor novo** e dão o
valor inicial da produção no painel da fábrica. Alterar os padrões **não** mexe
nos setores já cadastrados — cada um mantém o que foi salvo nele, e os campos
continuam todos editáveis por setor.

**Postos × variantes:** quem manda no cálculo são as variantes do catálogo do
setor. O campo "postos" é referência; quando os dois divergem, a tela de
setores mostra o número de variantes em amarelo.

O setor escolhido acompanha a navegação por cookie e por `?setor=` nos links.

## Funções de MIX (tela Filtro MIX)

A tela `/filtro-mix` atende todas as funções. O seletor no topo troca a função
e o resto da tela se adapta ao que ela declara em `processadores.py`:

| Função | Entradas | Parâmetros | Lista de colunas |
|---|---|---|---|
| FA2 | MIX | — | sim |
| FA1.2 | MIX | — | não |
| Fibras | MIX geral | — | não |
| Função 4 | MIX + MONA (+ 2º MIX opcional) | TAKT, bloco, produção, POPID de corte, início do turno | não |

Acrescentar uma função é acrescentar uma entrada no `CATALOGO` — sem rota nova
nem template novo. O contrato: a função recebe os caminhos de entrada, depois
os parâmetros, e por último o caminho de saída; devolve `(ok, mensagem)`.

**O que muda conforme o seletor**

- os campos de arquivo (a Função 4 pede três)
- os parâmetros (só a Função 4 tem)
- a lista de colunas a manter — guardada **por função**, em
  `colunas_por_funcao` no `filtro_mix.json`. FA1.2, Fibras e Função 4 têm
  layout escrito na própria função, então no lugar da lista aparece uma nota
  explicando; configurar algo ali não teria efeito.
- a captura do Outlook, que só aparece para funções de um arquivo só — ela
  traz um anexo, e a Função 4 precisa de MIX e MONA.

Os filtros de e-mail (assunto, remetente, dias, palavras-chave) e a pasta de
saída valem para todas as funções.

Campos com `exigido_com` só são cobrados quando o arquivo opcional
correspondente é enviado: sem a 2ª planilha MIX, os parâmetros de CPM não são
pedidos e a aba `ANALISE CPM` não é gerada.


## Painel da fábrica

`/painel` — todos os setores numa tela, a partir da **mesma** planilha.

Fluxo: carregar as planilhas (Outlook ou envio manual), preencher **produção do
dia** e **takt** no topo — valem para a fábrica toda — e o **POPID de corte** de
cada setor. Setor sem POPID fica de fora, que é como se olha só parte da linha.

O takt em branco faz cada setor usar o seu, cadastrado em `/setores`.
Preenchido, vale para todos e ignora o cadastro.

**O que aparece**

- quatro números da fábrica: setores analisados, gargalos (alta/média),
  variantes em 100% ou mais, e postos ativos contra o total
- os alertas de absenteísmo de todos os setores, com o nome do setor na frente
- um cartão por setor: contagem de gargalos, barra de carga por variante e uma
  curva de demanda com a linha de capacidade
- um gráfico com uma linha por **variante**, no eixo da **hora do dia**
- a tabela de gargalos da linha inteira, ordenada por severidade e depois pela
  posição na sequência

As curvas são SVG geradas no servidor — sem biblioteca de gráfico e sem
JavaScript, o que importa numa rede onde instalar coisa nova é caro.

**Erro de um setor não derruba os outros:** POPID que não existe naquele
intervalo vira mensagem no cartão daquele setor, e o resumo informa quantos
ficaram com erro.

Os POPIDs, produções e faltas ficam num cookie (`menumix_painel`), então voltam
preenchidos no dia seguinte.


### Relatório da fábrica em PDF

O botão **RELATÓRIO EM PDF**, no cabeçalho da tabela de gargalos do painel,
gera um documento com a fábrica inteira:

1. capa com os números do dia — setores, gargalos, variantes em 100% ou mais,
   postos ativos e ausências — mais os alertas de absenteísmo
2. o mesmo gráfico da tela: uma linha por função no eixo do relógio, com as
   trilhas de gargalo embaixo
3. carga por setor, com o percentual de cada variante colorido por faixa
4. a tabela de gargalos de toda a linha, com POPID/SEQ dos veículos

O PDF refaz a análise a partir dos parâmetros do link (turno, produção, takt,
POPIDs e faltas — inclusive quando o campo é um horário, não um número), então
bate com o que está na tela. Traz também a tabela **Momentos mais críticos do
dia**, logo abaixo do gráfico, quando algum setor tem horário de falta
configurado. Precisa de `reportlab`;
sem a biblioteca, o botão não aparece.

O gráfico é redesenhado para caber na largura da página — na tela ele rola na
horizontal, o que no papel não existe.

### Gráfico de demanda ao longo do dia

Uma linha por **função**, no eixo da **hora do dia**.

**Turno.** O seletor no topo escolhe entre o 1º turno — que usa o horário
cadastrado em cada setor — e o 2º, das **17:30 às 02:08**, igual para a fábrica
toda. O 2º turno atravessa a meia-noite e o eixo trata isso: as horas vão de
18:00 a 01:00 na ordem certa.

No 2º turno o **takt é obrigatório**: o valor do cadastro é do 1º turno, e usar
ele ali daria números errados sem avisar. As pausas cadastradas são de manhã,
caem fora da janela e são descartadas sozinhas — se o 2º turno tiver pausas
próprias, precisam ser cadastradas no setor.

**Produção, takt e POPIDs são guardados por turno.** O 2º turno não produz a
mesma quantidade nem começa no mesmo ponto de corte, então cada turno tem os
seus números. Trocar no seletor navega e traz de volta os dados daquele turno,
sem sobrescrever os do outro.

O turno vale só para a análise: não altera nada no cadastro dos setores.

**Eixo horizontal: hora do dia.** Cada setor começa num POPID diferente e pode
ter turno diferente; só o relógio permite dizer "às 10h esses dois apertam ao
mesmo tempo". Os horários saem do turno de cada setor com as pausas
descontadas. O gráfico cresce ~190px por hora e rola na horizontal.

**Eixo vertical: % da capacidade daquela função.** 100% é o limite dos postos
ativos dela, então um setor de 4 postos e outro de 1 cabem na mesma escala.

O ponto cheio marca o pico de cada função, e a legenda diz a que horas ele
acontece, a média do turno e qual variante puxa o pico — o detalhe por variante
fica no cartão do setor, não no gráfico geral.

**Exportar as regras.** No cabeçalho da tabela de gargalos do painel há
`regras · .tsv` e `regras · .json`, com as regras de **todos os setores** num
arquivo só — o `.tsv` traz o setor na primeira coluna e a regra em texto legível;
o `.json` é backup completo, um bloco por setor. A tela `/gargalos/regras`
continua exportando e importando o setor atual.

**Gargalos no gráfico.** Abaixo das curvas, uma trilha por setor, na cor dele,
marca quando as regras dispararam — cor cheia para severidade alta, esmaecida
para média. Os trechos de alta também aparecem sombreados sobre as curvas.

Gargalo e pico de demanda são coisas diferentes: um trecho pode disparar regra
sem a curva subir (carros pesados chegando juntos, sem folga entre eles) e um
pico pode não ter regra nenhuma. Ver os dois no mesmo eixo é o que mostra
quando os problemas se somam.


## Absenteísmo por horário e momentos críticos

O campo **Faltas** de cada setor aceita dois formatos:

- um número — o absenteísmo vale para o turno inteiro, como sempre
- um horário por linha — `10:00-12:00=2` — quando o número de ausentes muda ao
  longo do turno

Com horário, a capacidade do setor passa a variar por veículo: fora das
janelas, o setor está completo; dentro delas, os postos ativos caem conforme a
tabela de absenteísmo do setor (`postos_por_faltas`).

**Momentos mais críticos do dia** é a tabela que cruza essa informação com os
gargalos: para cada regra que disparou, se o trecho cai dentro de uma janela de
falta, o momento entra na lista, com a sobrecarga real — demanda dividida pela
capacidade reduzida daquele instante, não pela capacidade plena. Ordenado do
pior para o melhor. Um gargalo que aconteceu com o setor completo não entra;
só interessa aqui o cruzamento montagem-complexa + menos-gente.

O gráfico de demanda ao longo do dia também reflete isso: durante uma janela de
falta, a mesma sequência de veículos faz a linha daquele setor subir, porque a
capacidade (o denominador) caiu.

Setor sem horário de falta configurado continua com o comportamento de sempre
— um número fixo, sem entrar na tabela de momentos críticos, porque não há
"momento" de falta para cruzar.


## Presenteísmo na tela de análise

Ao lado do campo **Absenteísmo** (que falta o turno inteiro), o checkbox
**Presenteísmo** cobre o outro caso: pessoas que estão no turno mas saem por um
período. Marcado, abre um campo de quantidade; digitando um número, abrem esse
tanto de pares "hora início / hora fim" — um por pessoa.

Cada janela soma 1 ao absenteísmo **só durante aquele intervalo**; janelas que
se sobrepõem se somam (duas pessoas fora ao mesmo tempo pesam como duas
faltas). Fora das janelas, o posto continua completo. O restante do turno usa
o absenteísmo fixo do campo ao lado, normalmente.

Isso alimenta dois lugares:

- o resumo "Absenteísmo: N pessoa(s)" e a redistribuição de tarefas passam a
  usar o **pior instante do turno** — o maior número de gente fora ao mesmo
  tempo, mesmo que seja por poucos minutos
- a tabela **Momentos com presenteísmo**, que cruza os gargalos da sequência
  com as janelas: um gargalo que caiu num trecho de posto reduzido entra na
  lista, com a sobrecarga real (demanda dividida pela capacidade daquele
  instante, não a plena). Um gargalo que aconteceu com o posto completo não
  entra — só interessa aqui o cruzamento montagem-complexa + gente fora.

É a mesma lógica do horário de falta do painel da fábrica (ver seção acima),
aplicada ao setor único desta tela, com uma interface pensada para poucas
pessoas em vez de um texto livre — e com a diferença de que aqui as janelas
somam entre si, já que cada uma representa uma pessoa distinta.

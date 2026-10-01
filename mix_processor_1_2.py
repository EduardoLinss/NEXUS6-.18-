import logging
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)




COLUNAS_PADRAO = (
    "POPID", "CHASSI", "MODELO", "SUSP TRASEIRA", "COMPRIM", "MU", "SORDER", "Variantes",
)

# Itens que a MONA da FA1.2 registra mas que já existem em alguma coluna do
# mix geral da fábrica (GÁS já vem de uma coluna própria do MIX e cai em
# Variantes por uma regra abaixo; 6X6/4X4 já é distinguível pelo MODELO) —
# por isso ficam de fora do cruzamento. Ao contrário do FA2 (poucas
# variantes novas, lista permitida em cpm_core.tags_mona_ativas), na MONA da
# FA1.2 a maioria dos itens é nova, então aqui é mais simples excluir só o
# que já existe do que listar tudo que deve entrar.
ITENS_MONA_FA12_JA_NO_MIX: tuple[str, ...] = (
   # r"\bG[ÁA]S\b",
   # r"6X6\D*4X4",
   # r"8X2\D*8X2",
   # r"8X4\D*8X4",
   # r"LTDA\D*LTDA",

    r"\bG[ÁA]S\b",
    r"6X6\D*4X4",
    r"\b8X2\b",
    r"\b8X4\b",
    r"\bLTDA\b",
)


def _chave_coluna(nome) -> str:
    """Compara nome de coluna ignorando caixa e espaco extra."""
    return " ".join(str(nome).strip().upper().split())


def formatar_excel_mix_logic_funcao_um_dois(caminho_arquivo_original,
                                            caminho_arquivo_novo,
                                            colunas_desejadas=None,
                                            caminho_mona=None,
                                            texto_mona=None):
    """Função principal que formata o arquivo Excel MIX (comum a ambas as versões)"""
    try:
       # ==============================
        # Carregar a planilha MIX
        # ==============================
        df = pd.read_excel(caminho_arquivo_original)

        variantes_lista = []

        valores_da_coluna_opcionais = ['N_MODULE']

        for index, row in df.iterrows():
            variantes = []
            
            # Debug: verificar o que tem nas colunas
            if index < 3:  # Mostrar apenas as 3 primeiras linhas para debug
                logger.debug("linha %s:", index)
                if 'GÁS' in df.columns:
                    logger.debug("  GÁS: '%s'", row['GÁS'])
                if 'LNG' in df.columns:
                    logger.debug("  LNG: '%s'", row['LNG'])
                if 'ABS' in df.columns:
                    logger.debug("  ABS: '%s'", row['ABS'])
                if 'BEHIND CAB' in df.columns:
                    logger.debug("  BEHIND CAB: '%s'", row['BEHIND CAB'])
            
            # Adicionar valores não vazios
            if 'GÁS' in df.columns and pd.notna(row['GÁS']) and str(row['GÁS']).strip() != '':
                variantes.append(str(row['GÁS']).strip())
            
            if 'LNG' in df.columns and pd.notna(row['LNG']) and str(row['LNG']).strip() != '':
                variantes.append(str(row['LNG']).strip())
            
            if 'ABS' in df.columns and pd.notna(row['ABS']) and str(row['ABS']).strip() != '':
                variantes.append(str(row['ABS']).strip())
            
            if 'BEHIND CAB' in df.columns and pd.notna(row['BEHIND CAB']) and str(row['BEHIND CAB']).strip() != '':
                variantes.append(str(row['BEHIND CAB']).strip())
                
            if 'SCR HIDDEN' in df.columns and pd.notna(row['SCR HIDDEN']) and str(row['SCR HIDDEN']).strip() != '':
                variantes.append(str(row['SCR HIDDEN']).strip())

            # ADICIONAR APENAS VALORES ESPECÍFICOS DA COLUNA OPCIONAIS
            if 'OPCIONAIS' in df.columns and pd.notna(row['OPCIONAIS']) and str(row['OPCIONAIS']).strip() != '':
                valor_opcional = str(row['OPCIONAIS']).strip().upper()
                
                # Verificar se o valor está na lista dos desejados
                for valor_desejado in valores_da_coluna_opcionais:
                    if valor_desejado in valor_opcional:
                        variantes.append(valor_desejado)
                        break  # Para não duplicar se encontrar múltiplos

            # ENTRADA LINHA REAL: se tiver "TVL" em qualquer parte do valor,
            # grava so "TVL" em Variantes — no lugar do valor inteiro da
            # celula, igual o OPCIONAIS acima faz com LTDA/N_MODULE.
            if 'ENTRADA LINHA REAL' in df.columns and pd.notna(row['ENTRADA LINHA REAL']) and str(row['ENTRADA LINHA REAL']).strip() != '':
                valor_entrada_linha = str(row['ENTRADA LINHA REAL']).strip().upper()
                if 'TVL' in valor_entrada_linha:
                    variantes.append('TVL')

            # VERIFICAR SE O CARRO É LONGO (COMPRIMENTO > 7400)
            if 'COMPRIM' in df.columns and pd.notna(row['COMPRIM']):
                try:
                    comprimento = float(row['COMPRIM'])
                    if comprimento > 8100:
                        variantes.append("LONGO")
                except (ValueError, TypeError):
                    # Se não for possível converter para número, ignora
                    pass
            
           # if 'COMPRIM' in df.columns and pd.notna(row['COMPRIM']):
             #   try:
             #       comprimento = float(row['COMPRIM'])
             #       if comprimento == 6525:
             #           variantes.append("CALHA INVERTIDA")
             #   except (ValueError, TypeError):
             #       # Se não for possível converter para número, ignora
             #      pass

            resultado = ' - '.join(variantes) if variantes else ''
            variantes_lista.append(resultado)
            
            if index < 3:  # Debug
                logger.debug("  Variantes: '%s'", resultado)

        df['Variantes'] = variantes_lista

        # Cruzamento opcional com a MONA (mesmo cruzamento da Função 4): o
        # texto encontrado por POPID entra junto de Variantes, e não como
        # coluna separada — aqui Variantes já é o lugar de tudo isso.
        if caminho_mona:
            from mix_processor_4 import cruzar_com_mona
            df = cruzar_com_mona(df, caminho_mona, coluna_saida='_MONA_TEXTO',
                                 texto_lista_simples=texto_mona,
                                 itens_excluidos=ITENS_MONA_FA12_JA_NO_MIX)

            def _juntar_com_mona(linha):
                partes = []
                for valor in (linha.get('Variantes'), linha.get('_MONA_TEXTO')):
                    texto = str(valor).strip() if pd.notna(valor) else ''
                    if texto and texto.lower() != 'nan':
                        partes.append(texto)
                return ' | '.join(partes)

            df['Variantes'] = df.apply(_juntar_com_mona, axis=1)
            df = df.drop(columns=['_MONA_TEXTO'], errors='ignore')

        # LISTA BRANCA: mantem so o que o usuario pediu (ver a tela do
        # Filtro MIX). A lista negra antiga deixava passar qualquer coluna
        # nova que o sistema de origem criasse.
        desejadas = list(colunas_desejadas or COLUNAS_PADRAO)
        conhecidas = {_chave_coluna(c) for c in desejadas}
        manter = [c for c in df.columns if _chave_coluna(c) in conhecidas]
        descartadas = [str(c) for c in df.columns if c not in manter]
        if descartadas:
            logger.info("fa12: %s coluna(s) descartada(s): %s",
                       len(descartadas), ', '.join(descartadas))
        df = df[manter]

        # Remover linhas que contêm "Hide" em qualquer coluna
        mask = df.astype(str).apply(lambda row: row.str.contains('Hide', case=False, na=False).any(), axis=1)
        df = df[~mask]

    
        

        #Reescre a planilha
        book = load_workbook(caminho_arquivo_original)
        sheet = book.active
        
        # Limpar todas as células da planilha
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
            for cell in row:
                cell.value = None

        # Escrever cabeçalhos
        for col_idx, cabecalho in enumerate(df.columns, start=1):
            sheet.cell(row=1, column=col_idx, value=cabecalho)

        
        Fontes = {
            
            'POPID': 10,
            'CHASSI': 10,
            'MODELO': 10,
            'CHASSIS HEIGHT': 10,
            'SUSP TRASEIRA': 10,
            'MU': 9,
            'COR': 10,
            'GÁS':10,
          
        }
        Bordas = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin'))

        fonte_negrito = Font(bold=True)

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
        
        # Definição de cores e fontes das celulas das outras colunas
        cor_destaque_azul2 = PatternFill(start_color="163E64", end_color="163E64", fill_type="solid")
        cor_destaque2 = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        cor_destaque_azul3 = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")
        cor_destaque_azul_claro = PatternFill(start_color="CDEAF7", end_color="CDEAF7", fill_type="solid")
        cor_destaque_verde2 = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
        cor_destaque_cinza2 = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
        cor_destaque_vermelho2 = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        cor_destaque_preto2 = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
        cor_destaque_Verde_Escuro2 = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
        cor_destaque_laranja2 = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
        cor_destaque = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        cor_destaque_rouxo = PatternFill(start_color="993399", end_color="993399", fill_type="solid")

        fonte_branca = Font(color="FFFFFF", bold=True)
        fonte_vermelha = Font(color="FF0000", bold=True)
        fonte_azul = Font(color="0000FF", bold=True, italic=True)

        # Definição das regras de formatação
        # coluna_verificada: onde o sistema procura o valor (default: Variantes)
        # coluna_alvo:       onde a cor entra quando NÃO é linha inteira
        # linha_inteira:     True pinta toda a linha; False só a coluna_alvo
        # prioridade:        quando duas regras batem na MESMA célula, vence
        #                    a de prioridade MAIOR (aplicada por último).
        #                    Regras de linha inteira ficam com prioridade
        #                    baixa de propósito, pra uma regra de célula
        #                    (mais específica) nunca ficar por baixo.
        regras_formatacao = [
            {
                'valor_procurado': '8X2',
                'coluna_verificada': 'MODELO',
                'coluna_alvo': 'MODELO',
                'cor_fill': cor_destaque_preto2,
                'fonte': fonte_branca,
                'tipo_comparacao': 'contem',
                'linha_inteira': True,
                'prioridade': 10,
            },
            {
                'valor_procurado': '8X4',
                'coluna_verificada': 'MODELO',
                'coluna_alvo': 'MODELO',
                'cor_fill': cor_destaque_preto2,
                'fonte': fonte_branca,
                'tipo_comparacao': 'contem',
                'linha_inteira': True,
                'prioridade': 10,
            },
            {
                'valor_procurado': 'GÁS',
                'coluna_verificada': 'Variantes',
                'coluna_alvo': 'Variantes',
                'cor_fill': cor_destaque_azul3,
                'fonte': fonte_branca,
                'tipo_comparacao': 'contem',
                'linha_inteira': True,
                'prioridade': 10,
            },
            {
                'valor_procurado': '4X4',
                'coluna_verificada': 'MODELO',
                'coluna_alvo': 'MODELO',
                'cor_fill': cor_destaque_vermelho2,
                'fonte': fonte_branca,
                'tipo_comparacao': 'contem',
                'linha_inteira': False,
                'prioridade': 20,
            },
            {
                'valor_procurado': 'CALHA INVERTIDA',
                'coluna_verificada': 'Variantes',
                'coluna_alvo': 'Variantes',  # Coluna que será destacada
                'cor_fill': cor_destaque2,
                'fonte': None,
                'tipo_comparacao': 'contem',  # 'exato', 'contem', 'inicia_com'
                'linha_inteira': False,
                'prioridade': 20,
            },
            {
                'valor_procurado': 'LTDA',
                'coluna_verificada': 'Variantes',
                'coluna_alvo': 'Variantes',  # Coluna que será destacada
                'cor_fill': cor_destaque_laranja2,
                'fonte': None,
                'tipo_comparacao': 'contem',  # 'exato', 'contem', 'inicia_com'
                'linha_inteira': False,
                'prioridade': 20,
            },
            {
                'valor_procurado': 'LONGO',
                'coluna_verificada': 'Variantes',
                'coluna_alvo': 'Variantes',  # Coluna que será destacada
                'cor_fill': cor_destaque_cinza2,
                'fonte': fonte_branca,
                'tipo_comparacao': 'contem',  # 'exato', 'contem', 'inicia_com'
                'linha_inteira': True,
                'prioridade': 20,
            },
            {
                'valor_procurado': 'K1',
                'coluna_verificada': 'Variantes',
                'coluna_alvo': 'Variantes',  # Coluna que será destacada
                'cor_fill': cor_destaque_vermelho2,
                'fonte': None,
                'tipo_comparacao': 'contem',  # 'exato', 'contem', 'inicia_com'
                'linha_inteira': False,
                'prioridade': 20,
            },
            {
                'valor_procurado': 'DIR. ELETRICA',
                'coluna_verificada': 'Variantes',
                'coluna_alvo': 'Variantes',  # Coluna que será destacada
                'cor_fill': cor_destaque_Verde_Escuro2,
                'fonte': None,
                'tipo_comparacao': 'contem',  # 'exato', 'contem', 'inicia_com'
                'linha_inteira': False,
                'prioridade': 20,
            },
            {
                'valor_procurado': 'TVL',
                'coluna_verificada': 'Variantes',
                'coluna_alvo': 'Variantes',  # Coluna que será destacada
                'cor_fill': None,
                'fonte': fonte_branca,
                'tipo_comparacao': 'contem',  # 'exato', 'contem', 'inicia_com'
                'linha_inteira': True,
                'prioridade': 10,
            },




            # Adicione quantas regras quiser aqui
        ]

        # Aplica na ordem da prioridade: menor primeiro, maior por último —
        # assim quem tem prioridade maior "vence" quando duas regras
        # pintariam a mesma célula.
        regras_formatacao = sorted(
            regras_formatacao, key=lambda r: r.get('prioridade', 0))

        # Aplicar todas as regras de formatação
        for regra in regras_formatacao:
            coluna_verificada = regra.get('coluna_verificada', 'Variantes')
            coluna_alvo = regra['coluna_alvo']

            if coluna_verificada not in df.columns or coluna_alvo not in df.columns:
                continue

            col_idx_verificada = df.columns.get_loc(coluna_verificada) + 1
            col_idx_alvo = df.columns.get_loc(coluna_alvo) + 1

            logger.debug("aplicando regra: '%s' -> %s", regra['valor_procurado'], coluna_alvo)

            contador = 0
            for row_idx in range(2, sheet.max_row + 1):
                cell_verificada = sheet.cell(row=row_idx, column=col_idx_verificada)

                if cell_verificada.value:
                    valor_str = str(cell_verificada.value).strip().upper()
                    valor_procurado = str(regra['valor_procurado']).upper()

                    # Aplicar diferentes tipos de comparação
                    corresponde = False
                    if regra['tipo_comparacao'] == 'exato':
                        corresponde = (valor_str == valor_procurado)
                    elif regra['tipo_comparacao'] == 'contem':
                        corresponde = (valor_procurado in valor_str)
                    elif regra['tipo_comparacao'] == 'inicia_com':
                        corresponde = valor_str.startswith(valor_procurado)

                    if corresponde:
                        contador += 1

                        if regra.get('linha_inteira'):
                            # Destacar a linha inteira (todas as colunas)
                            for col_name in df.columns:
                                col_idx = df.columns.get_loc(col_name) + 1
                                cell = sheet.cell(row=row_idx, column=col_idx)
                                if regra['cor_fill']:
                                    cell.fill = regra['cor_fill']
                                if regra['fonte']:
                                    cell.font = regra['fonte']
                        else:
                            # Destacar só a coluna alvo
                            cell_alvo = sheet.cell(row=row_idx, column=col_idx_alvo)
                            if regra['cor_fill']:
                                cell_alvo.fill = regra['cor_fill']
                            if regra['fonte']:
                                cell_alvo.font = regra['fonte']

                        # Opcional: também destacar a célula Variantes de origem
                        if regra.get('destacar_origem', False) and coluna_verificada != coluna_alvo:
                            cell_verificada.fill = regra['cor_fill']
                            if regra['fonte']:
                                cell_verificada.font = regra['fonte']

            logger.debug("  → %s células destacadas na coluna %s", contador, coluna_alvo)

    
        # ==============================================================
        # DEFINIÇÃO DAS LARGURAS DAS COLUNAS
        # ==============================================================
        
        # Dicionário com as larguras desejadas para cada coluna
        larguras_colunas = {
            
            'POPID': 7.20,
            'CHASSI': 7.71,
            'MODELO': 22.00,
            'SUSP TRASEIRA': 3.14,
            'MU': 5.00,
            'Variantes': 60.00
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

        # Destacar comprimentos maiores que 7400 com cores diferentes por modelo
        if "COMPRIM" in df.columns and "MODELO" in df.columns:
            
            col_idx_comprim = df.columns.get_loc("COMPRIM") + 1
            col_idx_modelo = df.columns.get_loc("MODELO") + 1
            
            for row_idx in range(2, sheet.max_row + 1):
                cell_comprim = sheet.cell(row=row_idx, column=col_idx_comprim)
                cell_modelo = sheet.cell(row=row_idx, column=col_idx_modelo)
                
                if (isinstance(cell_comprim.value, (int, float)) and cell_comprim.value > 8100):
                    
                    modelo_str = str(cell_modelo.value).upper() if cell_modelo.value else ""
                    
                    # Aplicar formatação em toda a linha
                    for col_idx in range(1, len(df.columns) + 1):
                        cell = sheet.cell(row=row_idx, column=col_idx)
                        
                        if "6X2" in modelo_str:
                            # Formatação para 6X2: linha cinza, COMPRIM com fonte vermelha
                            cell.fill = cor_destaque_cinza2
                            cell.font = fonte_vermelha if col_idx == col_idx_comprim else fonte_branca
                        elif "6X4" in modelo_str:
                            cell.fill = cor_destaque_cinza2
                            cell.font = fonte_vermelha if col_idx == col_idx_comprim else fonte_branca
                        elif "4X2" in modelo_str:
                            cell.fill = cor_destaque_cinza2
                            cell.font = fonte_vermelha if col_idx == col_idx_comprim else fonte_branca

                            
                        else:
                            # Formatação para outros modelos: linha preta, todas as células com fonte branca
                            cell.fill = cor_destaque_preto2
                            cell.font = fonte_branca


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
        book.close()
        return True, "Arquivo formatado e salvo com sucesso!"
    
    except Exception as e:
        return False, f"Erro ao formatar o arquivo: {e}"

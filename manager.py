import pandas as pd
import os

# Marcas proibidas por Marketplace
marcas_proibidas_dafiti = [
    'Abercrombie', 'Adidas', 'Aeropostale', 'Colcci', 'Crocs', 'Diesel', 'Disky', 'Guess', 'Hollister', 'Individual', 'Lacoste', 'Levis', 'New Era', 'Nike', 'Osklen', 'Polo Ralph Lauren', 'Replay', 'Victory Eagle'
]
marcas_proibidas_meli = [
    'Adidas', 'Crocs', 'Nike', 'Lacoste'
]
marcas_proibidas_zattini = [
    'Aeropostale', 'Lacoste', 'Levis', 'Nike', 'Osklen', 'Polo Ralph Lauren', 'Reserva'
]

# Diretórios
diretorio_bling = "C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\planilhas\\plan_bling.xlsx"
diretorio_bagy = "C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\planilhas\\plan_bagy.xlsx"
diretorio_dafiti = "C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\planilhas\\plan_dafiti.xlsx"
diretorio_meli = "C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\planilhas\\plan_meli.xlsx"
direftorio_zattini =  "C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\planilhas\\plan_zattini.xlsx"

# Iniciando o comparador
print('Você deseja criar uma planilha para ver os produtos sem cadastro na Bagy ou nos Marketplaces?\n1 - Bagy\n2 - Marketplaces')
canal_escolhido = int(input('Digite o número correspondente a sua escolha:'))

# Função para criar a tabela e retornar o nome do arquivo
if canal_escolhido == 1:
    print('Perfeito! Vamos gerar uma planilha para ver os produtos sem cadastro e os que estão com estoque divergente.')
    nome_arquivo = 'sem_cadastro_bagy.xlsx'
    # Carregar as tabelas
    tabela_bling = pd.read_excel(diretorio_bling)
    tabela_bagy = pd.read_excel(diretorio_bagy)

    # Crie uma cópia da tabela do site para adicionar o status de cadastro
    tabela_final = tabela_bling.copy()

    # Adicionando a coluna 'Status de cadastro' com o valor padrão 'Sem Cadastro'
    tabela_final['Status de cadastro'] = 'Sem Cadastro'

    # Atualize o status para 'cadastrado' onde os códigos estão presentes no marketplace
    tabela_final.loc[tabela_final['Código'].isin(tabela_bagy['Código']), 'Status de cadastro'] = 'Cadastrado'

    # Adicione a coluna 'Divergência de Estoque'

    for index, row in tabela_bagy.iterrows():
        codigo = row['Código']
        estoque_bagy = row['Estoque']

        # Encontrar a linha correspondente em tabela_final com o mesmo código
        linha_correspondente = tabela_final[tabela_final['Código'] == codigo]

        if not linha_correspondente.empty:
            # Calcular a divergência de estoque
            estoque_final = linha_correspondente.iloc[0]['Estoque']
            divergencia = estoque_final - estoque_bagy

            # Atualizar a coluna 'Divergência de Estoque' em tabela_final
            tabela_final.loc[tabela_final['Código'] == codigo, 'Divergência de Estoque'] = divergencia

    # Salvar a nova tabela em um arquivo Excel
    # Diretório de saída pré-definido
    output_directory = 'C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\criadas\\'

    # Construa o caminho completo do arquivo de saída
    output_path = os.path.join(output_directory, nome_arquivo)

    # Salvar a nova tabela em um arquivo Excel no diretório escolhido
    tabela_final.to_excel(output_path, index=False)

    # Mensagem personalizada de impressão
    print(f"Uma nova tabela com o status de cadastro dos produtos foi criada: '{output_path}'")
elif canal_escolhido == 2:
    print('Para qual marketplace você deseja gerar a planilha?\n1 - Dafiti\n2 - Mercado Livre\n3 - Zattini - Netshoes')
    mktplc_selecionado = int(input('Digite o número que corresponde a sua escolha: '))
    if mktplc_selecionado == 1:
        print('Perfeito, vamos gerar a tabela de produtos para cadastrar da Dafiti...')
        marcas_proibidas = marcas_proibidas_dafiti
        nome_arquivo = 'produtos_sem_cadastro_dafiti.xlsx'
        tabela_bagy = pd.read_excel(diretorio_bagy)
        tabela_dafiti = pd.read_excel(diretorio_dafiti)
        nome_arquivo = 'Produtos_sem_cadastro_dafiti.xlsx'
        # Filtragem da tabela Bagy
        tabela_bagy_filtrada = tabela_bagy[~tabela_bagy['Marca'].isin(marcas_proibidas)]
        # Criação de uma cópia da tabela Bagy para adicionar o status de cadastro
        tabela_final = tabela_bagy_filtrada.copy()
        # Adição da coluna 'Status de cadastro' com o valor padrão 'Sem Cadastro'
        tabela_final['Status de cadastro'] = 'Sem Cadastro'
        # Atualização do status para 'cadastrado' onde os códigos estão presentes no marketplace
        tabela_final.loc[tabela_final['Código'].isin(tabela_dafiti['Código']), 'Status de cadastro'] = 'Cadastrado'
        # Merge com a tabela do marketplace para preço e estoque
        tabela_final = tabela_final.merge(tabela_dafiti[['Código', "Preço de", 'Preço por']], on='Código', how='left', suffixes=("", '_DFT'))
        # Cálculo da rentabilidade
        tabela_final['Preço Final'] = tabela_final['Preço por'].apply(lambda x: round((x/10), 0) * 10 - 0.01)
        # Remoção de colunas desnecessárias
        tabela_final.drop(columns=['Código', 'Marca'], inplace=True)
        # Agrupamento por 'SKU Pai', soma do estoque e manutenção das outras informações
        agg_dict = {'Estoque': 'sum',  # Soma do estoque
                    'Nome': 'first',  # Manter a primeira ocorrência do nome do produto
                    'Preço de': 'first',  # Manter a primeira ocorrência do preço
                    'Preço por': 'first',  # Manter a primeira ocorrência do preço
                    'Preço Final': 'first',  # Manter a primeira ocorrência do preço final
                    'Preço de_MELI': 'first',  # Manter a primeira ocorrência do preço do marketplace
                    'Preço por_MELI': 'first',  # Manter a primeira ocorrência do preço do marketplace
                    'Status de cadastro': 'first',  # Manter a primeira ocorrência do status
                    'SKU Pai': 'first'}  # Manter a primeira ocorrência da rentabilidade
        tabela_final = tabela_final.groupby('SKU Pai').agg(agg_dict)
        # Reorganização das colunas (opcional, se a ordem desejada for diferente)
        colunas_ordenadas = ['SKU Pai', 'Nome', 'Estoque', 'Preço de', 'Preço por', 'Preço Final', 'Preço de_MELI', 'Preço por_MELI', 'Status de cadastro']
        tabela_final = tabela_final.reindex(columns=colunas_ordenadas)
        # Salvamento em Excel
        output_directory = 'C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\criadas\\'
        output_path = os.path.join(output_directory, nome_arquivo)
        tabela_final.to_excel(output_path, index=False)
        print(f"Uma nova tabela com o status de cadastro dos produtos foi criada: '{nome_arquivo}'")
    elif mktplc_selecionado == 2:
        print('Perfeito, vamos gerar a tabela de produtos para cadastrar do Mercado Livre...')
        marcas_proibidas = marcas_proibidas_meli
        nome_arquivo = 'produtos_sem_cadastro_meli.xlsx'
        tabela_bagy = pd.read_excel(diretorio_bagy)
        tabela_meli = pd.read_excel(diretorio_meli)
        nome_arquivo = 'Produtos_sem_cadastro_meli.xlsx'
        # Filtragem da tabela Bagy
        tabela_bagy_filtrada = tabela_bagy[~tabela_bagy['Marca'].isin(marcas_proibidas)]
        # Criação de uma cópia da tabela Bagy para adicionar o status de cadastro
        tabela_final = tabela_bagy_filtrada.copy()
        # Adição da coluna 'Status de cadastro' com o valor padrão 'Sem Cadastro'
        tabela_final['Status de cadastro'] = 'Sem Cadastro'
        # Atualização do status para 'cadastrado' onde os códigos estão presentes no marketplace
        tabela_final.loc[tabela_final['Código'].isin(tabela_meli['Código']), 'Status de cadastro'] = 'Cadastrado'
        # Merge com a tabela do marketplace para preço e estoque
        tabela_final = tabela_final.merge(tabela_meli[['Código', "Preço de", 'Preço por']], on='Código', how='left', suffixes=("", '_MELI'))
        # Cálculo da rentabilidade
        tabela_final['Preço Final'] = tabela_final['Preço por'].apply(lambda x: round((x/10), 0) * 10 - 0.01)
        rentabilidade = (tabela_final['Preço Final'] - tabela_final['Preço por_MELI']) / tabela_final['Preço por_MELI'] * 100
        tabela_final['Rentabilidade'] = rentabilidade.round(2)
        # Remoção de colunas desnecessárias
        tabela_final.drop(columns=['Código', 'Marca'], inplace=True)
        # Agrupamento por 'SKU Pai', soma do estoque e manutenção das outras informações
        agg_dict = {'Estoque': 'sum',  # Soma do estoque
                    'Nome': 'first',  # Manter a primeira ocorrência do nome do produto
                    'Preço de': 'first',  # Manter a primeira ocorrência do preço
                    'Preço por': 'first',  # Manter a primeira ocorrência do preço
                    'Preço Final': 'first',  # Manter a primeira ocorrência do preço final
                    'Preço de_MELI': 'first',  # Manter a primeira ocorrência do preço do marketplace
                    'Preço por_MELI': 'first',  # Manter a primeira ocorrência do preço do marketplace
                    'Status de cadastro': 'first',  # Manter a primeira ocorrência do status
                    'SKU Pai': 'first'}  # Manter a primeira ocorrência da rentabilidade
        tabela_final = tabela_final.groupby('SKU Pai').agg(agg_dict)
        # Reorganização das colunas (opcional, se a ordem desejada for diferente)
        colunas_ordenadas = ['SKU Pai', 'Nome', 'Estoque', 'Preço de', 'Preço por', 'Preço Final', 'Preço de_MELI', 'Preço por_MELI', 'Status de cadastro']
        tabela_final = tabela_final.reindex(columns=colunas_ordenadas)
        # Salvamento em Excel
        output_directory = 'C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\criadas\\'
        output_path = os.path.join(output_directory, nome_arquivo)
        tabela_final.to_excel(output_path, index=False)
        print(f"Uma nova tabela com o status de cadastro dos produtos foi criada: '{nome_arquivo}'")
    elif mktplc_selecionado == 3:
        print('Perfeito, vamos gerar a tabela de produtos para cadastrar da Zattini...')
        marcas_proibidas = marcas_proibidas_zattini
        nome_arquivo = 'produtos_sem_cadastro_zattini.xlsx'
        tabela_bagy = pd.read_excel(diretorio_bagy)
        tabela_zattini = pd.read_excel(direftorio_zattini)
        nome_arquivo = 'Produtos_sem_cadastro_zattini.xlsx'
        # Filtragem da tabela Bagy
        tabela_bagy_filtrada = tabela_bagy[~tabela_bagy['Marca'].isin(marcas_proibidas)]
        # Criação de uma cópia da tabela Bagy para adicionar o status de cadastro
        tabela_final = tabela_bagy_filtrada.copy()
        # Adição da coluna 'Status de cadastro' com o valor padrão 'Sem Cadastro'
        tabela_final['Status de cadastro'] = 'Sem Cadastro'
        # Atualização do status para 'cadastrado' onde os códigos estão presentes no marketplace
        tabela_final.loc[tabela_final['Código'].isin(tabela_zattini['Código']), 'Status de cadastro'] = 'Cadastrado'
        # Merge com a tabela do marketplace para preço e estoque
        tabela_final = tabela_final.merge(tabela_zattini[['Código', "Preço de", 'Preço por']], on='Código', how='left', suffixes=("", '_'))
        # Cálculo da rentabilidade
        tabela_final['Preço Final'] = tabela_final['Preço por'].apply(lambda x: round((x/10), 0) * 10 - 0.01)
        # Remoção de colunas desnecessárias
        tabela_final.drop(columns=['Código', 'Marca'], inplace=True)
        # Agrupamento por 'SKU Pai', soma do estoque e manutenção das outras informações
        agg_dict = {'Estoque': 'sum',  # Soma do estoque
                    'Nome': 'first',  # Manter a primeira ocorrência do nome do produto
                    'Preço de': 'first',  # Manter a primeira ocorrência do preço
                    'Preço por': 'first',  # Manter a primeira ocorrência do preço
                    'Preço Final': 'first',  # Manter a primeira ocorrência do preço final
                    'Preço de_MELI': 'first',  # Manter a primeira ocorrência do preço do marketplace
                    'Preço por_MELI': 'first',  # Manter a primeira ocorrência do preço do marketplace
                    'Status de cadastro': 'first',  # Manter a primeira ocorrência do status
                    'SKU Pai': 'first'}  # Manter a primeira ocorrência da rentabilidade
        tabela_final = tabela_final.groupby('SKU Pai').agg(agg_dict)
        # Reorganização das colunas (opcional, se a ordem desejada for diferente)
        colunas_ordenadas = ['SKU Pai', 'Nome', 'Estoque', 'Preço de', 'Preço por', 'Preço Final', 'Preço de_MELI', 'Preço por_MELI', 'Status de cadastro']
        tabela_final = tabela_final.reindex(columns=colunas_ordenadas)
        # Salvamento em Excel
        output_directory = 'C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\criadas\\'
        output_path = os.path.join(output_directory, nome_arquivo)
        tabela_final.to_excel(output_path, index=False)
        print(f"Uma nova tabela com o status de cadastro dos produtos foi criada: '{nome_arquivo}'")
    else:
        print('Você deve selecionar uma opção válida!')
else:
    print('Você deve selecionar uma opção válida!')

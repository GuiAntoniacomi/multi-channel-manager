import pandas as pd
import re
import os

def base_bagy():
    """
    Processa um arquivo JSON da Bagy e retorna um DataFrame pandas formatado.

    Este método solicita ao usuário que insira o caminho para um arquivo JSON exportado da plataforma Bagy. 
    Em seguida, ele lê o arquivo JSON, seleciona e renomeia colunas específicas, reorganiza as colunas e 
    remove quaisquer linhas que não possuam um código de variação ('Código'). Finalmente, converte a coluna 
    'Código' para o tipo inteiro e retorna o DataFrame resultante.

    Retorna:
        pandas.DataFrame: DataFrame contendo os dados processados do arquivo JSON da Bagy com as colunas:
            - 'SKU Pai': ID externo do produto
            - 'Código': SKU da variação do produto, convertido para inteiro
            - 'Marca': Nome da marca do produto
            - 'Nome': Nome do produto
            - 'Estoque': Quantidade em estoque
            - 'Preço De': Preço original do produto (antes de desconto)
            - 'Preço Por': Preço atual do produto (após desconto)
    """
    arquivo_json = input('Digite o caminho do seu arquivo da Bagy: ').strip('\'"')
    bagy_json = pd.read_json(arquivo_json)
    bagy_json = bagy_json[['Brands → Name', 'Variations → Sku', 'Price', 'Price Compare', 'Name', 'Stocks → Balance', 'External ID']]
    bagy_json = bagy_json.rename(columns={'Brands → Name': 'Marca', 'Variations → Sku': 'Código', 'Price Compare': 'Preço De', 'Price': 'Preço Por', 'Name': 'Nome', 'Stocks → Balance': 'Estoque', 'External ID': 'SKU Pai'})
    bagy_json = bagy_json.reindex(columns=["SKU Pai", "Código", "Marca", "Nome", "Estoque", "Preço De", "Preço Por"])
    bagy_json = bagy_json.dropna(subset=['Código'])
    bagy_json['Código'] = bagy_json['Código'].astype(int)
    return bagy_json

def base_dafiti():
    """
    Processa um arquivo CSV da Dafiti e retorna um DataFrame pandas formatado.

    Este método solicita ao usuário que insira o caminho para um arquivo CSV exportado da plataforma Dafiti. 
    Em seguida, ele lê o arquivo CSV, remove colunas desnecessárias, renomeia colunas específicas, reorganiza 
    as colunas e retorna o DataFrame resultante.

    Retorna:
        pandas.DataFrame: DataFrame contendo os dados processados do arquivo CSV da Dafiti com as colunas:
            - 'Código': SKU do vendedor
            - 'Nome': Nome do produto
            - 'Estoque': Quantidade em estoque
            - 'Preço De': Preço original do produto (antes de desconto)
            - 'Preço Por': Preço atual do produto (após desconto)
    """
    arquivo_dafiti = input('Digite o caminho do seu arquivo da Dafiti: ').strip('\'"')
    dafiti_csv = pd.read_csv(arquivo_dafiti, delimiter=';')
    dafiti_df = pd.DataFrame(dafiti_csv)
    colunas_para_remover = ['CatalogProductId', 'Dafiti SKU', 'Ncm', 'Brand', 'Color',
                            'ColorFamily', 'PrimaryCategory', 'Categories', 'Gender', 'Warranty',
                            'Origin', 'Origincountry', 'ObservacoesImportantes', 'RegistroDoProduto',
                            'Description', 'BrowseNodes', 'SaleEndDate', 'SaleStartDate', 'Variation',
                            'ProductId', 'Occasion', 'Character', 'ModalidadeEsporte', 'Waist',
                            'Length', 'Specifications', 'Print', 'Style', 'Kit', 'Sleeve',
                            'ModelClothes', 'Composition', 'Washing', 'WashingType', 'ClothesMaterial',
                            'BoxHeight', 'BoxLength', 'BoxWidth', 'Weight', 'CreatedAt', 'UpdatedAt', 'Status', 'ParentSku']
    
    colunas_existentes = [coluna for coluna in colunas_para_remover if coluna in dafiti_df.columns]
    dafiti_df = dafiti_df.drop(columns=colunas_existentes)
    dafiti_df = dafiti_df.rename(columns={'SellerSku': 'Código', 'Name': 'Nome', 'Quantity': 'Estoque', 'Price': 'Preço De', 'SalePrice': 'Preço Por'})
    dafiti_df = dafiti_df.reindex(columns=['Código', 'Nome', 'Estoque', 'Preço De', 'Preço Por'])
    dafiti_df['Código'] = dafiti_df['Código'].astype(int)
    dafiti_df = dafiti_df.dropna(subset=['Código'])
    return dafiti_df

def base_meli():
    arquivo_meli = input('Digite o caminho do arquivo do Mercado Livre: ').strip('\'"')
    df_meli = pd.read_excel(arquivo_meli, sheet_name='Anúncios')
    df_meli = df_meli.drop([0, 1])
    df_meli.reset_index(drop=True, inplace=True)
    df_meli = df_meli[['ITEM_ID', 'SKU', 'PRICE']]
    
    # Preenchendo valores NaN com 0 antes de converter os tipos
    df_meli['SKU'] = df_meli['SKU'].fillna(0)
    df_meli['PRICE'] = df_meli['PRICE'].fillna(0)

    # Convertendo os tipos de dados
    df_meli['SKU'] = df_meli['SKU'].astype(int)
    df_meli['PRICE'] = df_meli['PRICE'].astype(float)
    
    #Agrupando os dados
    df_meli = df_meli.groupby('ITEM_ID').agg({
        'SKU': 'last',
        'PRICE': 'first'
    }).reset_index()
    
    # Ajustando o dataframe final
    df_meli = df_meli.drop(columns=['ITEM_ID'])
    df_meli = df_meli.rename(columns={'SKU': 'Código', 'PRICE': 'Preço Por Meli' })
    
    return df_meli

def base_zattini():
    arquivo_zattini = input('Digite o caminho do seu arquivo da zattini: ').strip('\'"')
    zattini_excel = pd.read_excel(arquivo_zattini)
    zattini_df = pd.DataFrame(zattini_excel)
    zattini_df = zattini_df.rename(columns={'Sku Seller': 'Código'})
    
    # Função para extrair números de uma string
    def extract_numbers(text):
        match = re.search(r'\d+', text)
        if match:
            return int(match.group())
        return None
    
    # Aplicar a função personalizada à coluna 'Código'
    zattini_df['Código'] = zattini_df['Código'].apply(extract_numbers)
    
    return zattini_df

def exportar_para_dafiti(bagy, marketplace):
    """
    Filtra e processa dados de produtos para exportação para a Dafiti.

    Este método recebe dois DataFrames: um contendo dados da Bagy e outro contendo uma tabela do marketplace. 
    Ele filtra produtos com marcas proibidas pela Dafiti, marca o status de cadastro dos produtos, mescla dados 
    com a tabela do marketplace, renomeia e organiza colunas, agrupa os dados por 'SKU Pai' e realiza uma série 
    de operações de agregação e ordenação. O resultado é um DataFrame formatado e pronto para exportação.

    Args:
        bagy (pandas.DataFrame): DataFrame contendo os dados da Bagy.
        tabela_marketplace (pandas.DataFrame): DataFrame contendo a tabela do marketplace da Dafiti.

    Retorna:
        pandas.DataFrame: DataFrame contendo os dados processados e prontos para exportação para a Dafiti, com as colunas:
            - 'SKU Pai': ID externo do produto
            - 'Marca': Nome da marca do produto
            - 'Nome': Nome do produto
            - 'Estoque': Quantidade total em estoque
            - 'Preço De Bagy': Preço original do produto na Bagy
            - 'Preço Por Bagy': Preço atual do produto na Bagy
            - 'Status Cadastro': Status de cadastro do produto na Dafiti (Cadastrado ou Sem Cadastro)
            - 'Preço De Dft': Preço original do produto na Dafiti
            - 'Preço Por Dft': Preço atual do produto na Dafiti
    """
    marcas_proibidas_dafiti = ['Abercrombie', 'Adidas', 'Aeropostale', 'Colcci', 'Crocs', 'Diesel', 'Disky', 'Guess', 'Hollister', 'Individual', 'Lacoste', 'Levis', 'New Era', 'Nike', 'Osklen', 'Polo Ralph Lauren', 'Replay', 'Victory Eagle']
    df_exportar_dafiti = bagy[~bagy['Marca'].isin(marcas_proibidas_dafiti)]
    df_exportar_dafiti['Status Cadastro'] = 'Sem Cadastro'
    df_exportar_dafiti.loc[df_exportar_dafiti['Código'].isin(marketplace['Código']), 'Status Cadastro'] = 'Cadastrado'
    df_exportar_dafiti = df_exportar_dafiti.merge(marketplace, on='Código', how='left')
    df_exportar_dafiti = df_exportar_dafiti.drop(['Nome_y', 'Estoque_y', 'Código'], axis=1)
    df_exportar_dafiti = df_exportar_dafiti.rename(columns={'Nome_x': 'Nome', 'Estoque_x': 'Estoque','Preço De_y': 'Preço De Dft', 'Preço Por_y': 'Preço Por Dft', 'Preço De_x': 'Preço De Bagy', 'Preço Por_x': 'Preço Por Bagy' })
    df_dafiti_final = df_exportar_dafiti.groupby('SKU Pai').agg({
        'Marca': 'first',
        'Nome': 'first',
        'Estoque': 'sum',
        'Preço De Bagy': 'first',
        'Preço Por Bagy': 'first',
        'Status Cadastro': 'first',
        'Preço De Dft': 'first',
        'Preço Por Dft': 'first'
    }).reset_index()
    df_dafiti_final = df_dafiti_final.sort_values(by='Estoque', ascending=False)
    df_dafiti_final = df_dafiti_final.dropna(subset=['Marca'])
    return df_dafiti_final

def exportar_meli(bagy, marketplace):
    marcas_proibidas_meli = ['Adidas', 'Crocs', 'Nike', 'Lacoste']
    df_exportar_meli = bagy[~bagy['Marca'].isin(marcas_proibidas_meli)]
    df_exportar_meli.loc[:, 'Status Cadastro'] = 'Cadastrado'
    df_exportar_meli = df_exportar_meli.merge(marketplace, on='Código', how='left')
    df_exportar_meli = df_exportar_meli.drop(['Marca'], axis=1)
    df_exportar_meli = df_exportar_meli.groupby('SKU Pai').agg({
        'Nome': 'first',
        'Estoque': 'sum',
        'Preço De': 'first',
        'Preço Por': 'first',
        'Status Cadastro': 'first',
        'Preço Por Meli': 'max'
    }).reset_index()
    df_exportar_meli = df_exportar_meli.sort_values(by='Estoque', ascending=False)
    df_exportar_meli.loc[df_exportar_meli['Preço Por Meli'].isna(), 'Status Cadastro'] = 'Sem Cadastro'
    df_exportar_meli.drop([6362, 5007], inplace=True)
   
    return df_exportar_meli

def exportar_zattini(bagy, marketplace):
    """
    Exporta dados de produtos para a Zattini, filtrando marcas proibidas e ajustando status de cadastro.

    Esta função realiza os seguintes passos:
    1. Remove produtos de marcas proibidas específicas.
    2. Define o status de cadastro para 'Sem Cadastro'.
    3. Atualiza o status de cadastro para 'Cadastrado' para produtos cujos códigos estão presentes no dataframe do marketplace.
    4. Mescla dados adicionais do marketplace no dataframe principal com base no código do produto.
    5. Remove as colunas 'Código' e 'Marca'.
    6. Renomeia as colunas de preços para distinguir entre preços do Bagy e da Zattini.
    7. Agrega os dados por 'SKU Pai', somando os estoques e mantendo o primeiro valor encontrado para as demais colunas.
    8. Ordena os dados pelo estoque em ordem decrescente.
    9. Remove linhas específicas com os índices 4604 e 3555, que correspondem a sacola plástica e embalagem de presente.
    
    Parâmetros:
    bagy (pd.DataFrame): DataFrame contendo os dados dos produtos do Bagy.
    marketplace (pd.DataFrame): DataFrame contendo os dados dos produtos do marketplace da Zattini.

    Retorna:
    pd.DataFrame: DataFrame processado e pronto para exportação para a Zattini, com as marcas proibidas removidas, status de cadastro atualizado e dados agregados por 'SKU Pai'.
    """
    marcas_proibidas_zattini = ['Aeropostale', 'Lacoste', 'Levis', 'Nike', 'Osklen', 'Polo Ralph Lauren', 'Reserva']
    df_exportar_zattini = bagy[~bagy['Marca'].isin(marcas_proibidas_zattini)]
    df_exportar_zattini.loc[:, 'Status Cadastro'] = 'Sem Cadastro'
    df_exportar_zattini.loc[df_exportar_zattini['Código'].isin(marketplace['Código']), 'Status Cadastro'] = 'Cadastrado'
    df_exportar_zattini = df_exportar_zattini.merge(marketplace, on='Código', how='left')
    df_exportar_zattini = df_exportar_zattini.drop(columns=['Código', 'Marca'])
    df_exportar_zattini = df_exportar_zattini.rename(columns= {'Preço De_x': 'Preço De Bagy', 'Preço Por_x': 'Preço Por Bagy', 'Preço De_y': 'Preço De Ztn', 'Preço Por_y': 'Preço Por Ztn' })
    df_exportar_zattini = df_exportar_zattini.groupby('SKU Pai').agg({
        'Nome': 'first',
        'Estoque': 'sum',
        'Preço De Bagy': 'first',
        'Preço Por Bagy': 'first',
        'Status Cadastro': 'first',
        'Preço De Ztn': 'first',
        'Preço Por Ztn': 'first'
    }).reset_index()
    df_exportar_zattini = df_exportar_zattini.sort_values(by='Estoque', ascending=False)
    df_exportar_zattini.drop([4604, 3555], inplace=True) #removendo sacola plastica e embalagem de presente
    return df_exportar_zattini

def salvar_arquivo(tabela, nome_arquivo):
    local = input(f'Onde deseja salvar o arquivo {nome_arquivo}? ').strip('\'"')
    local = os.path.abspath(local)  # Converte para um caminho absoluto, se necessário
    if not os.path.exists(local):
        os.makedirs(local)  # Cria o diretório, se não existir
    caminho_completo = os.path.join(local, nome_arquivo)
    tabela.to_excel(caminho_completo, index=False)
    print(f"Arquivo salvo em {caminho_completo}")

def salvar_todas_as_tabelas(bagy, opcoes):
    for opcao, (nome, func_base, func_exportar, nome_arquivo) in opcoes.items():
        base = func_base()
        tabela = func_exportar(bagy, base)
        salvar_arquivo(tabela, nome_arquivo)

def main():
    bagy = base_bagy()
    opcoes = {
        '1': ('Dafiti', base_dafiti, exportar_para_dafiti, 'tabela_dafiti.xlsx'),
        '2': ('Mercado Livre', base_meli, exportar_meli, 'tabela_meli.xlsx'),
        '3': ('Zattini/Netshoes', base_zattini, exportar_zattini, 'tabela_zattini.xlsx'),
        '4': ('Todos', None, None, None)  # Placeholder para indicar a opção '4'
    }
    
    while True:
        opcao = input('De qual marketplace deseja gerar tabela de exportação?\n1. Dafiti\n2. Mercado Livre\n3. Zattini/Netshoes\n4. Todos\n')
        opcao = opcao.strip()
        
        if opcao in opcoes and opcao != '4':
            nome, func_base, func_exportar, nome_arquivo = opcoes[opcao]
            base = func_base()
            tabela = func_exportar(bagy, base)
            salvar_arquivo(tabela, nome_arquivo)
            break
        elif opcao == '4':
            salvar_todas_as_tabelas(bagy, {k: v for k, v in opcoes.items() if k != '4'})
            break
        else:
            print("Opção inválida. Por favor, escolha uma das opções listadas.")

if __name__ == '__main__':
    main()
import pandas as pd
import re

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

def exportar_para_dafiti(bagy, tabela_marketplace):
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
    df_exportar_dafiti.loc[df_exportar_dafiti['Código'].isin(tabela_marketplace['Código']), 'Status Cadastro'] = 'Cadastrado'
    df_exportar_dafiti = df_exportar_dafiti.merge(tabela_marketplace, on='Código', how='left')
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


bagy = base_bagy()
dafiti = base_dafiti()
zattini = base_zattini()
meli = pd.read_excel(r'src\planilhas\plan_meli.xlsx')



print(exportar_para_dafiti(bagy, dafiti))
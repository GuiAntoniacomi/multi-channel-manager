import pandas as pd
import os

# Diretórios
diretorio_bling = "C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\planilhas\\plan_bling.xlsx"
diretorio_bagy = "C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\planilhas\\plan_bagy.xlsx"
diretorio_dafiti = "C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\planilhas\\plan_dafiti.xlsx"
diretorio_meli = "C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\planilhas\\plan_meli.xlsx"
direftorio_zattini = "C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\planilhas\\plan_zattini.xlsx"

# Marcas Proibidas
marcas_proibidas_bagy = []
marcas_proibidas_dafiti = [
    'Abercrombie', 'Adidas', 'Aeropostale', 'Colcci', 'Crocs', 'Diesel', 'Disky', 'Guess', 'Hollister', 'Individual', 'Lacoste', 'Levis', 'New Era', 'Nike', 'Osklen', 'Polo Ralph Lauren', 'Replay', 'Victory Eagle'
]
marcas_proibidas_meli = [
    'Adidas', 'Crocs', 'Nike', 'Lacoste'
]
marcas_proibidas_zattini = [
    'Aeropostale', 'Lacoste', 'Levis', 'Nike', 'Osklen', 'Polo Ralph Lauren', 'Reserva'
]

def carregar_tabela(diretorio):
    return pd.read_excel(diretorio)

def merge_com_marketplace(tabela_bagy, tabela_marketplace):
    return tabela_bagy.merge(tabela_marketplace[['Código', 'Preço de', 'Preço por']], on='Código', how='left')

def adicionar_preco_final(tabela):
    tabela['Preço Final'] = tabela['Preço por'].apply(lambda x: round((x/10), 0)*10 - 0.01)
    return tabela

def remover_colunas(tabela, colunas):
    return tabela.drop(columns=colunas)

def agrupar_por_sku_pai(tabela):
    agg_dict = {'Estoque': 'sum',
                'Nome': 'first',
                'Preço De': 'first',
                'Preço Por': 'first',
                'Preço Final': 'first',
                'Preço de': 'first',
                'Preço por': 'first',
                'Status de cadastro': 'first',
                'SKU Pai': 'first'}
    return tabela.groupby('SKU Pai').agg(agg_dict)

def reordenar_colunas(tabela, colunas_ordenadas):
    return tabela.reindex(columns=colunas_ordenadas)

def salvar_tabela(tabela, nome_arquivo, output_directory):
    output_directory='C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\criadas\\'
    output_path = os.path.join(output_directory, nome_arquivo)
    tabela.to_excel(output_path, index=False)

def gerar_tabela_produtos_sem_cadastro(diretorio_bagy, marcas_proibidas, nome_arquivo):
    tabela_bagy = carregar_tabela(diretorio_bagy)
    tabela_bagy_filtrada = tabela_bagy[~tabela_bagy['Marca'].isin(marcas_proibidas)]

    # Crie uma cópia da tabela do site para adicionar o status de cadastro
    tabela_final = tabela_bagy_filtrada.copy()

    # Adicionando a coluna 'Status de cadastro' com o valor padrão 'Sem Cadastro'
    tabela_final['Status de cadastro'] = 'Sem Cadastro'

    # Salvar a nova tabela em um arquivo Excel
    output_directory = 'C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\criadas\\'
    salvar_tabela(tabela_final, nome_arquivo, output_directory)
    # Mensagem personalizada de impressão
    print(f"Uma nova tabela com o status de cadastro dos produtos foi criada: '{nome_arquivo}'")

def main():
    output_directory = 'C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\criadas\\'
    print('Você deseja criar uma planilha para ver os produtos sem cadastro na Bagy ou nos Marketplaces?\n1 - Bagy\n2 - Marketplaces')
    canal_escolhido = int(input('Digite o número correspondente a sua escolha: '))

    if canal_escolhido == 1:
        print('Perfeito! Vamos gerar uma planilha para ver os produtos sem cadastro e os que estão com estoque divergente.')
        nome_arquivo = 'sem_cadastro_bagy.xlsx'

        # Carregar as tabelas
        tabela_bling = carregar_tabela(diretorio_bling)
        tabela_bagy = carregar_tabela(diretorio_bagy)

        # Processamento da Bagy
        tabela_final = processar_tabela_site(tabela_bling, tabela_bagy)

        # Salvar e imprimir mensagem
        salvar_tabela(tabela_final, nome_arquivo, output_directory)

    elif canal_escolhido == 2:
        print('Para qual marketplace você deseja gerar a planilha?\n1 - Dafiti\n2 - Mercado Livre\n3 - Zattini - Netshoes')
        mktplc_selecionado = int(input('Digite o número que corresponde a sua escolha: '))

        if mktplc_selecionado in (1, 2, 3):
            # Seleciona marcas proibidas baseado na escolha
            if mktplc_selecionado == 1:
                marcas_proibidas = marcas_proibidas_dafiti
                nome_arquivo = 'produtos_sem_cadastro_dafiti.xlsx'
                tabela_marketplace = carregar_tabela(diretorio_dafiti)
            elif mktplc_selecionado == 2:
                marcas_proibidas = marcas_proibidas_meli
                nome_arquivo = 'Produtos_sem_cadastro_meli.xlsx'
                tabela_marketplace = carregar_tabela(diretorio_meli)
            else:
                marcas_proibidas = marcas_proibidas_zattini
                nome_arquivo = 'produtos_sem_cadastro_zattini.xlsx'
                tabela_marketplace = carregar_tabela(direftorio_zattini)

            # Processamento do Marketplace
            tabela_bagy = carregar_tabela(diretorio_bagy)
            tabela_final = processar_tabela_marketplace(tabela_bagy, marcas_proibidas, tabela_marketplace)

            # Salvar e imprimir mensagem
            salvar_tabela(tabela_final, nome_arquivo, output_directory)
        else:
            print('Você deve selecionar uma opção válida!')
    else:
        print('Você deve selecionar uma opção válida!')
    print('Sua tabela foi gerada com sucesso!')

def processar_tabela_site(tabela_bling, tabela_bagy):
    tabela_final = tabela_bling.copy()
    tabela_final['Status de cadastro'] = 'Sem Cadastro'
    tabela_final.loc[tabela_final['Código'].isin(tabela_bagy['Código']), 'Status de cadastro'] = 'Cadastrado'

    for index, row in tabela_bagy.iterrows():
        codigo = row['Código']
        estoque_bagy = row['Estoque']
        linha_correspondente = tabela_final[tabela_final['Código'] == codigo]
        if not linha_correspondente.empty:
            estoque_final = linha_correspondente.iloc[0]['Estoque']
            divergencia = estoque_final - estoque_bagy
            tabela_final.loc[tabela_final['Código'] == codigo, 'Divergência de Estoque'] = divergencia

    return tabela_final

def processar_tabela_marketplace(tabela_bagy, marcas_proibidas, tabela_marketplace):
    tabela_bagy_filtrada = tabela_bagy[~tabela_bagy['Marca'].isin(marcas_proibidas)]
    tabela_final = tabela_bagy_filtrada.copy()
    tabela_final['Status de cadastro'] = 'Sem Cadastro'
    tabela_final.loc[tabela_final['Código'].isin(tabela_marketplace['Código']), 'Status de cadastro'] = 'Cadastrado'

    tabela_final = merge_com_marketplace(tabela_final, tabela_marketplace)
    tabela_final = adicionar_preco_final(tabela_final)
    tabela_final = remover_colunas(tabela_final, ['Código', 'Marca'])
    tabela_final = agrupar_por_sku_pai(tabela_final)
    colunas_ordenadas = ['SKU Pai', 'Nome', 'Estoque', 'Preço De', 'Preço Por', 'Preço Final', 'Preço de', 'Preço por', 'Status de cadastro']
    tabela_final = reordenar_colunas(tabela_final, colunas_ordenadas)

    return tabela_final

if __name__ == '__main__':
    main()

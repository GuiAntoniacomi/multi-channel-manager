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

# Diretórios dos arquivos
diretorio_bagy = "C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\planilhas\\plan_bagy.xlsx"
diretorio_dafiti = "C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\planilhas\\plan_dafiti.xlsx"
diretorio_meli = "C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\planilhas\\plan_meli.xlsx"
direftorio_zattini =  "C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\planilhas\\plan_zattini.xlsx"

# Função para criar a tabela de produtos sem cadastro
def criar_tabela_sem_cadastro(diretorio_produto, diretorio_marketplace, marcas_proibidas, nome_arquivo):
    # Carregar tabelas
    tabela_bagy = pd.read_excel(diretorio_bagy)
    tabela_marketplace = pd.read_excel(diretorio_marketplace)
    
    # Filtrar tabela Bagy
    tabela_bagy_filtrada = tabela_bagy[~tabela_bagy['Marca'].isin(marcas_proibidas)]
    
    # Criar cópia da tabela Bagy para adicionar status de cadastro
    tabela_final = tabela_bagy_filtrada.copy()
    
    # Adicionar coluna 'Status de cadastro' com valor padrão 'Sem Cadastro'
    tabela_final['Status de cadastro'] = 'Sem Cadastro'
    
    # Atualizar status para 'cadastrado' onde códigos estão presentes no marketplace
    tabela_final.loc[tabela_final['Código'].isin(tabela_marketplace['Código']), 'Status de cadastro'] = 'Cadastrado'
    
    # Merge com tabela do marketplace para preço e estoque
    tabela_final = tabela_final.merge(tabela_marketplace[['Código', "Preço de", 'Preço por']], on='Código', how='left', suffixes=("", '_MARKETPLACE'))
    
    # Calcular preço final e outras operações necessárias
    tabela_final['Preço Final'] = tabela_final['Preço por'].apply(lambda x: round((x/10), 0) * 10 - 0.01)
    
    # Remover colunas desnecessárias
    tabela_final.drop(columns=['Código', 'Marca'], inplace=True)
    
    # Agrupar por 'SKU Pai', somar estoque e manter outras informações
    agg_dict = {'Estoque': 'sum',  # Soma do estoque
                'Nome': 'first',  # Manter a primeira ocorrência do nome do produto
                'Preço de': 'first',  # Manter a primeira ocorrência do preço
                'Preço por': 'first',  # Manter a primeira ocorrência do preço
                'Preço Final': 'first',  # Manter a primeira ocorrência do preço final
                'Preço de_MELI': 'first',  # Manter a primeira ocorrência do preço do marketplace
                'Preço por_MELI': 'first',  # Manter a primeira ocorrência do preço do marketplace
                'Status de cadastro': 'first',  # Manter a primeira ocorrência do status
                'Rentabilidade': 'first'}  # Manter a primeira ocorrência da rentabilidade
    tabela_final = tabela_final.groupby('SKU Pai').agg(agg_dict)
    
    # Reorganizar colunas (se necessário)
    # ...
    
    # Salvar em Excel
    output_directory = 'C:\\Users\\anton\\OneDrive\\Documents\\GitHub\\SecretShop\\src\\criadas\\'
    output_path = os.path.join(output_directory, nome_arquivo)
    tabela_final.to_excel(output_path, index=False)
    
    print(f"Uma nova tabela com o status de cadastro dos produtos foi criada: '{nome_arquivo}'")

# Função principal
def main():
    print('Você deseja criar uma planilha para ver os produtos sem cadastro na Bagy ou nos Marketplaces?\n1 - Bagy\n2 - Marketplaces')
    canal_escolhido = int(input('Digite o número correspondente a sua escolha:'))
    
    if canal_escolhido == 1:
        criar_tabela_sem_cadastro(diretorio_bagy, diretorio_dafiti, marcas_proibidas_dafiti, 'sem_cadastro_bagy.xlsx')
    elif canal_escolhido == 2:
        print('Para qual marketplace você deseja gerar a planilha?\n1 - Dafiti\n2 - Mercado Livre\n3 - Zattini - Netshoes')
        mktplc_selecionado = int(input('Digite o número que corresponde a sua escolha: '))
        
        if mktplc_selecionado == 1:
            criar_tabela_sem_cadastro(diretorio_bagy, diretorio_dafiti, marcas_proibidas_dafiti, 'produtos_sem_cadastro_dafiti.xlsx')
        elif mktplc_selecionado == 2:
            criar_tabela_sem_cadastro(diretorio_bagy, diretorio_meli, marcas_proibidas_meli, 'produtos_sem_cadastro_meli.xlsx')
        elif mktplc_selecionado == 3:
            criar_tabela_sem_cadastro(diretorio_bagy, direftorio_zattini, marcas_proibidas_zattini, 'produtos_sem_cadastro_zattini.xlsx')
        else:
            print('Você deve selecionar uma opção válida!')
    else:
        print('Você deve selecionar uma opção válida!')

if __name__ == "__main__":
    main()

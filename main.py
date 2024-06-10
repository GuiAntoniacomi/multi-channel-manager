import pandas as pd
import numpy as np
import re
import os
from tkinter import *
from tkinter import ttk
import tkinter.filedialog
from tkinter import filedialog, messagebox
import sys

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def base_bagy(caminho):
    bagy_json = pd.read_json(caminho)
    bagy_json = bagy_json[['Brands → Name', 'Variations → Sku', 'Price', 'Price Compare', 'Name', 'Stocks → Balance', 'External ID']]
    bagy_json = bagy_json.rename(columns={'Brands → Name': 'Marca', 'Variations → Sku': 'Código', 'Price Compare': 'Preço De', 'Price': 'Preço Por', 'Name': 'Nome', 'Stocks → Balance': 'Estoque', 'External ID': 'SKU Pai'})
    bagy_json = bagy_json.reindex(columns=["SKU Pai", "Código", "Marca", "Nome", "Estoque", "Preço De", "Preço Por"])
    bagy_json = bagy_json.dropna(subset=['Código'])
    bagy_json['Código'] = bagy_json['Código'].astype(int)
    return bagy_json

def base_dafiti(caminho):
    dafiti_csv = pd.read_csv(caminho, delimiter=';', low_memory=False)
    colunas_para_remover = ['CatalogProductId', 'Dafiti SKU', 'Ncm', 'Brand', 'Color', 'ColorFamily', 'PrimaryCategory', 'Categories', 'Gender', 'Warranty', 'Origin', 'Origincountry', 'ObservacoesImportantes', 'RegistroDoProduto', 'Description', 'BrowseNodes', 'SaleEndDate', 'SaleStartDate', 'Variation', 'ProductId', 'Occasion', 'Character', 'ModalidadeEsporte', 'Waist', 'Length', 'Specifications', 'Print', 'Style', 'Kit', 'Sleeve', 'ModelClothes', 'Composition', 'Washing', 'WashingType', 'ClothesMaterial', 'BoxHeight', 'BoxLength', 'BoxWidth', 'Weight', 'CreatedAt', 'UpdatedAt', 'Status', 'ParentSku']
    colunas_existentes = [coluna for coluna in colunas_para_remover if coluna in dafiti_csv.columns]
    dafiti_df = dafiti_csv.drop(columns=colunas_existentes)
    dafiti_df = dafiti_df.rename(columns={'SellerSku': 'Código', 'Name': 'Nome', 'Quantity': 'Estoque', 'Price': 'Preço De', 'SalePrice': 'Preço Por'})
    dafiti_df = dafiti_df.reindex(columns=['Código', 'Nome', 'Estoque', 'Preço De', 'Preço Por'])
    dafiti_df['Código'] = dafiti_df['Código'].astype(int)
    dafiti_df = dafiti_df.dropna(subset=['Código'])
    return dafiti_df

def base_meli(caminho):
    df_meli = pd.read_excel(caminho, sheet_name='Anúncios')
    df_meli = df_meli.drop([0, 1])
    df_meli.reset_index(drop=True, inplace=True)
    df_meli = df_meli[['ITEM_ID', 'SKU', 'PRICE']]
    df_meli['SKU'] = df_meli['SKU'].fillna(0)
    df_meli['PRICE'] = df_meli['PRICE'].fillna(0)
    df_meli['SKU'] = df_meli['SKU'].astype(int)
    df_meli['PRICE'] = df_meli['PRICE'].astype(float)
    df_meli = df_meli.groupby('ITEM_ID').agg({'SKU': 'last', 'PRICE': 'first'}).reset_index()
    df_meli = df_meli.drop(columns=['ITEM_ID'])
    df_meli = df_meli.rename(columns={'SKU': 'Código', 'PRICE': 'Preço Por Meli'})
    return df_meli

def base_zattini(caminho):
    zattini_excel = pd.read_excel(caminho)
    zattini_df = pd.DataFrame(zattini_excel)
    zattini_df = zattini_df.rename(columns={'Sku Seller': 'Código'})
    
    def extract_numbers(text):
        match = re.search(r'\d+', text)
        if match:
            return int(match.group())
        return None
    
    zattini_df['Código'] = zattini_df['Código'].apply(extract_numbers)
    return zattini_df

def exportar_para_dafiti(bagy, marketplace):
    marcas_proibidas_dafiti = ['Abercrombie', 'Adidas', 'Aeropostale', 'Colcci', 'Crocs', 'Diesel', 'Disky', 'Guess', 'Hollister', 'Individual', 'Lacoste', 'Levis', 'New Era', 'Nike', 'Osklen', 'Polo Ralph Lauren', 'Replay', 'Victory Eagle']
    df_exportar_dafiti = bagy[~bagy['Marca'].isin(marcas_proibidas_dafiti)].copy()
    df_exportar_dafiti.loc[:, 'Status Cadastro'] = 'Sem Cadastro'
    df_exportar_dafiti.loc[df_exportar_dafiti['Código'].isin(marketplace['Código']), 'Status Cadastro'] = 'Cadastrado'
    df_exportar_dafiti = df_exportar_dafiti.merge(marketplace, on='Código', how='left')
    df_exportar_dafiti = df_exportar_dafiti.drop(['Nome_y', 'Estoque_y', 'Código'], axis=1)
    df_exportar_dafiti = df_exportar_dafiti.rename(columns={'Nome_x': 'Nome', 'Estoque_x': 'Estoque', 'Preço De_y': 'Preço De Dft', 'Preço Por_y': 'Preço Por Dft', 'Preço De_x': 'Preço De Bagy', 'Preço Por_x': 'Preço Pix Bagy'})
    df_dafiti_final = df_exportar_dafiti.groupby('SKU Pai').agg({'Marca': 'first', 'Nome': 'first', 'Estoque': 'sum', 'Preço De Bagy': 'first', 'Preço Pix Bagy': 'first', 'Status Cadastro': 'first', 'Preço De Dft': 'first', 'Preço Por Dft': 'first'}).reset_index()
    df_dafiti_final['Preço Por Bagy'] = df_dafiti_final['Preço Pix Bagy'].apply(lambda x: np.ceil(x / 10) * 10 - 0.01)
    df_dafiti_final = df_dafiti_final.sort_values(by='Estoque', ascending=False)
    df_dafiti_final = df_dafiti_final.dropna(subset=['Marca'])
    colunas = ['Marca', 'SKU Pai', 'Nome', 'Estoque', 'Preço De Bagy', 'Preço Pix Bagy', 'Preço Por Bagy', 'Preço De Dft', 'Preço Por Dft', 'Status Cadastro']
    df_dafiti_final = df_dafiti_final[colunas]
    return df_dafiti_final

def exportar_meli(bagy, marketplace):
    marcas_proibidas_meli = ['Adidas', 'Crocs', 'Nike', 'Lacoste']
    df_exportar_meli = bagy[~bagy['Marca'].isin(marcas_proibidas_meli)].copy()
    df_exportar_meli.loc[:, 'Status Cadastro'] = 'Cadastrado'
    df_exportar_meli = df_exportar_meli.merge(marketplace, on='Código', how='left')
    df_exportar_meli = df_exportar_meli.groupby('SKU Pai').agg({'Marca': 'first', 'Nome': 'first', 'Estoque': 'sum', 'Preço De': 'first', 'Preço Por': 'first', 'Status Cadastro': 'first', 'Preço Por Meli': 'max'}).reset_index()
    df_exportar_meli = df_exportar_meli.sort_values(by='Estoque', ascending=False)
    df_exportar_meli.loc[df_exportar_meli['Preço Por Meli'].isna(), 'Status Cadastro'] = 'Sem Cadastro'
    df_exportar_meli.drop([6362, 5007], inplace=True)
    df_exportar_meli = df_exportar_meli.rename(columns={'Preço Por': 'Preço Pix'})
    df_exportar_meli['Preço Por'] = df_exportar_meli['Preço Pix'].apply(lambda x: np.ceil(x / 10) * 10 - 0.01)
    colunas = ['Marca', 'SKU Pai', 'Nome', 'Estoque', 'Preço De', 'Preço Pix', 'Preço Por', 'Preço Por Meli', 'Status Cadastro']
    df_exportar_meli = df_exportar_meli[colunas]
    return df_exportar_meli

def exportar_zattini(bagy, marketplace):
    marcas_proibidas_zattini = ['Aeropostale', 'Lacoste', 'Levis', 'Nike', 'Osklen', 'Polo Ralph Lauren', 'Reserva']
    df_exportar_zattini = bagy[~bagy['Marca'].isin(marcas_proibidas_zattini)].copy()
    df_exportar_zattini.loc[:, 'Status Cadastro'] = 'Sem Cadastro'
    df_exportar_zattini.loc[df_exportar_zattini['Código'].isin(marketplace['Código']), 'Status Cadastro'] = 'Cadastrado'
    df_exportar_zattini = df_exportar_zattini.merge(marketplace, on='Código', how='left')
    df_exportar_zattini = df_exportar_zattini.drop(columns=['Código'])
    df_exportar_zattini = df_exportar_zattini.rename(columns={'Preço De_x': 'Preço De Bagy', 'Preço Por_x': 'Preço Pix Bagy', 'Preço De_y': 'Preço De Ztn', 'Preço Por_y': 'Preço Por Ztn'})
    df_exportar_zattini = df_exportar_zattini.groupby('SKU Pai').agg({'Marca': 'first', 'Nome': 'first', 'Estoque': 'sum', 'Preço De Bagy': 'first', 'Preço Pix Bagy': 'first', 'Status Cadastro': 'first', 'Preço De Ztn': 'first', 'Preço Por Ztn': 'first'}).reset_index()
    df_exportar_zattini = df_exportar_zattini.sort_values(by='Estoque', ascending=False)
    df_exportar_zattini.drop([4604, 3555], inplace=True)
    df_exportar_zattini['Preço Por Bagy'] = df_exportar_zattini['Preço Pix Bagy'].apply(lambda x: np.ceil(x / 10) * 10 - 0.01)
    colunas = ['Marca', 'SKU Pai', 'Nome', 'Estoque', 'Preço De Bagy', 'Preço Pix Bagy', 'Preço Por Bagy', 'Preço De Ztn', 'Preço Por Ztn', 'Status Cadastro']
    df_exportar_zattini = df_exportar_zattini[colunas]
    return df_exportar_zattini

def select_json_file():
    file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
    if file_path:
        json_file_entry.delete(0, END)
        json_file_entry.insert(0, file_path)

def select_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ('CSV files', '*.csv')])
    if file_path:
        excel_file_entry.delete(0, END)
        excel_file_entry.insert(0, file_path)

def caminho_salvar():
    caminho = tkinter.filedialog.askdirectory(title='Selecione onde salvar o arquivo')
    if caminho:
        diretorio_salvar.delete(0, END)
        diretorio_salvar.insert(0, caminho)

def executar_app():
    caminho_bagy = json_file_entry.get()
    caminho_marketplace = excel_file_entry.get()
    mktplc_selecionado = mktplc.get()
    
    if not caminho_bagy or not caminho_marketplace:
        messagebox.showwarning('Aviso', 'Por favor, selecione os arquivos JSON e Excel.')
        return
    
    arquivo_bagy = base_bagy(caminho_bagy)
    
    if mktplc_selecionado == 'Dafiti':
        arquivo_marketplace = base_dafiti(caminho_marketplace)
        df_final = exportar_para_dafiti(arquivo_bagy, arquivo_marketplace)
        nome_arquivo = 'tabela_dafiti.xlsx'
    elif mktplc_selecionado == 'Mercado Livre':
        arquivo_marketplace = base_meli(caminho_marketplace)
        df_final = exportar_meli(arquivo_bagy, arquivo_marketplace)
        nome_arquivo = 'tabela_meli.xlsx'
    elif mktplc_selecionado == 'Netshoes':
        arquivo_marketplace = base_zattini(caminho_marketplace)
        df_final = exportar_zattini(arquivo_bagy, arquivo_marketplace)
        nome_arquivo = 'tabela_zattini.xlsx'
    else:
        messagebox.showwarning('Aviso', 'Por favor, selecione um marketplace.')
        return
    
    caminho = diretorio_salvar.get()
    if not caminho:
        messagebox.showwarning('Aviso', 'Por favor, selecione o diretório para salvar o arquivo.')
        return
    
    caminho_completo = os.path.join(caminho, nome_arquivo)
    df_final.to_excel(caminho_completo, index=False)
    messagebox.showinfo('Sucesso', f'Arquivo salvo em {caminho_completo}')
    print(f"Arquivo salvo em {caminho_completo}")

def encerrar_app():
    window.quit()

window = Tk()
window.title("Gerador de Planilha")
window.geometry("700x400")
window.configure(bg = "#ffffff")

canvas = Canvas(window, bg = "#ffffff", height = 400, width = 700, bd = 0, highlightthickness = 0, relief = "ridge")
canvas.place(x = 0, y = 0)

background_img = PhotoImage(file = resource_path("front_end\\background.png"))
background = canvas.create_image(365.0, 200.0, image=background_img)

entry0_img = PhotoImage(file = resource_path("front_end\\img_textBox0.png"))
entry0_bg = canvas.create_image(517.5, 160.5, image = entry0_img)

entry0 = Label(bd = 0, bg = "#ffffff", highlightthickness = 0)

json_file_entry = Entry(window, width=50)
json_file_entry.place(x = 430, y = 150, width = 175, height = 19)
entry0.place(x = 430, y = 150, width = 175, height = 19)

entry1_img = PhotoImage(file = resource_path("front_end\\img_textBox1.png"))
entry1_bg = canvas.create_image(517.5, 229.5, image = entry1_img)

entry1 = Entry(bd = 0, bg = "#ffffff", highlightthickness = 0)

excel_file_entry = Entry(window, width=50)
excel_file_entry.place(x = 430, y = 219, width = 175, height = 19)
entry1.place(x = 430, y = 219, width = 175, height = 19)

entry2_img = PhotoImage(file = resource_path("front_end\\img_textBox2.png"))
entry2_bg = canvas.create_image(517.5, 298.5, image = entry2_img)

entry2 = Entry(bd = 0, bg = "#ffffff", highlightthickness = 0)

diretorio_salvar = Entry(window, width=50)
diretorio_salvar.place(x = 430, y = 288, width = 175, height = 19)
entry2.place(x = 430, y = 288, width = 175, height = 19)

img0 = PhotoImage(file = resource_path("front_end\\img0.png"))
b0 = Button(image = img0, borderwidth = 0, highlightthickness = 0, command = select_json_file, relief = "flat")
b0.place(x = 620, y = 150, width = 56, height = 21)

img1 = PhotoImage(file = resource_path("front_end\\img1.png"))
b1 = Button(image = img1, borderwidth = 0, highlightthickness = 0, command = select_excel_file, relief = "flat")
b1.place(x = 620, y = 219, width = 56, height = 21)

img2 = PhotoImage(file = resource_path("front_end\\img2.png"))
b2 = Button(image = img2, borderwidth = 0, highlightthickness = 0, command = caminho_salvar, relief = "flat")
b2.place(x = 620, y = 288, width = 56, height = 21)

img3 = PhotoImage(file = resource_path("front_end\\img3.png"))
b3 = Button(image = img3, borderwidth = 0, highlightthickness = 0, command = encerrar_app, relief = "flat")
b3.place(x = 430, y = 352, width = 55, height = 21)

img4 = PhotoImage(file = resource_path("front_end\\img4.png"))
b4 = Button(image = img4, borderwidth = 0, highlightthickness = 0, command = executar_app, relief = "flat")
b4.place(x = 605, y = 352, width = 71, height = 21)

lista_mktplc =['Dafiti', 'Mercado Livre', 'Netshoes']
mktplc = ttk.Combobox(window, values=lista_mktplc)
mktplc.place(x = 430, y = 100, width = 175, height = 19)

window.resizable(False, False)
window.mainloop()

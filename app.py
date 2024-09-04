from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.options import Options

from selenium.webdriver.common.by import By

from time import sleep
import random

import openpyxl

import PySimpleGUI as sg

sg.theme('Reddit')
#sg.Push() serve para centralizar o elemento, para isto coloque em ambos os lados
layout = [
    [sg.Push(),sg.Text('Digite o nome do produto'),sg.Push()],    
    [sg.Push(),sg.Input(key='nome_produto', size= 30),sg.Push()],
    [sg.Push(),sg.Button(button_text='Pesquisar Produto'),sg.Push()]    
]

window = sg.Window('Buscar Preço', layout)

edge_options = Options()
arguments = ['--lang=pt-BR', 'window-size=1280,960', '--incognito']
for argument in arguments:
    edge_options.add_argument(argument)

caminho_padrao_para_download = 'C:\\Users\\rscop\\Desktop'

edge_options.add_experimental_option("prefs", {
    'download.default_directory': caminho_padrao_para_download,
    'download.directory_upgrade': True,
    'download.prompt_for_download': False,
    'profile.default_content_setting_values.notifications': 2,
    'profile.default_content_setting_values.automatic_downloads': 1
})

def digitar_naturalmente(nome_produto, campo_pesquisa):
    for letra in nome_produto:
        campo_pesquisa.send_keys(letra)
        sleep(random.randint(1,5)/30)

while True:
    event,values = window.read()
    if event == sg.WIN_CLOSED:
        break
    elif event =='Pesquisar Produto': 

        driver = webdriver.Edge(options=edge_options)
        driver.get('https://www.buscape.com.br/')

        campo_pesquisa = driver.find_element(By.XPATH, "//div[@class='AutoCompleteStyle_autocomplete__BvELB']/input[@data-test='input-search']")
        digitar_naturalmente(values['nome_produto'], campo_pesquisa)
        sleep(5)
        botao_pesquisar = driver.find_element(By.XPATH, "//div[@class='AutoCompleteStyle_autocomplete__BvELB']/button[@class='AutoCompleteStyle_submitButton__VwVxN']")
        driver.execute_script('arguments[0].click()', botao_pesquisar)
        sleep(5)

        titulos = driver.find_elements(By.XPATH, "//h2[@data-testid='product-card::name']")        
        precos = driver.find_elements(By.XPATH, "//p[@data-testid='product-card::price']")
        parcelamentos = driver.find_elements(By.XPATH, "//span[@data-testid='product-card::installment']")
        links = driver.find_elements(By.XPATH, "//a[@data-testid='product-card::card']")
        numero_de_anuncios = driver.find_elements(By.XPATH, "//div[@class='Hits_ProductCard__Bonl_']")     

        workbook = openpyxl.Workbook()
        del workbook['Sheet']
        workbook.create_sheet('Produtos')
        sheet_produto = workbook['Produtos']
        sheet_produto.append(['Nome do Produto', 'Preço', 'Valor parcelado', 'Link da loja'])
        sheet_produto.column_dimensions['A'].width = 60
        sheet_produto.column_dimensions['B'].width = 15
        sheet_produto.column_dimensions['C'].width = 40
        sheet_produto.column_dimensions['D'].width = 300

        for produto in range (numero_de_anuncios.__len__()):            
            sheet_produto.append([titulos[produto].text, precos[produto].text, parcelamentos[produto].text, links[produto].get_attribute('href')])    

        workbook.save('Pesquisa_de_Preco-'+ values['nome_produto'] +'.xlsx')
        driver.close()
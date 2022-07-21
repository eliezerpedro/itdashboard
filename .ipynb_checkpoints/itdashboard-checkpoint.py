import json
import os
import shutil
import pandas as pd
from pathlib import Path
from time import sleep
import re
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException

class Itdashboard:
    def __init__(self):
        self.dirpath = os.getcwd()+'\\chromedriver.exe' #define o diretorio atual
        self.download_path = os.getcwd()+'\\output' #define o diretorio de download
        self.chrome_options = Options() 
        self.chrome_options.add_argument('--start-maximized') #ativa o modo full screen do navegador
        self.chrome_options.add_argument("--enable-javascript") #habilita o javascript
        self.appState = {
           "recentDestinations": [
                {
                    "id": "Save as PDF",
                    "origin": "local",
                    "account": ""
                }
            ],
            "selectedDestinationId": "Save as PDF",
            "version": 2
        }
        self.prefs = {
            'printing.print_preview_sticky_settings.appState': json.dumps(self.appState),
            'savefile.default_directory': self.download_path
                }

        self.chrome_options.add_experimental_option('prefs', self.prefs) #seta as preferências na tela de imprimir PDFS
        self.chrome_options.add_argument('--kiosk-printing')
        self.browser = webdriver.Chrome(executable_path=self.dirpath, options=self.chrome_options) #instancia o driver
        
    #Coleta os dados das agencias
    def get_agencys(self):
        #cria um dicionário que será transformado em um dataframe
        agency_dic = {"Agency": [],
       "FY 2022 IT Spending": [],
       "Spending on Major Investments": []
              }
        self.browser.get('https://www.itdashboard.gov/itportfoliodashboard') #acessa a página IT Portfolio Dashboard
        for agency in self.browser.find_element_by_id('agency-select').text.split('\n'): #itera no select de agencias, agencia por agencia
            print(f'Colhendo dados da empresa {agency}')
            self.browser.find_element_by_xpath(f'*//option[contains(text(), "{agency}")]').click() #clica na agencia selecionada
            #aguarda as labels contendo os spendings
            while 1:
                try:
                    self.browser.find_element_by_xpath('*//div[@class="spending-overview"]//h3')
                    break
                except NoSuchElementException:
                    pass
            #aguarda o nome da agencia aparecer em caixa alta na tela
            while 1:
                try:
                    while not self.browser.find_element_by_xpath('*//div[@class="spending-overview"]//h3').text == agency.upper(): 
                        sleep(1)
                    break
                except:
                    pass
                
            it_spending = self.browser.find_element_by_xpath('*//div[@class="it-spending"]').text.split('\n')[1] #captura os dados da label it-spending
            major_spending = self.browser.find_element_by_xpath('*//div[@class="major-spending"]').text.split('\n')[1] #captura os dados da label major-spending
            
            #adiciona os dados capturados no dicionário
            agency_dic['Agency'].append(agency)
            agency_dic['FY 2022 IT Spending'].append(it_spending)
            agency_dic['Spending on Major Investments'].append(major_spending)

        df_agencias = pd.DataFrame(agency_dic) #salva o dicionário com os dados das agências em um dataframe
        df_agencias.to_csv('df_agencias.csv') #salva o dataframe em um arquivo .csv
        return df_agencias
        
    #coleta os montantes individuais da empresa escolhida
    def get_individual_spendings(self, choose_agency):
        self.browser.get('https://www.itdashboard.gov/index.php/search-advanced') #Abre a página de busca avançada
        self.browser.find_element_by_id('edit-keywords').send_keys(choose_agency) #digita o nome da agencia escolhida
        self.browser.find_element_by_id('edit-submit').click() #clica em 'search"
        #Escolhe a opção de 40 itens por pagina, somente quando o select está disponível
        while 1:
            try:
                self.browser.find_element_by_xpath(f'*//option[contains(text(), "40")]').click()
                break
            except:
                sleep(0.1)
        #aguarda a pagina atualizar para poder seguir
        while not re.search('40', self.browser.find_elements_by_xpath('*//div[@class="number-of-results"]')[1].text):
            sleep(1)
        pages = int(self.browser.find_elements_by_xpath('*//div[@class="pager-nav"]')[1].text.split()[3]) #captura o numero de paginas
        #cria um dicionário que será transformado em um dataframe com o montante individual e o tipo do investimento
        spending_dic={"Spending $ FY2022": [],
                      "Type of Investment":[]
                      }
        #cria um dicionário onde serão adicionados os links de download dos business cases
        link_dic={"Link": []
        }
        #itera na página completa de uma só vez
        for page in range(1, pages + 1):
            print('Página {}. Linha {} de {}'.format(page, page, pages))
            #aguarda a página atualizar
            while not re.search(f"Page {page} of", self.browser.find_elements_by_xpath('*//div[@class="pager-nav"]')[1].text):
                sleep(1)
            #captura os montantes individuais iterando nos elementos que contenha os montantes
            for elemento in self.browser.find_elements_by_xpath('*//span[contains(text(), "Spending $ FY2022")]//..'):
                if len(elemento.text.split('\n')) == 2:
                    spending_dic["Spending $ FY2022"].append(elemento.text.split('\n')[1])
                    
            #captura os tipos de investimento iterando nos elementos que contenha os tipos de investimento
            for elemento in self.browser.find_elements_by_xpath('*//span[contains(text(), "Type of Investment")]//..'):
                if len(elemento.text.split('\n')) == 2:
                    spending_dic["Type of Investment"].append(elemento.text.split('\n')[1])
                    
            #cria uma lista de elementos que contem o link de download dos business cases
            elems = self.browser.find_elements_by_xpath('*//span[@class="download-tag"]//a')
            links = [elem.get_attribute('href') for elem in elems] #retorna uma lista com todos os links da pagina
            #itera na lista de links e adiciona no dicionário de links
            for link in links:
                link_dic["Link"].append(link)
            self.browser.find_element_by_xpath('*//a[@class="next"]').click() #clica no botão de next page
        
        df_spending = pd.DataFrame(spending_dic) #transforma o spending_dic em um dataframe
        df_spending.to_csv('df_spending.csv') #salva o dataframe em um arquivo .csv
        
        df_links = pd.DataFrame(link_dic) #transforma o link_dic em um dataframe
        df_links['Check'] = '-' #Cria uma coluna que servirá para armazenar a informação se o arquivo foi baixado ou não
        df_links.to_csv('links.csv') #salva o dataframe em um arquivo .csv
        
        return df_spending
        
    #Salva as informações das agencias e os dados da agencia escolhida em planilhas diferentes no mesmo arquivo
    def save_excel(self, df_agencias, df_spending):
        #define diretório e o nome do excel a ser salvo
        writer = pd.ExcelWriter(self.download_path+'\\'+'Agencys and Spending.xlsx', engine='xlsxwriter')
        #Salva cada dataframe em uma planilha diferente
        df_agencias.to_excel(writer, sheet_name='Agências')
        df_spending.to_excel(writer, sheet_name='Individual Investments')
        #Fecha o escritor de excel do pandas
        writer.save()
        os.remove('df_spending.csv')
        os.remove('df_agencias.csv')
        return True
    
    #faz o download dos business cases coletados no método get_individual_spendings
    def download_business_cases(self, df_links):
        it = 1
        for index, row in df_links.iterrows():
            print('baixando arquivo {}. Linha {} de {}'.format(index, it, len(df_links.index)))
            it = it+1
            #checa se o arquivo já foi baixado
            if row["Check"] == 'Baixado':
                print('Arquivo já baixado')
                continue

            self.browser.get(row["Link"]) #abre a página de download
            self.browser.find_element_by_xpath('*//a[@class="print-link"]').click() #clica no botão de download
            df_links['Check'][index] = 'Baixado' #marca o link como baixado
            df_links.to_csv('links.csv') #salva no arquivo
        self.browser.quit() #fecha a instancia do browser.
            

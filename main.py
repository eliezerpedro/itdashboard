from itdashboard import Itdashboard
import pandas as pd

def main():
    itdashboard = Itdashboard() #instancia a classe
    df_agencias = itdashboard.get_agencys() #Coleta os dados das agencias
    choose_agency = 'Department of Commerce' #escolhe a agencia Department of Commerce
    df_spending = itdashboard.get_individual_spendings(choose_agency) #coleta os montantes individuais da empresa escolhida
    itdashboard.save_excel(df_agencias, df_spending) #Salva as informações das agencias e os dados da agencia escolhida em planilhas diferentes no mesmo arquivo
    df_links = pd.read_csv('links.csv') #lê o arquivo links.csv criado no método get_individual_spendings
    itdashboard.download_business_cases(df_links) #faz o download dos business cases coletados no método get_individual_spendings
    
main()
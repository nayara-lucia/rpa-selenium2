from selenium import webdriver
from selenium.webdriver.common.by import By
import pyautogui as pa
from openpyxl import load_workbook


tabela_excel = 'C:\\Users\\znaya\\Desktop\\Python\\Pycharm\\Desafio final\\challenge.xlsx' #Direcionamos o caminho onde a tabela se encontra

dados_registros = load_workbook(tabela_excel) #Load Workbook para abrir o arquivo excel e realizar as operações


aba_selecionada = dados_registros['Sheet1']

navegador = webdriver.Chrome()
navegador.get('https://rpachallenge.com')
pa.sleep(1)

for linha in range(2, len(aba_selecionada["A"]) + 1): #Pra cada linha num RANGE que começa na linha 2 da planilha (pulando o cabeçalho) que tem o LEN(tamanho) 10(aba selecionada) + 1 11 linhas da coluna "A"

    # % linha pega o valor da linha quando estiver passando na linha da coluna A
    # %s transforma em string
    FirstName = aba_selecionada['A%s' % linha].value
    LastName = aba_selecionada['B%s' % linha].value
    Email = aba_selecionada['F%s' % linha].value
    Cargo = aba_selecionada['D%s' % linha].value
    Phone = aba_selecionada['G%s' % linha].value
    Empresa = aba_selecionada['C%s' % linha].value
    Endereco = aba_selecionada['E%s' % linha].value

# ------------------------------------------------------------------------ #

    navegador.find_element(By.XPATH, ".//*[@ng-reflect-name='labelFirstName']").send_keys(FirstName)
    pa.sleep(1)

    navegador.find_element(By.XPATH, ".//*[@ng-reflect-name='labelEmail']").send_keys(Email)
    pa.sleep(1)

    navegador.find_element(By.XPATH, ".//*[@ng-reflect-name='labelRole']").send_keys(Cargo)
    pa.sleep(1)

    navegador.find_element(By.XPATH, ".//*[@ng-reflect-name='labelLastName']").send_keys(LastName)
    pa.sleep(1)

    navegador.find_element(By.XPATH, ".//*[@ng-reflect-name='labelPhone']").send_keys(Phone)
    pa.sleep(1)

    navegador.find_element(By.XPATH, ".//*[@ng-reflect-name='labelCompanyName']").send_keys(Empresa)
    pa.sleep(1)

    navegador.find_element(By.XPATH, ".//*[@ng-reflect-name='labelAddress']").send_keys(Endereco)
    pa.sleep(1)

    navegador.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input').click()
    pa.sleep(1)

print("Cadastro Finalizado")
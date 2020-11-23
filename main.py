from selenium.webdriver import Firefox
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import setupInfo
from openpyxl import Workbook
import csv
from unidecode import unidecode
import time

idPedido = []
nome = []
cep = []
endereco = []
numero = []
complemento = []
bairro = []
cidade = []
estado = []
telefone = []
celular = []
email = []
cpf = []
conteudo = []
valor = []

def getConteudo(id, csv):
    found = ([linha[3], linha[10], linha[11]] for linha in reader2 if linha[0] == id)
    csv.seek(0)
    return list(found)

with open('pedidos.csv', newline='', encoding='iso-8859-1') as csvfile:
    reader = csv.reader(csvfile, delimiter=';')
    for rowNum, row in enumerate(reader):
        if(rowNum != 0 and row[8] != 'Motoboy'):
            idPedido.append(row[0])
            nome.append(row[27])
            endereco.append(row[14])
            numero.append(row[15])
            complemento.append(row[16])
            bairro.append(row[17])
            cidade.append(row[18])
            estado.append(row[19])
            cep.append(row[20])
            email.append(row[22])
            cpf.append(row[23].split('"')[1])
            valor.append(row[6])

with open('conteudo.csv', newline='', encoding='iso-8859-1') as csvfile:
    reader2 = csv.reader(csvfile, delimiter=';')
    conteudo = [getConteudo(id, csvfile) for id in idPedido]

##Editar!
def waitByXpath(xpath):
    return WebDriverWait(browser, 20).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )

def clickByXpath(xpath):
    elemento = waitByXpath(xpath)
    elemento.click()

def clickById(id):
    elemento = WebDriverWait(browser, 20).until(
            EC.element_to_be_clickable((By.ID, id))
        )
    elemento.click()

browser = Firefox()
browser.get('https://vipp.visualset.com.br/vipp/inicio/index.php')
browser.find_element_by_id("txtUsr").send_keys(setupInfo.userName)

browser.find_element_by_id('txtPwd').send_keys(setupInfo.userPassword)
clickById('btnEfetuarLogin')
clickByXpath("//a[@title='Ir Para CheckList Isolando Este Status' and @onclick=\"IrParaCheckListStatus('8');\"]")

browser.switch_to.window(browser.window_handles[1])

clickByXpath("(//th/div[contains(@class,'th-inner') and text()='Nº ViPP'])[1]")

WebDriverWait(browser, 20).until(
        EC.element_to_be_clickable((By.XPATH, "(//button[contains(@class,'btnedtobj')])[1]"))
    )

editarConteudos = browser.find_elements_by_xpath("//button[contains(@class,'btnedtobj')]")

for index in range(len(idPedido)):
    editarConteudos[index].click()
    frameEditar = browser.find_element_by_id('ifrEditarObjeto')
    browser.switch_to.frame(frameEditar)
    print(browser.find_element_by_id('txtNomeDestinatario').get_attribute("value"))
    clickById("btnDeclaracaoConteudo")
    browser.switch_to.frame(browser.find_element_by_id('IFrmDecConteudo'))
    time.sleep(2)
    waitByXpath("(//div[contains(@class,'ui-widget-content')]/div[contains(@class, 'CssPes')])[1]")
    print(browser.find_element_by_xpath("(//div[contains(@class,'ui-widget-content')]/div[1])[1]").get_attribute("text"))

    browser.switch_to.default_content()
    #browser.switch_to.frame(browser.find_element_by_id('ifrEditarObjeto'))
    clickById('BtnSalvarDeclaracao')
    clickByXpath("//div[contains(@class, 'form-group')]/button[contains(@class, 'cmdGravar')]")
    browser.switch_to.default_content()

# btnDeclaracaoConteudo
# IFrmDecConteudo

#Declarar conteudo //a[@class='Editor btnEditarDeclaracaoConteudo']
#descrição //div[contains(@class,'ui-widget-content')]/div[1]
#Quantidade //div[contains(@class,'ui-widget-content')]/div[contains(@class, 'CssPes')]
#valor //div[contains(@class,'ui-widget-content')]/div[contains(@class, 'CssVlr')]
#salvar //button[@id='BtnSalvarDeclaracao']
#importar: btnImportarArquivo

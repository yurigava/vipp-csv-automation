from selenium.webdriver import Firefox
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from unidecode import unidecode
import setupInfo
import pickle
import time

idPedido = 'idPedido'
nomeClientes = 'nome'
listaConteudos = 'conteudo'
pacName = 'pac'
sedexName = 'sedex'

dictKeys = [
    idPedido,
    nomeClientes,
    listaConteudos
]

INDICE_DESC_PRODUTO = 0
INDICE_QTY_PRODUTO = 1
INDICE_VALOR_PRODUTO = 2

with open('conteudo.pickle', 'rb') as conteudoDb:
    recoveredData = pickle.load(conteudoDb, encoding='bytes')

pedidosInfo = {
    key: [value for shipping in recoveredData for value in recoveredData[shipping][key]]
    for key in dictKeys
}

pedidosInfo[nomeClientes] = [unidecode(nome) for nome in pedidosInfo[nomeClientes]]
print(pedidosInfo[nomeClientes])


def waitBy(byType, value):
    return WebDriverWait(browser, 20).until(
            EC.element_to_be_clickable((byType, value))
        )


def clickByXpath(xpath):
    waitBy(By.XPATH, xpath)
    browser.find_element_by_xpath(xpath).click()


def clickById(id):
    waitBy(By.ID, id)
    browser.find_element_by_id(id).click()


def waitInvisibility(xpath):
    WebDriverWait(browser, 10).until(
        EC.invisibility_of_element_located((By.XPATH, xpath))
    )


def waitTillPageLoads():
    waitInvisibility("//div[contains(@class, 'modal in')]/div/div/div/h3[text()='Aguarde...']")
    time.sleep(1)
    waitInvisibility("//div[contains(@class, 'modal in')]/div/div/div/h3[text()='Aguarde...']")


def waitTillTableLoads():
    waitInvisibility("//div[@class='fixed-table-loading' and following-sibling::table[@id='TabelaPostagens']]")
    time.sleep(1)
    waitInvisibility("//div[@class='fixed-table-loading' and following-sibling::table[@id='TabelaPostagens']]")


def fillInContentValue(xpath, valueText):
    browser.find_element_by_xpath(xpath).click()
    browser.find_element_by_xpath(xpath+"/input").clear()
    browser.find_element_by_xpath(xpath+"/input").send_keys(valueText)


browser = Firefox()
browser.get('https://vipp.visualset.com.br/vipp/inicio/index.php')
browser.find_element_by_id("txtUsr").send_keys(setupInfo.userName)

browser.find_element_by_id('txtPwd').send_keys(setupInfo.userPassword)
clickById('btnEfetuarLogin')
clickByXpath("//a[@title='Ir Para CheckList Isolando Este Status' and @onclick=\"IrParaCheckListStatus('8');\"]")

browser.switch_to.window(browser.window_handles[1])
waitTillPageLoads()
waitTillTableLoads()
clickByXpath("(//th/div[contains(@class,'th-inner') and text()='Nº ViPP'])[1]")

waitTillTableLoads()
waitBy(By.XPATH, "(//button[contains(@class,'btnedtobj')])[1]")

totalPedidos = len(pedidosInfo[idPedido])
for page in range(int(totalPedidos/10)):
    numElements = min(totalPedidos - (page * 10), 10)
    for index in range(numElements):
        waitTillTableLoads()
        waitBy(By.XPATH, "(//button[contains(@class,'btnedtobj')])[1]")
        clickByXpath(f"(//button[contains(@class,'btnedtobj')])[{index + 1}]")

        browser.switch_to.frame(browser.find_element_by_id('ifrEditarObjeto'))
        waitTillPageLoads()

        clienteNome = browser.find_element_by_id('txtNomeDestinatario').get_attribute("value")
        clienteIndex = pedidosInfo[nomeClientes].index(unidecode(clienteNome).upper())
        print(clienteIndex)
        waitBy(By.ID, "btnDeclaracaoConteudo")
        waitBy(By.ID, "btnDeclaracaoConteudo")
        clickById("btnDeclaracaoConteudo")
        try:
            waitBy(By.XPATH, "//iframe[@id='IFrmDecConteudo']")
        except:
            clickById("btnDeclaracaoConteudo")
        waitBy(By.XPATH, "//iframe[@id='IFrmDecConteudo']")
        browser.switch_to.frame(browser.find_element_by_id('IFrmDecConteudo'))

        waitTillPageLoads()
        waitBy(By.XPATH,
               "(//div[contains(@class,'ui-widget-content')]/div[contains(@class, 'CssPes')])[1]")
        for produtoIndex in range(len(pedidosInfo[listaConteudos][clienteIndex])):
            currentConteudos = pedidosInfo[listaConteudos][clienteIndex]
            produtoDescXpath = f"(//div[contains(@class,'ui-widget-content')]/div[1])[{produtoIndex + 1}]"
            produtoQtyXpath = f"(//div[contains(@class,'ui-widget-content')]/div[contains(@class, 'CssPes')])[{produtoIndex + 1}]"
            produtoValorXpath = f"(//div[contains(@class,'ui-widget-content')]/div[contains(@class, 'CssVlr')])[{produtoIndex + 1}]"

            fillInContentValue(produtoDescXpath,
                               currentConteudos[produtoIndex][INDICE_DESC_PRODUTO])
            fillInContentValue(produtoQtyXpath, currentConteudos[produtoIndex][INDICE_QTY_PRODUTO])
            fillInContentValue(produtoValorXpath,
                               currentConteudos[produtoIndex][INDICE_VALOR_PRODUTO])
            browser.find_element_by_xpath(produtoValorXpath + "/input").send_keys(Keys.ENTER)

        browser.switch_to.parent_frame()
        clickById('BtnSalvarDeclaracao')
        waitTillPageLoads()
        clickByXpath("//div[contains(@class, 'form-group')]/button[contains(@class, 'cmdGravar')]")
        browser.switch_to.default_content()

        for nextPageClick in range(page):
            waitTillPageLoads()
            waitTillTableLoads()
            clickByXpath("//div[@id='DivTabelaCheckList']//li[@class='page-next']/a")
            waitTillPageLoads()
            waitTillTableLoads()
            waitBy(By.XPATH, "(//button[contains(@class,'btnedtobj')])[1]")

browser.quit()

# btnDeclaracaoConteudo
# IFrmDecConteudo

#Declarar conteudo //a[@class='Editor btnEditarDeclaracaoConteudo']
#descrição //div[contains(@class,'ui-widget-content')]/div[1]
#Quantidade //div[contains(@class,'ui-widget-content')]/div[contains(@class, 'CssPes')]
#salvar //button[@id='BtnSalvarDeclaracao']
#importar: btnImportarArquivo

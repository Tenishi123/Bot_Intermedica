from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time

#Função criada para exercutar o login sempre q é redirecionado para um novo site
def login(drive): 
    
    #Encontra e aceita os coockies do site
    while True:
        try:
            driver.find_element(By.XPATH, '//div[@class="content__Content-ui__sc-85foch-0 LhAGP"]/div[@class="wrapper__Wrapper-ui__sc-17negc7-0 ixVRLj cookie-banner__Container-cmp__sc-1co6ie-2 euymWo"]/button[@class="button__Wrapper-ui__sc-a2a0dz-0 jusxHR"]').click()
            
        except:
            time.sleep(2)
        else:
            time.sleep(2)
            break

    #Encontra e clica no botão de login
    while True:
        try:
            driver.find_element(By.XPATH, '/html/body/div[2]/div[4]/div/nav/div[1]/div[3]/button').click()
        except:
            time.sleep(2)
        else:
            time.sleep(2)
            break

    email = ''
    senha = ''

    #Encontra e inseri o email para login
    while True:
        try:
            campoEmail = driver.find_element(By.XPATH, '//input[@id="lookup-email-input-id"]')
        except:
            time.sleep(2)
        else:
            for letra in email:
                campoEmail.send_keys(letra)
                #time.sleep(1)
            time.sleep(2)
            break

    #Encontra e clica no botão para aceitar o email
    while True:
        try:
            driver.find_element(By.XPATH, '//button[@type="submit"]').click()
        except:
            time.sleep(2)
        else:
            time.sleep(2)
            break

    #encontra e inseri a senha para login
    while True:
        try:
            campoSenha = driver.find_element(By.XPATH, '//input[@id="login-with-email-and-password-password-id"]')
        except:
            print('Erro ao preencher o dado senha no login')
            time.sleep(2)
        else:
            for letra in senha:
                campoSenha.send_keys(letra)
                #time.sleep(1)            
            time.sleep(2)
            break
            
    #encontra e clica no botão de login
    x = 10
    while x != 0:
        try:
            driver.find_element(By.XPATH, '//button[@type="submit"]').click()
        except:
            print('Erro ao clicar no botão para aceitar o senha')
            time.sleep(2)
            x -= 1
        else:
            time.sleep(2)
            break

#Inicia uma nova conexão
DRIVER_PATH = '/path/to/chromedriver'
options = Options()
#options.headless = True
options.add_argument("--window-size=1920,1200")
driver = webdriver.Chrome(options=options,executable_path=DRIVER_PATH)

driver.get('https://app.informamarkets.com.br/event/hospitalar/people/RXZlbnRWaWV3XzE3MzcwNQ==')
driver.maximize_window()

#executa a função de login
login(drive=driver)
        
#laço de repetição para gerar todos os links do site
x = 0
while x <= 1:
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(5)
    x+=1
    
#capturando a parte onde estão todos os links
repete = True
while repete:
    try:
        pag_principal = driver.find_element(By.XPATH, '//div[@class="infinite-scroll-component__outerdiv"]/div[@class="infinite-scroll-component "]')
    except:
        print("")
    else:
        repete = False
        
#gerando um array com todos os links
links_elements = pag_principal.find_elements(By.TAG_NAME, "a")
links = []
for links_elements_unidades in links_elements:
    links.append(links_elements_unidades.get_attribute('href'))
time.sleep(2)

links_tratados = []

for link in links:
    controle = True
    for link2 in links_tratados:
        if ((link2 + '?openConnection=true') == link) or ((link + '?openConnection=true') == link2):
            controle = False
    
    if controle:
        links_tratados.append(link)       

#executa um laço de repetição que ira passar por todos os links de login encontrados
for link in links_tratados:
    
    # Abre uma nova aba e vai para o site do SO
    driver.execute_script("window.open('" + link + "', '_blank')")
    
    #espera o novo carregamento da pagina
    time.sleep(5)

    #texto q sera enviado as pessoas
    string = "oi"
    
    driver.switch_to.window(driver.window_handles[1]) 
    
    #laço de repetição para encontrar e preencher o campo de texto q sera enviado
    x = 50
    while x >= 0:
        try:
            form = driver.find_element(By.XPATH, '//textarea[@class="textarea__BaseTextarea-ui__sc-1jr6d1d-4 hbngvE"]').send_keys(string)
        except:
            x -= 1
        else:
            time.sleep(5)
            break

    #laço de repetição para encontrar e clicar no botão para realizar o envio da menssagem
    x = 20
    while x >= 0:
        try:
            driver.find_element(By.XPATH, '//button[@type="submit"]').click()
        except:
            x -= 1
        else:
            time.sleep(5) 
            break
       
    driver.close() 
    driver.switch_to.window(driver.window_handles[0]) 



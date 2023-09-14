from time import sleep
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
import os
import glob

import shutil
import time
import datetime
root_dir = os.getcwd()

class Connect:
    def __init__(self) -> None:
        self.user = None
        self.download_dir = os.path.expanduser('~') + '\\Downloads'

    def get45(self, user, password):
        hoje = datetime.date.today()
        primeiro_dia_mes_atual = hoje.replace(day=1)
        ultimo_dia_mes_anterior = (primeiro_dia_mes_atual - datetime.timedelta(days=1)).strftime("%d%m%Y")
        self.user = user
        self.password = password


        URLSinqia = 'http://prdsinqia.ativainvestimentos.com.br:8180/AtivaPROD/servlet/SQCTRL;jsessionid=D1BLgTjmOx1EumlJOJ5uAGlpkRPhNkgzJ2MvL5A3.srvsinqiasp&jsessionid=ot2M9MxdGzmujqQNy-hG0IWd6ZB_Kj55HTb01PvB.srvsinqiasp&jsessionid=wM_nI7o7nPiBjwJQ1D7aHYmNamZuVpLutNQUqq8u.srvsinqiasp&jsessionid=7fIFg1QoKRJxtWmmsWte-PaGZkIxzApIBrtUVwmx.srvsinqiasp'
        try:
            service = Service(fr'{root_dir}\\Chromedriver\\chromedriver.exe')
            options = webdriver.ChromeOptions()
            #options.add_argument("--headless")
            options.add_argument("--start-minimized")
            driver = webdriver.Chrome(service=service, options=options)
            driver.switch_to.window(driver.current_window_handle)
        except:
            erro = 1
            print('Chrome error!')
            
        driver.get(URLSinqia)

        input_user = driver.find_element(by=By.XPATH, value='//*[@id="Usr"]')  # XPath do input de usuários - psette
        input_user.send_keys(self.user)
        inputPassword = driver.find_element(by=By.XPATH, value='//*[@id="Pwd"]')  # XPath do input de senha
        inputPassword.send_keys(self.password)
        driver.find_element("xpath", '//*[@id="btnSubmit"]').click() # clicando no botão "entrar" para acessar a Sinqia
        sleep(1)
        inputBuscaProcessos = driver.find_element(by=By.XPATH, value='//*[@id="txArgBusca"]')  # textbox busca de processos na Sinqia
        inputBuscaProcessos.send_keys("45") #enviando 224 para a textbox
        driver.find_element("xpath", '//*[@id="search-bar"]/table/tbody/tr/td[2]/a/img'  ).click() #clicando na lupa para a busca de processos

        # # # ***************************** Parametros da Carteira ***************************************************
        driver.find_element("xpath", '//*[@id="el_ip_lbList"]/a' ).click() #clicando em 'Selecionar por Lista'

        driver.find_element("xpath", '//*[@id="chk0"]' ).click() #clicando na checkbox Todas as Carteiras
        # driver.find_element("xpath", '//*[@id="chk0"]').click() 
        driver.find_element("xpath", '//*[@id="el_ip_cbList"]/span').click() #clicando na combobox de listas

        for _ in range(23):
            ActionChains(driver) \
                .send_keys(Keys.ARROW_UP) \
                .perform()

        ActionChains(driver) \
            .send_keys(Keys.ENTER) \
            .perform()

        driver.find_element(by=By.XPATH, value='//*[@id="el_cp_idtDataEmis_dateField"]').click() #input da data de Emissao
        inputData = driver.find_element(by=By.XPATH, value='//*[@id="el_cp_idtDataEmis_dateField"]') #input da data de Emissao

        sleep(2)
        ActionChains(driver)\
            .context_click(inputData)\
            .key_down(Keys.CONTROL)\
            .send_keys("a")\
            .key_up(Keys.CONTROL)\
            .perform()
        ActionChains(driver)\
            .send_keys(Keys.DELETE)\
            .perform()

        for num in ultimo_dia_mes_anterior:
            ActionChains(driver)\
                .send_keys(num)\
                .perform()
                
        driver.find_element("xpath", '//*[@id="el_irotOutputType_selROT"]').click() #caixa de seleção do formato de saída

        ActionChains(driver)\
            .send_keys(Keys.ARROW_DOWN)\
            .send_keys(Keys.ENTER)\
            .perform()
        driver.find_element("xpath", '//*[@id="chk11"]').click() #clicando na checkbox de Download
        driver.find_element("xpath", '/html/body/center/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr/td[3]/button').click() #clicando em 'OK'

        def arquivo_mais_recente(pasta):
            arquivos = glob.glob(os.path.join(pasta, '*'))
            arquivos.sort(key=os.path.getmtime, reverse=True)
            
            if arquivos:
                if ".crdownload" not in arquivos[0]:
                    return arquivos[0]
            else:
                return None

        #funçao para certificar de que será considerado o arquivo correto
        # o ultimo arquivo baixado precisa ter "Posição de Cotistas" e não ter
        # a extensão '.crdownload' no final. E também precisa ter sido baixado há menos
        # de 5 segundos atrás
        def atende_condicoes(arquivo):
            tempo_atual = time.time()
            cinco_segundos_atras = tempo_atual - 5
            if os.path.getmtime(arquivo) < cinco_segundos_atras:
                return False
            if "Posição de Cotistas" not in arquivo:
                return False
            if arquivo.endswith(".crdownload"):
                return False
            return True

        def verifica_diretorio_downloads(download_dir):
            while True:
                arquivo_mais_novo = arquivo_mais_recente(download_dir)
                #print(f"arquivo mais recente: {arquivo_mais_novo}")
                if arquivo_mais_novo and atende_condicoes(arquivo_mais_novo):
                    return arquivo_mais_novo
                sleep(1)
                    
        arquivo_45_pos_cot = verifica_diretorio_downloads(self.download_dir)
        return arquivo_45_pos_cot


# con = Connect()
# file = con.get45('psette', 'Ativa!@#2023')


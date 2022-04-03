import datetime
from lib2to3.pgen2 import driver
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
import time
import cx_Oracle
import keyboard
import numpy as np 
import pyautogui 



def enviamsg(Telefone,Cliente): 
    print('_______________________________________________________________________________________')
    try:
    ############# Parâmetros Chrome #############
        exec_path_driver = r'C:/chromedriver.exe'
        chrome_options = webdriver.ChromeOptions();
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--start-maximized')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_experimental_option('excludeSwitches', ['load-extension', 'enable-automation'])
        chrome_options.add_argument(r'user-data-dir=C:\Users\maycon.silva1\AppData\Local\Google\Chrome\User Data\Default')
        driver = webdriver.Chrome(executable_path = exec_path_driver,chrome_options=chrome_options)

        #ABRINDO LINK COM O FORMATO NECESSÁRIO.
        link = f"""https://web.whatsapp.com/send?phone={Telefone}&text=[Mensagem Automática]\n - GTFOODS GROUP - Olá, {str.strip(Cliente)}, Faturamos sua nota dentro de 24 horas será entregue, Produto a ser entregue: {str.strip(Produto)}. Endereço da entrega: {str.strip(Endereco)}."""
        driver.get(link)

        #ESPERA CARREGAR A PÁGINA PARA SEGUIR 
        while len(driver.find_elements_by_id("side")) < 1:
            time.sleep(1)

        #INFORMAÇÃO PARA SETAR A PÁGINA WHATSAPP
        pyautogui.click(500,500)
        keyboard.press_and_release('ENTER')
        keyboard.press_and_release('ENTER')
        pyautogui.click(500,500)
        time.sleep(10)
        keyboard.press_and_release('ENTER')
        pyautogui.click(500,500)
        keyboard.press_and_release('ENTER')
        

        #Reconhecimento de IMAGEM PARA IDENTIFICAR SE ENVIOU OU NÃO A IMAGEM.
        enviado = pyautogui.locateOnScreen('C:\\Users\\maycon.silva1\\Desktop\\Projeto 2.0 Whatsapp\\Log\\ENVIADO.png')
        whatsappnexite = pyautogui.locateOnScreen('C:\\Users\\maycon.silva1\\Desktop\\Projeto\\Log\\whatsappnexiste.png')
        chegou = pyautogui.locateOnScreen('C:\\Users\\maycon.silva1\\Desktop\\Projeto 2.0 Whatsapp\\Log\\CHEGOU.png')
        visto = pyautogui.locateOnScreen('C:\\Users\\maycon.silva1\\Desktop\\Projeto 2.0 Whatsapp\\Log\\VISTO.png')
        semnet = pyautogui.locateOnScreen('C:\\Users\\maycon.silva1\\Desktop\\Projeto 2.0 Whatsapp\\Log\\SEMNET.png')
        iniciandoconversa = pyautogui.locateOnScreen('C:\\Users\\maycon.silva1\\Desktop\\Projeto 2.0 Whatsapp\\Log\\INICIANDOCONVERSA.png')


        #VARIAVEIS CONTENDO OS STATUS A SER APRESENTADO
        arquivo = open('C:\\Users\\maycon.silva1\\Desktop\\Projeto 2.0 Whatsapp\\Logarq01.txt','a')
        var1 = 'STATUS 1 - CLIENTE/FORNECEDOR NÃO POSSUI WHATSAPP CADASTRADO CORRETAMENTE |'  + str.strip(Cliente) +' Telefone: '+ Telefone + ' incorreto !'
        var2 = 'STATUS 2 - MENSAGEM CHEGOU -> "' + str.strip(Cliente) +' Telefone: '+ Telefone + '!'
        var3 = 'STATUS 3 - MENSAGEM CHEGOU  E JÁ FOI VISUALIZADA |'  + str.strip(Cliente) +' Telefone: '+ Telefone + ' !'
        var4 = 'STATUS 4 - ESTAMOS SEM INTERNET - CONNECT A INTERNET DO CELULAR -> "' + str.strip(Cliente) +' Telefone: '+ Telefone + '!'
        var5 = 'STATUS 5 - FICOU INICIANDO CONVERSA E TRAVOU  " '+ str.strip(Cliente) +' Telefone: '+ Telefone + '!'
        var6 = 'STATUS 6 - Erro Indefinido Programador: @MayconMotta' +  str.strip(Cliente) +' Telefone: '+ Telefone + '!'

        if enviado:
            print("MENSAGEM ENVIADA COM SUCESSO!!!")
           
        elif whatsappnexite:
            arquivo.write(var1 + "\n")
            print (var1 + "\nLog Salvo:  " + arquivo.name + " | ") 
               
        elif chegou:
            arquivo.write(var2 + "\n")
            print (var2 + "\nLog Salvo:  " + arquivo.name + " | " ) 
            
        elif visto:
            arquivo.write(var3 + "\n")
            print (var3 + "\nLog Salvo:  " + arquivo.name + " | " ) 
                  
        elif semnet:
            arquivo.write(var4 + "\n")
            print (var4 + "\nLog Salvo:  " + arquivo.name + " | ") 

        elif iniciandoconversa:
            arquivo.write(var5 + "\n")
            print (var5 + "\nLog Salvo:  " + arquivo.name + " | " ) 
            
        else:
            arquivo.write(var6 + "\n")
            print (var6 + "\nLog Salvo: " + arquivo.name + " | " ) 
            
        keyboard.press_and_release('ctrl + w')
        keyboard.press_and_release('ctrl + w')
        time.sleep(10)
    
    except:

        print('Mensagem não Enviada Telefone: ' + str.strip(Telefone) + ' CLIENTE: '+ str.strip(Cliente) + ' |ERROR INDEFINIDO| - COLOQUE PARA RODAR NOVAMENTE QUANDO TERMINAR|;')
            
# -------------------------------------------------
#   REALIZANDO CONEXÃO COM O BANCO DE DADOS PROTHEUS / INTEGRAÇÃO / CX_ORACLE
# -------------------------------------------------


dsn_tns = cx_Oracle.makedsn('host','port',service_name='id') # if needed, place an 'r' before any parameter in order to address special characters such as '\'.
conexão = cx_Oracle.connect(user=r'user', password='pass', dsn=dsn_tns) # if needed, place an 'r' before any parameter in order to address special characters such as '\'. For example, if your user name contains '\', you'll need to place 'r' before the user name: user=r'User Name'

consulta = conexão.cursor()

consulta.execute("""SELECT DISTINCT
    TRIM(REGEXP_REPLACE(concat('+55'|| SA1.A1_DDD,SA1.A1_TEL),'\s*', '')) AS "DDD + TELEFONE", -- campos[0]
    SA1.A1_END AS ENDERECO, ------------------------------------------------------------------- campos[1]
    SA1.A1_NREDUZ AS NOME, -------------------------------------------------------------------- campos[2]
    TRIM(B1_DESC) AS DESCRICAO_02, ------------------------------------------------------------ campo [3]
    PAD.PAD_DATA
FROM
  TMPROD.PAD010 PAD
LEFT JOIN
      TMPROD.pa1010 pa1 ON
      PAD.PAD_LOTEAV = pa1.pa1_xltav
LEFT JOIN TMPROD.SA1010 SA1 ON
     SA1.A1_COD = PA1.PA1_FORNEC
INNER JOIN
      TMPROD.SB1010 SB1 ON
      SB1.B1_FILIAL = SB1.B1_FILIAL AND
      B1_COD = PAD.PAD_PRODU AND
      SB1.D_E_L_E_T_ = ' '
     
WHERE PAD.D_E_L_E_T_ = ' ' 
    AND TO_DATE (PAD.PAD_DATA, 'YYYYMMDD') = TO_DATE(SYSDATE)
    AND TRIM(PA1.PA1_DTFECH) IS NULL 
    AND PAD.PAD_PRODU IN ('SUBRACMI000031','SUBRACMI000004','SUBRACMI000001','SUBRACMI000002','SUBRACMI000028','SUBRACMI000005','SUBRACMI000003','SUBRACMI000017','MCINS000116','SUBRACMI000007','SUBRACMI000032','SUBRACMI000006','SUBRACMI000008','SUBRACMI000016','MCINS000019')
    AND TRIM(PA1.PA1_DTFECH) IS NULL""")

retorno = consulta.fetchall()

for campos in (retorno):      
   Telefone = '+5544998465693'#campos[0] # Primeiro campo da consulta SQL
   Endereco = campos[1] # Segundo campo da consulta SQL
   Cliente =  campos[2]
   Produto = campos [3] # Terceiro campo da consulta SQL
    
   if enviamsg(Telefone,Cliente): # Valida se o site do Whastapp está online
        print(retorno)# Envia a mensagem





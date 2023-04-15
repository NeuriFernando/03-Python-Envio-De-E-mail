# Envio de Menssagem via E-mail (Outlook)

### Importar Biblioteca

```import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import pyautogui
import time
import pyperclip
```

### Passo 1: buscar a base de dados para ser usada

```
contatos_df = pd.read_excel("Emails.xlsx")
display(contatos_df)
```
### Passo 2: transformar em lista

```
dados = contatos_df[['Chave','Email','Modelo' ]].values.tolist()
print (dados)
```

### Passo 3: definir um corpo para o e-mail

```
corpo_mensagem = """
Descrever aqui o corpo do e-mail
"""
```
### Passo 04: Abrir navegador

```
navegador = webdriver.Chrome()
navegador.maximize_window()
navegador.implicitly_wait(30)
navegador.set_page_load_timeout(220)
navegador.get("https://outlook.office.com")
```
### Passo 5: Inserir email e senha para acessar a aplicação

```
navegador.find_element('xpath', '//*[@id="i0116"]').send_keys("Seu email de acesso")
navegador.find_element('xpath', '//*[@id="i0116"]').send_keys(Keys.ENTER)
navegador.find_element('xpath', '//*[@id="passwordInput"]').send_keys("Sua senha de acesso")
navegador.find_element('xpath', '//*[@id="passwordInput"]').send_keys(Keys.ENTER)
navegador.find_element('xpath', '//*[@id="idSIButton9"]').send_keys(Keys.ENTER)
```

### Passo 6: Criando looping com base nos dados imputados

```
for dado in dados:
    navegador.find_element('xpath', '//*[@id="innerRibbonContainer"]/div[1]/div/div/div/div[2]/div/div/span/button[1]').click()
    time.sleep(2)
    navegador.find_element('xpath', '/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div[3]/div/div[3]/div[3]/div[1]/div/div/div/div[1]/div[1]/div/div[4]/div/div/div[1]').send_keys(dado)
    navegador.find_element('xpath', '/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div[3]/div/div[3]/div[3]/div[1]/div/div/div/div[1]/div[2]/div[2]/div/div/div/input').send_keys("Agendamento para Bitlocker - LUBRAX")
    navegador.find_element('xpath', '/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div[3]/div/div[3]/div[3]/div[1]/div/div/div/div[2]/div/div/div/div[1]').send_keys(corpo_mensagem)
    navegador.find_element('xpath', '/html/body/div[2]/div/div[2]/div[2]/div[2]/div/div/div/div[3]/div/div[3]/div[3]/div[1]/div/div/div/div[3]/div[3]/div[1]/div/div/span/button[1]').click()
    time.sleep(4)
```

import openpyxl
import pyperclip
import pyautogui
from time import sleep
# Entra na planilha 
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
#Copiar informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row = 2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(123,174,duration = 1)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(123,260,duration = 1)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(123,358,duration = 1)
    pyautogui.hotkey('ctrl','v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(123,444,duration = 1)
    pyautogui.hotkey('ctrl','v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(123,531,duration = 1)
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(123,618,duration = 1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(150,717,duration = 1)
    sleep(5)


    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(124,202,duration = 1)
    pyautogui.hotkey('ctrl','v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(124,292,duration = 1)
    pyautogui.hotkey('ctrl','v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(124,376,duration = 1)
    pyautogui.hotkey('ctrl','v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(124,464,duration = 1)
    pyautogui.hotkey('ctrl','v')

    tamanho = linha[10].value
    pyautogui.click(204,550,duration=1)
    
    if tamanho == 'Pequeno':
        pyautogui.click(261,580,duration=2)
    elif tamanho == 'Médio':
        pyautogui.click(261,604,duration=2)
    else:
        pyautogui.click(261,626,duration=2)
        

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(124,635,duration = 1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(154,695,duration = 1)
    sleep(5)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(124,224,duration = 1)
    pyautogui.hotkey('ctrl','v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(124,311,duration = 1)
    pyautogui.hotkey('ctrl','v')

    observacao = linha[14].value
    pyperclip.copy(observacao)
    pyautogui.click(124,397,duration = 1)
    pyautogui.hotkey('ctrl','v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(124,530,duration = 1)
    pyautogui.hotkey('ctrl','v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(material)
    pyautogui.click(124,617,duration = 1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(159,673,duration = 1)
    sleep(3)

    pyautogui.click(851,186,duration = 2)
    sleep(3)

    pyautogui.click(673,448,duration = 2)
    sleep(3)
   
#Repetir esses passos para outros campos até preencher campos daquela página
#Clicar em próxima
#Repetir os mesmos passos e ir para próxima página(página 2)
#Repetir os mesmo passos e finalizar o cadastro daquele produto e clicar em #concluir
#Clicar em ok, para finalizar processo
#Clicar no ok mais uma vez na mensagem de confirmação de salvamento no #banco de dados 
#CLICAR EM “dicionar mais um e repetir o processo até finalizar a planilha”

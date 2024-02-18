import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('Produtos_Ficticios.xlsx')
pagina_produtos= workbook['Planilha1']

for linha in pagina_produtos.iter_rows(min_row=3):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1119,279 , duration=1)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1114,363 , duration=1)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1123,500 , duration=1)
    pyautogui.hotkey('ctrl','v')
    
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1121,585 , duration=1)
    pyautogui.hotkey('ctrl','v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1125,673 , duration=1)
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1119,755 , duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1132,822, duration=1)
    sleep(1)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1117,311 , duration=1)
    pyautogui.hotkey('ctrl','v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(1102,392 , duration=1)
    pyautogui.hotkey('ctrl','v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(1099,477 , duration=1)
    pyautogui.hotkey('ctrl','v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1099,564 , duration=1)
    pyautogui.hotkey('ctrl','v')

    tamanho = linha[10].value
    pyautogui.click(1118,649)

    if tamanho == 'Pequeno':
        pyautogui.click(1159,675, duration=1)
        

    elif tamanho == 'Médio':
        pyautogui.click(1152,699, duration=1)

    else:
        pyautogui.click(1147,723, duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1108,731, duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1127,792, duration=1)
    sleep(1)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1116,317, duration=1)
    pyautogui.hotkey('ctrl','v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1110,410, duration=1)
    pyautogui.hotkey('ctrl','v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1107,490, duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1106,626, duration=1)
    pyautogui.hotkey('ctrl','v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1102,708, duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1123,766, duration=1)
    pyautogui.click(1510,191, duration=1)
    pyautogui.click(1337,549, duration=1)



   
   
   



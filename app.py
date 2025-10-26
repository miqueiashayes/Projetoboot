import openpyxl
import pyperclip
import pyautogui
import time

workbook = openpyxl.load_workbook('clientes.xlsx')
sheet_produtos = workbook['Produtos']
pyautogui.PAUSE = 0.5

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(str(nome_produto))
    pyautogui.click(1518, 305, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    descricao = linha[1].value
    pyperclip.copy(str(descricao))
    pyautogui.click(1518, 345, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    categoria = linha[2].value
    pyperclip.copy(str(categoria))
    pyautogui.click(1518, 385, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    codigo_produto = linha[3].value
    pyperclip.copy(str(codigo_produto))
    pyautogui.click(1518, 425, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    peso = linha[4].value
    pyperclip.copy(str(peso))
    pyautogui.click(1518, 465, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    dimensoes = linha[5].value
    pyperclip.copy(str(dimensoes))
    pyautogui.click(1518, 505, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    preco = linha[6].value
    pyperclip.copy(str(preco))
    pyautogui.click(1518, 545, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(str(quantidade_em_estoque))
    pyautogui.click(1518, 585, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    data_de_validade = linha[8].value
    pyperclip.copy(str(data_de_validade))
    pyautogui.click(1518, 625, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    cor = linha[9].value
    pyperclip.copy(str(cor))
    pyautogui.click(1518, 665, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    tamanho = linha[10].value
    pyperclip.copy(str(tamanho))
    pyautogui.click(1518, 705, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    material = linha[11].value
    pyperclip.copy(str(material))
    pyautogui.click(1518, 745, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    fabricante = linha[12].value
    pyperclip.copy(str(fabricante))
    pyautogui.click(1518, 785, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    pais_de_origem = linha[13].value
    pyperclip.copy(str(pais_de_origem))
    pyautogui.click(1518, 825, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    observacoes = linha[14].value
    pyperclip.copy(str(observacoes))
    pyautogui.click(1518, 865, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    codigo_de_barras = linha[15].value
    pyperclip.copy(str(codigo_de_barras))
    pyautogui.click(1518, 905, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    localizacao_armazem = linha[16].value
    pyperclip.copy(str(localizacao_armazem))
    pyautogui.click(1518, 945, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    time.sleep(1)


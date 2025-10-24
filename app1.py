import openpyxl
import pyperclip
import pyautogui
import time

# Carregar planilha
workbook = openpyxl.load_workbook('clientes.xlsx')
sheet_produtos = workbook['Produtos']

# Loop nas linhas
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    descricao = linha[1].value
    categoria = linha[2].value
    codigo_produto = linha[3].value
    preco = linha[6].value
    quantidade_em_estoque = linha[7].value
    data_de_validade = linha[8].value

    # --- Exemplo de automação preenchendo os campos ---

    # Nome do Produto
    pyperclip.copy(nome_produto)
    pyautogui.click(1518, 305, duration=1) # posição exata do campo
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)

    # Descrição
    pyperclip.copy(descricao)
    pyautogui.click(1520, 350, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)

    # Categoria
    pyperclip.copy(categoria)
    pyautogui.click(1520, 390, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)

    # Código do Produto
    pyperclip.copy(codigo_produto)
    pyautogui.click(1520, 430, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    pyperclip.copy(str(preco))
    pyautogui.click(1520, 470, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    pyperclip.copy(str(quantidade_em_estoque))
    pyautogui.click(1520, 510, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    if data_de_validade:
        pyperclip.copy(data_de_validade.strftime('%d/%m/%Y'))
    else:
        pyperclip.copy("")
    pyautogui.click(1520, 550, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    time.sleep(1) 

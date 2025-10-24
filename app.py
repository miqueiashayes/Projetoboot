import openpyxl
import pyperclip
import pyautogui

#Ler planilha e guardar informações sobre o nome, telefone e data de vencimento 



workbook = openpyxl.load_workbook('clientes.xlsx')
sheet_produtos = workbook['Produtos']
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1518, 305, duretion=1)
    pyautogui.hotkey('ctrl', 'v')
    descricao = linha[1].value
    categoria = linha[2].value
    codigo_produto = linha[3].value
    peso = linha[4].value
    dimensoes = linha[5].value
    preco = linha[6].value
    quantidade_em_estoque = linha[7].value
    data_de_validade = linha[8].value
    cor = linha[9].value
    tamanho = linha[10].value
    material = linha[11].value
    fabricante = linha[12].value
    pais_de_origem = linha[13].value
    observacoes = linha[14].value
    codigo_de_barras = linha[15].value
    localizacao_armazem = linha[16].value 


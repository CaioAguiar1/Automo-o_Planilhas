import pyautogui
import openpyxl
import pyperclip

# 2 - abrir a planilha
workbook = openpyxl.load_workbook(
    r'C:\Users\Caio\Desktop\Caio\automaçãoPython\produtos.xlsx')
sheet_produtos = workbook['produtos']
for linha in sheet_produtos.iter_rows(min_row=2, max_row=501):
    produto = linha[0].value
    fornecedor = linha[1].value
    categoria = linha[2].value
    quantidade = linha[3].value
    valor_unitario = linha[4].value
    notificar_venda = linha[5].value
    # colar dados campo produto
    pyautogui.click(877, 353, duration=1)
    pyautogui.write(produto)
    # colar dados campo fornecedor
    pyautogui.click(1130, 353, duration=1)
    pyautogui.write(fornecedor)
    # colar dados campo categoria
    pyautogui.click(877, 437, duration=1)
    pyperclip.copy(categoria)
    pyautogui.hotkey('ctrl', 'v')
    # colar dados campo valor unitário
    pyautogui.click(1095, 427, duration=1)
    pyperclip.copy(valor_unitario)
    pyautogui.hotkey('ctrl', 'v')
    # se notificar venda for igual a sim, marcar sim
    # se notificar venda for igual a não, marcar não
    if notificar_venda == "Sim":
        pyautogui.click(842, 507, duration=1)
    elif notificar_venda == "Não":
        pyautogui.click(919, 503, duration=1)
    # clicar em registrar produto
    pyautogui.click(914, 563, duration=1)
    # clicar em ok, na mensagem de cadastro com sucesso
    pyautogui.click(1219, 165, duration=1)

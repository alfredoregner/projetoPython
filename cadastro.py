import pyautogui
import time
import openpyxl
import pyperclip

# Formulário para cadastro dos funcionários, simulando a utilização de um sistema
# https://docs.google.com/forms/d/e/1FAIpQLSduehX0b4pdnbk1PBPxqWrynw4ADLntbk9uadJ-FMBPR3qPhg/viewform

# Declaração de abertura do documento e planilha correta do sistema
workbook = openpyxl.load_workbook('Funcionarios.xlsx')
pagina = workbook['Funcionarios']

# Acessando a linha 2 da planilha escolhida, para que comece a navegar entre as células, coletando as informações preenchidas
for linha in pagina.iter_rows(min_row = 2):     # Coleta a informação a partir da segunda linha da tabela (a primeira não é necessária nesse caso pois é o cabeçalho do arquivo)
    # Nome
    funcionario_nome = linha[0].value       # Acessa o campo selecionado de acordo com o índice ([0] é o primeiro campo da linha 2)
    pyperclip.copy(funcionario_nome)        # Copia a informação do campo selecionado, com a formatação original
    pyautogui.click(735,444, duration = 1)  # Clica no local indicado pelas coordenadas
    pyautogui.hotkey('ctrl', 'v')           # Cola a informação copiada anteriormente, no local selecionado
    
    # CPF
    funcionario_cpf = linha[1].value
    pyperclip.copy(funcionario_cpf)
    pyautogui.press('tab', interval = 0.5)  # Troca para o próximo campo
    pyautogui.hotkey('ctrl', 'v')

    # E-mail
    funcionario_email = linha[2].value
    pyautogui.press('tab', interval = 0.5)
    pyperclip.copy(funcionario_email)
    pyautogui.hotkey('ctrl', 'v')

    # Telefone
    funcionario_telefone = linha[3].value
    pyperclip.copy(funcionario_telefone)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

    # btn-Próximo
    pyautogui.press('tab', interval = 1)
    pyautogui.press('enter', interval = 1)  # Pressiona Enter para seguir para a próxima parte do formulário
    time.sleep(2)

    # CEP
    funcionario_cep = linha[4].value
    pyperclip.copy(funcionario_cep)
    pyautogui.click(783,471, duration = 1)
    pyautogui.hotkey('ctrl', 'v')

    # Rua
    funcionario_rua = linha[5].value
    pyperclip.copy(funcionario_rua)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

    # Número
    funcionario_numero = linha[6].value
    pyperclip.copy(funcionario_numero)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

    # Complemento
    funcionario_complemento = linha[7].value
    pyperclip.copy(funcionario_complemento)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

    # Bairro
    funcionario_bairro = linha[8].value
    pyperclip.copy(funcionario_bairro)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

    # Cidade
    funcionario_cidade = linha[9].value
    pyperclip.copy(funcionario_cidade)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

    # Estado
    
    funcionario_estado = linha[10].value
    pyperclip.copy(funcionario_estado)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

    # btn-Próximo
    pyautogui.press('tab', interval = 1)
    pyautogui.press('tab', interval = 1)
    pyautogui.press('enter', interval = 1)
    time.sleep(2)

    # Cargo
    funcionario_cargo = linha[11].value
    if funcionario_cargo == 'Gerente':         # Verifica o que está escrito na célular e compara com as condições para clicar na opção correta
        pyautogui.click(697,463, duration = 1)
    elif funcionario_cargo == 'Supervisor':
        pyautogui.click(696,503, duration = 1)
    elif funcionario_cargo == 'Repositor':
        pyautogui.click(695,545, duration = 1)
    else:
        pyautogui.click(694,583, duration = 1)
        
    # Salário
    funcionario_salario = linha[12].value
    pyperclip.copy(funcionario_salario)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

    # btn-Enviar
    pyautogui.press('tab', interval = 1)
    pyautogui.press('tab', interval = 1)
    pyautogui.press('enter', interval = 1)
    time.sleep(2)

    # Enviara outra resposta
    pyautogui.click(741,273, duration = 1)
    time.sleep(2)






# 
# pyautogui.click(735,444, duration = 1)

# 
# pyautogui.press('tab', duration = 0.1)

# 
# pyautogui.press('tab', duration = 0.1)

# 
# pyautogui.press('tab', duration = 0.1)


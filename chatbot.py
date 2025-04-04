import time
import random
import openpyxl 

print('Bem vindo ao Suporte da Maculados LTDA. Eu serei sua assistente virtual, estou pronta para te ajudar!')
print('Digite o número correspondente ao atendimento que deseja: ')
print('1- Cadastro de Funcionário \n2- Edição de Cadastro \n3- Encerrar atendimento')
atendimento = int(input('Insira sua opção: '))

planilha = openpyxl.load_workbook('Funcionarios.xlsx')
pagina = planilha['Funcionarios']

while atendimento < 1 and atendimento > 3:
    print('Opção inválida!')
    print('1- Cadastro de Funcionário \n2- Edição de Cadastro \n3- Encerrar atendimento')
    atendimento = int(input('Insira sua opção: '))

if atendimento == 1:
    pass





elif atendimento == 2:
    workbook = openpyxl.load_workbook('Funcionarios.xlsx')
    pagina = workbook['Funcionarios']

else :
    print('Agradecemos pelo seu contato. Tenha um ótimo dia!')



# nome = input("Olá, qual é o seu nome? ")

# time.sleep(2)
# print("Olá " + nome)

# feeling = input("Como está se sentindo hoje? ")

# time.sleep(2)
# if "Bem" in feeling:
#     print("Eu também estou bem :) ")
# else:
#     print("Sinto muito :( ")

# time.sleep(2)
# cor = input("Qual é sua cor favorita? ")

# colours = ["Vermelho", "Verde", "Azul", "Amarelo", "Rosa", "Roxo", "Preto", "Branco"]

# time.sleep(2)
# print("Minha cor favorita é: " + random.choice(colours))
import time
import random
import openpyxl 

print('Bem vindo ao Suporte da Maculados LTDA. Eu serei sua assistente virtual, estou pronta para te ajudar!')
print('Digite o número correspondente ao atendimento que deseja: ')
print('1- Cadastro de Funcionário \n2- Edição de Cadastro \n3- Encerrar atendimento')
atendimento = int(input('Insira sua opção: '))

# Acessa a planilha com os dados dos funcionários
planilha = openpyxl.load_workbook('Funcionarios.xlsx')
pagina = planilha['Funcionarios']

while atendimento < 1 and atendimento > 3:
    print('Opção inválida!')
    print('1- Cadastro de Funcionário \n2- Edição de Cadastro \n3- Encerrar atendimento')
    atendimento = int(input('Insira sua opção: '))

if atendimento == 1:
    # Coleta os dados que serão inseridos na planilha
    cadasto_nome = input('Insira o nome do funcionário: ')
    cadasto_cpf = input('Insira o CPF do funcionário: ')
    cadasto_email = input('Insira o E-mail do funcionário: ')
    cadasto_telefone = input('Insira o Telefone do funcionário: ')
    cadasto_cep = input('Insira o CEP do endereço do funcionário: ')
    cadasto_rua = input('Insira a Rua do endereço do funcionário: ')
    cadasto_numero = input('Insira o Número do endereço do funcionário: ')
    cadasto_complemento = input('Insira o Complemento do endereço do funcionário: ')
    cadasto_bairro = input('Insira o Bairro do endereço do funcionário: ')
    cadasto_cidade = input('Insira a Cidade do endereço do funcionário: ')
    cadasto_estado = input('Insira o Estado do endereço do funcionário: ')
    cadasto_cargo = input('Insira o Cargo do funcionário: ')
    cadasto_salario = input('Insira o Salário do funcionário: ')

    # Insere os dados na planilha
    pagina.append([cadasto_nome, cadasto_cpf, cadasto_email, cadasto_telefone, cadasto_cep, cadasto_rua, cadasto_numero, cadasto_complemento, cadasto_bairro, cadasto_cidade, cadasto_estado, cadasto_cargo, cadasto_salario])
    planilha.save('Funcionarios.xlsx') # Salva a planilha com os dados atualizados


elif atendimento == 2:
    # Inserir o CPF do cadastro desejado para alteração
    # Verificar se o CPF é valido, e acessar a linha de cadastro desse funcionário
    # Realizar a alteração de todos os campos de cadastro do usuário (ou talvez separar por tópico: Pessoal, Endereço, Função)
    for rows in pagina.iter_rows(min_row=2, max_row=None):
    # Coleta os dados que serão inseridos na planilha
        edicao_nome = input('Insira o nome do funcionário: ')
        edicao_cpf = input('Insira o nome do funcionário: ')
        edicao_email = input('Insira o nome do funcionário: ')
        edicao_telefone = input('Insira o nome do funcionário: ')
        edicao_cep = input('Insira o nome do funcionário: ')
        edicao_rua = input('Insira o nome do funcionário: ')
        edicao_numero = input('Insira o nome do funcionário: ')
        edicao_complemento = input('Insira o nome do funcionário: ')
        edicao_bairro = input('Insira o nome do funcionário: ')
        edicao_cidade = input('Insira o nome do funcionário: ')
        edicao_estado = input('Insira o nome do funcionário: ')
        edicao_cargo = input('Insira o nome do funcionário: ')
        edicao_salario = input('Insira o nome do funcionário: ')

    # Insere os dados na planilha
    pagina.append([edicao_nome, edicao_cpf, edicao_email, edicao_telefone, edicao_cep, edicao_rua, edicao_numero, edicao_complemento, edicao_bairro, edicao_cidade, edicao_estado, edicao_cargo, edicao_salario])
    planilha.save('Funcionarios.xlsx') # Salva a planilha com os dados atualizados
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
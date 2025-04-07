import time
import random
import openpyxl 

print('Bem vindo ao Suporte da Maculados LTDA. Eu serei sua assistente virtual, estou pronta para te ajudar!')
print('Digite o número correspondente ao atendimento que deseja: ')
print('1- Cadastro de Funcionário \n2- Edição de Cadastro \n3- Encerrar atendimento')
atendimento = int(input('Insira sua opção: '))
time.sleep(1)

# Acessa a planilha com os dados dos funcionários
planilha = openpyxl.load_workbook('Funcionarios.xlsx')
pagina = planilha['Funcionarios']

while atendimento < 1 and atendimento > 3:
    print('Opção inválida!')
    print('1- Cadastro de Funcionário \n2- Edição de Cadastro \n3- Encerrar atendimento')
    atendimento = int(input('Insira sua opção: '))
    time.sleep(1)

if atendimento == 1:
    # Coleta os dados que serão inseridos na planilha
    cadastro_nome = input('Insira o nome do funcionário: ')
    time.sleep(0.5)
    cadastro_cpf = input('Insira o CPF do funcionário: ')
    time.sleep(0.5)
    cadastro_email = input('Insira o E-mail do funcionário: ')
    time.sleep(0.5)
    cadastro_telefone = input('Insira o Telefone do funcionário: ')
    time.sleep(0.5)
    cadastro_cep = input('Insira o CEP do endereço do funcionário: ')
    time.sleep(0.5)
    cadastro_rua = input('Insira a Rua do endereço do funcionário: ')
    time.sleep(0.5)
    cadastro_numero = input('Insira o Número do endereço do funcionário: ')
    time.sleep(0.5)
    cadastro_complemento = input('Insira o Complemento do endereço do funcionário: ')
    time.sleep(0.5)
    cadastro_bairro = input('Insira o Bairro do endereço do funcionário: ')
    time.sleep(0.5)
    cadastro_cidade = input('Insira a Cidade do endereço do funcionário: ')
    time.sleep(0.5)
    cadastro_estado = input('Insira o Estado do endereço do funcionário: ')
    time.sleep(0.5)
    cadastro_cargo = input('Insira o Cargo do funcionário: ')
    time.sleep(0.5)
    cadastro_salario = input('Insira o Salário do funcionário: ')
    time.sleep(0.5)

    # Insere os dados na planilha
    pagina.append([cadastro_nome, cadastro_cpf, cadastro_email, cadastro_telefone, cadastro_cep, cadastro_rua, cadastro_numero, cadastro_complemento, cadastro_bairro, cadastro_cidade, cadastro_estado, cadastro_cargo, cadastro_salario])
    planilha.save('Funcionarios.xlsx') # Salva a planilha com os dados atualizados

    print(f'Dados do funcionário {cadastro_nome} cadastrados com sucesso!')


elif atendimento == 2:
    # Inserir o CPF do cadastro desejado para alteração
    cpf_usuario = input('Insira o CPF do cadastro que deseja alterar: ')

    
    for rows in pagina.iter_rows(min_row=2, max_row=None):
        # Verificar se o CPF é valido, e acessar a linha de cadastro desse funcionário
        if cpf_usuario == rows[1].value: 
            print('O que você deseja alterar?')
            print('1- Dados Básicos do Funcionário \n2- Endereço do Funcionário \n3- Função do Funcionário')
            editar = int(input('Insira sua opção: '))
            
            while editar < 1 and editar > 3: 
                print('Opção inválida!')
                editar = input('1- Dados Básicos do Funcionário \n2- Endereço do Funcionário \n3- Função do Funcionário')
                editar = int(input('Insira sua opção: '))

            # Realizar a alteração de todos os campos de dados básicos do funcionário
            if editar == 1:
                edicao_nome = input('Insira o nome do funcionário: ')
                time.sleep(0.5)
                edicao_cpf = input('Insira o nome do funcionário: ')
                time.sleep(0.5)
                edicao_email = input('Insira o nome do funcionário: ')
                time.sleep(0.5)
                edicao_telefone = input('Insira o nome do funcionário: ')
                time.sleep(0.5)

                # Insere os dados na planilha
                rows[0] = edicao_nome
                rows[1] = edicao_cpf
                rows[2] = edicao_email
                rows[3] = edicao_telefone
                planilha.save('Funcionarios.xlsx') # Salva a planilha com os dados atualizados

                print(f'Dados básicos do funcionário alterados com sucesso!')

            # Realizar a alteração de todos os campos de endereço do funcionário
            elif editar == 2:
                edicao_cep = input('Insira o nome do funcionário: ')
                time.sleep(0.5)
                edicao_rua = input('Insira o nome do funcionário: ')
                time.sleep(0.5)
                edicao_numero = input('Insira o nome do funcionário: ')
                time.sleep(0.5)
                edicao_complemento = input('Insira o nome do funcionário: ')
                time.sleep(0.5)
                edicao_bairro = input('Insira o nome do funcionário: ')
                time.sleep(0.5)
                edicao_cidade = input('Insira o nome do funcionário: ')
                time.sleep(0.5)
                edicao_estado = input('Insira o nome do funcionário: ')
                time.sleep(0.5)

                # Insere os dados na planilha
                rows[4] = edicao_cep
                rows[5] = edicao_rua
                rows[6] = edicao_numero
                rows[7] = edicao_complemento
                rows[8] = edicao_bairro
                rows[9] = edicao_cidade
                rows[10] = edicao_estado
                planilha.save('Funcionarios.xlsx') # Salva a planilha com os dados atualizados

                print(f'Dados de endereço do funcionário alterados com sucesso!')

            # Realizar a alteração de todos os campos de função do funcionário
            else:                    
                edicao_cargo = input('Insira o nome do funcionário: ')
                time.sleep(0.5)
                edicao_salario = input('Insira o nome do funcionário: ')
                time.sleep(0.5)

                # Insere os dados na planilha
                rows[11] = edicao_cargo
                rows[12] = edicao_salario
                planilha.save('Funcionarios.xlsx') # Salva a planilha com os dados atualizados

                print(f'Dados de função do funcionário alterados com sucesso!')
            
        else:
            print('CPF Inválido')
            time.sleep(1)

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
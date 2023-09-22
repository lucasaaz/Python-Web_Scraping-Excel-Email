import pandas as pd
import numpy as np
from datetime import date
import smtplib
import email.message

# Funçao na qual ira pegar os valores dentro da coluna Data Nascimento e formatar ela, deixando todas iguais
def formartarData():
    # Abrir excel e colocar na var tabela
    tabela = pd.read_excel('teste1.xlsx')

    # Lendo dados da coluna Data Nascimento
    dataNascimento = list(tabela.loc[tabela['Nº']>0, 'Data  de Nascimento'])

    # Fazer um for para pegar as datas que estao fora do padrao e coloca-las no padrao
    num = 1
    for i in dataNascimento:
        # Ver se i eh uma string
        resp = isinstance(i, str)
        if resp != True:
            i = i.strftime('%d/%m/%Y')
            # Alterar no execel com dados novos
            tabela.loc[tabela['Nº']==num, 'Data  de Nascimento'] = i
            # print(f'{i}')
        num = num + 1

    # Salvar os dados no excel
    tabela.to_excel('teste2.xlsx', index=False)  

# Pegar apenas o ano de Nascimento
def pegarAnoNasc():
    tabela = pd.read_excel('teste2.xlsx')

    global dataNascimento
    dataNascimento = list(tabela.loc[tabela['Nº']>0, 'Data  de Nascimento'])

    global anoNascimento
    anoNascimento = []
    for i in dataNascimento:
        i = i[6:]
        anoNascimento.append(i)



# -----------------------------------//----------------------------------------------



# Funçao na qual ira pegar os valores dentro da coluna E-mail e formatar ela, dando aos e-mails vazios um valor NAO EMAIL
def formartarEmail():
    tabela = pd.read_excel('teste1.xlsx')

    email = list(tabela.loc[tabela['Nº']>0, 'E-mail'])

    num = 1
    for i in email:
        resp = isinstance(i, float)
        if resp == True:
            i = str('NAO EMAIL')
            tabela.loc[tabela['Nº']==num, 'E-mail'] = i
        num = num + 1

    tabela.to_excel('teste2.xlsx', index=False) 

# Funcao para pegar apenas os email existente, ou pegar os email completos formatados 
def pegarEmail():
    tabela = pd.read_excel('teste2.xlsx')

    global email2
    email2 = list(tabela.loc[tabela['Nº']>0, 'E-mail'])

    global email3
    email3 = []
    for i in email2:
        if 'NAO EMAIL' not in i:
            email3.append(i)



# -----------------------------------//----------------------------------------------



# Funçao na qual ira pegar os valores dentro da coluna Tel Resp e formatar ela, dando aos Tel Resp vazios um valor NAO TEL e tirar carcteres especiais
def formartarTelResp():
    tabela = pd.read_excel('teste1.xlsx')

    telResp = list(tabela.loc[tabela['Nº']>0, 'Tele. Celular Resp.'])

    num = 1
    for i in telResp:
        resp = isinstance(i, str)
        if resp == True:
            i = i.replace('(','')
            i = i.replace(')',' ')
            i = i.replace('-','')
            tabela.loc[tabela['Nº']==num, 'Tele. Celular Resp.'] = i
        else:
            i = str('NAO TEL')
            tabela.loc[tabela['Nº']==num, 'Tele. Celular Resp.'] = i
        num = num + 1

    tabela.to_excel('teste2.xlsx', index=False) 

# Funcao para pegar apenas os Tel Resp existente, ou pegar os Tel Resp completos formatados
def pegarTelResp():
    tabela = pd.read_excel('teste2.xlsx')

    global TelResp
    TelResp = list(tabela.loc[tabela['Nº']>0, 'Tele. Celular Resp.'])

    global TelResp3
    TelResp3 = []
    for i in TelResp:
        if 'NAO TEL' not in i:
            TelResp3.append(i)



# -----------------------------------//----------------------------------------------



# Funçao na qual ira pegar os valores dentro da coluna Tel Aluno e formatar ela, dando aos Tel Aluno vazios um valor NAO TEL e tirar carcteres especiais
def formartarTelAluno():
    tabela = pd.read_excel('teste1.xlsx')

    telAluno = list(tabela.loc[tabela['Nº']>0, 'Tel. Celular'])

    num = 1
    for i in telAluno:
        resp = isinstance(i, str)
        if resp == True:
            i = i.replace('(','')
            i = i.replace(')',' ')
            i = i.replace('-','')
            tabela.loc[tabela['Nº']==num, 'Tel. Celular'] = i
        else:
            i = str('NAO TEL')
            tabela.loc[tabela['Nº']==num, 'Tel. Celular'] = i
        num = num + 1

    tabela.to_excel('teste2.xlsx', index=False) 

# Funcao para pegar apenas os Tel Aluno existente, ou pegar os Tel Aluno completos formatados
def pegarTelAluno():
    tabela = pd.read_excel('teste2.xlsx')

    global TelAluno
    TelAluno = list(tabela.loc[tabela['Nº']>0, 'Tel. Celular'])

    global TelAluno3
    TelAluno3 = []
    for i in TelAluno:
        if  'NAO TEL' not in i:
            TelAluno3.append(i)




# -----------------------------------//----------------------------------------------



# Funçao na qual ira pegar os valores dentro da coluna Nome Alunos e formatar, deixando todos o caracteres maiusculos
def formartarNomeAluno():
    tabela = pd.read_excel('teste1.xlsx')

    NomeAluno = list(tabela.loc[tabela['Nº']>0, 'Nome (aluno)'])

    num = 1
    for i in NomeAluno:
        i = i.upper()
        tabela.loc[tabela['Nº']==num, 'Nome (aluno)'] = i
        num = num + 1

    tabela.to_excel('teste2.xlsx', index=False) 

# Funcao para pegar apenas os Nome Alunos formatados
def pegarNomeAluno():
    tabela = pd.read_excel('teste2.xlsx')

    NomeAluno = list(tabela.loc[tabela['Nº']>0, 'Nome (aluno)'])

    global NomeAluno3
    NomeAluno3 = []
    for i in NomeAluno:
        if  'NAO NOME' not in i:
            NomeAluno3.append(i)



# -----------------------------------//----------------------------------------------


def chamarTodasFuncoes():
    formartarData()
    pegarAnoNasc()

    formartarEmail()
    pegarEmail()

    formartarTelResp()
    pegarTelResp()

    formartarTelAluno()
    pegarTelAluno()

    formartarNomeAluno()
    pegarNomeAluno()



def colherDados():
    lista = []
    global DadosAlunos
    DadosAlunos = []

    for i in range(0,20):
        lista.append(NomeAluno3[i])
        lista.append(TelAluno[i])
        lista.append(TelResp[i])
        lista.append(email2[i])
        lista.append(anoNascimento[i])
        DadosAlunos.append(lista)
        lista = []


def send_email():
    for i in DadosAlunos:
        nome = i[0]
        num = i[1]
        numResp = i[2]
        email1 = i[3]
        ano = i[4]

        email_content = f'''
        <p> Ola <b>{nome}</b>, tudo bem? </p>
        <p>  </p>
        <p> <b>Numero::</b> {num} </p>
        <p> <b>Numero Responsavel:</b> {numResp} </p>
        <p> <b>E-mail:</b> {email1} </p>
        <p> <b>Ano Nascimento:</b> {ano} </p>
        '''
        print(email_content)

        msg = email.message.Message()
        msg['Subject'] = 'E-mail enviado com sucesso'
        msg['From'] = 'seuEmail@gmail.com'
        msg['To'] = 'seuEmail@gmail.com'
        password = 'suasenha'
        msg.add_header('Content-Type', 'text/html')
        msg.set_payload(email_content)

        s = smtplib.SMTP('smtp.gmail.com: 587')
        s.starttls()
        s.login(msg['From'], password)
        s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))

        print('Enviado com sucesso!')


# -----------------------------------//----------------------------------------------


def main():
    chamarTodasFuncoes()
    colherDados()
    send_email()

# '''
# Main
# '''

main()

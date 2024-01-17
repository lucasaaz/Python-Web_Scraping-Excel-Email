# Python <img src="https://cdn.jsdelivr.net/gh/devicons/devicon/icons/python/python-original.svg" width="40px"/> 
## Automação, Excel e Email

### Resumo:
O código fornecido é um programa em Python que realiza diversas operações em um arquivo Excel (teste1.xlsx) contendo informações sobre alunos. O código manipula os dados, realiza formatações e salva o 
resultado em um novo arquivo Excel (teste2.xlsx). Além disso, o programa envia e-mails para uma lista de alunos utilizando a biblioteca smtplib e email.message.

<h1></h1>

### Vou fornecer um resumo mais detalhado das principais funcionalidades do código:

### formartarData():

- Lê um arquivo Excel (teste1.xlsx) usando a biblioteca Pandas.
- Formata a coluna 'Data de Nascimento', convertendo todas as datas para o mesmo padrão.
- Salva as alterações em um novo arquivo Excel (teste2.xlsx).

### pegarAnoNasc():

- Lê o arquivo Excel formatado (teste2.xlsx).
- Extrai o ano de nascimento de cada aluno.

### formartarEmail():

- Lê o arquivo Excel original.
- Formata a coluna 'E-mail', atribuindo "NAO EMAIL" a e-mails vazios.
- Salva as alterações em teste2.xlsx.

### pegarEmail():

- Lê o arquivo Excel formatado (teste2.xlsx).
- Extrai e armazena os e-mails formatados.

### formartarTelResp():

- Lê o arquivo Excel original.
- Formata a coluna 'Tel. Celular Resp.', atribuindo "NAO TEL" a telefones vazios e removendo caracteres especiais.
- Salva as alterações em teste2.xlsx.

### pegarTelResp():

- Lê o arquivo Excel formatado (teste2.xlsx).
- Extrai e armazena os telefones dos responsáveis formatados.

### formartarTelAluno():

- Lê o arquivo Excel original.
- Formata a coluna 'Tel. Celular', atribuindo "NAO TEL" a telefones vazios e removendo caracteres especiais.
- Salva as alterações em teste2.xlsx.

### pegarTelAluno():

- Lê o arquivo Excel formatado (teste2.xlsx).
- Extrai e armazena os telefones dos alunos formatados.

### formartarNomeAluno():

- Lê o arquivo Excel original.
- Formata a coluna 'Nome (aluno)', convertendo todos os caracteres para maiúsculas.
- Salva as alterações em teste2.xlsx.

### pegarNomeAluno():

- Lê o arquivo Excel formatado (teste2.xlsx).
- Extrai e armazena os nomes dos alunos formatados.

### chamarTodasFuncoes():

- Chama todas as funções de formatação em uma única chamada.

### colherDados():

- Cria uma lista contendo informações específicas (nome, telefone aluno, telefone responsável, e-mail, ano de nascimento) para cada aluno.

### send_email():

- Utiliza a biblioteca smtplib e email.message para enviar e-mails para cada aluno com base nas informações coletadas.

### main():

- Função principal que orquestra a execução de todas as outras funções.

# Observação: 
Para que o envio de e-mails funcione, é necessário substituir as credenciais de e-mail e senha no código. Além disso, o código é específico para a estrutura do arquivo Excel fornecido e pode precisar de ajustes dependendo da estrutura real do seu conjunto de dados

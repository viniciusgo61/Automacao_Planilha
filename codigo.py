#bibliotecas
# pandas -> bases de dados "pip install pandas numpy openpyxl"
# os ->trabalha com arquivos do computador
# pywin32 -> enviar email

#importação
import os
from datetime import datetime
import pandas as pd
import win32com.client as win32 #client para importar c o outlook

caminho = "bases/" #varivavel com o caminha completo da base de dados
arquivos = os.listdir(caminho) #listando todos os arquivos
print(arquivos) # todos os arquivos.

tabela_consolidada = pd.DataFrame() #tabela vazia
#tratando as datas da tabela
for nome_arquivo in arquivos: #percorrendo os arquivos e adicionando na tabela vazia
    tabela_vendas = pd.read_csv(os.path.join(caminho, nome_arquivo)) #lendo arquivos juntos pela biblioteca os
    tabela_vendas["Data de Venda"] = pd.to_datetime("01/01/1900") + pd.to_timedelta(tabela_vendas["Data de Venda"],
                                                                                    unit="d") # ajuste da coluna de vendas
    tabela_consolidada = pd.concat([tabela_consolidada, tabela_vendas]) #concatenando as 2 bases de dados

tabela_consolidada = tabela_consolidada.sort_values(by="Data de Venda") # ordenando a tabela pelas datas de venda
tabela_consolidada = tabela_consolidada.reset_index(drop=True) # resetando o index de 0 até a ultima linha
tabela_consolidada.to_excel("Vendas.xlsx", index=False) # salvando em um único arquivo excel e excluindo o index

#enviando o email
outlook = win32.Dispatch('outlook.application') #nome do outlook no computador/conexao
email = outlook.CreateItem(0) # criar um item c indice 0
email.To = "viniciusgoms831@gmail.com" #para quem iremos enviar
data_hoje = datetime.today().strftime("%d/%m/%Y") #variavel data de hoje/transformando data em string
email.Subject = f"Relatório de Vendas {data_hoje}" #assunto do email
email.Body = f"""
Prezados,

Segue em anexo o Relatório de Vendas de {data_hoje} atualizado.
Qualquer coisa estou à disposição.
Abs,
Lira Python
"""

caminho_arquivo = os.getcwd() #caminho do local do arquivo vendas.
anexo = os.path.join(caminho_arquivo, "Vendas.xlsx") #anexo de arquivo (caminho do os + nome do arquivo)
email.Attachments.Add(anexo) #adicionando o anexo no email que criamos.

email.Send()#enviar o email.
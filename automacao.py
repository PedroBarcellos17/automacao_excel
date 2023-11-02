# Nesse código vamos usar o pandas para tratar a base de dados, os para acessar arquivos e pywin32 para enviar pelo outlook

import os
import pandas as pd
import win32com.client as win32
from datetime import datetime

caminho = "bases/"
# Vai listar os arquivos de um diretório
arquivos = os.listdir(caminho)
print(arquivos)

tabela_consolidada = pd.DataFrame()

for nome in arquivos:
    # Vamos criar um leitor de csv para passar em cada csv
    # A bibliotecaa os vai unir o caminho e o nome da base, para acha o arquivo e ler ele
    tabela_vendas = pd.read_csv(os.path.join(caminho, nome))
    tabela_vendas['Data de Venda'] = pd.to_datetime(
        "01/01/1900") + pd.to_timedelta(tabela_vendas['Data de Venda'], unit='d')
    tabela_consolidada = pd.concat([tabela_consolidada, tabela_vendas])

tabela_consolidada = tabela_consolidada.sort_values(by="Data de Venda")
tabela_consolidada = tabela_consolidada.reset_index(drop=True)
tabela_consolidada.to_excel("Vendas.xlsx", index=False)


outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)
email.To = 'Insira o email que será enviado'
data_hoje = datetime.today().strftime("%d/%m/%Y")
email.Subject = f'Relatório de Vendas - {data_hoje}'
email.Body = f'''
Prezados,

Segue em anexo o Relatório de Vendas de {data_hoje} atualizado.
Qualquer problema estou à disposição.
Abs,
Pedro
'''

caminho = os.getcwd()
anexo = os.path.join(caminho, 'Vendas.xlsx')
email.Attachments.Add(anexo)

email.Send()

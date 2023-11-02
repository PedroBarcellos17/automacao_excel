## Automação de Relatório de Vendas
Este script Python automatiza a consolidação de dados de vendas de arquivos CSV em um único arquivo Excel e envia o relatório gerado por e-mail através do Outlook. O script faz uso das bibliotecas Pandas para manipulação de dados e pywin32 para interagir com o Outlook.

## Instalação
Antes de executar o script, as seguintes bibliotecas Python devem ser instaladas:

Pandas: Para manipulação de dados tabulares.
pywin32: Fornece acesso às funcionalidades COM do Windows.
Certifique-se de ter o Python instalado. Em seguida, instale as dependências com o seguinte comando:

pip install pandas
pip install pywin32

## Como Usar
Configuração do Ambiente:

## Clone ou faça o download deste repositório para o seu ambiente de trabalho.
Preparação dos Arquivos:

Os arquivos de vendas a serem consolidados devem ser colocados na pasta bases/.
Personalização do Código:

Abra o arquivo automacao_vendas.py em um editor de texto.
Modifique a variável email.To para o endereço de e-mail de destino.
Execução do Script:

Execute o script automacao_vendas.py através do terminal ou IDE Python.
Funcionamento do Código
Leitura dos Arquivos de Vendas:

O script busca arquivos na pasta bases/, lê os arquivos CSV e consolida os dados em um DataFrame Pandas.
Manipulação de Dados:

Os dados são tratados para converter a coluna de data para o formato apropriado.
Todos os dados são consolidados em um único arquivo Excel chamado Vendas.xlsx.
Envio do Relatório pelo Outlook:

O script utiliza o Outlook para enviar um e-mail com o relatório anexado.
O e-mail contém o assunto "Relatório de Vendas" seguido pela data atual.
## Observações
Certifique-se de possuir permissões para acessar o Outlook e enviar e-mails automaticamente.
Recomenda-se revisar e personalizar a mensagem de e-mail conforme necessário para sua situação específica.
Este script é fornecido como exemplo e deve ser utilizado com responsabilidade e respeitando as políticas de uso de sistemas e envio de e-mails da sua organização.

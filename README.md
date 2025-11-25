Extração de Conversas WhatsApp – Z-API (Python)
Descrição do Projeto

Este projeto tem como objetivo extrair todas as conversas da sua instância do WhatsApp conectada via Z-API.
O sistema acessa a API, coleta todos os chats (com paginação automática) e gera um arquivo Excel contendo as principais informações de cada conversa.

A ideia é facilitar o gerenciamento de contatos, análise de atendimento e organização interna, especialmente para quem utiliza WhatsApp como canal de vendas ou suporte.

Tecnologias Utilizadas

Python 3

Biblioteca requests (requisições à API)

Biblioteca openpyxl (criação e formatação da planilha Excel)

Plataforma Z-API (API utilizada para obter os chats)

Configuração Inicial

No código, basta substituir suas credenciais da Z-API:

INSTANCE_ID = "SUA_INSTANCE"
TOKEN = "SEU_TOKEN"
CLIENT_TOKEN = "SEU_CLIENT_TOKEN"


Após ajustar essas informações, o script já está pronto para uso.

Como Executar
1. Instale as bibliotecas necessárias

Use o terminal:

pip install requests openpyxl

2. Insira suas credenciais no arquivo Python

Substitua:

INSTANCE_ID = "..."
TOKEN = "..."
CLIENT_TOKEN = "..."

3. Execute o programa

No terminal, dentro da pasta do projeto:

python nome_do_arquivo.py

4. Resultado

O script irá gerar automaticamente um arquivo Excel com o nome:

conversas_whatsapp_YYYYMMDD_HHMMSS.xlsx


Esse arquivo contém todas as conversas organizadas com informações como telefone, nome, mensagens não lidas, data da última mensagem, entre outros dados relevantes.

O que o Script Faz

Conecta à Z-API utilizando suas credenciais

Busca todas as conversas em todas as páginas

Converte timestamps para datas legíveis

Gera uma planilha Excel estruturada e formatada

Salva o arquivo na pasta do projeto

Observações Importantes

Este projeto realiza apenas leitura dos contatos para envio de uma planilha.
Não envia mensagens, não altera chats e não realiza nenhuma ação além de extrair os dados e gerar a planilha.

Autor

Desenvolvido por Gustavo, utilizando Python e Z-API.
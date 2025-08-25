# Projeto de Notificação de Inconsistências em E-mails de Instrutores, Coordenadores e Orientador Pedagógico - FINDES
 
Este projeto tem como objetivo validar os dados de e-mails dos supervisores, instrutores e responsáveis pedagógicos do sistema educacional utilizado pela FINDES. O sistema realiza a validação de dados a partir de informações extraídas de uma planilha, que contém os dados de cada filial, e realiza chamadas para a API do RM, utilizando parâmetros de consulta que envolvem datas (DataAtual - 4 dias e DataAtual - 1 dia).
 
O processo de validação tem como objetivo identificar inconsistências nos dados de e-mail, retornando uma tabela com as informações relacionadas a instrutores, supervisores, responsáveis pedagógicos, turno e código de perito (CODPERLET). Quando uma inconsistência é detectada, uma notificação é enviada à filial responsável para que as correções necessárias sejam realizadas.
## Estrutura do Projeto
 
- **main.py**: Arquivo principal que executa o RPA (Robotic Process Automation).
- **RmAPI.py**: Contém a classe `RmAPI` com os métodos para interagir com a API externa.
- **Logger.py**: Contém a classe `LogGenerator` para configuração e registro de logs.
- **slack_notifier.py**: Implementa a classe `SlackNotifier` para enviar notificações para um canal do Slack.
- **.env** e **sample.env**: Arquivos de configuração contendo variáveis de ambiente como URLs de API e credenciais.
- **requirements.txt**: Lista de dependências do projeto Consulte **[Implantação](https://github.com/joaopbravo/12.b-Notificacao_erros_instrutores/blob/abaae334c8738a6685846eb1856f985a111d9bdd/requirements.txt)
** para saber como implantar o projeto
 
## Pré-requisitos
 
- Python 3.8 ou superior
- As bibliotecas listadas em `requirements.txt`
 
## Instalação
 
1. Clone o repositório para sua máquina local.
```bash
git clone https://github.com/joaopbravo/12.b-Notificacao_erros_instrutores.git
```
2. Navegue até o diretório do projeto.
```bash
cd seu-repositorio
```
3. Crie um ambiente virtual.
```bash
python -m venv venv
```
4. Ative o ambiente virtual.
- No Windows:
```bash
venv\Scripts\activate
```
5. Instale as dependências.
```bash
pip install -r requirements.txt
```
6. Configure as variáveis de ambiente no arquivo `.env` baseado no `sample.env`.
 
## Uso
 
1. Certifique-se de que as variáveis de ambiente estão corretamente configuradas no arquivo `.env`.
2. Execute o script principal para iniciar o processo de validação de e-mail.
```bash
python main.py
```
 
## Configuração
 
As seguintes variáveis de ambiente devem ser configuradas no arquivo `.env`:
 
- `API_URL_GLOBAL`: URL da API do sistema educacional - RM.
- `USER`: Nome de usuário para autenticação na API.
- `PASSWORD`: Senha para autenticação na API.
- `DOWLOAD_PATH_MODELS`: Caminho para a planilha de dados de Responsáveis Unidade.xlsx
 
## Funcionalidades
 
- **Consulta de Email**: O script lê os dados de uma planilha de Responsáveis da Unidade, localizada no SharePoint PedagogicoBot, e realiza uma chamada GET via API do sistema Educacional RM, com o objetivo de identificar inconsistências nos dados.
- **Logging**: Os logs das operações são registrados para auditoria e depuração.
- **Notificações no Slack**: Notificações são enviadas para um canal do Slack para informar o status das operações
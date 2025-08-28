# import os
# from office365.sharepoint.client_context import ClientContext
# from office365.runtime.auth.user_credential import UserCredential
# from src.functions.SlackFunctions import SlackNotifier
# from src.functions.Logger import LogGenerator
# # from datetime import datetime
# from dotenv import load_dotenv
# import pandas as pd

# from src.functions.Logger import LogGenerator
# from dotenv import load_dotenv
# import os
# from datetime import datetime, timedelta
# from office365.runtime.auth.authentication_context import AuthenticationContext
# from office365.sharepoint.client_context import ClientContext
# from office365.runtime.auth.client_credential import ClientCredential
# from office365.runtime.auth.token_response import TokenResponse
# from office365.graph_client import GraphClient



# load_dotenv()
# class Sharepoint:
#     def __init__(self,slack_notifier, logger):
#         self.SHAREPOINT_EMAIL = os.getenv('SHAREPOINT_EMAIL')
#         self.SHAREPOINT_PASSWORD = os.getenv('SHAREPOINT_PASSWORD')
#         self.SHAREPOINT_URL_SITE = os.getenv('SHAREPOINT_URL_SITE')
#         self.SHAREPOINT_SITE_NAME = os.getenv('SHAREPOINT_SITE_NAME')
#         self.SHAREPOINT_DOC_LIBRARY = os.getenv('SHAREPOINT_DOC_LIBRARY')
#         self.DOWNLOAD_PATH = os.getenv('DOWNLOAD_PATH')
#         self.DOWNLOAD_PATH_MODELS = os.getenv('DOWNLOAD_PATH_MODELS')
#         self.slack = slack_notifier
#         self.logger = logger


#     def ConnectSharepoint(self):
#         return ClientContext(self.SHAREPOINT_URL_SITE).with_credentials(UserCredential(self.SHAREPOINT_EMAIL, self.SHAREPOINT_PASSWORD))

#     def DownloadResponsaveisUnidade(self):
#         ctx = self.ConnectSharepoint()

#         # Caminho no SharePoint
#         sharepointFilePath = f"/sites/{self.SHAREPOINT_SITE_NAME}/Documentos%20Compartilhados/01%20-%20Verifica%C3%A7%C3%A3o%20Frequ%C3%AAncia%20e%20Plano%20de%20Aula/2025/Respons%C3%A1vel%20x%20Unidade/Respons%C3%A1veis%20Unidade.xlsx"
        
#         # Caminho de salvamento local, no diretório correto
#         fileSavePath = os.path.join(self.DOWNLOAD_PATH_MODELS, "Responsáveis Unidade.xlsx")

#         print(f"Arquivo será salvo em: {fileSavePath}")

#         try:
#             # Baixando o arquivo do SharePoint e salvando localmente
#             with open(fileSavePath, "wb") as local_file:
#                 ctx.web.get_file_by_server_relative_url(sharepointFilePath).download(local_file).execute_query()

#             # Verificando se o arquivo foi baixado corretamente
#             if os.path.exists(fileSavePath):
#                 print(f"Arquivo baixado com sucesso: {fileSavePath}")
#             else:
#                 print(f"Erro: O arquivo não foi baixado corretamente.")
#             return "Success"
#         except Exception as e:
#             print(f"Erro: {e}")
#             return f"Error: {str(e)}"

#     def carregar_unidades(self):
#         # Chama a função de download antes de carregar o arquivo
#         download_status = self.DownloadResponsaveisUnidade()

#         if download_status != "Success":
#             print(f"DownloadResponsaveisUnidade : Erro ao baixar o arquivo: {download_status}")
#             self.slack.post_message(f"DownloadResponsaveisUnidade : Houve um erro ao baixar o arquivo Responsáveis Unidade.xlsx. : {download_status}")

#             return None

#         # Caminho correto para o arquivo já baixado
#         file_path_unidades = os.path.join(self.DOWNLOAD_PATH_MODELS, "Responsáveis Unidade.xlsx")
        
#         # Verificando se o arquivo realmente existe
#         if not os.path.exists(file_path_unidades):
#             print(f"Erro: O arquivo {file_path_unidades} não foi encontrado.")
#             self.logger.info(f"Erro: O arquivo {file_path_unidades} não foi encontrado.")
#             return None
        
#         try:
#             # Carregando o arquivo Excel usando o Pandas
#             df_unidades = pd.read_excel(file_path_unidades)
#             return df_unidades
#         except Exception as e:
#             print(f"Erro ao carregar o arquivo Excel: {e}")
#             self.logger.info(f"Erro ao carregar o arquivo Excel: {e}")
#             self.slack.post_message(f"DownloadResponsaveisUnidade : Erro ao carregar o arquivo Excel Responsáveis Unidade.xlsx. : {e}")

#             return None



from src.functions.Logger import LogGenerator
from dotenv import load_dotenv
import os
from datetime import datetime, timedelta
# from office365.runtime.auth.authentication_context import AuthenticationContext
# from office365.sharepoint.client_context import ClientContext
# from office365.runtime.auth.client_credential import ClientCredential
import pandas as pd

from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext

load_dotenv()

class Sharepoint:
    def __init__(self, slack_notifier, logger):
        self.CLIENT_ID = os.getenv('ID_CLIENT')
        self.CLIENT_SECRET = os.getenv('SECRET_TD')
        self.SHAREPOINT_SITE_URL = os.getenv('SHAREPOINT_URL_SITE')
        self.logger = logger
        self.slack = slack_notifier               
        # Adicione esta linha para definir o atributo faltante
        self.SHAREPOINT_SITE_NAME = os.getenv('SHAREPOINT_SITE_NAME', 'PedagogicoBot')  # Valor padrão caso não exista       
        # Outras configurações
        self.DOWNLOAD_PATH_MODELS = os.getenv('DOWNLOAD_PATH_MODELS')
        self.logger = logger
        self.slack = slack_notifier

    def get_sharepoint_context(self):
        """Método confiável para autenticação moderna"""
        try:
            # Método 1: Recomendado para versões mais novas
            credentials = ClientCredential(self.CLIENT_ID, self.CLIENT_SECRET)
            ctx = ClientContext(self.SHAREPOINT_SITE_URL).with_credentials(credentials)           
            # Teste rápido para verificar se a autenticação funcionou
            web = ctx.web
            ctx.load(web)
            ctx.execute_query()           
            return ctx
            
        except Exception as e:
            self.logger.error(f"Falha na autenticação: {str(e)}")
            raise

    def DownloadResponsaveisUnidade(self):
        """Download do arquivo usando autenticação moderna"""
        try:
            ctx = self.get_sharepoint_context()
            
            sharepointFilePath = f"/sites/{self.SHAREPOINT_SITE_NAME}/Documentos%20Compartilhados/SESI/01%20-%20Verifica%C3%A7%C3%A3o%20Frequ%C3%AAncia%20e%20Plano%20de%20Aula/2025/Respons%C3%A1vel%20x%20Unidade/Respons%C3%A1veis%20Unidade.xlsx"
            fileSavePath = os.path.join(self.DOWNLOAD_PATH_MODELS, "Responsáveis Unidade.xlsx")
# R
            self.logger.info(f"Tentando baixar arquivo para: {fileSavePath}")

            with open(fileSavePath, "wb") as local_file:
                ctx.web.get_file_by_server_relative_url(sharepointFilePath).download(local_file).execute_query()

            if os.path.exists(fileSavePath):
                self.logger.info("Arquivo baixado com sucesso")
                return "Success"
            else:
                error_msg = "O arquivo não foi baixado corretamente"
                self.logger.error(error_msg)
                self.slack.post_message(error_msg)
                return f"Error: {error_msg}"
                
        except Exception as e:
            error_msg = f"Erro no download: {str(e)}"
            self.logger.error(error_msg)
            self.slack.post_message(error_msg)
            return f"Error: {error_msg}"

    def carregar_unidades(self):
        """Método mantido igual, apenas ajustando chamada interna"""
        download_status = self.DownloadResponsaveisUnidade()

        if download_status != "Success":
            error_msg = f"Erro ao baixar o arquivo: {download_status}"
            self.logger.error(error_msg)
            self.slack.post_message(error_msg)
            return None

        file_path_unidades = os.path.join(self.DOWNLOAD_PATH_MODELS, "Responsáveis Unidade.xlsx")
        
        if not os.path.exists(file_path_unidades):
            error_msg = f"Arquivo {file_path_unidades} não encontrado"
            self.logger.error(error_msg)
            self.slack.post_message(error_msg)
            return None
        
        try:
            df_unidades = pd.read_excel(file_path_unidades)
            return df_unidades
        except Exception as e:
            error_msg = f"Erro ao carregar Excel: {e}"
            self.logger.error(error_msg)
            self.slack.post_message(error_msg)
            return None
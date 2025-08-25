import os
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from src.functions.SlackFunctions import SlackNotifier
from src.functions.Logger import LogGenerator
from datetime import datetime
from dotenv import load_dotenv
import pandas as pd

load_dotenv()
class Sharepoint:
    def __init__(self,slack_notifier, logger):
        self.SHAREPOINT_EMAIL = os.getenv('SHAREPOINT_EMAIL')
        self.SHAREPOINT_PASSWORD = os.getenv('SHAREPOINT_PASSWORD')
        self.SHAREPOINT_URL_SITE = os.getenv('SHAREPOINT_URL_SITE')
        self.SHAREPOINT_SITE_NAME = os.getenv('SHAREPOINT_SITE_NAME')
        self.SHAREPOINT_DOC_LIBRARY = os.getenv('SHAREPOINT_DOC_LIBRARY')
        self.DOWNLOAD_PATH = os.getenv('DOWNLOAD_PATH')
        self.DOWNLOAD_PATH_MODELS = os.getenv('DOWNLOAD_PATH_MODELS')
        self.slack = slack_notifier
        self.logger = logger


    def ConnectSharepoint(self):
        return ClientContext(self.SHAREPOINT_URL_SITE).with_credentials(UserCredential(self.SHAREPOINT_EMAIL, self.SHAREPOINT_PASSWORD))

    def DownloadResponsaveisUnidade(self):
        ctx = self.ConnectSharepoint()

        # Caminho no SharePoint
        sharepointFilePath = f"/sites/{self.SHAREPOINT_SITE_NAME}/Documentos%20Compartilhados/01%20-%20Verifica%C3%A7%C3%A3o%20Frequ%C3%AAncia%20e%20Plano%20de%20Aula/2025/Respons%C3%A1vel%20x%20Unidade/Respons%C3%A1veis%20Unidade.xlsx"
        
        # Caminho de salvamento local, no diretório correto
        fileSavePath = os.path.join(self.DOWNLOAD_PATH_MODELS, "Responsáveis Unidade.xlsx")

        print(f"Arquivo será salvo em: {fileSavePath}")

        try:
            # Baixando o arquivo do SharePoint e salvando localmente
            with open(fileSavePath, "wb") as local_file:
                ctx.web.get_file_by_server_relative_url(sharepointFilePath).download(local_file).execute_query()

            # Verificando se o arquivo foi baixado corretamente
            if os.path.exists(fileSavePath):
                print(f"Arquivo baixado com sucesso: {fileSavePath}")
            else:
                print(f"Erro: O arquivo não foi baixado corretamente.")
            return "Success"
        except Exception as e:
            print(f"Erro: {e}")
            return f"Error: {str(e)}"

    def carregar_unidades(self):
        # Chama a função de download antes de carregar o arquivo
        download_status = self.DownloadResponsaveisUnidade()

        if download_status != "Success":
            print(f"DownloadResponsaveisUnidade : Erro ao baixar o arquivo: {download_status}")
            self.slack.post_message(f"DownloadResponsaveisUnidade : Houve um erro ao baixar o arquivo Responsáveis Unidade.xlsx. : {download_status}")

            return None

        # Caminho correto para o arquivo já baixado
        file_path_unidades = os.path.join(self.DOWNLOAD_PATH_MODELS, "Responsáveis Unidade.xlsx")
        
        # Verificando se o arquivo realmente existe
        if not os.path.exists(file_path_unidades):
            print(f"Erro: O arquivo {file_path_unidades} não foi encontrado.")
            self.logger.info(f"Erro: O arquivo {file_path_unidades} não foi encontrado.")
            return None
        
        try:
            # Carregando o arquivo Excel usando o Pandas
            df_unidades = pd.read_excel(file_path_unidades)
            return df_unidades
        except Exception as e:
            print(f"Erro ao carregar o arquivo Excel: {e}")
            self.logger.info(f"Erro ao carregar o arquivo Excel: {e}")
            self.slack.post_message(f"DownloadResponsaveisUnidade : Erro ao carregar o arquivo Excel Responsáveis Unidade.xlsx. : {e}")

            return None
    # def DownloadTabelaAuxiliar(self, nomeUnidade):
    #     ctx = self.ConnectSharepoint()

    #     sharepointFilePath = f"{self.SHAREPOINT_DOC_LIBRARY}{datetime.now().year}/{nomeUnidade}/Tabela Auxiliar/Tabela Auxiliar.xlsx"
    #     fileSavePath = self.DOWNLOAD_PATH + f"Tabela Auxiliar - {nomeUnidade}.xlsx"

    #     print(fileSavePath)
    #     try:

    #         with open(fileSavePath, "wb") as local_file:
    #                 ctx.web.get_file_by_server_relative_url(sharepointFilePath).download(local_file).execute_query()

    #         return "Success"
    #     except Exception as e:
    #         print('>>>>>>>>>>>>>>>>>> ', e)
    #         return f"Error: {str(e)}"


    # def UploadFile(self, filePath):
    #     ctx = self.ConnectSharepoint()
    #     folderPath = f"{self.SHAREPOINT_DOC_LIBRARY}{datetime.now().year}/Compilado Geral"
    #     try:
    #         folder = ctx.web.get_folder_by_server_relative_url(folderPath).get().execute_query()

    #         with open(filePath, "rb") as content_file:
    #             file_content = content_file.read()
            
    #         folder.upload_file(os.path.basename(filePath), file_content).execute_query()

    #         return "Success"
    #     except Exception as e:
    #         return f"Error: {str(e)}"

    # def DeleteCompiladoGeral(self):
    #     ctx = self.ConnectSharepoint()

    #     try:
    #         fileUrl = f"{self.SHAREPOINT_DOC_LIBRARY}{datetime.now().year}/Compilado Geral/Compilado.xlsx"
    #         file = ctx.web.get_file_by_server_relative_url(fileUrl)
    #         file.delete_object().execute_query()

    #         return "Success"
    #     except Exception as e:
    #         return f"Error: {str(e)}"
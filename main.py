from src.functions.RmAPI import RmAPI
from src.functions.ExcelFunctions import ExcelFunctions
from src.functions.MailFunctions import MailFunctions
from src.functions.SlackFunctions import SlackNotifier
from src.functions.DatabaseFunctions import Database
from src.functions.SharepointFunctions import Sharepoint
from src.functions.cycle_time import Cycletime
from src.functions.Logger import LogGenerator
import os
import sys
from time import sleep
import pandas as pd
from datetime import datetime, timedelta
from dotenv import load_dotenv
import numpy as np
import requests
import base64
from PIL import Image
import io
from O365 import Account, FileSystemTokenBackend
import math

load_dotenv()
class main:
    def __init__(self, slack_notifier, logger):
        self.baseDirectory = os.getcwd()
        self.dataDirectory = self.baseDirectory + "\\data\\tempFiles\\"
        self.slack = slack_notifier
        self.logger = logger
        self.cycle_time = Cycletime()
        # Carregar o caminho da imagem do .env
        self.image_path = os.getenv('IMAGE_PATH')
        self.image_resized_base64 =None  # Obtém o caminho da imagem da variável do .env

        if self.image_path:
            print(f"Imagem será carregada de: {self.image_path}")
            self.logger.info("Imagem de Assinatura encontrada e carregada")
        else:
            print("Erro: Caminho da imagem não encontrado no .env")

        #Limpa todos os arquivos da tempFiles
        try:
            tempFiles = os.listdir(self.dataDirectory)
            for file in tempFiles:
                    caminho_completo = os.path.join(self.dataDirectory, file)
                    if os.path.isfile(caminho_completo):
                        os.remove(caminho_completo)
        except:
            pass

#FUNÇÃO PARA TESTE EM HOMOLOGAÇÃO - CARREGAR ARQUIVOS MANUALMENTE
    # def carregar_unidades(self):
    #     # Aqui você carrega o dfUnidades (substitua com seu código real)
    #     file_path_unidades = os.getenv('DOWNLOAD_PATH_MODELS') + r"\Responsáveis Unidade.xlsx"
    #     df_unidades = pd.read_excel(file_path_unidades)
    #     return df_unidades
    

    def main(self):
        start_time = self.cycle_time.get_start_time()
        rmApi = RmAPI(self.logger)
        dt_inicio_d = (datetime.now() - timedelta(days=4)).strftime("%Y%m%d")
        dt_fim_d = (datetime.now() - timedelta(days=1)).strftime("%Y%m%d")

        self.logger.info("Pegando dados da API do RM...")

        status, dfRM = rmApi.GetConsultaSQL(dt_inicio_d, dt_fim_d)
        print(dfRM)

        if "Error" in status:
            self.logger.info("Erro ao pegar dados da API.")
            self.slack.post_message("Erro ao pegar dados da API.")
            sys.exit()

        excel = ExcelFunctions()
        # sharepoint = Sharepoint()

        self.logger.info("Iniciando etapa de excel/emails.")
        try:
            self.logger.info("Lendo arquivo RM.")
            #dfRM = pd.read_excel(self.dataDirectory + "arquivoRM.xlsx")
            dfRM["STATUS"] = "PENDENTE"
            dfRM.rename(columns={'PROFESSOR': 'INSTRUTOR'}, inplace=True)
            #Filtro de unidade para piloto
            dfRM.sort_values(by="UNIDADE")

            # ** Remover comentário após testes de novo SQL.
            dfInstrutores = excel.GetInstrutores(dfRM)
            print(dfInstrutores)
            dfInstrutores = dfInstrutores.drop_duplicates(subset="CODPERLET")
            
        except Exception as e:
             self.logger.info(f"Erro ao ler dataframe: {e}")
             self.slack.post_message(f"Error in get DataFrame: {e}")
             sys.exit()

        self.logger.info("Gerando token e-mail.")
        mail = MailFunctions()
        token = mail.GenerateToken()
        self.logger.info("Começando Loop instrutores")

        # Chamando a função carregar_unidades para obter o DataFrame do Sharepoint
        downloader = Sharepoint(slack_notifier,logger)

        dfUnidades = downloader.carregar_unidades()
        #Carregar um arquivo de teste self.carregar_unidades()
        # dfUnidades = self.carregar_unidades()

        # Verificando o resultado
        if dfUnidades is not None:
            print("Arquivo Responsáveis Unidade.xlsx - carregado com sucesso!")
            print(dfUnidades)     
            self.logger.info("Arquivo Responsáveis Unidade.xlsx - carregado com sucesso!")
        else:
            print("Houve um erro ao carregar o arquivo Responsáveis Unidade.xlsx.")
            self.logger.info("Houve um erro ao carregar o arquivo Responsáveis Unidade.xlsx")
            self.slack.post_message(f"Houve um erro ao carregar o arquivo Responsáveis Unidade.xlsx.")
            sys.exit()

        # Lista de domínios permitidos
        dominios_permitidos = ['@sesi-es.org.br', '@senai-es.org.br', '@findes.org.br', '@docente.senai.br']

        # DataFrame para armazenar os resultados
        #resultado_validacao = []

        # Loop sobre as linhas do dfUnidades
        for index, row in dfUnidades.iterrows():
            try:
                # Obter o valor de 'CODFILIAL' da linha de dfUnidades (isso vai ser usado no assunto)
                codfilial = row['CODFILIAL']
                unidade_desc = row['UNIDADE']
                # codfilial = 4
               # Filtrar dfInstrutores com base no 'CODFILIAL'
                df_filtrado_instrutores = dfInstrutores[dfInstrutores['CODFILIAL'] == codfilial]

                # Verificar se existem resultados após o filtro
                if not df_filtrado_instrutores.empty:
                    # Gerar a tabela com as validações dos instrutores
                    tabela_validacao = self.gerar_tabela_validacao(df_filtrado_instrutores, dominios_permitidos)

                    if tabela_validacao.shape[0] > 0:  # verifica se há linhas no DataFrame
                        # Criar o corpo do e-mail em HTML
                        # body = self.create_html_body(tabela_validacao, row)
                        body = self.create_html_body(tabela_validacao, row['UNIDADE'])
                        # Enviar o e-mail para os responsáveis (ADM1, ADM2, ADM3)
                        mail.SendMail(
                            row['Responsáveis ADM1'], 
                            row['Responsáveis ADM2'], 
                            row['Responsáveis ADM3'], 
                            body, 
                            codfilial,
                            token
                        )
                    else:
                        print(f"Nenhuma correção pendente para a unidade {codfilial}")
                        self.logger.info("Nenhuma correção pendente para a unidade")
                   
                else:
                    print(f"Nenhum instrutor encontrado para CODFILIAL = {codfilial}")
            except Exception as e:
                print(f"Ocorreu um erro para o CODFILIAL {codfilial}: {e}")

    def gerar_tabela_validacao(self, df_instrutores, dominios_permitidos):
        resultado = []
        for _, instrutor_row in df_instrutores.iterrows():
            emailSupervisor = instrutor_row["SUPIMED_EMAIL"]
            emailInstrutor = instrutor_row["EMAIL"]
            emailOrientador = instrutor_row["RESP_PED_EMAIL"]

            # Validar e-mails
            emailSupervisor_valido = self.validar_email(emailSupervisor, dominios_permitidos)
            emailInstrutor_valido = self.validar_email(emailInstrutor, dominios_permitidos)
            emailOrientador_valido = self.validar_email(emailOrientador, dominios_permitidos)

            # Verificar se todos os e-mails são válidos
            if emailSupervisor_valido == 'Válido' and emailInstrutor_valido == 'Válido' and emailOrientador_valido == 'Válido':
                self.logger.info(f"Email validos SUPERVISOR: {emailSupervisor} INSTRUTOR: {emailInstrutor} ORIENTADOR{emailOrientador} ")
                print(f"Email validos SUPERVISOR: {emailSupervisor} INSTRUTOR: {emailInstrutor} ORIENTADOR {emailOrientador} ")
                continue  # Pula a adição dessa linha no resultado

            # Adicionar as validações na tabela
            resultado.append({
                'Instrutor': str(instrutor_row["INSTRUTOR"]),  # Garantir que todos sejam strings
                'E-mail do Instrutor': str(emailInstrutor_valido),
                'E-mail do Supervisor': str(emailSupervisor_valido),
                'E-mail do Responsável Pedagógico': str(emailOrientador_valido),
                'TURNO': str(instrutor_row['TURNO']),
                'CODPERLET': str(instrutor_row['CODPERLET'])})

        df_resultado = pd.DataFrame(resultado)
        return df_resultado
   
    def validar_email(self, email, dominios_permitidos):
        # Verifica se o valor do e-mail é NaN, None, vazio ou se é "nan" (string)
        if pd.isna(email) or email == '' or (isinstance(email, str) and email.lower() == 'nan'):
            return 'Não Preenchido'

        # Verifica se o e-mail termina com um dos domínios permitidos
        for dominio in dominios_permitidos:
            if email.lower().endswith(dominio):
                return 'Válido'

        return 'Inválido'

    def CreateHTMLTable(self, df_resultado):
        # Começa criando a tabela e adicionando o cabeçalho
        return_str = '<table style="border-collapse: collapse; border: 1px solid #333333; width: 100%;">'
        
        # Cabeçalho da tabela - centralizando o conteúdo
        return_str += '<tr>'
        for column in df_resultado.columns:
            return_str += f'<th class="header" style="background-color: #333333; color: #FFFFFF; padding: 8px; text-align: center;">{column}</th>'  # Centralizando o texto do cabeçalho
        return_str += '</tr>'
        
        # Corpo da tabela - percorrendo os dados e aplicando as condições de formatação
        for _, row in df_resultado.iterrows():
            return_str += '<tr>'
            
            for column in df_resultado.columns:
                value = str(row[column])
                style = 'border: 1px solid #333333; padding: 8px; text-align: center;'  # Centralizando o conteúdo das células

                # Alterando o fundo para amarelo em caso de E-mail inválido
                if 'E-mail' in column and value in ['Inválido', 'Não Preenchido']:
                    style = 'background-color: yellow; color: black; border: 1px solid #333333; padding: 8px; text-align: center;'
                
                # Alterando o fundo para azul claro caso seja "Noturno" no "TURNO"
                if column == "TURNO" and value == "Noturno":
                    style = 'background-color: lightblue; border: 1px solid #333333; padding: 8px; text-align: center;'
                
                return_str += f'<td style="{style}">{value}</td>'
            
            return_str += '</tr>'
        
        # Fechar a tabela
        return_str += '</table>'
        
        return return_str
    

    def create_html_body(self, df_resultado, unidade_desc):
        # Gerar a tabela HTML
        tabela_html = self.CreateHTMLTable(df_resultado)
        
        # Carregar e redimensionar a imagem
        if self.image_path and os.path.exists(self.image_path):
            with Image.open(self.image_path) as img:
                # Verificar se a imagem é JPG ou PNG
                if img.format not in ['JPEG', 'PNG']:
                    raise ValueError("O arquivo de imagem não é um JPG ou PNG válido.")
                
                # Redimensionar a imagem para 740x220
                img_resized = img.resize((500, 200), Image.Resampling.LANCZOS)

                # Salvar a imagem redimensionada em um buffer de memória (em formato JPG ou PNG)
                img_byte_arr = io.BytesIO()
                if img.format == 'JPEG':
                    img_resized.save(img_byte_arr, format="JPEG")  # Salvar como JPG
                else:
                    img_resized.save(img_byte_arr, format="PNG")  # Salvar como PNG
                img_byte_arr = img_byte_arr.getvalue()
                
                # Codificar a imagem em base64
                img_base64 = base64.b64encode(img_byte_arr).decode('utf-8')
                
                # Atualizar a variável de imagem codificada
                self.image_resized_base64 = img_base64

            # Criar o corpo do e-mail com a imagem da assinatura
            body = f"""
            <html>
            <body>
                <p>Prezado(s),</p>
                <p>Foi identificado que o(s) seguinte(s) instrutor(es) possuem uma ou mais correções pendentes. Solicitamos que realizem as correções.</p>
                <p><strong>FILIAL: {unidade_desc}</strong></p>
                {tabela_html}
                <p><strong>Atenciosamente,</strong></p>
                <p><strong>Pedagogico Bot</strong></p>
                <p>Equipe de Validação</p>
                
                <!-- Inserir a assinatura com a imagem fixa à esquerda -->
                <div style="text-align: left;">
                <p><strong>Classificação: Interno</strong></p> <!-- Informar que o e-mail é interno -->

                <img src="data:image/{img.format.lower()};base64,{self.image_resized_base64}" alt="Assinatura" style="display: block;" />
                </div>
                <!-- Adiciona uma linha de espaço (escape) -->
                <p> </p>
                <!-- Espaço maior com margem -->
                <p style="margin-top: 20px;"> </p>

                <!-- Centralizar Visão e Missão -->
                <div style="text-align: left;">
                <p><strong>Visão:</strong> “Ser referência como Departamento Econômico entre as Federações de Indústria até 2030”</p>
                <p><strong>Missão:</strong> “Fortalecer o desenvolvimento da indústria do Estado do Espírito Santo por meio de pesquisas, estudos e análises de dados”</p>
                </div>
            </body>
            </html>
            """
            return body
        
        # database = Database()
        # database.UploadDFToTable(dfRM)
        # database.ExportToExcel(f"{self.dataDirectory}\\compilado.xlsx")
        # sharepoint.UploadFile(f"{self.dataDirectory}\\compilado.xlsx")
        # end_time = self.cycle_time.get_start_time()
        # self.cycle_time.gerar_cycle_time(start_time, end_time)

        # self.slack.post_message(f"Finalizado com sucesso, arquivo no Sharepoint.")

if __name__ == "__main__":
    slack_notifier = SlackNotifier(os.getenv("ENDPOINT_SLACK"), os.getenv("CHANNEL_SLACK"), os.getenv("NAME_ALERT"))
    log_instance = LogGenerator()
    logger = log_instance.setup_logger()
    
    start = main(slack_notifier, logger)
    start.main()
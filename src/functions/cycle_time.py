import os

import sqlite3
from datetime import datetime
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import pandas as pd
from dotenv import load_dotenv

load_dotenv()


class Cycletime:
    """
    Classe para gerenciar o cycletime de processos automatizados, incluindo
    registro de informações no banco de dados, exportação de dados para CSV e
    upload de relatórios para o SharePoint.
    """

    def __init__(self):
        """
        Inicializa o caminho para o banco de dados e configurações do SharePoint.
        Os caminhos são definidos direto no código, pois é padrão para todo projeto.
        Caso haja necssidade, alterar.
        """
        self.DB_PATH = "C:/Findes_Deploy/Cycletime DB/bd_cycletime.db"
        self.SHAREPOINT_DOC_LIBRARY_CYCLE = "Documentos Partilhados/2020-2023_Processos/33 - Transformação Digital & Inovação/Hiperautomação/10 - Robotização"
        self.SHAREPOINT_URL_SITE_CYCLE = "https://sesisenaies.sharepoint.com/sites/PMO-Findes"

        #Credenciais devem ser configuradas no arquivo .env
        self.SHAREPOINT_EMAIL = os.getenv('SHAREPOINT_EMAIL_C')
        self.SHAREPOINT_PASSWORD = os.getenv('SHAREPOINT_PASSWORD_C')

        #Caminho local do script
        self.LOCAL_SCRIPT_PATH = os.getenv("LOCAL_SCRIPT_PATH")
    

    def gerar_cycle_time(self, hora_inicio, hora_fim):
        """
            Calcula o tempo de ciclo de um processo, armazena os dados no banco de dados, 
            exporta o banco para um arquivo CSV e faz upload do arquivo para o SharePoint.

            Parâmetros:
            -----------
            hora_inicio : datetime
                Hora de início do processo.
            hora_fim : datetime
                Hora de término do processo.
        """
        cycletime = self.get_cycle_time(hora_inicio, hora_fim)
        self.inserir_cycle_time("N/A", "Robô 12 - Controle frequencia", hora_inicio, hora_fim, cycletime)
        self.exportar_db_to_csv('processos', f'{self.LOCAL_SCRIPT_PATH}/rpa_cycletime.csv')
        self.upload_cycle_time(f'{self.LOCAL_SCRIPT_PATH}/rpa_cycletime.csv')

    def exportar_db_to_csv(self, table, csv_path, db_path=None):
        """
        Exporta dados de uma tabela SQLite para um arquivo CSV.

        Parâmetros:
        -----------
        table : str
            Nome da tabela no banco de dados SQLite.
        csv_path : str
            Caminho onde o arquivo CSV será salvo.
        db_path : str, opcional
            Caminho para o banco de dados SQLite (Se nenhum valor for passado,
            será usado o valor declarado dentro da classe (self.DB_PATH)).

        Retorna:
        --------
        Nennhum retorno, o CSV é salvo na pasta indicada.
        """
        if db_path is None:
            db_path = self.DB_PATH

        conn = sqlite3.connect(db_path)
        df = pd.read_sql_query(f"SELECT * FROM {table}", conn)
        conn.close()
        df.to_csv(csv_path, index=False)

    def inserir_cycle_time(self, escopo, nome_robo, start_time, end_time, cycletime):
        """
        Insere dados de ciclo de tempo no banco de dados.

        Parâmetros:
        -----------
        escopo : str
            Descrição do escopo do processo.
        nome_robo : str
            Nome do robô que executou o processo.
        start_time : datetime
            Data e hora de início do processo.
        end_time : datetime
            Data e hora de término do processo.
        cycletime : timedelta
            Tempo total do ciclo de processo.

        Retorna:
        --------
        None
        """
        conn = sqlite3.connect(self.DB_PATH)
        cursor = conn.cursor()
        cursor.execute('''
                        INSERT INTO processos (Escopo, Robo, DataHoraInicio, DataHoraFim, Cycletime)
                        VALUES (?, ?, ?, ?, ?)
        ''', (escopo, nome_robo, str(start_time), str(end_time), str(cycletime)))
        conn.commit()
        conn.close()

    def get_start_time(self):
        """
        Obtém a data e hora atual como o início do processo.

        Retorna:
        --------
        datetime
            A data e hora atual.
        """
        return datetime.now()

    def get_end_time(self):
        """
        Obtém a data e hora atual como o término do processo.

        Retorna:
        --------
        datetime
            A data e hora atual.
        """
        return datetime.now()

    def get_cycle_time(self, start_time, end_time):
        """
        Calcula o tempo de ciclo entre o início e o término do processo.

        Parâmetros:
        -----------
        start_time : datetime
            Data e hora de início do processo.
        end_time : datetime
            Data e hora de término do processo.

        Retorna:
        --------
        timedelta
            Diferença de tempo entre o início e o término.
        """
        return end_time - start_time

    def upload_cycle_time(self, file_path):
        """
        Faz o upload de um arquivo para o SharePoint.

        Parâmetros:
        -----------
        file_path : str
            Caminho completo para o arquivo que será enviado.

        Retorna:
        --------
        str
            "Success" se o upload foi bem-sucedido ou uma mensagem de erro se falhou.
        """

        ctx = ClientContext(self.SHAREPOINT_URL_SITE_CYCLE).with_credentials(
            UserCredential(self.SHAREPOINT_EMAIL, self.SHAREPOINT_PASSWORD))

        folder_path = self.SHAREPOINT_DOC_LIBRARY_CYCLE
        try:
            folder = ctx.web.get_folder_by_server_relative_url(folder_path).get().execute_query()

            with open(file_path, "rb") as content_file:
                file_content = content_file.read()

            folder.upload_file(os.path.basename(file_path), file_content).execute_query()

            return "Success"
        except Exception as e:
            return f"Error: {str(e)}"

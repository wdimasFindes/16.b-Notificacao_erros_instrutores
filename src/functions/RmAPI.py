import pandas as pd
import requests
import os
from dotenv import load_dotenv

load_dotenv()


class RmAPI:
    def __init__(self, logger):
        self.url = os.getenv('API_URL_GLOBAL')
        self.authorization = os.getenv('AUTH_TOKEN')
        self.logger = logger

    def GetConsultaSQL(self, dt_inicio_d, dt_fim_d):
        try:
            headers = {
            'Authorization': self.authorization
            }

            fullUrl = self.url + f"$CODCOLIGADA=3;DT_INICIO_D={dt_inicio_d};DT_FIM_D={dt_fim_d}"
            # fullUrl = self.url + f"DT_INICIO_D={data};DT_FIM_D={data}"

            self.logger.info(f"URL Criada: {fullUrl}")

            self.logger.info("Começando Response")
            response = requests.get(fullUrl, headers=headers)

            # Verifica se a requisição foi bem-sucedida (código 200)
            if response.status_code == 200:
                # Converte o JSON para um DataFrame
                self.logger.info("Chamada de API realizada com sucesso!")
                data = response.json()
                df = pd.json_normalize(data)  # Se estiver usando uma versão do pandas que suporta json_normalize

                return "Sucesso", df
            else:
                self.logger.info(f"Erro na chamada de API: {response.status_code}")
                return "Error", None
        except Exception as e:
            self.logger.info(f"Erro na chamada de API: Exception: {e}")
            return f"Error: {e}", None
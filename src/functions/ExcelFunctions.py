import os
from dotenv import load_dotenv
import pandas as pd

# Carregar as variáveis de ambiente do arquivo .env
load_dotenv()

class ExcelFunctions:

    def validar_email(self, email, dominios_permitidos):
        # Função para validar e-mails
        if pd.isna(email) or email == '' or (isinstance(email, str) and email.lower() == 'nan'):
            return 'Não Preenchido'

        for dominio in dominios_permitidos:
            if email.lower().endswith(dominio):
                return 'Válido'
        return 'Inválido'


    def gerar_tabela_validacao(self, df_instrutores, dominios_permitidos):
        # Função para gerar a tabela com a validação dos instrutores
        resultado = []
        for _, instrutor_row in df_instrutores.iterrows():
            emailSupervisor = instrutor_row["SUPIMED_EMAIL"]
            emailInstrutor = instrutor_row["EMAIL"]
            emailOrientador = instrutor_row["RESP_PED_EMAIL"]

            # Validar e-mails
            emailSupervisor_valido = self.validar_email(emailSupervisor, dominios_permitidos)
            emailInstrutor_valido = self.validar_email(emailInstrutor, dominios_permitidos)
            emailOrientador_valido = self.validar_email(emailOrientador, dominios_permitidos)

            # Adicionar as validações na tabela
            resultado.append({
                'Instrutor': instrutor_row["INSTRUTOR"],
                'E-mail do Supervisor': emailSupervisor_valido,
                'E-mail do Instrutor': emailInstrutor_valido,
                'E-mail do Responsável Pedagógico': emailOrientador_valido,
            })

        # Converter para um DataFrame para fácil exibição como tabela
        df_resultado = pd.DataFrame(resultado)
        return df_resultado




    def GetInstrutores(self, df):
        # Lista de domínios permitidos
        dominios_permitidos = ['@sesi-es.org.br', '@senai-es.org.br', '@findes.org.br', '@docente.senai.br']
        
        # Filtrando os e-mails que terminam com os domínios permitidos
        df_filtrado = df[df['EMAIL'].str.endswith(tuple(dominios_permitidos))]
        
        # Selecionando as colunas necessárias e removendo duplicatas
        #self.dfInstrutores = df_filtrado[["CODFILIAL","INSTRUTOR", "EMAIL", "SUPIMED_EMAIL", "RESP_PED_EMAIL","TURNO","CODPERLET"]].drop_duplicates()
        self.dfInstrutores = df_filtrado[["CODFILIAL","INSTRUTOR", "EMAIL", "SUPIMED_EMAIL", "RESP_PED_EMAIL","TURNO","CODPERLET"]]

        return self.dfInstrutores

    # def GetInstrutores(self, df):
    #     self.dfInstrutores = df[["INSTRUTOR", "EMAIL", "SUPIMED_EMAIL", "RESP_PED_EMAIL"]].drop_duplicates()
    #     return self.dfInstrutores

    def CreateHTMLTable(self, dicData):

        return_str = '<table style="border-collapse: collapse; border: 1px solid #333333;"><tr>'

        for key in dicData[0].keys():
            if key == "FREQUENCIALIBERADA":
                return_str = return_str + '<th class="header" style="background-color: #333333; color: #FFFFFF;">' + "FREQUENCIA LIBERADA" + '</th>'
            elif key == "CONTEUDOREALIZADO":
                return_str = return_str + '<th class="header" style="background-color: #333333; color: #FFFFFF;">' + "CONTEUDO REALIZADO" + '</th>'
            elif key == "CONTEUDOPREVISTO":
                return_str = return_str + '<th class="header" style="background-color: #333333; color: #FFFFFF;">' + "CONTEUDO PREVISTO" + '</th>'
            else:
                return_str = return_str + '<th class="header" style="background-color: #333333; color: #FFFFFF;">' + key + '</th>'


        return_str = return_str + '</tr>'

        for key in dicData[1].keys():
            return_str = return_str + '<tr>'
            for subkey in dicData[1][key]:
                if subkey == "FREQUENCIALIBERADA" and dicData[1][key][subkey] == "NÃO":
                    return_str = return_str + '<td style="background-color: yellow; border: 1px solid #333333; padding: 8px;">' + str(dicData[1][key][subkey]) + '</td>'
                elif subkey == "CONTEUDOREALIZADO" and dicData[1][key][subkey] == "VAZIO":
                    return_str = return_str + '<td style="background-color: yellow; border: 1px solid #333333; padding: 8px;">' + str(dicData[1][key][subkey]) + '</td>'
                else:
                    return_str = return_str + '<td style="border: 1px solid #333333; padding: 8px;">' + str(dicData[1][key][subkey]) + '</td>'

        return_str = return_str + '</tr></table>'

        return return_str
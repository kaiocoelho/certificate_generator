import pandas as pd
from mailmerge import MailMerge
from datetime import datetime

class CertificateGenerator:

    def __init__(self):
        self.atletas = []
        
        self.certificados = []

        self.preposicoes = ["de", "da", "dos", "do", "e"]

        self.data_convert = {
            1: "janeiro",
            2: "fevereiro",
            3: "março",
            4: "abril",
            5: "maio",
            6: "junho",
            7: "julho",
            8: "agosto",
            9: "setembro",
            10: "outubro",
            11: "novembro",
            12: "dezembro"
        }
        try:
            self.atletas = pd.read_excel(r'/code/generator/table/atletas.xlsx').dropna(how='all').values.tolist()
        except Exception as error:
            print(error)
            return None


    def get_register(self, nome, data):
        nomes = nome.split(' ')
        registro = ''
        for i in nomes:
            if len(i) > 0:
                registro += i[0].upper() if i.lower() not in self.preposicoes else ''
        registro += f"{data.strftime('%m')}{data.strftime('%Y')[2:]}"

        return registro

    def generate(self):
        current_atleta = []
        try:
            for atleta in self.atletas:
                current_atleta = atleta
                info_atleta = {
                    "graduacao": atleta[1],
                    "mes": self.data_convert[atleta[5]],
                    "nome": atleta[0],
                    "local": atleta[3],
                    "ano": str(atleta[6]),
                    "dia": str(atleta[4]),
                    "registro": self.get_register(atleta[0], datetime.strptime(atleta[2].strip(), '%d/%m/%Y').date() if type(atleta[2]) == str else atleta[2])
                }

                self.certificados.append(info_atleta)

            template = "/code/generator/model/certificados_modelo.docx"
            document = MailMerge(template)
            document.merge_pages(self.certificados)
            document.write('/code/generator/certificates/certificados_final.docx')
            return True
        except Exception as error:
            print(error)
            with open('/code/generator/logs/error_log.txt', 'w') as file:
                file.write(
                f"""General error:
                {str(error)}
                    Local do erro: 
                    Erro na linha do atleta: {current_atleta[0]}
                    
                
            """)
            return False



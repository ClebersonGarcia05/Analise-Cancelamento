
import os.path
import pandas as pd
import matplotlib.pyplot as plt

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

#Se modificar esses escopos, exclua o arquivo token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"] #Define o escopo de autorização para acessar as planilhas do Google.

def check_sheet_exists(service, spreadsheet_id, month, analisy_month):
    #obter os metadados de todas as planilhas no arquivo
    spreadsheet = (
        service.spreadsheets().get(spreadsheetId = spreadsheet_id).execute()
    )
    
    sheets = spreadsheet.get('sheets',[]) #Obtem uma lista de todas as planilhas dentro da planilha especificada.
    sheet_names = [sheet['properties']['title'] for sheet in sheets] #Cria uma lista contendo os nomes de todas as planilhas presentes na planilha
    
    #Verifica se o nome do mês passado corresponde a algum nome de planilha
    if month in sheet_names:
        #Verifica se já existe uma planilha de Análise correspondente ao mês
        if analisy_month in sheet_names:
            range_name = f"{analisy_month}!A1:ZZ"
            service.spreadsheets().values().clear(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                body={}
            ).execute()
            return True
        else:
            new_sheet = {
                "properties": {
                    "title": analisy_month
                }
            }
            service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests":[{"addSheet": new_sheet}]}).execute()
        return True
    else:
        return False

def find_next_empty_row(service,spreadsheet_id, sheet_name):
    range_name = f"{sheet_name}!A:A" # Define a faixa de células para a primeira coluna
    
    #Faz a solicitação para obter os valores na primeira coluna
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    
    values = result.get("values", []) # Obtem os valores da primeira coluna
    
    if not values:
        return 1 # Se a primeira coluna estiver vazia, retorna a primeira linha
    for i, row in enumerate(values):
        if not row:
            return i + 1 # Retorna o número da próxima linha vazia
    return len(values) + 1 # Se todas as linhas estiverem preenchidas, retorna o número da próxima linha após a última

def write_dataframe_sheet(service, spreadsheet_id, sheet_name, dataframe):
    #Encontra a próxima linha vazia na primeira coluna
    next_empty_row = find_next_empty_row(service=service, spreadsheet_id=spreadsheet_id, sheet_name=sheet_name)
    
    #Para cada DataFrame, encontra a próxima linha vazia e escrever os dados
    for i, df in enumerate(dataframe):
        #Converter o DataFrame para o formato que possa ser enviado para a planilha
        columns = list(df.columns)
        data = [columns] + df.values.tolist()
        
        #Define a faixa onde os dados serão inseridos começãndo pela próxima linha vazia
        range_name = f"{sheet_name}!A{next_empty_row}"
        
        #Criar o corpo da solicitação para escrever os dados na planilha
        value_range_body = {
            'range': range_name,
            'majorDimension': 'ROWS',
            'values': data,
        }
        
        #Executar a solicitação para escrever os dados na planilha
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption='RAW',
            body=value_range_body,
        ).execute()
        
        next_empty_row += len(df) + 3

def write_value(service, spreadsheet_id, sheet_name, valor, quantidade):
    column = 'C' #Coluna vai receber C
    line = '1'# Linha vai receber 1
    range_name = f"{sheet_name}!{column}{line}" #Define que o intervalo vai ser "Nome da planilha"!C1
    value_range_body = {
        'range': range_name,
        'majorDimension': 'ROWS',
        'values': [["Valor total do mes", "Valor total no ano", "Total de clientes"],
            [valor, valor*12, quantidade],
            ]
    }
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption='RAW',
        body=value_range_body,
        ).execute()

def main(): # Esta é a função principal do programa. Ela coordena todas as operações, desde a autenticação até a escrita dos dados na planilha.
    creds = None
    """ O arquivo token.json armazena os tokens de acesso e atualização do usuário e é criado automaticamente quando o fluxo de autorização é concluído pela primeira vez tempo. """
    if os.path.exists("token.json"): #Verifica se o arquivo token.json, que armazena as credenciais de acesso, existe no diretório atual.
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    #Se não houver credenciais (válidas) disponíveis, deixe o usuário fazer login.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("CAMINHO DO SEU ARQUIVO CLIENT_SECRET AQUI.", SCOPES) #Cria um fluxo de autorização a partir do arquivo de segredos do cliente.
            creds = flow.run_local_server(port=0) # Inicia o fluxo de autorização em um servidor local e obtém as credenciais válidas.
        #Salve as credenciais para a próxima execução
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    try:
        service = build("sheets", "v4", credentials=creds) #Inicializa o serviço do Google Sheets usando as credenciais autorizadas.
        spreadsheet_id = 'COLOQUE O ID DA SUA PLANILHA AQUI.' #Define o ID da planilha onde os dados serão escritos.
        months = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        mes = str(input('Qual mês deseja analisar?')).capitalize().strip() #Pede ao usuário para inserir o mês que deseja analisar e armazena a entrada em mes.
        name_analisy_month = 'Análise ' + mes #Cria o nome da planilha de análise com base no mês selecionado.
        while mes not in months:
            print('Opção Inválida, digite novamente!')
            mes = str(input('Qual mês deseja analisar?')).capitalize().strip()
        if check_sheet_exists(service=service, spreadsheet_id=spreadsheet_id, month=mes, analisy_month=name_analisy_month): # Verifica se a planilha de análise para o mês já existe na planilha principal.
            #Se a Planilha existir seguir com a operação...
            #Chama a API do sheets
            sheet = service.spreadsheets()
            #Ler as informações do Google Sheet
            result = (
                sheet.values().get(spreadsheetId=spreadsheet_id, range=f"{mes}!A:F").execute() #Obtém os valores da planilha principal para o mês especificado.
            )
            values = result.get("values", [])
            df = pd.DataFrame(values[1:], columns=values[1]) # Cria um DataFrame pandas com os valores obtidos da planilha.
            df = df.drop(df[df['Cidade'].isnull() | (df['Cidade'] == 'Cidade')].index) # Deleta as linhas da coluna cidade que estiverem vazias ou preenchidas como "Cidade"
            
            df_quantity_client_city = df.groupby('Cidade').agg({'Cliente': 'count'}).reset_index() # Agrupa os dados do DataFrame por cidade e conta o número de clientes em cada cidade.
            df_quantity_client_city.columns = ['Cidade', 'Quantidade de cliente por cidade'] # Renomeia as colunas do DataFrame resultante.
            df_quantity_client_city = df_quantity_client_city.sort_values('Quantidade de cliente por cidade', ascending=False) #Alinha a tabela por ordem decrescente de acordo com a coluna "Quantidade de cliente"
            quantity_client = df['Cliente'].count().astype(str) # Conta o número total de clientes e converte o resultado para uma string.
            
            df_reason = df.groupby('Descrição/Justificativa').agg({'Cliente': 'count'}).reset_index() # Agrupa os dados do DataFrame por justificativa e conta o número de clientes em cada justificativa.
            df_reason.columns = ['Justificativa', 'Quantidade de clientes por justificativa']
            df_reason = df_reason.sort_values('Quantidade de clientes por justificativa', ascending=False)
            
            df_reason_city = df.groupby(['Cidade', 'Descrição/Justificativa']).agg({'Cliente': 'count'}).reset_index() #Agrupa os dados do DataFrame por cidade e justificativa, contando o número de clientes para cada combinação.
            df_reason_city.columns = ['Cidade', 'Justificativa', 'Quantidade de cliente por cidade e justificativa']
            df_reason_city = df_reason_city.sort_values(['Quantidade de cliente por cidade e justificativa', 'Cidade'], ascending=False)
            
            #Remove o símbolo de moeda e converte os valores dos serviços para o tipo float.
            df['Valor do Serviço'] = df['Valor do Serviço'].str.replace("R$ ", "").str.replace(",", ".")
            df['Valor do Serviço'] = df['Valor do Serviço'].astype(float)
            amount = df['Valor do Serviço'].sum()
            
            write_value(service, spreadsheet_id,name_analisy_month, valor=amount, quantidade=quantity_client) # Escreve o valor total do mês e o número total de clientes na planilha de análise.
            
            df_value_city = df.groupby('Cidade').agg({"Valor do Serviço": "sum"}).reset_index() # Agrupa os dados do DataFrame por cidade e calcula o total do valor dos serviços em cada cidade.
            df_value_city.columns = ['Cidade', 'Valor por cidade']
            df_value_city = df_value_city.sort_values('Valor por cidade', ascending=False)
            
            df_list = [df_quantity_client_city, df_value_city, df_reason, df_reason_city, ] #Cria uma lista de DataFrames que serão escritos na planilha de análise.
            write_dataframe_sheet(service, spreadsheet_id, name_analisy_month, df_list) #Escreve os DataFrames na planilha de análise.
        else:
            #Executa este bloco de código se a planilha de análise para o mês ainda não existir.
            new_sheet = {
                  "properties": {
                      "title": 'Análise ' + mes
                  }
              }
            service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": [{"addSheet": new_sheet}]}).execute()
    except HttpError as err:
        print(err)
if __name__=="__main__": #Verifica se o script está sendo executado como o programa principal e, nesse caso, chama a função main().
    main()
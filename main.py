import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import json
import time
import numpy as np

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1qh85a8HdbhPPPnO37E16XokLto9Zh9ixIhia535wFhc'

 
def main():
    
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                './client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

            
    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()

        
        #-----------------------------------------------------------------------------------------------#
        # CRIAR ABA
        #-----------------------------------------------------------------------------------------------#
        
        # NOMES=['Mateus Lima','Vitor Saraiva', 'Guilherme PensaBem', 'Thiago Paixão']
        
        # NOMES.sort(reverse=True)
        
        # p=json.dumps(service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute())
        # p=json.loads(p)

        # sourcce_spreadsheet_id=p['sheets'][0]['properties']['sheetId']

        # for NOME_ABA in NOMES:

        #     body={
        #         'requests':[
        #             {
        #                 'duplicateSheet':{
        #                     'sourceSheetId':sourcce_spreadsheet_id,
        #                     'newSheetName':NOME_ABA
        #                 }
        #             }
        #         ]
        #     }
        #     result = service.spreadsheets().batchUpdate(
        #                                                 spreadsheetId=SPREADSHEET_ID, 
        #                                                 body=body
        #                                                 ).execute()
            
        #     # ESCREVER O NOME NO C2:E2
            
        #     result = sheet.values().update(
        #                 spreadsheetId=SPREADSHEET_ID,
        #                 range=NOME_ABA+'!C2:E2',
        #                 valueInputOption='USER_ENTERED',
        #                 body={"values":[[NOME_ABA]]},
        #             ).execute()
        
        
        
        #-----------------------------------------------------------------------------------------------#
        # Verificar Horarios
        #-----------------------------------------------------------------------------------------------#

        p=json.dumps(service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute())
        p=json.loads(p)

        NOMES=[]
        
        for i in range(len(p['sheets'])):
            NOMES.append(p['sheets'][i]['properties']['title'])

        Quadro=[[0,0,0,0,0],
        [0,0,0,0,0],
        [0,0,0,0,0],
        [0,0,0,0,0],
        [0,0,0,0,0],
        [0,0,0,0,0],
        [0,0,0,0,0],
        [0,0,0,0,0],
        [0,0,0,0,0],
        [0,0,0,0,0],
        [0,0,0,0,0],
        [0,0,0,0,0],
        [0,0,0,0,0],
        [0,0,0,0,0],
        [0,0,0,0,0],
        [0,0,0,0,0],
        [0,0,0,0,0],
        [0,0,0,0,0],
        [0,0,0,0,0],
        [0,0,0,0,0]]

        PessoasLivre=[[[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]]]

        PessoasOcupadas=[[[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]],
        [[],[],[],[],[]]]
        

        for NOME_ABA in NOMES:
            RANGE_NAME = f'{NOME_ABA}!B6:F25'
            result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                        range=RANGE_NAME).execute()
            values = result.get('values', [])

            if len(values)<20:
                for x in range(20-len(values)):
                    values.append(['','','','',''])

            l=0
            c=0
            for linhas in values:
                if len(linhas)<5:
                    for i in range(5-len(linhas)):
                       linhas.append('') 
                
                c=0
                for colunas in range(5): 
                    
                    if len(values[l][c])==0:
                        Quadro[l][c]=Quadro[l][c]+1
                        PessoasLivre[l][c].append(NOME_ABA)
                    else:
                        PessoasOcupadas[l][c].append(NOME_ABA)
                    c=c+1
                l=l+1
            time.sleep(60/len(NOMES))
        
        l=0
        c=0

        ## Opçao de mostrar quantos estao com o tempo ocupado

        # for x in Quadro:
        #     c=0
        #     for y in Quadro[c]:
        #         Quadro[l][c]=len(NOMES)-Quadro[l][c]
        #         c=c+1
        #     l=l+1

        #'   HORÁRIO  ',
        DIAS_DA_SEMANA=['Seg','Ter','Qua','Qui','Sex']

        # HORARIOS=['M1 - 07:00/07:50',
        #           'M2 - 07:50/08:40',
        #           '    INTERVALO   ',
        #           'M3 - 08:50/09:40',
        #           'M4 - 09:40/10:30',
        #           '    INTERVALO   ',
        #           'M5 - 10:40/11:30',
        #           'M6 - 11:30/12:20',
        #           '      ALMOÇO    ',
        #           'T1 - 12:30/13:20',
        #           'T2 - 13:20/14:10',
        #           '    INTERVALO   ',
        #           'T3 - 14:20/15:10',
        #           'T4 - 15:10/16:00',
        #           '    INTERVALO   ',
        #           'T5 - 16:10/17:00',
        #           'T6 - 17:00/17:50',
        #           '    INTERVALO   ',
        #           'N1 - 18:00/18:45',
        #           'N2 - 18:45/19:30'
        #           ]
        
        
        # print(DIAS_DA_SEMANA)
        # for x,y in zip(HORARIOS,Quadro):
        #     print(x,y)
        
        # print('')
        # print(f'Total de Membros: {len(NOMES)}')

        # for x,y in zip(HORARIOS,PessoasLivre):
        #     print(x,y,f"\n")
        
        # for x,y in zip(HORARIOS,PessoasOcupadas):
        #     print(x,y,f"\n")


        txt='{'
        txt=txt+f"""
        "quantidade":{len(NOMES)},
        "horarios":"""
        txt=txt+'{'
        for x,y in zip(DIAS_DA_SEMANA,np.array(Quadro, dtype=object).transpose()):
            txt=txt+f"""
                "{x}":{str(y).replace(' ',', ')},"""
        txt=txt.rstrip(',')
        txt=txt+'\n        },'

        txt=txt+"""\n        "disponivel":{"""
        for x,y in zip(DIAS_DA_SEMANA,np.array(PessoasLivre, dtype=object).transpose()):
            txt=txt+f"""
                "{x}":{str(y).replace('list(','').replace(')',',')},"""
        txt=txt.rstrip(',')  
        txt=txt+'\n        },'

        txt=txt+"""\n        "indisponivel":{"""
        for x,y in zip(DIAS_DA_SEMANA,np.array(PessoasOcupadas, dtype=object).transpose()):
            txt=txt+f"""
                "{x}":{str(y).replace('list(','').replace(')',',')},"""
        txt=txt.rstrip(',')
        txt=txt+'\n        }\n'
        
        txt=txt+'\n}\n'

        with open(os.path.join(".","Saida.json"),"w",encoding='utf8') as my_file:
            my_file.write(txt.replace("'",'"').replace("],]",']]'))

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()
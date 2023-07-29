from Biblioteca import Database
from datetime import datetime
import pandas as pd
import shutil
import os

table = 'tbl_excel_plano_contas'
direx = '.\\Arquivos\\Excel\\PlanoDeContas\\'

def dfToSql(df,table,is_increment): 
    db = Database()
    db.sqlSet(df=df, table=table, is_increment=is_increment)
    
def start(direx):
    arquivos_xlsx = [arquivo for arquivo in os.listdir(direx) if arquivo.endswith('.xlsx')]    
    for arquivo in arquivos_xlsx:
        try:
            caminho_arquivo = os.path.join(direx, arquivo)
            df_final = pd.read_excel(caminho_arquivo, header=[0])

            df_final.insert(0, 'data_import', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            df_final.insert(1, 'arquivo', arquivo)
            df_final['data_import'] = pd.to_datetime(df_final['data_import'])
            
            dfToSql(df=df_final,table=table,is_increment=True)
            print('Arquivo importado no SQL com sucesso! ', arquivo)
                # Mover o arquivo para a pasta sucesso
            shutil.move(caminho_arquivo, os.path.join(direx + 'sucesso\\', arquivo))

        except Exception as e:
            print('### Erro ao importar o arquivo! ', arquivo)
            print(e)

            # Mover o arquivo para a pasta erro
            shutil.move(caminho_arquivo, os.path.join(direx + 'erro\\', arquivo))
        
start(direx)
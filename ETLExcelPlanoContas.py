from Biblioteca import Database
from datetime import datetime
import pandas as pd
import shutil
import os

tplan = 'tbl_excel_plano_contas'
tpsub = 'tbl_excel_plano_contas_subgrupo'
direx = '.\\Arquivos\\Excel\\PlanoDeContas\\'

def dfToSql(df,table,is_increment): 
    db = Database()
    db.sqlSet(df=df, table=table, is_increment=is_increment)
    
def start(direx):
    arquivos_xlsx = [arquivo for arquivo in os.listdir(direx) if arquivo.endswith('.xlsx')]    
    for arquivo in arquivos_xlsx:
        try:
            caminho_arquivo = os.path.join(direx, arquivo)
            
            df_plano = pd.read_excel(caminho_arquivo, sheet_name='PlanoDeContas', header=[0])
            df_plano.insert(0, 'data_import', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            df_plano.insert(1, 'arquivo', arquivo)
            df_plano['data_import'] = pd.to_datetime(df_plano['data_import'])            
            dfToSql(df=df_plano,table=tplan,is_increment=True)
             
            df_plano = pd.read_excel(caminho_arquivo, sheet_name='PlanoDeContasMargemContrib', header=[0])
            df_plano.insert(0, 'data_import', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            df_plano.insert(1, 'arquivo', arquivo)
            df_plano['data_import'] = pd.to_datetime(df_plano['data_import'])            
            dfToSql(df=df_plano,table=tpsub,is_increment=True)            

            df_plano = pd.read_excel(caminho_arquivo, sheet_name='PlanoDeContasSubGrupo', header=[0])
            df_plano.insert(0, 'data_import', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            df_plano.insert(1, 'arquivo', arquivo)
            df_plano['data_import'] = pd.to_datetime(df_plano['data_import'])            
            dfToSql(df=df_plano,table=tpsub,is_increment=True)    

            print('Arquivo importado no SQL com sucesso! ', arquivo)            
            shutil.move(caminho_arquivo, os.path.join(direx + 'sucesso\\', arquivo))

        except Exception as e:
            print('### Erro ao importar o arquivo! ', arquivo)
            print(e)
            shutil.move(caminho_arquivo, os.path.join(direx + 'erro\\', arquivo))
        
start(direx)
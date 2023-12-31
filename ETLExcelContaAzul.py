from Biblioteca import Database
from datetime import datetime
import pandas as pd
import shutil
import os

table = 'tbl_excel_conta_azul'
direx = '.\\Arquivos\\Excel\\ContaAzul\\'

def dfToSql(df,table,is_increment): 
    db = Database()
    db.sqlSet(df=df, table=table, is_increment=is_increment)
    
def tratarDf(df):
    df_final = pd.DataFrame() 
    for  coluna in df.columns:
        col = coluna[0].replace('(','').replace(')','').replace('R$','').replace(' ','')
        mes = coluna[1]

        if coluna[1] == 'CATEGORIAS': #verifica se o valor da segunda linha é categorias
            df_categorias = pd.DataFrame({
                'categoria': df[coluna]
            })
            
        else:
            df_metricas = pd.DataFrame({
                'ano_mes_':mes,
                'metrica':col,
                'valor': df[coluna]
            })
            
            df_metricas = df_categorias.merge(df_metricas, left_index=True, right_index=True)
            df_final = pd.concat([df_final, df_metricas], ignore_index=True)

    return df_final

def start(direx):
    arquivos_xlsx = [arquivo for arquivo in os.listdir(direx) if arquivo.endswith('.xlsx')]    
    for arquivo in arquivos_xlsx:
        try:
            caminho_arquivo = os.path.join(direx, arquivo)
            df = pd.read_excel(caminho_arquivo, header=[0, 1])
            df_final = tratarDf(df)

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
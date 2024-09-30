import os, glob
import pandas as pd
import traceback


def merge(lista):

    
    try:
        
        caminho = lista[0]
        nome_arquivo = lista[1] 

        raw_string = caminho.replace("@","\\\\")
        lst_raw_string = raw_string.split(";")

        raw_string= os.path.join(*lst_raw_string)

        all_files = glob.glob(os.path.join(raw_string, "temp-*.csv"))        

        df_merge= [pd.read_csv(i, sep=',',encoding = 'latin-1') for i in all_files]
        
        tamanho = len(df_merge)
        #print(all_files)

        df_concat = df_merge[0]

        for i in range(tamanho-1):
            df_concat = pd.concat([df_concat,df_merge[i+1]], ignore_index=True)
        
        df_concat = df_concat[~df_concat["Contrato R/3"].isnull()]

        df_concat.to_excel(raw_string+"\\"+nome_arquivo+".xlsx",sheet_name="SAP",index=False)
        
        retorno="0"
        return retorno        

    except:
        retorno = "1"","+str(traceback.format_exc())
        return retorno

#print(merge(["@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho","Planilha_Planejamento_Contrato_RPA_30042021"]))






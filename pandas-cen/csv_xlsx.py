import pandas as pd
import traceback
import os,glob
from openpyxl import load_workbook
import re

"""
caminho="@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho;desenbolso.csv"
caminho_save = "@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho"
nome_arquivo= "Planilha_Planejamento_Contrato_RPA_12052021"

teste=[caminho,caminho_save,nome_arquivo]

"""

def filtra_desembolso(lista):
    try:
        caminho = lista[0]
        caminho_save = lista[1]
        nome_arquivo = lista[2]

        caminho= caminho.replace("@","\\\\") # Substitui o @
        caminho_lst= caminho.split(";")

        caminho_save= caminho_save.replace("@","\\\\")
        caminho_save_lst= caminho_save.split(";")

        df_desembolso= pd.read_csv(os.path.join(*caminho_lst),sep=',',encoding="ANSI")

        
        # retorna uma lista com o nome de todas as colunas
        primeira_linha = list(df_desembolso.columns)

        # retorna a primeira linha em uma Série
        segunda_linha = df_desembolso.iloc[0]


        # valor de entrada: 63
        # range : 0 ao 62
        # percorre todas as colunas
        for i in range(len(primeira_linha)): 

            if bool(re.search(r"Unnamed",primeira_linha[i])):
                primeira_linha[i]= primeira_linha[i-1]

                

        #print(primeira_linha)
        count= 0
        for i,j in zip(primeira_linha,segunda_linha):    
            if str(j) != "nan":
                primeira_linha[count] = str(i) +" | "+str(j)

            count+=1    


        df_desembolso.columns = primeira_linha

        df_desembolso=df_desembolso.drop(0)


        path_save = os.path.join(*caminho_save_lst)

        book = load_workbook(path_save+"\\"+nome_arquivo+".xlsx")

        writer = pd.ExcelWriter(path_save+"\\"+nome_arquivo+".xlsx", engine='openpyxl') 
        writer.book = book   

        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

        for chave in writer.sheets:
            if chave == "Desembolso":                
                std=book[chave]                 
                book.remove(std)
                break
                
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)        

        df_desembolso.to_excel(writer, "Desembolso",index=False)        

        writer.save()
        
        retorno = "0"+","
        return retorno       

    except:
        retorno = "1"+","+str(traceback.format_exc())
        return retorno

#filtra_desembolso(teste)

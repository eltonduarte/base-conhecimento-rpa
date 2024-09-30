
import os,glob
import pandas as pd 
import traceback
from openpyxl import load_workbook
import numpy as np


#dependencias pandas, openpyxl

"""
gerencia_geral = ";CENPES/PDGP;CENPES/PDDP;CENPES/PDIDP"
gerencia_tecnica = ";CENPES/PDGP/IRF;CENPES/PDGP/PCP;CENPES/PDDP/FCE;CENPES/PDDP/PCP;CENPES/PDIDP/EPOCOS;CENPES/PDIDP/EPOCOS/PERF;CENPES/PDIDP/EPOCOS/COMP;"
caminho="@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho;arquivo extraido sigitec.csv"
caminho_save = "@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho"
nome_arquivo= "Planilha_Planejamento_Contrato_RPA_24062021"
data_hoje = "24//06//2021"

teste = [gerencia_geral,gerencia_tecnica,caminho,caminho_save,nome_arquivo,data_hoje]
"""


def filtro(lista):
    
    try:

        gerencia_geral = lista[0]
        gerencia_tecnica = lista[1]
        caminho = lista[2]
        caminho_save  = lista[3]
        nome_arquivo = lista[4]
        data_hoje = lista[5]

        data_hoje = data_hoje.replace("//","/")

        raw_string = caminho.replace("@","\\\\")
        lst_raw_string = raw_string.split(";")

        raw_string_save = caminho_save.replace("@","\\\\")
        lst_raw_string_save = raw_string_save.split(";")

        lst_gerencia_geral = gerencia_geral.split(";")
        lst_gerencia_tecnica = gerencia_tecnica.split(";")

        path_save = os.path.join(*lst_raw_string_save)

        gerencia_geral_limpa = [string for string in lst_gerencia_geral if string != ""]

        gerencia_tecnica_limpa = [string for string in lst_gerencia_tecnica if string != ""]

        df_gerencias= pd.read_csv(os.path.join(*lst_raw_string),sep=',',encoding="ANSI")#Antes era UTF-8

        

        df_filtro = df_gerencias[df_gerencias["Gerência Geral"].isin(gerencia_geral_limpa)]  

            
        
        df_filtro = df_filtro[df_filtro["Gerência Técnica"].isin(gerencia_tecnica_limpa)]

        
        
        df_filtro_3 = df_filtro.loc[df_filtro["Data de Término"].isnull()] 

        data_hoje = pd.Timestamp(data_hoje)        
       
        df_filtro_2 = df_filtro.loc[pd.to_datetime(df_filtro["Data de Término"],format="%d/%m/%Y") >= data_hoje]    
        
        

        df_filtro_merge = pd.concat([df_filtro_2,df_filtro_3])

        df_filtro_merge_2= df_filtro_merge[~df_filtro_merge["Tipo de Instrumento Contratual"].str.contains("Acordo de Sigilo")]

        df_filtro_merge_3= df_filtro_merge_2[~df_filtro_merge_2["Estado do Processo"].str.contains("Suspenso")]

       

        book = load_workbook(path_save+"\\"+nome_arquivo+".xlsx")

        writer = pd.ExcelWriter(path_save+"\\"+nome_arquivo+".xlsx", engine='openpyxl') 
        writer.book = book   

        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

        for chave in writer.sheets:
            if chave == "SIGITEC":                
                std=book[chave]                 
                book.remove(std)
                break
        
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)        

        df_filtro_merge_3.to_excel(writer, "SIGITEC",index=False)        

        writer.save()

        retorno="0"
        return retorno

    except:
        retorno = "1"+","+str(traceback.format_exc())
        return retorno

    


#print(filtro(teste))




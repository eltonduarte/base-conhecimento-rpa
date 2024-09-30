import os, glob
import pandas as pd
import traceback
from openpyxl import load_workbook

"""
caminho="@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho;Planilha_Planejamento_Contrato_RPA_18062021.xlsx"
caminho_save = "@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho"
nome_arquivo= "Planilha_Planejamento_Contrato_RPA_18062021"

teste= [caminho,caminho_save,nome_arquivo]
"""

def filtro_tabela(lista):
    try:      
        caminho = lista[0]
        caminho_save = lista[1]
        nome_arquivo= lista[2]

        caminho= caminho.replace("@","\\\\")
        caminho_lst= caminho.split(";")

        caminho_save= caminho_save.replace("@","\\\\")
        caminho_save_lst= caminho_save.split(";")

        df_sap = pd.read_excel(os.path.join(*caminho_lst),sheet_name="SAP")

        df_sigitec = pd.read_excel(os.path.join(*caminho_lst),sheet_name="SIGITEC")

        df_sigitec_sem_sap= df_sigitec.loc[df_sigitec['Número do SAP'].isnull()]

        #df_sigitec_e_sap= df_sigitec.loc[~df_sigitec['Número do SAP'].isnull()]

        df_sigitec_sem_sap_filtro= df_sigitec_sem_sap.filter(["Número do SIC/AEP / Processo","Número do SAP","Proponente","Número de Parcelas Previstas"],axis=1)

        df_merge = pd.merge(df_sap, df_sigitec, left_on='Contrato R/3',right_on='Número do SAP',how='left',indicator=True) #outer, indicator=True

        df_sap_sem_sigitec = df_merge.query('_merge=="left_only"') 

        df_sigitec_e_sap= df_merge.query('_merge=="both"')

        df_sigitec_e_sap_filtro= df_sigitec_e_sap.filter(["Número do SIC/AEP / Processo","Número do SAP","Proponente","Número de Parcelas Previstas"],axis=1)

        #.query('_merge=="left_only"')

        #df_sap_sem_sigitec = df_sap[~df_sap["Contrato R/3"].isin(df_sigitec_e_sap['Número do SAP'])].dropna(how='all')

        #.query('_merge=="left_only"')

        #df_sap_sem_sigitec.to_excel("C:\\Users\\cz5d\\Desktop\\teste.xlsx", "SAP sem SIGITEC",index=False)

        #del df_sap_sem_sigitec["_merge"]

        df_sap_sem_sigitec_filtro = df_sap_sem_sigitec.filter(["Inst Contr Jurídico","Contrato R/3","Objeto Contratual"],axis=1)

        path_save = os.path.join(*caminho_save_lst)


        book = load_workbook(path_save+"\\"+nome_arquivo+".xlsx")

        writer = pd.ExcelWriter(path_save+"\\"+nome_arquivo+".xlsx", engine='openpyxl') 
        writer.book = book   

        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

        for chave in writer.sheets:
            if chave == "SAP sem SIGITEC" or chave == "SIGITEC sem SAP" or chave == "SIGITEC e SAP":                
                std=book[chave]                 
                book.remove(std)
                #break
                
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)        

        df_sap_sem_sigitec_filtro.to_excel(writer, "SAP sem SIGITEC",index=False)  

        df_sigitec_sem_sap_filtro.to_excel(writer, "SIGITEC sem SAP",index=False) 

        df_sigitec_e_sap_filtro.to_excel(writer,"SIGITEC e SAP",index=False)

        writer.save()

        retorno = "0"+","
        return retorno 
              
    except:
        retorno = "1"+","+str(traceback.format_exc())
        return retorno


#print(filtro_tabela(teste))
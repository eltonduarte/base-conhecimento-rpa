import os, glob
import pandas as pd
import traceback
from openpyxl import load_workbook

"""
caminho="@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho;Planilha_Planejamento_Contrato_RPA_13052021.xlsx"
caminho_save = "@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho"
nome_arquivo= "lista_icj_2_13052021"
caminho_template="@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho;template_planejamento.xlsx"

teste = [caminho,caminho_save,nome_arquivo,caminho_template]
"""

def filtro_icj_nove(lista):
    try:
        caminho = lista[0]
        caminho_save = lista[1]
        nome_arquivo = lista[2]

        caminho= caminho.replace("@","\\\\")
        caminho_lst= caminho.split(";")

        caminho_save= caminho_save.replace("@","\\\\")
        caminho_save_lst= caminho_save.split(";")

        path_save = os.path.join(*caminho_save_lst)

        df_sap_sem_sigitec = pd.read_excel(os.path.join(*caminho_lst), sheet_name='SAP sem SIGITEC')

        #print(df_sap_sem_sigitec)

        df_sap_sem_sigitec_filtro = df_sap_sem_sigitec[df_sap_sem_sigitec['Inst Contr Jurídico'].str.contains(r'9$',regex=True)]

        #print(df_sap_sem_sigitec_filtro)

        writer = pd.ExcelWriter(path_save+"\\"+nome_arquivo+".xlsx", engine='openpyxl') 

        df_sap_sem_sigitec_filtro.to_excel(writer, "ICJ9",index=False)  

        writer.save()
        
        retorno = "0"+","
        return retorno

    except:
        retorno = "1"+","+str(traceback.format_exc())
        return retorno


#filtro_icj_nove(teste)

def filtro_icj_dois(lista):
    try:
        caminho = lista[0]
        caminho_save = lista[1]
        nome_arquivo = lista[2]
        caminho_template = lista[3]

        caminho= caminho.replace("@","\\\\")
        caminho_lst= caminho.split(";")

        caminho_save= caminho_save.replace("@","\\\\")
        caminho_save_lst= caminho_save.split(";")

        caminho_template= caminho_template.replace("@","\\\\")
        caminho_template_lst= caminho_template.split(";")

        path_save = os.path.join(*caminho_save_lst)

        df_sap_sem_sigitec = pd.read_excel(os.path.join(*caminho_lst), sheet_name='SAP sem SIGITEC')

        df_template = pd.read_excel(os.path.join(*caminho_template_lst))


        #print(df_sap_sem_sigitec)

        df_sap_sem_sigitec_filtro = df_sap_sem_sigitec[df_sap_sem_sigitec['Inst Contr Jurídico'].str.contains(r'2$',regex=True)]

        df_template["Descrição"] = df_sap_sem_sigitec_filtro["Contrato R/3"]

        df_template["SAP"] = df_sap_sem_sigitec_filtro["Contrato R/3"]

        df_template["Parceiro Dev"] = df_sap_sem_sigitec_filtro["Objeto Contratual"]

        df_template["Status"] = "Contratado"

        #print(df_sap_sem_sigitec_filtro)

        writer = pd.ExcelWriter(path_save+"\\"+nome_arquivo+".xlsx", engine='openpyxl') 

        df_sap_sem_sigitec_filtro.to_excel(writer, "ICJ2",index=False)  

        writer.save()

        book = load_workbook(os.path.join(*caminho_lst))

        writer_2 = pd.ExcelWriter(os.path.join(*caminho_lst), engine='openpyxl')

        writer_2.book = book

        writer_2.sheets = dict((ws.title, ws) for ws in book.worksheets)

        for chave in writer_2.sheets:
            if chave == "Planejamento Plurianual":                
                std=book[chave]                 
                book.remove(std)
                break
        
        writer_2.sheets = dict((ws.title, ws) for ws in book.worksheets)
         

        df_template.to_excel(writer_2, "Planejamento Plurianual",index=False)  

        writer_2.save()

        #print("foi 2")
        retorno = "0"+","
        return retorno

    except:
        retorno = "1"+","+str(traceback.format_exc())
        return retorno


#print(filtro_icj_dois(teste))

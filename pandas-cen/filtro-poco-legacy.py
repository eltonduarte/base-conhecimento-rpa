import os, glob
import pandas as pd
import traceback
from openpyxl import load_workbook

"""

caminho="@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho;Planilha_Planejamento_Contrato_RPA_23072021.xlsx"
caminho_save = "@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho"
nome_arquivo_principal= "Planilha_Planejamento_Contrato_RPA_23072021"
nome_arquivo_email = "CTO_divergente_23072021"
arquivo_rateio= "@petrobras.biz;petrobras;CENPES;CENPES_PDIDP_EPOCOS;NP-2;06. GestaoPDI;03. Paineis;Dados;Painel Consistencia - POCOS - 202101.xlsx"

lista = [caminho,caminho_save,nome_arquivo_principal,nome_arquivo_email,arquivo_rateio]
"""

def poco(lista):
    try:

        caminho = lista[0]
        caminho_save = lista[1]
        nome_arquivo_principal = lista[2]
        nome_arquivo_email = lista[3]
        arquivo_rateio = lista[4]

        caminho= caminho.replace("@","\\\\")
        caminho_lst= caminho.split(";")

        caminho_save= caminho_save.replace("@","\\\\")
        caminho_save_lst= caminho_save.split(";")

        arquivo_rateio= arquivo_rateio.replace("@","\\\\")
        arquivo_rateio_lst= arquivo_rateio.split(";")

        df_planejamento = pd.read_excel(os.path.join(*caminho_lst), sheet_name='Planejamento Plurianual')

        df_pocos = pd.read_excel(os.path.join(*arquivo_rateio_lst), sheet_name="PAINEL CTO", skiprows=15)

        df_pocos.rename(columns={'Mês':'Mês|poco', "Ano":"Ano|poco","Valor Parcela":"Valor Parcela|poco", "%Rateio":"%Rateio|poco"}, inplace=True)

        #Mês , Ano e Valor Parcela 

        df_planejamento["multiplica"] = df_planejamento["Valor Parcela"].str.replace("\.","",regex=True)

        df_planejamento["multiplica"] = df_planejamento["multiplica"].str.replace(",",".",regex=True)

        df_planejamento["multiplica"] = df_planejamento["multiplica"].astype("float64")

        df_planejamento["mes tratado"] = df_planejamento["Mês"].str.replace("(?i)Mês|mes","",regex=True)

        df_planejamento["mes tratado"] = df_planejamento["mes tratado"].astype("float64")

        #df_planejamento.to_excel("C:\\Users\\cz5d\\Desktop\\kappa.xlsx")
        """
        df_merge_ok = pd.merge(df_planejamento,df_pocos,
         how= "inner",
         on= "Descrição")
        """
        #right_on=["Descrição","Parcela","Mês|poco","Ano|poco","Valor Parcela|poco"],
        #left_on=["Descrição","Parcela","mes tratado","Ano","multiplica"] )

        
        #df_merge_ok.to_excel("C:\\Users\\cz5d\\Desktop\\kappa_ok.xlsx")

        """
        df_merge_div = pd.merge(df_planejamento,df_pocos,
         how= "inner",
         on=["Descrição","Parcela"])
        """

        #df_merge_div = df_merge_div.merge(df_merge_ok,how= "left", on=["Descrição","Parcela"], indicator=True ).query('_merge=="left_only"')

        #df_merge_div.to_excel("C:\\Users\\cz5d\\Desktop\\kappa_div.xlsx")

        for i,j,k,d,r in zip(df_pocos["Mês|poco"], 
            df_pocos["Ano|poco"],
            df_pocos["Valor Parcela|poco"],
            df_pocos["Descrição"],
            df_pocos["%Rateio|poco"],
            ):
            
            df_planejamento.loc[(df_planejamento["Descrição"]== d )&            
            (df_planejamento["mes tratado"]== i )&
            (df_planejamento["Ano"]==j) & 
            (df_planejamento["multiplica"]==k) &
            (df_planejamento["%Rateio"]==r),"Validacao Painel CTO"] = "OK"

        """
        for d,p in zip(df_merge_div["Descrição"],df_merge_div["Parcela"]):
            
            df_planejamento.loc[(df_planejamento["Descrição"]== d )&
                (df_planejamento["Parcela"]== p ) ,"Validacao Painel CTO"] = "DIVERGENTE"
            
        """       
        
        del df_planejamento["mes tratado"]

        del df_planejamento["multiplica"]

        df_planejamento.loc[df_planejamento["Validacao Painel CTO"].isna(),"Validacao Painel CTO"] = "DIVERGENTE"

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
                                    
        df_planejamento.to_excel(writer_2, "Planejamento Plurianual",index=False)  
        writer_2.save()


        df_divergente = df_planejamento[df_planejamento["Validacao Painel CTO"]== "DIVERGENTE"] 

        writer = pd.ExcelWriter(os.path.join(*caminho_save_lst)+"\\"+nome_arquivo_email+".xlsx", engine='openpyxl') 

        df_divergente.to_excel(writer,"divergente",index=False)  

        writer.save()


        retorno = "0"+","
        return retorno
    except:
        retorno = "1"+","+str(traceback.format_exc())
        return retorno




#print(poco(lista))

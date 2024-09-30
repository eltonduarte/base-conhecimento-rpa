import os, glob, math
import pandas as pd
import traceback
import numpy as np
from openpyxl import load_workbook


caminho="@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho;Planilha_Planejamento_Contrato_RPA_15092021.xlsx"
caminho_save = "@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho"
nome_arquivo_principal= "Planilha_Planejamento_Contrato_RPA_15092021"
nome_arquivo_email = "CTO_divergente_15092021"
arquivo_rateio= "@petrobras.biz;petrobras;CENPES;CENPES_PDIDP_EPOCOS;NP-2;06. GestaoPDI;03. Paineis;Dados;Painel Consistencia - POCOS - 202101.xlsx"

lista = [caminho,caminho_save,nome_arquivo_principal,nome_arquivo_email,arquivo_rateio]


def poco(lista):
    try:

        pd.options.mode.chained_assignment = None

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

        df_pocos.rename(columns={'Mês':'Mês|poco', "Ano":"Ano|poco","Valor Pagamento":"Valor Pagamento|poco"}, inplace=True)

        #Mês , Ano e Valor Pagamento
        #df_planejamento["Valor Pagamento|str"] = df_planejamento["Valor Pagamento"].astype(str)

        """
        df_planejamento["multiplica"] = df_planejamento["Valor Pagamento|str"].str.replace("\.","",regex=True)

        df_planejamento["multiplica"] = df_planejamento["multiplica"].str.replace(",",".",regex=True)

        df_planejamento["multiplica"] = df_planejamento["multiplica"].astype("float64")

        """

        df_planejamento["mes tratado"] = df_planejamento["Mês"].str.replace("(?i)Mês|mes","",regex=True)

        #df_planejamento["mes tratado"] = df_planejamento["mes tratado"].astype("float64")

        #del df_planejamento["Valor Pagamento|str"]

        #df_planejamento.to_excel("C:\\Users\\cz5d\\Desktop\\kappa.xlsx")
        """
        df_merge_ok = pd.merge(df_planejamento,df_pocos,
         how= "inner",
         on= "Descrição")
        """
        #right_on=["Descrição","Parcela","Mês|poco","Ano|poco","Valor Parcela|poco"],
        #left_on=["Descrição","Parcela","mes tratado","Ano","multiplica"] )

        
        #df_merge_ok.to_excel("C:\\Users\\cz5d\\Desktop\\kappa_ok.xlsx")

        df_planejamento["Descrição"] = df_planejamento["Descrição"].astype(str)
        df_planejamento["Descrição"] = df_planejamento["Descrição"].str.replace("\s","")
        #df_planejamento["Parcela"]= df_planejamento["Parcela"].astype(str)
        #df_planejamento["Parcela"]= df_planejamento["Parcela"].str.replace("\s","")
        print(df_planejamento["Parcela"].dtype)
        df_planejamento["mes tratado"]= df_planejamento["mes tratado"].astype(str)
        df_planejamento["mes tratado"]= df_planejamento["mes tratado"].str.replace("\s","")


        """
        df_planejamento["Ano"]= df_planejamento["Ano"].astype(str)
        df_planejamento["Ano"]= df_planejamento["Ano"].str.replace("\s","")
        df_planejamento["Valor Pagamento"]= df_planejamento["Valor Pagamento"].astype(str)
        df_planejamento["Valor Pagamento"]= df_planejamento["Valor Pagamento"].str.replace("\s","")
        """

        df_pocos["Descrição"] = df_pocos["Descrição"].astype(str)
        df_pocos["Descrição"] = df_pocos["Descrição"].str.replace("\s","")
        #df_pocos["Parcela"]= df_pocos["Parcela"].astype(str)
        #df_pocos["Parcela"]= df_pocos["Parcela"].str.replace("\s","")
        df_pocos["Mês|poco"]= df_pocos["Mês|poco"].astype(str)
        df_pocos["Mês|poco"]= df_pocos["Mês|poco"].str.replace("\s","")
        """
        df_pocos["Ano|poco"]= df_pocos["Ano|poco"].astype(str)
        df_pocos["Ano|poco"]= df_pocos["Ano|poco"].str.replace("\s","")
        df_pocos["Valor Pagamento|poco"]= df_pocos["Valor Pagamento|poco"].astype(str)
        df_pocos["Valor Pagamento|poco"]= df_pocos["Valor Pagamento|poco"].str.strip()
        """
        df_pocos.to_excel("C:\\Users\\cz5d\\Desktop\\poco.xlsx")
        df_planejamento.to_excel("C:\\Users\\cz5d\\Desktop\\planejament.xlsx")

        df_merge_ok = df_planejamento.merge(df_pocos, how="inner",
        left_on=["Descrição","Parcela","mes tratado","Ano","Valor Pagamento"],
        right_on=["Descrição","Parcela","Mês|poco","Ano|poco","Valor Pagamento|poco"]).drop_duplicates(subset=["entrega"])

        df_merge_ok.to_excel("C:\\Users\\cz5d\\Desktop\\merge_ok.xlsx")

        df_planejamento.loc[(df_planejamento["Descrição"].isin(df_merge_ok["Descrição"]))  & 
            (df_planejamento["Parcela"].isin(df_merge_ok["Parcela"]))
            ,["Validacao Painel CTO"] ] = "OK"       
             
        
        """
        df_planejamento.loc[(df_planejamento["Descrição"].isin(df_merge_ok["Descrição"]))  & 
            (df_planejamento["Parcela"].isin(df_merge_ok["Parcela"])) &
            (df_planejamento["mes tratado"].isin(df_merge_ok["Mês|poco"])) &
            (df_planejamento["Ano"].isin(df_merge_ok["Ano|poco"])) &
            (df_planejamento["Valor Pagamento"].isin(df_merge_ok["Valor Pagamento|poco"])), ["Validacao Painel CTO"] ] = "OK"
        """
        """        

        df_merge_ok.to_excel("C:\\Users\\cz5d\\Desktop\\merge_ok.xlsx")

        df_merge_anti_mes = df_planejamento.merge(df_pocos,
            how="left",
            left_on="mes tratado",
            right_on="Mês|poco", indicator=True).query("_merge != 'both'")

        print(df_merge_anti_mes)

        df_merge_mes_divergente = pd.merge(df_merge_anti_mes,df_pocos,
            how="inner",
            left_on=["Descrição_x","Parcela_x","Ano","Valor Pagamento"],
            right_on=["Descrição","Parcela","Ano|poco","Valor Pagamento|poco"]).drop_duplicates(subset=["entrega"])

        df_merge_mes_divergente.to_excel("C:\\Users\\cz5d\\Desktop\\merge_mes_div.xlsx")  
        """      

        """
        for i in df_planejamento["Valor Pagamento"]:
            print(i)
        

        for i,j,k in zip(df_pocos["Mês|poco"], 
            df_pocos["Ano|poco"],
            df_pocos["Valor Pagamento|poco"]                    
            ):                     
                
                df_merge.loc[        
                (df_merge["mes tratado"]== i )&
                (df_merge["Ano"]==j) & 
                (df_merge["Valor Pagamento"]==k),"Validacao Painel CTO"] = "OK"              
                
                
        df_planejamento_restante = df_merge.loc[df_merge["Validacao Painel CTO"].isna()]

        df_planejamento_ok = df_merge.loc[df_merge["Validacao Painel CTO"] == "OK"]
            
        for i,j,k in zip(df_pocos["Mês|poco"], 
            df_pocos["Ano|poco"],
            df_pocos["Valor Pagamento|poco"]                     
            ):            
                
                df_planejamento_restante.loc[        
                (df_planejamento_restante["mes tratado"]!= i )&
                (df_planejamento_restante["Ano"]==j) & 
                (df_planejamento_restante["Valor Pagamento"]==k),"Validacao Painel CTO"] = "MES DIVERGENTE"              
                
       
        df_planejamento_mes = df_planejamento_restante.loc[df_planejamento_restante["Validacao Painel CTO"] == "MES DIVERGENTE"]   

        df_planejamento_div = df_planejamento_restante.loc[df_planejamento_restante["Validacao Painel CTO"].isna()]     

        df_merge = pd.concat([df_planejamento_div,df_planejamento_mes,df_planejamento_ok])          
                
        df_planejamento.loc[(df_planejamento["Descrição"].isin(df_merge["Descrição"]))  & 
            (df_planejamento["Parcela"].isin(df_merge["Parcela"])), ["Validacao Painel CTO"] ] = df_merge["Validacao Painel CTO"]      

          
        
        #print(df_planejamento["Validacao Painel CTO"])
        
        del df_planejamento["mes tratado"]

        #del df_planejamento["multiplica"]


        """

        del df_planejamento["mes tratado"]

        #df_planejamento.loc[df_planejamento["Validacao Painel CTO"].isna(),"Validacao Painel CTO"] = "DIVERGENTE"   

        #df_planejamento_restante.loc[df_planejamento_restante["Validacao Painel CTO"].isna(),"Validacao Painel CTO"] = "DIVERGENTE"

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




print(poco(lista))

import os, glob
import pandas as pd
import traceback
import numpy as np
from openpyxl import load_workbook

#arquivo_rateio= "@petrobras.biz;petrobras;CENPES;CENPES_PDIDP_EPOCOS;NP-2;06. GestaoPDI;03. Paineis;Dados;rateio_ent_processo.xlsx"

"""
caminho="@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho;Planilha_Planejamento_Contrato_RPA_15092021.xlsx"
caminho_save = "@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho"
nome_arquivo_principal= "Planilha_Planejamento_Contrato_RPA_15092021"
arquivo_rateio= "@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho;rateio_ent_processo.xlsx"

teste = [caminho,caminho_save,nome_arquivo_principal,arquivo_rateio]
"""


def filtro(lista):
    try:

        pd.options.mode.chained_assignment = None

        caminho = lista[0]
        caminho_save = lista[1]
        nome_arquivo_principal = lista[2]
        arquivo_rateio = lista[3]   

        caminho= caminho.replace("@","\\\\")
        caminho_lst= caminho.split(";")

        caminho_save= caminho_save.replace("@","\\\\")
        caminho_save_lst= caminho_save.split(";")

        arquivo_rateio= arquivo_rateio.replace("@","\\\\")
        arquivo_rateio_lst= arquivo_rateio.split(";")

        path_save = os.path.join(*caminho_save_lst)

        
        df_planejamento = pd.read_excel(os.path.join(*caminho_lst), sheet_name='Planejamento Plurianual')

        df_rateio = pd.read_excel(os.path.join(*arquivo_rateio_lst))

        

        df_rateio.rename(columns={'%Rateio':'%Rateio|tab_rateio'}, inplace=True)

        df_rateio.rename(columns={'Parcela':'Parcela|tab_rateio'}, inplace=True)
        
        #df_rateio.rename(columns={'Descrição':'Descrição|left'}, inplace=True)

        #.drop_duplicates(subset=["Descrição","Parcela","Mês","Ano"])
        """
        df_merge_duplicado= pd.merge(df_rateio,    
                    df_planejamento,
                    how="inner", 
                    left_on= ["Descrição","Parcela|tab_rateio"],
                    right_on=["Descrição","Parcela"])   
        """  

        df_merge= pd.merge(df_rateio,    
                    df_planejamento,
                    how="inner", 
                    left_on= ["Descrição","Parcela|tab_rateio"],
                    right_on=["Descrição","Parcela"]).drop_duplicates(subset=["Descrição","Mês","Ano"])


        #print(df_merge_duplicado)
        #print(df_merge)            
                    
                                      
                    
        #df_merge.to_excel("D:\\resultado py\\teste.xlsx", "SAP sem SIGITEC",index=False)

        df_merge["%Rateio|tab_rateio"]= df_merge["%Rateio|tab_rateio"].astype("str")


        df_merge.loc[(df_merge["%Rateio|tab_rateio"].str.match("0\.99*")),"%Rateio|tab_rateio" ] = "1"


        df_merge["%Rateio|tab_rateio"]= df_merge["%Rateio|tab_rateio"].astype("float64")


        #print(len(df_merge.index))

        #print(df_merge)

        

        for i in range(len(df_merge.index)):    

            df_planejamento.loc[(df_planejamento["Descrição"]== df_merge.iloc[i]["Descrição"] ),"entrega"] = df_merge.iloc[i]["Entrega"]

            df_planejamento.loc[(df_planejamento["Descrição"]== df_merge.iloc[i]["Descrição"] ),"tipo"] = df_merge.iloc[i]["Tipo"]
            
            df_planejamento.loc[(df_planejamento["Descrição"]== df_merge.iloc[i]["Descrição"]),"gasto"] = df_merge.iloc[i]["Gasto"]

            df_planejamento.loc[(df_planejamento["Descrição"]== df_merge.iloc[i]["Descrição"] ),"%Rateio"] = df_merge.iloc[i]["%Rateio|tab_rateio"]

            df_planejamento.loc[(df_planejamento["Descrição"]== df_merge.iloc[i]["Descrição"] ),"cod entrega"] = df_merge.iloc[i]["Cod. Entrega"]

            #df_planejamento.loc[(df_planejamento["Descrição"]== df_merge.iloc[i]["Parcela"] ),"Parcela"] = df_merge.iloc[i]["Parcela|tab_rateio"]


            

        
        


        df_planejamento.loc[df_planejamento["%Rateio"]==1,"Compartilhamento"] = "SIM"

        df_planejamento.loc[df_planejamento["%Rateio"]!=1,"Compartilhamento"] = "NÃO"

        df_compartilhamento =  df_planejamento
        #df_planejamento.loc[(df_planejamento["%Rateio"]!=1) & (df_planejamento["%Rateio"].notnull())|(df_planejamento["%Rateio"]==0)]

        #df_compartilhamento.to_excel("D:\\resultado py\\kappa.xlsx")

        df_rateio_realizado = pd.DataFrame(data=None, columns=df_planejamento.columns)

        
        nome_colunas = df_planejamento.columns.to_numpy()

        for i in range(len(df_compartilhamento.index)):

            array_temp = df_compartilhamento.iloc[i].to_numpy()

            filtro_tab_rateio = df_rateio.loc[(df_rateio["Descrição"]== array_temp[0])&(df_rateio["Parcela|tab_rateio"]== array_temp[8])]

            tamanho_df = len(filtro_tab_rateio)            

            if tamanho_df == 0 :
                

                mat_array_temp = np.tile(array_temp, (1,1))    
                           
                
                if not filtro_tab_rateio.empty:

                    mat_array_temp[0][12] = filtro_tab_rateio.iloc[0]["Cod. Entrega"]
                    mat_array_temp[0][13] = filtro_tab_rateio.iloc[0]["Entrega"]
                    mat_array_temp[0][14] = filtro_tab_rateio.iloc[0]["Tipo"]
                    mat_array_temp[0][15] = filtro_tab_rateio.iloc[0]["Gasto"]
                    mat_array_temp[0][16] = filtro_tab_rateio.iloc[0]["%Rateio|tab_rateio"]
                    #mat_array_temp[0][12] = filtro_tab_rateio.iloc[0]["Parcela"]
                               
                df_rateio_realizado = df_rateio_realizado.append(pd.DataFrame([mat_array_temp[0]], columns=nome_colunas), ignore_index=True)

                continue
                
            else:    
                mat_array_temp = np.tile(array_temp, (tamanho_df,1))                                 
            
            for j in range(len(mat_array_temp)):     
                           
                
                mat_array_temp[j][12] = filtro_tab_rateio.iloc[j]["Cod. Entrega"]
                mat_array_temp[j][13] = filtro_tab_rateio.iloc[j]["Entrega"]
                mat_array_temp[j][14] = filtro_tab_rateio.iloc[j]["Tipo"]
                mat_array_temp[j][15] = filtro_tab_rateio.iloc[j]["Gasto"]
                mat_array_temp[j][16] = filtro_tab_rateio.iloc[j]["%Rateio|tab_rateio"]
                #mat_array_temp[j][12] = filtro_tab_rateio.iloc[j]["Parcela"]

                               
                df_rateio_realizado = df_rateio_realizado.append(pd.DataFrame([mat_array_temp[j]], columns=nome_colunas), ignore_index=True)
                
                
        


        #df_compartilhamento_inv = df_planejamento.loc[(df_planejamento["%Rateio"]==1) | (df_planejamento["%Rateio"].isna())]

        

        #df_compartilhamento_inv.to_excel("D:\\resultado py\\kappa.xlsx")

        df_planejamento = df_rateio_realizado
        
        df_planejamento.loc[df_planejamento["Valor Parcela"] == "", "Valor Parcela"] = 0

        #df_planejamento["multiplica"] = df_planejamento["Valor Parcela"].str.replace("\.","",regex=True)

        #df_planejamento["multiplica"] = df_planejamento["multiplica"].str.replace(",",".",regex=True)
        
        df_planejamento["multiplica"] = df_planejamento["Valor Parcela"].str.replace(",","",regex=True)

        df_planejamento["Valor Pagamento"]= df_planejamento["multiplica"].astype("float") *df_planejamento["%Rateio"].astype("float")

        del df_planejamento["multiplica"]


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

        retorno = "0"+","
        return retorno

    except:
        

        retorno = "1"+","+str(traceback.format_exc())
        return retorno



#print(filtro(teste))


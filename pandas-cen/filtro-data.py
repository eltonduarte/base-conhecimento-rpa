import os, glob
import pandas as pd
import traceback
from openpyxl import load_workbook

"""

caminho="@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho;Planilha_Planejamento_Contrato_RPA_13052021.xlsx"
caminho_save = "@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho"
nome_arquivo_principal= "Planilha_Planejamento_Contrato_RPA_13052021"


teste = [caminho,caminho_save,nome_arquivo_principal]

"""

def filtro_data(lista):
    try:
        caminho = lista[0]
        caminho_save = lista[1]
        nome_arquivo_principal = lista[2]

        caminho= caminho.replace("@","\\\\")
        caminho_lst= caminho.split(";")

        caminho_save= caminho_save.replace("@","\\\\")
        caminho_save_lst= caminho_save.split(";")

        path_save = os.path.join(*caminho_save_lst)

        df_planejamento = pd.read_excel(os.path.join(*caminho_lst), sheet_name='Planejamento Plurianual')

        df_desembolso = pd.read_excel(os.path.join(*caminho_lst), sheet_name='Desembolso')


        df_merge= pd.merge(df_desembolso,    
            df_planejamento,
            how="inner", 
            left_on=["Número do Processo","Desembolsos Planejados | Número da Parcela"],
            right_on=["SIGITEC","Parcela"]
            )


        #print(df_merge)

        #df_merge.to_excel("C:\\Users\\cz5d\\Desktop\\teste_data.xlsx")


        df_merge_filtro_1 = df_merge.loc[df_merge["Desembolsos Planejados | Data de Pagamento"].str.match("(?i)Mês|mes")==True]

        df_merge_filtro_2 = df_merge[df_merge["Desembolsos Planejados | Data de Pagamento"].str.match("(?i)Mês|mes")==False]



        for i,j,k,v in zip(df_merge_filtro_1["Número do Processo"],
            df_merge_filtro_1["Desembolsos Planejados | Número da Parcela"],
            df_merge_filtro_1["Desembolsos Planejados | Data de Pagamento"],
            df_merge_filtro_1["Desembolsos Planejados | Valor Previsto"]):

            df_planejamento.loc[(df_planejamento["SIGITEC"]== i )&(df_planejamento["Parcela"]==j),"Mês"] = k
            df_planejamento.loc[(df_planejamento["SIGITEC"]== i )&(df_planejamento["Parcela"]==j),"Valor Parcela"] = v

        

        pd.options.mode.chained_assignment = None

        df_merge_filtro_2[["dia pagamento","mes pagamento","ano pagamento"]] = df_merge_filtro_2["Desembolsos Planejados | Data de Pagamento"].str.split("/",expand=True)



        for i,j,m,a,v in zip(df_merge_filtro_2["Número do Processo"],
                        df_merge_filtro_2["Desembolsos Planejados | Número da Parcela"],               
                        df_merge_filtro_2["mes pagamento"],
                        df_merge_filtro_2["ano pagamento"],
                        df_merge_filtro_2["Desembolsos Planejados | Valor Previsto"]):

            df_planejamento.loc[(df_planejamento["SIGITEC"]== i )&(df_planejamento["Parcela"]==j),"Mês"] = m
            df_planejamento.loc[(df_planejamento["SIGITEC"]== i )&(df_planejamento["Parcela"]==j),"Ano"] = a
            df_planejamento.loc[(df_planejamento["SIGITEC"]== i )&(df_planejamento["Parcela"]==j),"Valor Parcela"] = v


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


#print(filtro_data(teste))

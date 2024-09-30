import os, glob
import pandas as pd
import traceback
from openpyxl import load_workbook


"""
caminho="@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho;Planilha_Planejamento_Contrato_RPA_31052021.xlsx"
caminho_save = "@s6006vfs0601;prd_robotic_process_aa_002$;RPA_ROBOS;CENPES;CENPES01;Trabalho"
nome_arquivo_principal= "Planilha_Planejamento_Contrato_RPA_31052021"
arquivo_rateio= "@petrobras.biz;petrobras;CENPES;CENPES_PDIDP_EPOCOS;NP-2;06. GestaoPDI;03. Paineis;Dados;rateio_ent_processo.xlsx"

teste = [caminho,caminho_save,nome_arquivo_principal,arquivo_rateio]
"""

def filtro(lista):
    try:

        retorno = "0"+","
        return retorno
    except:

        retorno = "1"+","+str(traceback.format_exc())
        return retorno


-- STRING DE CONEXÃO PLANILHAS

-- Connection string para planilhas, Connection mode: Default
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$iStrCaminhoBase$;Extended Properties="Excel 12.0 Xml;HDR=YES;IMEX=1";

-- Com IMEX=0, a planilha aceita updates
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$variavel$;Extended Properties="Excel 12.0 Xml;HDR=YES;IMEX=0";

-- Para selects com indice da coluna (F1, F2) e por range de celulas.
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$iStrPlanilha$;Extended Properties="Excel 12.0 Xml;HDR=NO;IMEX=1";


-- SELECT

SELECT DISTINCT F1 AS Fornecedor, F2 AS Empresa
FROM [Exportação SAPUI5$$A2:K] 
WHERE ISNULL(F1) = False

Select [SST], Replace([Data], '.', '/') As Data, [Usuário]
From [Sheet1$$] 
Where Len([SST]) > 0

SELECT Distinct [Imobilizado], [Pasta_Controle_Imobiliario] AS Codigo FROM [Base$$] WHERE LEN([Imobilizado]) > 0 AND LEN([Pasta_Controle_Imobiliario]) = 6
  
Select * From [Base$$] Where [Imobilizado] Like '$iStrImobilizado$' and [Status] IN('Pendente', 'Falha') 
  
Select * From [$iStrNomeAba$$$] Where IsNull([Usuário]) = False
  
Select [Asset ID] From [$vListaAbas[1]$$$] Where Len([Asset ID]) > 0

Select Top 1 * From [$vListaAbas[3]$$$]

SELECT
IIF(ISNUll([Data de pagamento]) = False, Replace([Data de pagamento], '/', '.'),'') AS 'Data de pagamento', 
IIF(ISNULL([Data de compra]) = False, Replace([Data de compra], '/', '.'),'') AS 'Data de compra', 
IIF(ISNULL([Data Liquidacao]) = False, Replace([Data Liquidacao], '/', '.'),'') AS 'Data de Liquidacao', 
[Asset ID] AS 'Asset ID',
IIF(ISNULL([Valor de face]) = False, Format([Valor de face], '#,##0.00'),'') AS 'Valor de face', 
IIF(ISNULL([Preco]) = False, Format([Preco], '#,##0.00'),'') AS 'Preco', 
IIF(ISNULL([Parceiro]) = False, Ucase([Parceiro]),'') AS 'Parceiro', 
IIF(ISNULL([Juros acruados]) = False, Format([Juros acruados], '#,##0.00'),'') AS 'Juros acruados', 
IIF(ISNULL([Agio Desagio]) = False, Format([Agio Desagio], '#,##0.00'),'') AS 'Agio Desagio', 
IIF(ISNULL([Deposito]) = False, [Deposito], '') AS 'Deposito', 
[SAP RECOMPRA], [STATUS]
FROM [$vListaAbas[0]$$$] 
WHERE LEN([Asset ID]) > 0 

SELECT `Materiais Óleo` FROM [Parâmetros$$] WHERE [Materiais Óleo] IS NOT NULL or [Materiais Óleo] <> ""

SELECT DISTINCT CENTRO FROM [Entradas$$] WHERE len([CENTRO]) > 0 


-- UPDATE
  
Update [Base$$] Set [Status] = '$oStrStatus$', 
[Mensagem_SGTP] = '$oStrMensagemSgtp$',
[Data_Processamento_SGTP] = '$vStrDataExec$'
Where [Imobilizado] = '$vRcdImob{Imobilizado}$'

  
UPDATE [Entradas$$] 
SET [STATUS] = 'Transferido' 
WHERE [CENTRO] LIKE '$vLinhaCentros{CENTRO}$'
  

-- INSERT

Insert Into [Consolidado_ECOMEX$$] (SST, DATA, USUARIO, ARQUIVO, ID) Values(
'$vArrayConsulta{SST}$', 
'$vArrayConsulta{"Data"}$', 
'$vArrayConsulta{"Usuário"}$', 
'$vStrArquivo$.xlsx',
'$vNumId.Number:toString$')

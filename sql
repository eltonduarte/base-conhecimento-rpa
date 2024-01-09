-- Connection string para planilhas, Connection mode: Default
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$iStrCaminhoBase$;Extended Properties="Excel 12.0 Xml;HDR=YES;IMEX=1";

-- com IMEX=0, a planilha aceita updates
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$variavel$;Extended Properties="Excel 12.0 Xml;HDR=YES;IMEX=0";



-- Exemplos de updates
Update [Base$$] Set [Status] = '$oStrStatus$', 
[Mensagem_SGTP] = '$oStrMensagemSgtp$',
[Data_Processamento_SGTP] = '$vStrDataExec$'
Where [Imobilizado] = '$vRcdImob{Imobilizado}$'

  

-- Exemplos de consultas válidas

SELECT Distinct [Imobilizado], [Pasta_Controle_Imobiliario] AS Codigo 
FROM [Base$$] 
WHERE LEN([Imobilizado]) > 0 AND LEN([Pasta_Controle_Imobiliario]) = 6


Select * 
From [Base$$]
Where [Imobilizado] Like '$iStrImobilizado$' and [Status] IN('Pendente', 'Falha') 

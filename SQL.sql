-- Connection string para planilhas, Connection mode: Default
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$iStrCaminhoBase$;Extended Properties="Excel 12.0 Xml;HDR=YES;IMEX=1";








-- Exemplos de consultas válidas

SELECT Distinct [Imobilizado], [Pasta_Controle_Imobiliario] AS Codigo 
FROM [Base$$] 
WHERE LEN([Imobilizado]) > 0 AND LEN([Pasta_Controle_Imobiliario]) = 6


Select * 
From [Base$$]
Where [Imobilizado] Like '$iStrImobilizado$' and [Status] IN('Pendente', 'Falha') 

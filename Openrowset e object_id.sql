--Excel
FROM OPENROWSET
('Microsoft.ACE.OLEDB.12.0', 
'Excel 12.0;Database=E:\Impplan\FMF - Tabela Matriz x Desconto Pontualidade - update.XLS;HDR=YES', 
'SELECT * FROM [Sheet$]')

--txt/csv precisa do arquivo .ini
FROM OPENROWSET ('MICROSOFT.ACE.OLEDB.12.0', 'TEXT;DATABASE=E:\MSSQL\SAP\DASHBOARD;HDR=YES;FMT=DELIMITED(;)', 'SELECT * FROM [BSAD_1000.CSV]')

--CREATE TEMP 1
IF Object_id('tempdb..#ColigadaJob') IS NOT NULL 
BEGIN
	DROP TABLE #COLIGADAJOB
END

CREATE TABLE #COLIGADAJOB (

)

--CREATE TEMP 2
SELECT COLIGADA AS [CODCOLIGADA] 
       ,CODFILIAL 
INTO   #COLIGADAJOB 
FROM   [dbo].[Coligada_job]()


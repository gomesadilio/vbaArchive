Attribute VB_Name = "mdl_InsertAccessFromExcel"
Option Explicit

Option Private Module

'---------------------------------------------------------------------------------------
' By..........: SILVA, ADILIO
' Contact.....: gomesadilio@outlook.com
' Date........: 1/1/2021
' Description.: INSERT INTO full data
'---------------------------------------------------------------------------------------

'                         .
'                     /   ))     |\         )               ).
'               c--. (\  ( `.    / )  (\   ( `.     ).     ( (
'               | |   ))  ) )   ( (   `.`.  ) )    ( (      ) )
'               | |  ( ( / _..----.._  ) | ( ( _..----.._  ( (
' ,-.           | |---) V.'-------.. `-. )-/.-' ..------ `--) \._
' | /===========| |  (   |      ) ( ``-.`\/'.-''           (   ) ``-._
' | | / / / / / | |--------------------->  <-------------------------_>=-
' | \===========| |                 ..-'./\.`-..                _,,-'
' `-'           | |-------._------''_.-'----`-._``------_.-----'
'               | |         ``----''            ``----''
'               | |
'               c--`

Sub SentDataToAccess()

	If MsgBox("Agree?", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
  
    Dim rst             	As Object  
	Dim cnn 				As Object
    Dim sql             	As String
	Dim strProvider 		As String
	Dim strDatabase 		As String

	Set rst = CreateObject("ADODB.Recordset")
	Set cnn = CreateObject("ADODB.Connection")
	
	strDatabase = "C:\Users\MyUser\Desktop\database.accdb"

	strProvider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strDatabase

	sql = "DELETE * FROM TBL_MAIN WHERE 1 = 1; "
        
	cnn.Open strProvider

	cnn.Execute sql
        
    sql = __
		"INSERT INTO TBL_MAIN " & _ 
		"SELECT [Field1], [Field2] FROM [Sheet$] " & _
		" IN '" & ThisWorkbook.FullName & "' 'Excel 12.0;'", _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & strDatabase
        
    Set rst = CreateObject("ADODB.Recordset")
 
    rst.Open sql
	
    Set rst = Nothing
	set cnn = Nothing

End Sub                

Attribute VB_Name = "mdl_InsertAccessFromExcel"
Option Explicit

Option Private Module

'---------------------------------------------------------------------------------------
' Autor.....: ADILIO GOMES
' Contato...: gomesadilio@outlook.com
' Data......: 19/01/2021
' Descricao.: Get data from excel file or sql server database
'---------------------------------------------------------------------------------------

Sub SentDataToAccess()
  
    Dim rst             As Object  
    Dim sql             As String
        
    sql = _
        "INSERT INTO TBL_MAIN " & _
        "SELECT [Field1], [Field2], 
        "FROM [Sheet$]"
        
    Set rst = CreateObject("ADODB.Recordset")
 
    rst.Open _
        sql & " IN '" & ThisWorkbook.FullName & "' 'Excel 12.0;'", _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=C:\Users\MyUser\Desktop\database.accdb"
 
    Set rst = Nothing

End Sub                

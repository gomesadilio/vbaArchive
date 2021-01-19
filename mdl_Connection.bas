Attribute VB_Name = "mdl_Connection"
Option Explicit

Option Private Module

'---------------------------------------------------------------------------------------
' Autor.....: ADILIO GOMES
' Contato...: gomesadilio@outlook.com
' Data......: 19/01/2021
' Descricao.: Get data from excel file or sql server database
'---------------------------------------------------------------------------------------

Sub get_data_from_table( _
        rng As Range, _
            sql As String, _
                header As Boolean, _
                    dsn As String, _
                        Optional uid As String, _
                            Optional pwd As String)
  
    On Error GoTo ErrorHandler
    
    Dim cnn                     As Object
    Dim rst                     As Object
    Dim i                       As Long
    
    Set cnn = CreateObject("ADODB.Connection")
    Set rst = CreateObject("ADODB.Recordset")
    
    If Not VBA.InStr(1, sql, "$") Then
        
        'Excel 2013+ provider
        cnn.Provider = "Microsoft.ACE.OLEDB.12.0"
        
        'Workbook full name
        cnn.ConnectionString = "Data Source=" & dsn
        
        cnn.Properties("Extended Properties") = "Excel 12.0 Xml;HDR=YES"
        
    Else
        
        'ODBC for sql server sgbd
        cnn.ConnectionString = _
            "DSN=   " & dsn & "; DRIVER=Client Access ODBC Driver (32-bit); " & _
            "UID =  " & uid & _
            ";PWD = " & pwd
    
    End If
  
  cnn.Open
  
  Set rst = cn.Execute(sql)
  
  i = rst.Fields.Count
  
  While header
    If i = 0 Then header = False Else rng.Cells(1, i) = rst.Fields(i - 1).name: i = i - 1
  Wend
  
  rng.Offset(1, 0).CopyFromRecordset rst
  
  rst.Close
  cnn.Close
  
  Exit Sub
  
ErrorHandler:
  
  MsgBox Err.Description, vbOKOnly + vbCritical
  
  If cnn.State = 1 Then cnn.Close
  
End Sub

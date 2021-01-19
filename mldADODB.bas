Attribute VB_Name = "mldADODB"
Option Explicit

Option Private Module

Sub GetData(rng As Range, sql As String, header As Boolean, dsn As String, Optional uid As String, Optional pwd As String)
  
    On Error GoTo ErrorHandler
    
    Dim cn                                 As Object
    Dim rs                                 As Object
    Dim i                                  As Long
    
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    If Not VBA.InStr(1, sql, "$") Then
        
        cn.Provider = "Microsoft.ACE.OLEDB.12.0"
        cn.ConnectionString = "Data Source=" & dsn
        cn.Properties("Extended Properties") = "Excel 12.0 Xml;HDR=YES"
        
    Else
        
        cn.ConnectionString = _
            "DSN=   " & dsn & "; DRIVER=Client Access ODBC Driver (32-bit); " & _
            "UID =  " & uid & _
            ";PWD = " & pwd
    
    End If
  
  cn.Open
  
  Set rs = cn.Execute(sql)
  
  i = rs.Fields.Count
  
  While header
    If i = 0 Then header = False Else rng.Cells(1, i) = rs.Fields(i - 1).name: i = i - 1
  Wend
  
  rng.Offset(1, 0).CopyFromRecordset rs
  
  rs.Close
  cn.Close
  
  Exit Sub
  
ErrorHandler:
  
  MsgBox Err.Description, vbOKOnly + vbCritical
  
  If cn.State = 1 Then cn.Close
  
End Sub

Attribute VB_Name = "mdl_Refresh"
Option Explicit

Option Private Module

'---------------------------------------------------------------------------------------
' Autor.....: ADILIO GOMES
' Contato...: gomesadilio@outlook.com
' Data......: 25/01/2021
' Descricao.: Refresh all connections in tables and pivot tables
'---------------------------------------------------------------------------------------

Sub refresh_reports()
Attribute refresh_reports.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Dim sht             As Worksheet
    Dim pvt             As PivotTable
    Dim lst             As ListObject
    
    Application.StatusBar = "Wait finish...."
    
    For Each sht In ThisWorkbook.Worksheets
    
        For Each lst In sht.ListObjects
        
            lst.QueryTable.Refresh False
            
        Next
        
    Next

    For Each sht In ThisWorkbook.Worksheets
    
        For Each pvt In sht.PivotTables
        
            pvt.RefreshTable
            
        Next
        
    Next
    
    Application.StatusBar = ""
    
    MsgBox ("Updated!"), vbInformation

End Sub

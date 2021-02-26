Attribute VB_Name = "mdl_ExportGraphAsImage"
Option Explicit

Option Private Module

'---------------------------------------------------------------------------------------
' Autor.....: ADILIO GOMES
' Contato...: gomesadilio@outlook.com
' Data......: 2021
' Descricao.: Export as image
'---------------------------------------------------------------------------------------
 
Sub ExportGraphAsImage()   
    
    Dim objChrt             As ChartObject
    Dim myChart             As Chart
    Dim myFileName          As String
    Dim i                   As Byte
    
    Calculate
    
    For i = 1 To Sheets("Aux").Shapes.Count
        
        Set objChrt = Sheets("Aux").ChartObjects(i)
    
        objChrt.Activate
        
        Set myChart = objChrt.Chart

        myFileName = "myChart" & i & ".jpeg"

        On Error Resume Next
        Kill ThisWorkbook.Path & "\" & myFileName
        On Error GoTo 0

        myChart.Export Filename:=ThisWorkbook.Path & "\" & myFileName, Filtername:="JPEG"

    Next
   
End Sub

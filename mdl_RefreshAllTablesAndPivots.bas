Attribute VB_Name = "mdl_RefreshAllTablesAndPivots"
Option Explicit

Option Private Module

'---------------------------------------------------------------------------------------
' By..........: SILVA, ADILIO
' Contact.....: gomesadilio@outlook.com
' Date........: 1/1/2021
' Description.: Refresh data from report
'---------------------------------------------------------------------------------------

'(\ 
'\'\ 
' \'\     __________  
' / '|   ()_________)
' \ '/    \ ~~~~~~~~ \
'   \       \ ~~~~~~   \
'   ==).      \__________\
'  (__)       ()__________)

Sub RefreshAllTablesAndPivots()

    With Application
        .ScreenUpdating = False
        .StatusBar = "Wait...updating..."
    End With
    
    Dim sht             	As Worksheet
    Dim pvt             	As PivotTable
    Dim lastRow    		As Long
    Dim oList           	As Object

    'Refreshing tables
    For Each sht In ThisWorkbook.Worksheets
        For Each oList In sht.ListObjects
            oList.QueryTable.Refresh BackgroundQuery:=False
        Next
    Next

    'Refreshing pivots tables
    For Each sht In ThisWorkbook.Worksheets
        For Each pvt In sht.PivotTables
            pvt.RefreshTable
        Next
    Next

    lastRow = shtData.Cells(4, shtData.Cells.Columns.Count).End(xlToLeft).Column + 1    
    
    shtResume.Range("I5:I15").Copy
    shtData.Cells(5, lastRow).Resize(10, 1).PasteSpecial xlPasteValues    
    
    With Application
        .ScreenUpdating = True
        .StatusBar = ""
    End With
    
    Application.CutCopyMode = False
    
    shtManager.Activate

End Sub

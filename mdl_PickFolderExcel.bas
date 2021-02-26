Attribute VB_Name = "mdl_PickFolderExcel"
Option Explicit

Option Private Module

'---------------------------------------------------------------------------------------
' By..........: SILVA, ADILIO
' Contact.....: gomesadilio@outlook.com
' Date........: 1/1/2021
' Description.: Pick a folder
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

Public Function PickFolderExcel() As String

    ChDrive "Z"
    ChDir "Z:\MyFolder\"

    With Application.FileDialog(msoFileDialogFolderPicker)
        
		.AllowMultiSelect = False
        
		.InitialFileName = "Z:\MyFolder\"
        
		If .Show = -1 Then ' if OK is pressed
            pickFolder = .SelectedItems(1)
        End If
		
    End With
   
End Function 

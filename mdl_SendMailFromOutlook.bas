Attribute VB_Name = "mdl_SendMailFromOutlook"
Option Explicit

Option Private Module

'---------------------------------------------------------------------------------------
' By..........: SILVA, ADILIO
' Contact.....: gomesadilio@outlook.com
' Date........: 1/1/2021
' Description.: Send mail from outlook with signature, attachments, image, hyperlink
'---------------------------------------------------------------------------------------


Sub send_outlook_mail_with_signature()
       
    Dim objOutlook 			As Object        				'Dim objOutlook As Outlook.Application
    Dim objMessage 			As Object        				'Dim objMessage As Outlook.MailItem
    Dim htmlText 			As String
    
    Set objOutlook = CreateObject("Outlook.Application")    	'New Outlook.Application	
    Set objMessage = objOutlook.CreateItem(0)               	'olMailItem
    
    With objMessage
        
        .Display
        
        'Remittee
        .To = "<somebody@mymail.com>; "
        
        .CC = "<people@mymail.com>; "
		
		.BCC = "<anotherpeople@mymail.com>; "
       
        .Subject = "Reports - " & VBA.Format$(Now, "dd/mm/yy")
        
        .Attachments.Add ThisWorkbook.FullName
       
        htmlText = _
            "<!DOCTYPE html><html><body>" & _
            "<p font-family:'Segoe UI', Calibri, Arial, Helvetica; font-size: 20px;>" & _
            "Hello,<br><br>Dear,<br><br>" & _
            "Follow Report.</p>" & _
            "<h3 font-family:'Segoe UI', Calibri, Arial, Helvetica;>" & _
            "<a color: blue; href='" & ThisWorkbook.FullName & "'>" & _
            "Click to open file!</a></h3>" & _
            "<img src='" & "C:\Users\MyUser\Desktop\MyImage.jpeg" & "' width ='1200' ><br><br>" & _                            
            "</p></body></html>"
            
        .HTMLBody = htmlText & .HTMLBody
        
    End With
    
    Set objMessage = Nothing
    Set objOutlook = Nothing
    
End Sub
                     

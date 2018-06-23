# ExcelMacros
Automatic send email if change is not updated more then 7 days

Sub SendEmail()

Dim olApp As Outlook.Application
Set olApp = CreateObject("Outlook.Application")

    Dim olMail As Outlook.MailItem
    Set olMail = olApp.CreateItem(olMailItem)
    
    olMail.To = "alexandr.hadjioglo@mail.ru"
    olMail.Subject = "test"
    olMail.Body = "Please update documentation"
    olMail.Send

End Sub

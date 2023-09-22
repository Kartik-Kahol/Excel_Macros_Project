Attribute VB_Name = "Module4"
Sub Send_File()
Attribute Send_File.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Send_File Macro
'

'

Dim OutApp As Object
Dim OutMail As Object
Dim strbody As String



Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

strbody = ""

On Error Resume Next
With OutMail
.To = "Rahul@gmail.com; Carl@gmail.com; rohit@gmail.com"
.CC = "manisha@gmail.com"
.BCC = ""
.Subject = "Main_File_1" & Format(Date, " mm/dd/yyyy")
.Body = "Hi," & vbNewLine & vbNewLine & _
"Please find attached reports." & vbNewLine & vbNewLine & _
"Regards" & vbNewLine & _
"Kartik Kahol"
.Display
.Attachments.Add " **Locationofattachment** "
.Attachments.Add " **Locationofattachment** "

End With
On Error GoTo 0

Set OutMail = Nothing








Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

strbody = ""

On Error Resume Next
With OutMail
.To = "Abhinav@gmail.com"
.CC = "Shreya@gmail.com"
.BCC = ""
.Subject = "Main_File_2" & Format(Date, " mm/dd/yyyy")
.Body = "Hi," & vbNewLine & vbNewLine & _
"Please find attached reports." & vbNewLine & vbNewLine & _
"Regards" & vbNewLine & _
"Kartik Kahol"
.Display
.Attachments.Add " **Locationofattachment** "

End With
On Error GoTo 0

Set OutMail = Nothing


Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

strbody = ""

On Error Resume Next
With OutMail
.To = "Nishank@gmail.com"
.CC = "Shreyas@gmail.com"
.BCC = ""
.Subject = "Main_File_3" & Format(Date, " mm/dd/yyyy")
.Body = "Hi," & vbNewLine & vbNewLine & _
"Please find attached reports." & vbNewLine & vbNewLine & _
"Regards" & vbNewLine & _
"Kartik Kahol"
.Display
.Attachments.Add " **Locationofattachment** "

End With
On Error GoTo 0

Set OutMail = Nothing



Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

strbody = ""

On Error Resume Next
With OutMail
.To = "Kevin@gmail.com "
.CC = "Sam@gmail.com"
.BCC = ""
.Subject = "Main_File_4" & Format(Date, " mm/dd/yyyy")
.Body = "Hi," & vbNewLine & vbNewLine & _
"Please find attached reports." & vbNewLine & vbNewLine & _
"Regards" & vbNewLine & _
"Kartik Kahol"
.Display
.Attachments.Add " **Locationofattachment** "

End With
On Error GoTo 0

Set OutMail = Nothing

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

strbody = ""

On Error Resume Next
With OutMail
.To = "Tom@gmail.com"
.CC = "Rayan@gmail.com"
.BCC = ""
.Subject = "Main_File_5" & Format(Date, " mm/dd/yyyy")
.Body = "Hi," & vbNewLine & vbNewLine & _
"Please find attached reports." & vbNewLine & vbNewLine & _
"Regards" & vbNewLine & _
"Kartik Kahol"
.Display
.Attachments.Add " **Locationofattachment** "
.Attachments.Add " **Locationofattachment** "

End With
On Error GoTo 0

Set OutMail = Nothing
End Sub

Sub WhatsAppMsg()
Dim LastRow As Long
Dim i As Integer
Dim strip As String
Dim strPhoneNumber As String
Dim strmessage As String
Dim strPostData As String
Dim IE As Object

LastRow = Range("A" & Rows.Count).End(xlUp).Row
For i = 2 To LastRow

strPhoneNumber = Sheets("Data").Cells(i, 1).Value
strmessage = Sheets("Data").Cells(i, 2).Value
ActiveSheet.Shapes(1).Copy

'IE.navigate "whatsapp://send?phone=phone_number&text=your_message"

strPostData = "whatsapp://send?phone=" & strPhoneNumber & "&text=" & strmessage
Set IE = CreateObject("InternetExplorer.Application")

IE.navigate strPostData
Application.Wait (Now + TimeValue("00:00:05"))

Call SendKeys("{Enter}", True)
'Application.Wait Now() + TimeSerial(0, 0, 5)
'SendKeys "~"

Next i
End Sub

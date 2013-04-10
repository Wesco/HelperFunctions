Attribute VB_Name = "Examples"
Option Explicit

Sub Gaps()
    On Error GoTo ErrHandler
    ImportGaps
    On Error GoTo 0
    Exit Sub

ErrHandler:
End Sub

Sub SendMail()
    Email SendTo:="treische@wesco.com", _
          Subject:="Attachment test", _
          Body:="Multiple attachment test", _
          Attachment:=Array("C:\Users\treische\Desktop\007.png", "C:\Users\treische\Desktop\trollface.png")
End Sub

Sub test()
    Import117byISN ReportType.DS, Sheets("Sheet2").Range("A1")
End Sub

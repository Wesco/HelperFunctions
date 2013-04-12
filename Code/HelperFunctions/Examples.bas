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
    Email SendTo:="treischewesco.com", _
          Subject:="Attachment test", _
          Body:="Multiple attachment test", _
          Attachment:=Array("C:\Users\treische\Desktop\007.png", "C:\Users\treische\Desktop\trollface.png")
End Sub

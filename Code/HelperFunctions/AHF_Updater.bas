Attribute VB_Name = "AHF_Updater"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : CheckForUpdates
' Date : 4/24/2013
' Desc : Checks to see if the macro is up to date
'---------------------------------------------------------------------------------------
Sub CheckForUpdates(URL As String)
    Dim Ver As String
    Dim LocalVer As String
    Dim Path As String
    Dim LocalPath As String
    Dim FileNum As Integer
    Dim RegEx As Variant

    Set RegEx = CreateObject("VBScript.RegExp")
    Ver = Left(DownloadTextFile(URL), 5)
    RegEx.Pattern = "[0-9]\.[0-0]\.[0-9]"
    Path = GetWorkbookPath & "Version.txt"
    FileNum = FreeFile

    Open Path For Input As #FileNum
    Line Input #FileNum, LocalVer
    Close FileNum

    If RegEx.test(Ver) Then
        If CInt(Replace(Ver, ".", "")) > CInt(Replace(LocalVer, ".", "")) Then
            MsgBox Prompt:="An update is available. Please close the macro and get the latest version!", Title:="Update Available"
        End If
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : DownloadTextFile
' Date : 4/25/2013
' Desc : Returns the contents of a text file from a website
'---------------------------------------------------------------------------------------
Private Function DownloadTextFile(URL As String) As String
    Dim success As Boolean
    Dim responseText As String
    Dim oHTTP As Variant

    Set oHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")

    oHTTP.Open "GET", URL, False
    oHTTP.Send
    success = oHTTP.WaitForResponse()

    If Not success Then
        DownloadTextFile = ""
        Exit Function
    End If

    responseText = oHTTP.responseText
    Set oHTTP = Nothing

    DownloadTextFile = responseText
End Function

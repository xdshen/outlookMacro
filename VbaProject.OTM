Option Explicit

Private Declare PtrSafe Function GetPrivateProfileString _
    Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function WritePrivateProfileString _
    Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
        ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
        ByVal lpString As Any, ByVal lpFileName As String) As Long


' read
Function ReadStringFromIni(ByVal fileName As String, ByVal section As String, ByVal key As String) As String
    Dim x As Long
    Dim xBuff As String * 1024
    GetPrivateProfileString section, key, "", xBuff, 1024, fileName
    x = InStr(xBuff, Chr(0))
    ReadStringFromIni = Trim(Left(xBuff, x - 1))
End Function
 
' write
Sub WriteStringToIni(ByVal fileName As String, ByVal section As String, ByVal key As String, ByVal value As String)
    Dim xBuff As String * 1024
    xBuff = value + Chr(0)
    WritePrivateProfileString section, key, xBuff, fileName
End Sub


Function getIniFile() As String
    Dim ini_file_path As String
    ini_file_path = GetSetupPath("outlook.exe") & "\outlookFilter.ini"
    'if ini file is not existed then create one
    If Dir(ini_file_path) = "" Then
        Dim Fso, sFile
        Set Fso = CreateObject("Scripting.FileSystemObject")
        Set sFile = Fso.CreateTextFile(ini_file_path, True)
        Call sFile.Write(defaultIniFileContent)
        sFile.Close
        Set Fso = Nothing
        Set sFile = Nothing
    End If
    getIniFile = ini_file_path
End Function

Function defaultIniFileContent() As String
    defaultIniFileContent = "[attachmentExtFilter]" & vbCrLf & _
                            "# list the file extensions which is allowed, comma seperated for each" & vbCrLf & _
                            "ALLOWED_TYPE = pdf, doc, docx, zip, rar" & vbCrLf & _
                            "# list the file extensions which is not allowed, comma seperated for each" & vbCrLf & _
                            "NOT_ALLOWED_TYPE = xls,xlsx,xlsm, ppt, pptx" & vbCrLf & _
                            "[operations]" & vbCrLf & _
                            "# if you don't want to display any yes/no choice and directly block the email if the attachement type is not allowed, set to following parameter to false" & vbCrLf & _
                            "YESNO_WHEN_NOT_ALLOWED_TYPE = true"

End Function
Function GetIniVal(ByVal section As String, ByVal key As String)
    Dim ini_file_path As String
    ini_file_path = getIniFile()
    GetIniVal = ReadStringFromIni(ini_file_path, section, key)
End Function

Sub SetIniVal(ByVal section As String, ByVal key As String, ByVal value As String)
    Dim ini_file_path As String
    ini_file_path = getIniFile()
    Call WriteStringToIni(ini_file_path, section, key, value)
End Sub

Function GetSetupPath(ByVal AppName As String)
    Dim Wsh As Object
    Set Wsh = CreateObject("Wscript.Shell")
    GetSetupPath = Wsh.RegRead("HKEY_LOCAL_MACHINE\Software" _
        & "\Microsoft\Windows\CurrentVersion\App Paths\" _
        & AppName & "\Path")
    Set Wsh = Nothing
End Function

Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)

On Error Resume Next

Dim message As Outlook.MailItem

Set message = Item


If Not checkAttachmentExtFilter(message) Then

    Cancel = True
    
    Exit Sub

End If

End Sub

Private Function checkAttachmentExtFilter(message As Outlook.MailItem) As Boolean
    checkAttachmentExtFilter = True ' True means everything is OK.
    If message.Attachments.Count > 0 Then
        Dim allowedType As String, notAllowedType As String
        allowedType = GetIniVal("attachmentExtFilter", "ALLOWED_TYPE")
        notAllowedType = GetIniVal("attachmentExtFilter", "NOT_ALLOWED_TYPE")
        Dim notAllowedFileList As New Collection
        If Not notAllowedType = "" Then
            Dim i As Integer
            For i = 1 To message.Attachments.Count
                If isExtFilterList(message.Attachments.Item(i).fileName, notAllowedType) <> "" Then
                    notAllowedFileList.Add (message.Attachments.Item(i).fileName)
                End If
            Next i
        End If
        If notAllowedFileList.Count > 0 Then
            Dim answer As VbMsgBoxResult
            answer = vbNo
            Dim warningString As String
            Dim s
            For Each s In notAllowedFileList
                warningString = warningString & s & vbCrLf
            Next
            If LCase(GetIniVal("operations", "YESNO_WHEN_NOT_ALLOWED_TYPE")) = "true" Then

                warningString = "[Alert]The following files are not allowed to be sent:" & vbCrLf & warningString & vbCrLf & "Are you sure you want it to be sent? seriously!"
                answer = MsgBox(warningString, vbYesNo + vbQuestion, "Microsoft Office Outlook")
            Else
                warningString = "[Warning]The following files are not allowed to be sent:" & vbCrLf & warningString
                MsgBox warningString, vbCritical
            End If
            
            If Not (answer = vbYes) Then checkAttachmentExtFilter = False
        End If
    End If
End Function

Private Function isExtFilterList(ByVal fileName As String, ByVal listString As String) As String
    isExtFilterList = ""
    If Not listString = "" Then
        Dim delimiter
        delimiter = Split(listString, ",")
        Dim ext
        For Each ext In delimiter
            If InStrRev(fileName, "." & Trim(ext)) > 0 Then
                isExtFilterList = Trim(ext)
                Exit Function
            End If
        Next
    End If
End Function



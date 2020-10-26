Attribute VB_Name = "Sheet2Funcs"
Function FillLink(rngLink As Range, sPath As String, sLinkText As String, Optional sKeyword As String)
    FillLink = False
    Dim sTestDir As String 'network Dir returns error if empty
    On Error GoTo 0 'errhandler
    If InStr(UCase(sPath), "HTTP") > 0 Then 'online file
        If Testlink(sPath) Then 'sees if link is valid (valid=True)
            rngLink.Formula = "=HYPERLINK(" & """" & sPath & """" & "," & """" & sLinkText & """" & ")"
            rngLink.Font.Color = vbBlue
            rngLink.Font.Underline = xlUnderlineStyleSingle
            rngLink.Font.Strikethrough = False
        Else
            rngLink.Value = sLinkText
            rngLink.Font.Color = vbBlack
            rngLink.Font.Underline = xlUnderlineStyleNone
            rngLink.Font.Strikethrough = True
        End If
    ElseIf InStr(UCase(sPath), "EPDM") > 0 Then 'vault file
        Dim sTest As String
        sTest = GetLatest(sPath) 'looks for filepath using vault/get latest
        If sTest <> "" Then
            rngLink.Formula = "=HYPERLINK(" & """" & sTest & """" & "," & """" & sLinkText & """" & ")"
            rngLink.Font.Color = vbBlue
            rngLink.Font.Underline = xlUnderlineStyleSingle
            rngLink.Font.Strikethrough = False
        Else
            rngLink.Value = sLinkText
            rngLink.Font.Color = vbBlack
            rngLink.Font.Underline = xlUnderlineStyleNone
            rngLink.Font.Strikethrough = True
        End If
    Else 'folder directory file
        On Error Resume Next
        sTestDir = Dir(sPath, vbDirectory)
        If Err.Number <> 0 Then Debug.Print sPath
        Err.Clear
        On Error GoTo errhandler
        If (sTestDir <> "" And UCase(sKeyword) = "FOLDER") Or Dir(sPath) <> "" Then 'fails if SharePoint not logged in
            'On Error Resume Next 'why is this necessary??
            rngLink.Formula = "=HYPERLINK(" & """" & sPath & """" & "," & """" & sLinkText & """" & ")"
            rngLink.Font.Color = vbBlue
            rngLink.Font.Underline = xlUnderlineStyleSingle
            rngLink.Font.Strikethrough = False
        Else
            Dim sTmpPath As String
            sTmpPath = sPath
            If InStr(UCase(sPath), "SHAREPOINT") > 0 Then 'sharepoint location
                If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
                sTmpPath = sPath & "*" & sKeyword & "*" 'add "\" & skeyword in case it's a containing folder
            End If
            If Dir(sTmpPath) <> "" Then 'only happens for sharepoint links
                rngLink.Formula = "=HYPERLINK(" & """" & sPath & Dir$(sTmpPath) & """" & "," & """" & sLinkText & """" & ")"
                rngLink.Font.Color = vbBlue
                rngLink.Font.Underline = xlUnderlineStyleSingle
                rngLink.Font.Strikethrough = False
            Else
                rngLink.Value = sLinkText
                rngLink.Font.Color = vbBlack
                rngLink.Font.Underline = xlUnderlineStyleNone
                rngLink.Font.Strikethrough = True
            End If
        End If
    End If
    Exit Function
errhandler:
    Call ErrorRep("FillLink", "Function", "N/A", Err.Number, Err.Description, "")
End Function

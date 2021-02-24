Attribute VB_Name = "Sheet1Funcs"
Public arrRngInput() As Range 'used in subs on sheet 1 & in function below
Public rngBrt As Range, rngCar As Range, rngMat As Range 'used in subs on sheet 1 & in function below
Public rngPhonePer As Range, rngPhoneRm As Range 'used in subs on sheet 1 & in function below
Public rngSN As Range, rngCoInit As Range, rngCOlat As Range
Public rngLayout As Range, rngCust As Range, rngModel As Range
Public rngPhoneName As Range, rngPhoneExt As Range
Public rngFast As Range, rngLinks As Range
Function Sheet1Vars()
    '''''''hardcoded''''''''''''
    Set rngBrt = Sheet1.Range("A2")
    Set rngCar = Sheet1.Range("B2")
    Set rngMat = Sheet1.Range("C2")
    Set rngPhonePer = Sheet1.Range("F2")
    Set rngPhoneRm = Sheet1.Range("F3")
    Set rngFast = Sheet1.Range("G10")
    Set rngLinks = Sheet1.Range("E10")
    ReDim arrRngInput(4) 'collection of ranges that cause a macro when changed
    Set arrRngInput(0) = rngBrt
    Set arrRngInput(1) = rngCar
    Set arrRngInput(2) = rngMat
    Set arrRngInput(3) = rngPhonePer
    Set arrRngInput(4) = rngPhoneRm
    Set rngSN = Sheet1.Range("A6")
    Set rngCoInit = Sheet1.Range("B6")
    Set rngCOlat = Sheet1.Range("C6")
    Set rngLayout = Sheet1.Range("A8")
    Set rngCust = Sheet1.Range("B8")
    Set rngModel = Sheet1.Range("A10")
    Set rngPhoneName = Sheet1.Range("E6")
    Set rngPhoneExt = Sheet1.Range("E8")
    ''''''''''''''''''''''''''''
End Function
Function DetectInputType(sInput As String, sProdLine As String) As String()
'outputs CO, SERIAL, CUST, or NONE ; and outputs the value (for CO, would output 434343 instead of CO434343)
'shouldn't ever output "NONE" unless an error occurs (?)
    'On Error GoTo errhandler
    Dim arrOutput(1) As String
    arrOutput(0) = "NONE" 'default
    If UCase(sInput) Like "CO?######" Or UCase(sInput) Like "CO######" Then 'strip CO and use that
        sInput = Right(sInput, 6) 'sinput is now CO number
    End If
    If IsNumeric(sInput) And Len(sInput) = 6 Then 'probably CO
        If Trim(CInt(Left(sInput, 1))) > 1 And Trim(CInt(Left(sInput, 1))) < 5 Then
            arrOutput(0) = "CO"
        End If
    End If
    If arrOutput(0) = "NONE" Then
        If UCase(sProdLine) = "BURT" Then
            If IsNumeric(sInput) And Len(sInput) > 2 Then 'probably SN
                arrOutput(0) = "SERIAL"
            Else 'may be SN, may be customer
                If UCase(sInput) Like "*-###" Then
                    arrOutput(0) = "SERIAL"
                    sInput = Right(sInput, 3)
                Else
                    arrOutput(0) = "CUST"
                End If
            End If
        ElseIf UCase(sProdLine) = "CARR" Then
            If IsNumeric(sInput) Then
                If (Len(sInput) > 2 And Len(sInput) < 5) Or (Len(sInput) > 6 And Len(sInput) < 9) Then 'probably SN
                    arrOutput(0) = "SERIAL"
                    If Len(sInput) = 3 Then sInput = "0" & sInput
                End If
            Else 'may be SN, may be customer
                If UCase(sInput) Like "C#######*" Then 'SN -> Centritech
                    sInput = Mid(sInput, 2, 7)
                    arrOutput(0) = "SERIAL"
                ElseIf UCase(sInput) Like "C####*" Then 'SN -> Carr
                    sInput = Mid(sInput, 2, 4)
                    arrOutput(0) = "SERIAL"
                Else 'Customer
                    arrOutput(0) = "CUST"
                End If
            End If
        ElseIf UCase(sProdLine) = "MATEER" Then
            If IsNumeric(sInput) Then
                If Len(sInput) = 5 Or (Len(sInput) = 6 And Trim(CInt(Left(sInput, 1))) = 8) Then
                    arrOutput(0) = "SERIAL"
                Else 'not SN, not CO
                    arrOutput(0) = "CUST" 'probably not, but might as well look for matching customers
                End If
            Else 'non-numeric
                If UCase(sInput) Like "W####-#####*" Or UCase(sInput) Like "SN #####" Or _
                    UCase(sInput) Like "SN#####" Or UCase(sInput) Like "W######" Then 'strip SN and use that
                    arrOutput(0) = "SERIAL"
                    If InStr(sInput, "-") > 0 Then
                        sInput = Mid(sInput, InStr(sInput, "-"), 5)
                    Else 'replace non numerical characters with nothing -> S, N, W, C," "
                        sInput = Trim(Replace(Replace(Replace(Replace(Replace(UCase(sInput), "S", ""), "N", ""), "W", ""), " ", ""), "C", ""))
                    End If
                Else 'not SN, not CO
                    arrOutput(0) = "CUST" 'probably not, but might as well look for matching customers
                End If
            End If
        End If
    End If
    arrOutput(1) = sInput
    DetectInputType = arrOutput
    Exit Function
    
errhandler:
    MsgBox "Error in DetectInputType function"
    Call ErrorRep("DetectInputType", "Function", DetectInputType, Err.Number, Err.Description, "")
End Function

Function GivenInfo(sInfoOrig As String, sInfoType As String, sProdLine As String) As String()
'(0)=SN, (1)=initial CO, (2)=latest CO, (3)=customer, (4)=model
    
    Dim arrOutput(4) As String, arrTmp As Variant, wbSnLog As Workbook, wbAmLog As Workbook, wbCent As Workbook
    Dim arrCols() As Variant
    Dim sSerNum As String, sCOinit As String, sCOlatest As String, sCOtest As String, sRng As String
    Dim sCustomer As String, sModel As String, sWbName As String, sWBpath As String
    Dim varVar As Variant, varRng As Variant 'need a 2nd level variant for a nested loop
    Dim bWbSNOpen As Boolean, bWbOtOpen As Boolean, bEmptyRow As Boolean
    Dim rngResult As Range, arrRngs() As Variant 'search results
    Dim i As Integer, j As Integer, x As Integer
    Dim dDateLatest As Date
    
    On Error GoTo errhandler

    Application.ScreenUpdating = False
    
    For i = 0 To 4 'in case an error occurs
        arrOutput(i) = "Not found"
    Next
    GivenInfo = arrOutput
    sCOinit = "Not found"
    sCOlatest = "Not found"
    sCustomer = "Not found"
    sModel = "Not found"
    Call GlobalVariables
    
    If InStr(UCase(sProdLine), "MATEER") > 0 And Not SharePointAccess Then 'sign into sharepoint to make sure those connections work
        Dim sSpDriveLet As String
        sSpDriveLet = MapSharepoint(sSharePointLink) 'req'd to avoid opening IE for sharepoint
        If sSpDriveLet = "" Then Exit Function
        Call UnmapSharepoint(sSpDriveLet) 'unmap the sharepoint drive
    End If
    On Error GoTo errhandler
    
    If UCase(sInfoType) = "SERIAL" Then
        sSerNum = sInfoOrig
    Else 'find SN
        sCOlatest = "0"
        If UCase(sProdLine) = "BURT" Then
            sWbName = Right(sBurtLogPath, Len(sBurtLogPath) - InStrRev(sBurtLogPath, "\"))
            sWBpath = Left(sBurtLogPath, Len(sBurtLogPath) - Len(sWbName))
        ElseIf UCase(sProdLine) = "CARR" Then
            sWbName = Right(sCarrLogPath, Len(sCarrLogPath) - InStrRev(sCarrLogPath, "\"))
            sWBpath = Left(sCarrLogPath, Len(sCarrLogPath) - Len(sWbName))
        Else
            sWBpath = sMateerLogPath 'FindMostCurrent(sMateerLogFolder, "Mateer") 'find filename
            sWbName = Right(sMateerSnLogPath, Len(sMateerSnLogPath) - InStrRev(sMateerSnLogPath, "\"))
            'sWBpath = Left(sWBpath, Len(sWBpath) - Len(sWbName))
        End If
        For Each varVar In Workbooks 'get SN log (if open)
            If UCase(varVar.Name) = UCase(sWbName) Then
                bWbSNOpen = True
                Set wbSnLog = varVar
            End If
        Next
        If wbSnLog Is Nothing And Dir(sWBpath & sWbName) <> "" Then Set wbSnLog = Workbooks.Open(sWBpath & sWbName, , True)   'open SN log (read only)
        sLookupCol = GetLookupColumn(sProdLine, "SN Log", sInfoType)
        arrOutCols = GetOutputCols(sProdLine, "SN Log")
        If Not wbSnLog Is Nothing Then
            For Each varVar In wbSnLog.Worksheets
                If UCase(sInfoType) = "CO" And sSerNum <> "" Then Exit For
                arrRngs = varVar.Range("A1:M2500").Value
                If UCase(sInfoType) = "CO" Then '-> find related SN
                    For i = 1 To 2450
                        If UCase(arrRngs(i, sLookupCol)) Like "*" & UCase(sInfoOrig) & "*" Then
                            sSerNum = arrRngs(i, arrOutCols(0))
                            If IsNumeric(sSerNum) Then
                                If sSerNum > 0 Then Exit For
                            Else
                                If sSerNum > "" Then Exit For
                            End If
                        End If
                    Next
                Else 'customer -> find latest CO -> find related SN
                    For i = 1 To 2450
                        If UCase(arrRngs(i, sLookupCol)) Like "*" & UCase(sInfoOrig) & "*" Then
                            sCOtest = TrimCO(CStr(arrRngs(i, arrOutCols(1))), False)
                            If IsNumeric(sCOtest) And sCOtest > sCOlatest Then
                                sCOlatest = sCOtest
                                sSerNum = arrRngs(i, arrOutCols(0))
                            End If
                        End If
                    Next
                End If
            Next
        End If
        
        If UCase(sInfoType) <> "CO" Or sSerNum = "" Then 'check other log(s) -> compare COlat's & input CO (if appl)
            If UCase(sProdLine) = "BURT" Then
                sWbName = Right(sBurtPathAM, Len(sBurtPathAM) - InStrRev(sBurtPathAM, "\"))
                sWBpath = Left(sBurtPathAM, Len(sBurtPathAM) - Len(sWbName))
            ElseIf UCase(sProdLine) = "CARR" Then
                sWbName = Right(sCentLogPath, Len(sCentLogPath) - InStrRev(sCentLogPath, "\"))
                sWBpath = Left(sCentLogPath, Len(sCentLogPath) - Len(sWbName))
            Else
                sWbName = Right(sMateerPmPath, Len(sMateerPmPath) - InStrRev(sMateerPmPath, "\"))
                sWBpath = Left(sMateerPmPath, Len(sMateerPmPath) - Len(sWbName))
            End If
            Set wbAmLog = Nothing
            For Each varVar In Workbooks 'get AM log (if open)
                If UCase(varVar.Name) = UCase(sWbName) Then
                    bWbOtOpen = True
                    Set wbAmLog = varVar
                End If
            Next
            If wbAmLog Is Nothing And Dir(sWBpath & sWbName) <> "" Then Set wbAmLog = Workbooks.Open(sWBpath & sWbName, , True) 'open AM log (read only)
            sLookupCol = GetLookupColumn(sProdLine, "AM Log", sInfoType)
            arrOutCols = GetOutputCols(sProdLine, "AM Log")
            If Not wbAmLog Is Nothing Then
                For Each varVar In wbAmLog.Worksheets
                    If UCase(sInfoType) = "CO" And sSerNum <> "" Then Exit For 'cust needs to keep checking
                    arrRngs = varVar.Range("A1:M2500").Value
                    If UCase(sInfoType) = "CO" Then '-> find related SN
                        For i = 1 To 2450
                            If UCase(arrRngs(i, sLookupCol)) Like "*" & UCase(sInfoOrig) & "*" Then
                                sSerNum = arrRngs(i, arrOutCols(0))
                                If sSerNum > 0 Then Exit For
                            End If
                        Next
                    Else 'customer -> find latest CO -> find related SN
                        For i = 1 To 2450
                            If UCase(arrRngs(i, sLookupCol)) Like "*" & UCase(sInfoOrig) & "*" Then
                                sCOtest = TrimCO(CStr(arrRngs(i, arrOutCols(3))), False)
                                If IsNumeric(sCOtest) And sCOtest > sCOlatest Then
                                    sCOlatest = TrimCO(CStr(arrRngs(i, arrOutCols(3))), False)
                                    sSerNum = arrRngs(i, arrOutCols(0))
                                End If
                            End If
                        Next
                    End If
                Next
            End If
        End If
    End If
    
    If sSerNum <> "" Then 'find all the info for that SN
        arrOutput(0) = sSerNum
        'check SN Log -> get description/customer/COinit
        If UCase(sProdLine) = "BURT" Then
            sWbName = Right(sBurtLogPath, Len(sBurtLogPath) - InStrRev(sBurtLogPath, "\"))
            sWBpath = Left(sBurtLogPath, Len(sBurtLogPath) - Len(sWbName))
        ElseIf UCase(sProdLine) = "CARR" Then
            If UCase(sSerNum) Like "C#######*" Or sSerNum Like "#######*" Then
                sWbName = Right(sCentLogPath, Len(sCentLogPath) - InStrRev(sCentLogPath, "\"))
                sWBpath = Left(sCentLogPath, Len(sCentLogPath) - Len(sWbName))
            Else
                sWbName = Right(sCarrLogPath, Len(sCarrLogPath) - InStrRev(sCarrLogPath, "\"))
                sWBpath = Left(sCarrLogPath, Len(sCarrLogPath) - Len(sWbName))
            End If
        Else
            sWBpath = FindMostCurrent(sMateerLogFolder, "Mateer") 'find filename
            sWbName = Right(sMateerSnLogPath, Len(sMateerSnLogPath) - InStrRev(sMateerSnLogPath, "\"))
            sWBpath = Left(sWBpath, Len(sWBpath) - Len(sWbName))
        End If
        For Each varVar In Workbooks 'get SN log (if open)
            If UCase(varVar.Name) = UCase(sWbName) Then
                If UCase(sInfoType) = "SERIAL" Then bWbSNOpen = True
                Set wbSnLog = varVar
            End If
        Next
        If wbSnLog Is Nothing And Dir(sWBpath & sWbName) <> "" Then Set wbSnLog = Workbooks.Open(sWBpath & sWbName, , True)  'open SN log (read only)
        If Not UCase(sSerNum) Like "C#######*" And Not sSerNum Like "#######" Then 'check for centritech
            sLookupCol = GetLookupColumn(sProdLine, "SN Log", "SERIAL")
            arrOutCols = GetOutputCols(sProdLine, "SN Log")
        Else
            sLookupCol = GetLookupColumn(sProdLine, "Other Log", "SERIAL") 'centritech
            arrOutCols = GetOutputCols(sProdLine, "Other Log") 'centritech
        End If
        If IsNumeric(sCOlatest) Then
            If sCOlatest > 0 Then arrOutput(2) = sCOlatest 'it's been found already
        Else
            sCOlatest = "0" 'starting point
        End If
        If Not wbSnLog Is Nothing Then
            For Each varVar In wbSnLog.Worksheets
                arrRngs = varVar.Range("A1:M2500").Value
                For i = 1 To 2450
                    If UCase(arrRngs(i, sLookupCol)) Like "*" & UCase(sSerNum) & "*" Then
                        arrOutput(0) = UCase(arrRngs(i, sLookupCol))
                        sCOtest = TrimCO(CStr(arrRngs(i, arrOutCols(1))), False)
                        arrOutput(1) = sCOtest
                        If IsDate(arrRngs(i, arrOutCols(2))) Then
                            arrOutput(1) = sCOtest & " (" & Format(arrRngs(i, arrOutCols(2)), "mm/dd/yy") & ")"
                        End If
                        arrOutput(3) = arrRngs(i, arrOutCols(5))
                        arrOutput(4) = arrRngs(i, arrOutCols(6))
                        Exit For
                    End If
                Next
            Next
        End If

        'check aftermarket log
        If UCase(sProdLine) = "BURT" Then
            sWbName = Right(sBurtPathAM, Len(sBurtPathAM) - InStrRev(sBurtPathAM, "\"))
            sWBpath = Left(sBurtPathAM, Len(sBurtPathAM) - Len(sWbName))
        ElseIf UCase(sProdLine) = "CARR" Then
            sWbName = Right(sCentLogPath, Len(sCentLogPath) - InStrRev(sCentLogPath, "\"))
            sWBpath = Left(sCentLogPath, Len(sCentLogPath) - Len(sWbName))
        Else
            sWbName = Right(sMateerPmPath, Len(sMateerPmPath) - InStrRev(sMateerPmPath, "\"))
            sWBpath = Left(sMateerPmPath, Len(sMateerPmPath) - Len(sWbName))
        End If
        For Each varVar In Workbooks 'get AM log (if open)
            If UCase(varVar.Name) = UCase(sWbName) Then
                If UCase(sInfoType) = "SERIAL" Then bWbOtOpen = True
                Set wbAmLog = varVar
            End If
        Next

        If wbAmLog Is Nothing And Dir(sWBpath & sWbName) <> "" Then Set wbAmLog = Workbooks.Open(sWBpath & sWbName, , True) 'open Am log (read only)
        arrOutCols = GetOutputCols(sProdLine, "AM Log")
        If Not wbAmLog Is Nothing Then
            If arrOutput(2) = "Not found" Then 'look through everything for the latest CO
                sLookupCol = GetLookupColumn(sProdLine, "AM Log", sInfoType)
                For Each varVar In wbAmLog.Worksheets
                    arrRngs = varVar.Range("A1:M2500").Value
                    If UCase(arrRngs(i, sLookupCol)) Like "*" & UCase(sSerNum) & "*" Then 'correct SN
                        sCOtest = TrimCO(CStr(arrRngs(i, arrOutCols(3))), False)
                        If UCase(sInfoType) = "CO" Then 'compare to requested CO
                            If sCOtest = sInfoOrig Then 'same CO
                                arrOutput(2) = sCOtest '"latest" is actually the req. CO
                                If IsDate(arrRngs(i, arrOutCols(4))) Then dDateLatest = arrRngs(i, arrOutCols(4))
                                If arrOutput(3) = "" Then arrOutput(3) = arrRngs(i, arrOutCols(5))
                                If arrOutput(4) = "" Then arrOutput(4) = arrRngs(i, arrOutCols(6))
                                Exit For
                            End If
                        Else 'compare to latest CO
                            If IsNumeric(sCOtest) And sCOtest > sCOlatest Then
                                sCOlatest = sCOtest
                                If IsDate(arrRngs(i, arrOutCols(4))) Then dDateLatest = arrRngs(i, arrOutCols(4))
                                If arrOutput(3) = "" Then arrOutput(3) = arrRngs(i, arrOutCols(5))
                                If arrOutput(4) = "" Then arrOutput(4) = arrRngs(i, arrOutCols(6))
                            End If
                        End If
                    End If
                Next
            Else 'definitive latest has been found
                sLookupCol = GetLookupColumn(sProdLine, "AM Log", "CO") 'search by CO, regardless of input type
                For Each varVar In wbAmLog.Worksheets
                    If arrOutput(3) <> "" And arrOutput(4) <> "" Then Exit For
                    arrRngs = varVar.Range("A1:M2500").Value
                    For i = 1 To 2450
                        If UCase(arrRngs(i, sLookupCol)) Like "*" & UCase(arrOutput(2)) & "*" Then 'correct CO
                            If IsDate(arrRngs(i, arrOutCols(4))) Then arrOutput(2) = arrOutput(2) & " (" & arrRngs(i, arrOutCols(4)) & ")"
                            If arrOutput(3) = "" Then arrOutput(3) = arrRngs(i, arrOutCols(5))
                            If arrOutput(4) = "" Then arrOutput(4) = arrRngs(i, arrOutCols(6))
                            Exit For
                        End If
                    Next
                Next
            End If
        End If
    End If

    If arrOutput(2) = "0" Then arrOutput(2) = "Not found"
    If Not bWbSNOpen And Not wbSnLog Is Nothing Then wbSnLog.Close savechanges:=False
    If Not bWbOtOpen And Not wbAmLog Is Nothing Then wbAmLog.Close savechanges:=False

    GivenInfo = arrOutput

    Exit Function

errhandler:
    GivenInfo = arrOutput
    MsgBox "Error in GivenInfo function"
    Call ErrorRep("GivenInfo", "Function", "array values: " & vbCrLf & Join(GivenInfo, ";"), Err.Number, Err.Description, "")
End Function

Function LinkEwsQuoteLayout(rngSN As Range, rngCoInit As Range, rngCOlat As Range, rngLayout As Range, sProdLine As String, bLinks As Boolean, bLinksAll As Boolean)
    Dim sCOnum As String, sLayoutName As String
    Dim i As Integer
    Dim sFileInfo As String, sFilePath As String, sFileName As String, sSheetName As String
    'find ews & link to it (replace CO text with link)
    'getlatest if necessary
    'get layout #
    'link to layout if possible
    'sCOnum = TrimCO(rngCoInit.Value, False) 'find initial CO EWS
    Debug.Print "in it"
    If IsNumeric(Left(rngCoInit.Value, 6)) Then 'faster than TrimCO
        sCOnum = Left(rngCoInit.Value, 6)
    End If
    If sCOnum Like "######" Or UCase(sProdLine) = "CARR" Then 'CO number found
        Application.StatusBar = "Looking for initial EWS..."
        If bLinks Then
            sFilePath = FindEWSPath(sCOnum, rngSN, sProdLine, bLinks, bLinksAll) 'entire path
        End If
        If InStr(sFilePath, "\") > 0 Then 'backward slashes
            sFileName = Right(sFilePath, Len(sFilePath) - InStrRev(sFilePath, "\"))
        ElseIf InStr(sFilePath, "/") > 0 Then 'forward slashes -> shouldn't ever happen
            sFileName = Right(sFilePath, Len(sFilePath) - InStrRev(sFilePath, "/"))
        End If
        If sFilePath <> "" Then
            If bLinks Then rngCoInit.Formula = "=HYPERLINK(" & """" & sFilePath & """" & ", " & """" & rngCoInit.Value & """" & ")"
            Application.StatusBar = "Looking for layout number..."
            Application.DisplayAlerts = False
            sFilePath = Left(sFilePath, Len(sFilePath) - Len(sFileName)) 'redefine path (remove file name)
            Sheet3.Range("M7").Formula = "=IFERROR('" & sFilePath & "[" & sFileName & "]Customer Order Worksheet'!$F$12,0)" 'layout#1
            If UCase(sProdLine) = "BURT" Then 'because god forbid we be consistent at all
                Sheet3.Range("M8").Formula = "=IFERROR('" & sFilePath & "[" & sFileName & "]M1'!$E$12,0)" 'layout#2
            Else
                Sheet3.Range("M8").Formula = "=IFERROR('" & sFilePath & "[" & sFileName & "]Seq. No. 1'!$E$12,0)" 'layout#2
            End If
            If Sheet3.Range("M7").Value Like "*####*" And Not UCase(Sheet3.Range("M7").Value) Like "W#*" Then
                sLayoutName = Sheet3.Range("M7").Value 'layout number
            ElseIf Sheet3.Range("M8").Value Like "*####*" And Not UCase(Sheet3.Range("M8").Value) Like "W#*" Then
                sLayoutName = Sheet3.Range("M8").Value 'layout number
            End If
        End If
    ElseIf UCase(Left(rngSN.Value, 1)) = "8" And UCase(sProdLine) = "MATEER" Then 'look for autocad layout
        sFileInfo = SearchVault("W" & rngSN.Value, "dwg")
        If sFileInfo <> "" Then 'dwg found
            rngLayout.Value = "=hyperlink(" & """" & sFileInfo & """" & ", " & """" & rngSN.Value & """" & ")"
        Else 'no dwg
            rngLayout.Value = "Not found"
        End If
    End If
    
    If sLayoutName = "" And UCase(sProdLine) = "CARR" Then 'assume layout number
        If UCase(rngSN.Value) Like "C#######*" Then
            sLayoutName = Left(rngSN.Value, 8) & "*"
        ElseIf UCase(rngSN.Value) Like "C####*" Then
            sLayoutName = Left(rngSN.Value, 5) & "*"
        ElseIf rngSN.Value Like "####*" Then
            sLayoutName = "C" & Left(rngSN.Value, 4) & "*"
        ElseIf rngSN.Value Like "*####*" Then
            i = 1
            Do While i < Len(rngSN.Value) - 4
                If IsNumeric(Mid(rngSN.Value, i, 4)) Then
                    sLayoutName = "C" & Mid(rngSN.Value, i, 4) & "*"
                    Exit Do
                End If
                i = i + 1
            Loop
        End If
    End If

    If sLayoutName <> "" Then
        If InStr(sLayoutName, "*") = 0 Then
            rngLayout.Value = sLayoutName 'set the number, no link
        Else
            rngLayout.Value = "Not found"
        End If
        If bLinks Then 'try to find the file
            Application.StatusBar = "Looking for layout as SLDDRW..."
            sFilePath = SearchVault(sLayoutName, "slddrw")
            Application.StatusBar = "Looking for layout as DWG..."
            If sFilePath = "" Then sFilePath = SearchVault(sLayoutName, "dwg") 'autocad
            Application.StatusBar = "Looking for layout as SLDASM..."
            If sFilePath = "" Then sFilePath = SearchVault(sLayoutName, "sldasm") 'assembly
            Application.StatusBar = False
            If InStr(sFilePath, "\") > 0 Then
                sLayoutName = Mid(sFilePath, InStrRev(sFilePath, "\") + 1, InStrRev(sFilePath, ".") - InStrRev(sFilePath, "\") - 1)
                rngLayout.Formula = "=hyperlink(" & """" & sFilePath & """" & ", " & """" & sLayoutName & """" & ")" 'results were found
            End If
        End If
    Else
        rngLayout.Value = "Not found"
    End If

    If bLinks Then
        If rngCOlat.Value <> "Not found" And TrimCO(rngCOlat.Value, False) <> sCOnum Then 'do the same for the latest CO's EWS
            Application.StatusBar = "Looking for latest EWS..."
            sCOnum = TrimCO(rngCOlat.Value, False)
            sFilePath = FindEWSPath(sCOnum, rngSN, sProdLine, bLinks, bLinksAll) 'entire path
            If InStr(sFilePath, "\") > 0 Then 'backward slashes
                sFileName = Right(sFilePath, Len(sFilePath) - InStrRev(sFilePath, "\"))
            ElseIf InStr(sFilePath, "/") > 0 Then 'forward slashes -> shouldn't ever happen
                sFileName = Right(sFilePath, Len(sFilePath) - InStrRev(sFilePath, "/"))
            End If
            If sFilePath <> "" Then
                rngCOlat.Formula = "=HYPERLINK(" & """" & sFilePath & """" & ", " & """" & rngCOlat.Value & """" & ")"
            End If
        End If
    End If
    Application.StatusBar = "Looking for quote..."
    sFileInfo = ""
    If bLinks Then sFileInfo = FindQuotePath(rngSN.Value, sProdLine) 'find quote
    If sFileInfo <> "" Then
        rngSN.Formula = "=HYPERLINK(" & """" & sFileInfo & """" & ", " & """" & rngSN.Value & """" & ")"
    End If
    Application.StatusBar = False
    Exit Function
errhandler:
    Application.StatusBar = False
    MsgBox "Error in LinkEwsQuoteLayout function"
    rngLayout.Value = "Not found"
    Call ErrorRep("LinkEwsQuoteLayout", "Function", "N/A", Err.Number, Err.Description, "")
End Function

Function FormatPlain(rng As Range)
    With rng.Font
        .Color = vbBlack
        .Underline = False
    End With
End Function


Function FindQuotePath(sSerNum As String, sProdLine As String) As String
    'returns path to quote file, or empty string
    Dim sUmbFolderPath(1) As String, sSubFolderName As String, sProjFolder As String
    Dim sFolderNamePatt As String, sProjFolderPath As String, sFileName As String
    Dim varVar As Variant, sOutput As String
    
    If sSerNum = "Not found" Then Exit Function
    If UCase(sProdLine) = "BURT" Then
        ''''''''''hardcoded'''''''
        sUmbFolderPath(0) = "\\PSACLW02\RELEASED\MB SOP\BURT CUSTOMER ORDER  FILES\"
        sUmbFolderPath(1) = "" 'no second path for Burt
        sSubFolderName = "ORDER INFORMATION (ODS - CHECK SHEETS - QUOTE - MACHINE DOC'S AFTERWARDS)"
        ''''''''''''''''''''''''''
        If Len(sSerNum) > 2 Then 'folders are named "672...", "673...", etc
            If Right(sSerNum, 3) Like "###" Then sFolderNamePatt = Right(sSerNum, 3) & "*"
        End If
        If sFolderNamePatt = "" Then sFolderNamePatt = sSerNum
    ElseIf UCase(sProdLine) = "CARR" Then
        ''''''''''hardcoded'''''''
        sUmbFolderPath(0) = "\\PSACLW02\HOME\SHARED\CARR-CENTRITECH\PROJECT FOLDERS\"
        sUmbFolderPath(1) = "" 'no second path for carr
        sSubFolderName = "Order Documents"
        ''''''''''''''''''''''''''
        If Len(sSerNum) > 2 Then 'folders are named "672...", "673...", etc
            If sSerNum Like "C#######*" Then
                sFolderNamePatt = Left(sSerNum, 8) & "*"
            ElseIf Left(sSerNum, 4) Like "####" Then
                sFolderNamePatt = "C" & Left(sSerNum, 4) & "*"
            ElseIf Left(sSerNum, 3) Like "###" Then
                sFolderNamePatt = "C0" & Left(sSerNum, 3) & "*"
            End If
        End If
        If sFolderNamePatt = "" Then sFolderNamePatt = sSerNum
    Else
        ''''''''''hardcoded'''''''
        sUmbFolderPath(0) = "\\PSACLW02\HOME\SHARED\MATEER\PROJECT MANAGEMENT\OPEN ORDERS\NEW MACHINES\"
        sUmbFolderPath(1) = "\\PSACLW02\HOME\SHARED\MATEER\PROJECT MANAGEMENT\CLOSED ORDERS\NEW MACHINES\"
        sSubFolderName = "Quotes"
        ''''''''''''''''''''''''''
        sFolderNamePatt = "*" & sSerNum & "*" 'folders have the SN in the name
    End If
    
    For Each varVar In sUmbFolderPath
        If Right(varVar, 1) <> "\" Then varVar = varVar & "\"
        sProjFolder = Dir(varVar & sFolderNamePatt, vbDirectory)
        If sProjFolder <> "" Then 'sometimes SN's for mateer are part of CO's for unrelated projects
            If Not sProjFolder Like "*#" & sSerNum & "*" And Not sProjFolder Like "*" & sSerNum & "#*" Then
                sProjFolderPath = varVar & sProjFolder
                Exit For
            End If
        End If
    Next
    
    If Right(sSubFolderName, 1) <> "\" Then sSubFolderName = sSubFolderName & "\"
    If Right(sProjFolderPath, 1) <> "\" Then sProjFolderPath = sProjFolderPath & "\"
    On Error Resume Next 'error 52 happens for failed network paths, which is dumb
    sFileName = Dir(sProjFolderPath & sSubFolderName & "*NM*.doc*")
    If sFileName <> "" Then
        sOutput = sProjFolderPath & sSubFolderName & sFileName
    ElseIf Dir(sProjFolderPath & sSubFolderName, vbDirectory) <> "" Then
        sOutput = sProjFolderPath & sSubFolderName
    ElseIf Dir(sProjFolderPath, vbDirectory) <> "" And Len(sProjFolderPath) > 5 Then
        sOutput = sProjFolderPath
    End If
    FindQuotePath = sOutput
errhandler:
End Function

Function FindEWSPath(sCO As String, rngSN As Range, sProdLine As String, bLinks As Boolean, bLinksAll As Boolean) As String
    Dim sOutput As String 'path to the doc
    Dim sSpDriveLet As String 'drive letter for sharepoint
    Dim sFolderPath As String, sFileName As String
    Dim sPath As String, sFilePatt As String

    Call GlobalVariables
    Application.StatusBar = Left(Application.StatusBar, InStrRev(Application.StatusBar, ".")) & " (Vault)"
    ''''''''hardcoded'''''''''''''''''
    If UCase(sProdLine) = "CARR" Then
        sPath = "C:\EPDM\PSA_Vault\CLW\CARR\C" & Left(rngSN.Value, 2) & "\"
        sFilePatt = "*" & Replace(rngSN.Value, "-", "") & "*EWS*.xls*"
    Else
        If UCase(sProdLine) = "BURT" Then
            sPath = "\\PSACLW02\RELEASED\MB SOP\EWS (ENGR WORK SHEET)\EWS200000\EWS" & CStr(Int((sCO / 1000)) * 1000) & "\"
        Else
            sPath = "C:\EPDM\PSA_Vault\CLW\Mateer\EWS\EWS" & CStr(Int((sCO / 1000)) * 1000) & "\"
        End If
        sFilePatt = "EWS" & sCO & "*.xls*"
    End If
    '''''''''''''''''''''''''''''''''
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    sOutput = Dir(sPath & sFilePatt)
    
    If sOutput <> "" Or Not bLinks Then
        sOutput = sPath & sOutput
        If InStr(sOutput, "EPDM") > 0 And InStr(UCase(sOutput), ".XLS") > 0 And bLinks Then sOutput = GetLatest(sOutput)
        FindEWSPath = sOutput
        Exit Function
    End If

    If UCase(sProdLine) = "CARR" Then 'check vault (sn filename)
        sOutput = SearchVault("C" & Replace(rngSN.Value, "-", "") & "*EWS*", "xls*")
    Else 'check  vault (co filename)
        sOutput = SearchVault("EWS" & sCO & "*", "xls*")
    End If
    If sOutput <> "" Then 'result in the Vault
        sOutput = GetLatest(sOutput) 'make sure user has latest version
    ElseIf bLinksAll Then 'no results -> check sharepoint
        Application.StatusBar = Left(Application.StatusBar, InStrRev(Application.StatusBar, ".")) & " (SharePoint)"
        On Error Resume Next
            If Dir(sSharePointMap, vbDirectory) = "" Then 'check for Sharepoint connection
                sSpDriveLet = MapSharepoint(sSharePointLink)
            End If
        Err.Clear
        On Error GoTo errhandler
        
        If Left(sCO, 1) = "3" Or Left(sCO, 1) = "4" Then 'EWS Page -> org structure A
            sFolderPath = sEWSListFolder
            If Right(sFolderPath, 1) <> "\" Then sFolderPath = sFolderPath & "\"
            If Left(sCO, 1) = "3" Then sFolderPath = sFolderPath & Dir(sFolderPath & "EWS3*", vbDirectory)
        Else 'EWS Archive [likely 1 or 2]
            sFolderPath = sEWSArchFolder
            If Right(sFolderPath, 1) <> "\" Then sFolderPath = sFolderPath & "\"
            sFolderPath = sFolderPath & Dir(sFolderPath & "EWS" & Left(sCO, 1) & "*", vbDirectory)
        End If
        If Right(sFolderPath, 1) <> "\" Then sFolderPath = sFolderPath & "\"
        sOutput = FindFullPath(sFolderPath, "EWS" & sCO, "xls*")
    End If
    If sOutput = "" Then 'check other locations
        Application.StatusBar = Left(Application.StatusBar, InStrRev(Application.StatusBar, ".")) & " (shared drive)"
        If UCase(sProdLine) = "MATEER" Then 'check AM folder
            ''''''''''hardcoded'''''''
            'sFolderPath = "\\PSACLW02\HOME\SHARED\MATEER\PROJECT MANAGEMENT\OPEN ORDERS\AFTERMARKET\"
            'AM folder is organized like garbage so it takes forever & rarely results in anything
            ''''''''''''''''''''''''''
        ElseIf UCase(sProdLine) = "BURT" Then 'check MK6 folder
            If Left(sCO, 1) <> "2" Then Exit Function
            ''''''''''hardcoded'''''''
            sFolderPath = "\\PSACLW02\RELEASED\MB SOP\EWS (ENGR WORK SHEET)\EWS200000\"
            ''''''''''''''''''''''''''
        End If
        If sFolderPath <> "" Then sOutput = FindFullPath(sFolderPath, "EWS" & sCO, "xls*")
    End If
    FindEWSPath = sOutput
errhandler:
End Function

Function CheckFastLinks() As Boolean()
'(0) = true if search type is "Fast"
'(1) = TRUE if any links are shown, (2) = true if sharepoint to be incl.

    Dim sOut(2) As Boolean
    On Error GoTo errhandler
    
    sOut(0) = True 'fast by default
    Call Sheet1Vars
    If UCase(rngFast.Value) = "TORTOISE" Then sOut(0) = False
    
    sOut(1) = False 'no links by default
    If InStr(UCase(rngLinks.Value), "NO") = 0 Then 'some links are desired
        sOut(1) = True
        If InStr(UCase(rngLinks.Value), "LOC") = 0 Then 'local
            sOut(2) = True
        End If
    End If
    
    CheckFastLinks = sOut
errhandler:
End Function

Function RetrieveInfo(wB As Workbook, iSheetInd As String, sLookupVal As String, sLookupCol As String, sColSN As String, _
                         sColCO As String, sColDate As String, sColCust As String, sColModel As String, bFindLatest As Boolean) As String()
'returns details for a job from a workbook, going by earliest or latest CO in the sheet
'(0)=SN, (1)=CO, (2)=Date, (3)=Customer, (4)=Model
    Dim rngResult As Range, arrRngs() As Range  'results of find
    Dim arrOut(4) As String
    Dim i As Integer, x As Integer
    Dim wSheet As Worksheet
    Dim varVar As Variant, varSheet As Variant
    Dim lCOToBeat As Long, lTestCO As Long
    Dim bNewResult As Boolean 'bFindLatest As Boolean 'if false, then find earliest
    Dim bExitDo As Boolean 'required because for...each can't be interrupted or array gets locked
    
    If bFindLatest Then
        lCOToBeat = 1 'co of 0 must not be accepted -> CO's must be > 1
    Else 'co will have to be lower than "lCOToBeat"
        lCOToBeat = 999999 'default val of 0 will always be "earlier" than a CO#
    End If
    
    ReDim arrRngs(0)
    If IsNumeric(iSheetInd) Then 'only lookup for 1 sheet
        i = CInt(iSheetInd)
        Set wSheet = wB.Worksheets(i)
        On Error Resume Next
            Set rngResult = wSheet.Range(sLookupCol & ":" & sLookupCol).Find(what:=sLookupVal)
        On Error GoTo 0 'errhandler
        x = 0
        Do While Not rngResult Is Nothing
            bNewResult = False
            If IsNumeric(rngResult.Offset(0, Asc(sColCO) - Asc(sLookupCol)).Value) Then 'valid CO
                lTestCO = rngResult.Offset(0, Asc(sColCO) - Asc(sLookupCol)).Value
                If bFindLatest And lTestCO > lCOToBeat Then 'better CO
                    bNewResult = True
                ElseIf (Not bFindLatest) And lTestCO < lCOToBeat Then 'better CO
                    bNewResult = True
                ElseIf lTestCO = lCOToBeat Then
                    bNewResult = True
                End If
            Else 'deals with bad data
                bNewResult = False
            End If
            
            If bNewResult Then 'better date -> make this the output
                arrOut(0) = rngResult.Offset(0, Asc(sColSN) - Asc(sLookupCol)).Value
                arrOut(1) = rngResult.Offset(0, Asc(sColCO) - Asc(sLookupCol)).Value
                arrOut(2) = rngResult.Offset(0, Asc(sColDate) - Asc(sLookupCol)).Value
                arrOut(3) = rngResult.Offset(0, Asc(sColCust) - Asc(sLookupCol)).Value
                arrOut(4) = rngResult.Offset(0, Asc(sColModel) - Asc(sLookupCol)).Value
            End If
            
            If x > 0 Then
                For Each varVar In arrRngs
                    If varVar.Address = rngResult.Address Then Exit Do
                Next
            End If
            
            ReDim Preserve arrRngs(x)
            Set arrRngs(x) = rngResult
            x = x + 1
            On Error Resume Next
                Set rngResult = wSheet.Range(sLookupCol & ":" & sLookupCol).FindNext(after:=rngResult)
            On Error GoTo errhandler
            
            If x > 500 Then Exit Do 'something bad
        Loop
    Else 'lookup on all sheets
        For Each varSheet In wB.Worksheets
            On Error Resume Next
                Set rngResult = varSheet.Range(sLookupCol & ":" & sLookupCol).Find(what:=sLookupVal)
            On Error GoTo 0 'errhandler
            
            x = 0
            ReDim arrRngs(0)
            Do While Not rngResult Is Nothing
                bNewResult = False
                If IsNumeric(rngResult.Offset(0, Asc(sColCO) - Asc(sLookupCol)).Value) Then 'valid CO
                    lTestCO = rngResult.Offset(0, Asc(sColCO) - Asc(sLookupCol)).Value
                Else 'deals with bad data
                    If bFindLatest Then
                        lTestCO = -99999
                    Else
                        lTestCO = 999999
                    End If
                End If
                
                If bFindLatest And lTestCO > lCOToBeat Then 'better CO
                    bNewResult = True
                ElseIf (Not bFindLatest) And lTestCO < lCOToBeat Then 'better CO
                    bNewResult = True
                End If
                
                If bNewResult Then 'better date -> make this the output
                    arrOut(0) = rngResult.Offset(0, Asc(sColSN) - Asc(sLookupCol)).Value
                    arrOut(1) = rngResult.Offset(0, Asc(sColCO) - Asc(sLookupCol)).Value
                    arrOut(2) = rngResult.Offset(0, Asc(sColDate) - Asc(sLookupCol)).Value
                    arrOut(3) = rngResult.Offset(0, Asc(sColCust) - Asc(sLookupCol)).Value
                    arrOut(4) = rngResult.Offset(0, Asc(sColModel) - Asc(sLookupCol)).Value
                End If
                
                bExitDo = False
                If x > 0 Then
                    For Each varVar In arrRngs
                        If varVar.Address = rngResult.Address Then bExitDo = True
                    Next
                End If
                
                If bExitDo Then 'this is required because for...each will lock the array if interrupted
                    Exit Do
                End If
                
                ReDim Preserve arrRngs(x)
                Set arrRngs(x) = rngResult
                x = x + 1
                
                On Error Resume Next
                    Set rngResult = wSheet.Range(sLookupCol & ":" & sLookupCol).FindNext(after:=rngResult)
                On Error GoTo errhandler
                If x > 50 Then Exit Do 'something bad
            Loop
        Next varSheet
    End If
    
    On Error Resume Next
    RetrieveInfo = arrOut
    Exit Function
    
errhandler:
    MsgBox "Error in RetrieveInfo function"
    Call ErrorRep("RetrieveInfo", "Function", "", Err.Number, Err.Description, "")
End Function

Function PhoneNumber(sInput As String, sType As String, rngName As Range, rngNum As Range) As Boolean
    PhoneNumber = False
    On Error GoTo errhandler
    
    Dim sNameOut As String, sNumOut As String, sExtOut As String, sPhoneOut As String, sCellOut As String, sWbName As String
    Dim wbPhoneListCLW As Workbook, wbPhoneListAKR As Workbook, arrSheets() As Worksheet
    Dim varVar As Variant
    Dim iPos As Integer
    Dim arrResults As Variant
    Dim bWbOpen As Boolean
    
    Call GlobalVariables
    
    If InStr(sPhonePathCLW, "\") > 0 Then
        sWbName = Right(sPhonePathCLW, Len(sPhonePathCLW) - InStrRev(sPhonePathCLW, "\"))
    Else
        sWbName = Replace(Right(sPhonePathCLW, Len(sPhonePathCLW) - InStrRev(sPhonePathCLW, "/")), "%20", " ")
    End If
    
    For Each varVar In Workbooks 'see if it's open
        If UCase(varVar.Name) = UCase(sWbName) Then
            Set wbPhoneListCLW = varVar
            bWbOpen = True
            Exit For
        End If
    Next
    If wbPhoneListCLW Is Nothing Then
        Set wbPhoneListCLW = Workbooks.Open(sPhonePathCLW, , True)
    End If
    
    arrResults = SearchPhoneList(sInput, sType, wbPhoneListCLW)
    sNameOut = arrResults(0)
    sExtOut = arrResults(1)
    sPhoneOut = arrResults(2)
    sCellOut = arrResults(3)
    
    If Not bWbOpen Then wbPhoneListCLW.Close savechanges:=False
    
    If sNameOut = "" Or (sExtOut = "" And sPhoneOut = "" And sCellOut = "") Then 'look in AKR list
        If InStr(sPhonePathAKR, "\") > 0 Then
            sWbName = Right(sPhonePathAKR, Len(sPhonePathAKR) - InStrRev(sPhonePathAKR, "\"))
        Else
            sWbName = Replace(Right(sPhonePathAKR, Len(sPhonePathAKR) - InStrRev(sPhonePathAKR, "/")), "%20", " ")
        End If
        bWbOpen = False
        For Each varVar In Workbooks 'see if it's open
            If UCase(varVar.Name) = UCase(sWbName) Then
                Set wbPhoneListAKR = varVar
                bWbOpen = True
                Exit For
            End If
        Next
        If wbPhoneListAKR Is Nothing Then
            Set wbPhoneListAKR = Workbooks.Open(sPhonePathAKR, , True)
        End If
        
        arrResults = SearchPhoneList(sInput, "PERSON", wbPhoneListAKR) 'AKR conf rooms aren't labeled
        sNameOut = arrResults(0)
        sExtOut = arrResults(1)
        sPhoneOut = arrResults(2)
        sCellOut = arrResults(3)

        If Not bWbOpen Then wbPhoneListAKR.Close savechanges:=False
    End If
    Debug.Print Join(arrResults, ",")
    If sNameOut <> "" Then 'put the values in the sheet
        rngName.Value = sNameOut
    Else
        rngName.Value = sInput
    End If
    
    If sPhoneOut <> "" Then
        sNumOut = sPhoneOut & " | "
    ElseIf sExtOut <> "" Then
        sNumOut = sExtOut & " | "
    End If
    If sCellOut <> "" Then sNumOut = sNumOut & sCellOut
    
    If Right(Trim(sNumOut), 1) = "|" Then sNumOut = Left(sNumOut, Len(sNumOut) - 3)
    
    rngNum.Value = sNumOut
    iPos = InStr(rngNum.Value, sExtOut)
    If iPos > 0 And (iPos < InStr(rngNum.Value, "|") Or InStr(rngNum.Value, "|") = 0) Then 'bold the extension
        With rngNum.Characters(Start:=iPos, Length:=4).Font
            .FontStyle = "Bold"
        End With
    End If
    
    PhoneNumber = True
    Exit Function
errhandler:
    MsgBox "Error finding phone number"
    Call ErrorRep("PhoneNumber", "Function", PhoneNumber, Err.Number, Err.Description, "")
End Function

Function SearchPhoneList(sInput As String, sType As String, wbPhone As Workbook) As String()
    
    Dim sOutput(3) As String '(0)=Name, (1)=Ext, (2)=Number, (3)=Cell
    Dim x As Integer, iNameOffset As Integer
    Dim bConfRng As Boolean, bRevSearch As Boolean
    Dim rngConfRms() As Range, rngResult As Range, rngStart As Range, rngEnd As Range 'ranges with conf room extensions
    Dim varVar As Variant
    Dim rngIsect As Range, arrRngs() As Range
    
    On Error GoTo errhandler
    
    x = 0
    For Each Sheet In wbPhone.Sheets
        Set rngResult = Sheet.Range("A:L").Find(what:="Conference", lookat:=xlPart)
        If Not rngResult Is Nothing Then
            Set rngStart = rngResult
            Do While Not rngResult Is Nothing
                Set rngEnd = rngStart
                Do While rngEnd.Value > 0
                    Set rngStart = rngStart.Resize(rngStart.Rows.Count + 1, rngStart.Columns.Count)
                    Set rngEnd = rngEnd.Offset(1, 0)
                Loop
                
                If x > 0 Then 'check if all ranges have been found (starting to repeat)
                    For Each varVar In rngConfRms
                        If rngStart.Address = varVar.Address Then 'repeating -> found all ranges
                            Exit Do
                        End If
                    Next
                End If
                
                ReDim Preserve rngConfRms(x)
                Set rngConfRms(x) = rngStart 'really has been resized to encompass start->end

                x = x + 1
                If x > 100 Then GoTo errhandler 'something is wrong... found >100 resulting ranges
                Set rngResult = Sheet.Range("A:L").FindNext(after:=rngEnd)
                Set rngStart = rngResult
            Loop
        End If
    Next
    
    Set rngResult = Nothing
    If UCase(sType) = "ROOM" Then 'look in conf room ranges
        For Each varVar In rngConfRms 'varvar = range with conf rooms
            On Error Resume Next 'in case "Find" fails
            Set rngResult = varVar.Find(what:=sInput, lookat:=xlPart)
            If Not rngResult Is Nothing Then 'inameoffset=0
                sOutput(0) = rngResult.Value
                sOutput(1) = rngResult.Offset(0, iNameOffset + 1).Value
                sOutput(2) = rngResult.Offset(0, iNameOffset + 2).Value
                sOutput(3) = rngResult.Offset(0, iNameOffset + 3).Value
                Exit For
            End If
        Next
    Else 'person or other non-conference room
        If IsNumeric(sInput) Then bRevSearch = True
        For Each Sheet In wbPhone.Sheets
            x = 0 'for array of ranges which have been found
            On Error Resume Next 'in case "Find" fails
            Set rngResult = Sheet.Range("A:L").Find(what:=sInput, lookat:=xlPart)
            Do While Not rngResult Is Nothing
                bConfRng = False 'check if the result is in conference room ranges
                For Each varVar In rngConfRms
                    Set rngIsect = Application.Intersect(varVar, rngResult)
                    If Not rngIsect Is Nothing Then 'conf room range
                        bConfRng = True
                    End If
                Next
                If Not bConfRng Then 'not part of any conf rm ranges
                    If bRevSearch Then
                        If rngResult.Value = sInput Then 'extension was found
                            iNameOffset = -1
                        Else 'phone # was found
                            iNameOffset = -2
                        End If
                        sOutput(0) = rngResult.Offset(0, iNameOffset).Value
                        sOutput(1) = rngResult.Offset(0, iNameOffset + 1).Value
                        If InStr(rngResult.Offset(0, iNameOffset + 2).Value, ",") = 0 Then 'not a person's name in the next column
                            sOutput(2) = rngResult.Offset(0, iNameOffset + 2).Value
                            If InStr(rngResult.Offset(0, iNameOffset + 3).Value, ",") = 0 Then
                                sOutput(3) = rngResult.Offset(0, iNameOffset + 3).Value
                            End If
                        End If
                    Else
                        sOutput(0) = rngResult.Value
                        sOutput(1) = rngResult.Offset(0, 1).Value
                        If InStr(rngResult.Offset(0, 2).Value, ",") = 0 Then 'not a person's name in the next column
                            sOutput(2) = rngResult.Offset(0, 2).Value
                            If InStr(rngResult.Offset(0, 3).Value, ",") = 0 Then
                                sOutput(3) = rngResult.Offset(0, 3).Value
                            End If
                        End If
                    End If
                End If
                If sOutput(0) <> "" And (sOutput(1) <> "" Or sOutput(2) <> "" Or sOutput(3) <> "") Then Exit For 'info has been found
                If x > 0 Then 'check for no results (keeps cycling, trying)
                    For Each varVar In arrRngs
                        If varVar.Address = rngResult.Address Then Exit Do 'no results
                    Next
                End If
                ReDim Preserve arrRngs(x)
                arrRngs(x) = rngResult
                x = x + 1
                Set rngResult = Sheet.Range("A:L").FindNext(after:=rngResult)
            Loop
        Next
    End If
    
    SearchPhoneList = sOutput
    Exit Function
errhandler:
    Call ErrorRep("SearchPhoneList", "Function", "N/A", Err.Number, Err.Description, "")
End Function

Sub ShowHistory()
Attribute ShowHistory.VB_ProcData.VB_Invoke_Func = "H\n14"
    Dim sAddress As String, sProdLine As String, sColLet As String, arrItems() As String
    Dim i As Integer, x As Integer
    On Error GoTo errhandler
    sAddress = ActiveCell.Address
    If sAddress = "$A$2" Or sAddress = "$B$2" Or sAddress = "$C$2" Then
        sProdLine = ActiveCell.Offset(-1, 0).Value
        '''''hardcoded''''''
            If UCase(sProdLine) = "BURT" Then
                sColLet = "M"
            ElseIf UCase(sProdLine) = "CARR" Then
                sColLet = "N"
            Else
                sColLet = "O"
            End If
        ''''''''''''''''''''
        ReDim arrItems(0)
        For i = 11 To 20
            If Sheet3.Range(sColLet & i).Value <> "" Then
                ReDim Preserve arrItems(x)
                arrItems(x) = Sheet3.Range(sColLet & i).Value
                x = x + 1
            End If
        Next
        If x > 0 Then MsgBox "Recent searches for " & sProdLine & ":" & vbCrLf & vbCrLf & _
                            Join(arrItems, vbCrLf)
    End If
errhandler:
End Sub

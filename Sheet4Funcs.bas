Attribute VB_Name = "Sheet4Funcs"
Function FillHistory(sInfo As String, sInfoType As String, sProdLine As String)
    'sinfo = SN or Customer, sinfotype = "SERIAL" or "CUST"
    
    Dim arrWbPaths(2) As String, arrWBs(2) As Workbook, arrWerentOpen(2) As String
    Dim iLookupCol As Integer, sWbName As String, sWbType As String
    Dim arrResultsCols As Variant, varVar As Variant, varVar2 As Variant, arrContent As Variant
    Dim i As Integer, j As Integer, k As Integer
    Dim arrRng As Variant, arrOutput() As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Searching history for " & """" & sInfo & """" & "..."
    On Error GoTo errhandler
    
    Call GlobalVariables
    
    If UCase(sProdLine) = "BURT" Then
        arrWbPaths(0) = sBurtLogPath
    ElseIf UCase(sProdLine) = "CARR" Then
        arrWbPaths(0) = sCarrLogPath
        arrWbPaths(1) = sCentLogPath
    Else
        arrWbPaths(0) = sMateerLogPath 'FindMostCurrent(sMateerLogFolder, "Mateer")
        arrWbPaths(1) = sMateerPmPath
    End If
    
    For i = 0 To 2
        If arrWbPaths(i) <> "" Then
            sWbName = Right(arrWbPaths(i), Len(arrWbPaths(i)) - InStrRev(arrWbPaths(i), "\"))
            For Each varVar In Workbooks
                If UCase(varVar.Name) = UCase(sWbName) Then
                    Set arrWBs(i) = varVar
                    Exit For
                End If
            Next
            If arrWBs(i) Is Nothing Then
                Set arrWBs(i) = Workbooks.Open(arrWbPaths(i), , True)
                arrWerentOpen(i) = sWbName
            End If
        End If
    Next
    
    ReDim arrOutput(4, 0)
    i = 0 'each row of output
    k = i 'count of row with worksheet name
    For Each varVar In arrWBs
        If Not varVar Is Nothing Then
            If (InStr(UCase(varVar.Name), "SERIAL") > 0 And InStr(UCase(varVar.Name), "CENT") = 0) _
                Or InStr(UCase(varVar.Path), "BURT") > 0 Then 'because Burt SN log has a bad name
                sWbType = "SN Log"
            Else
                sWbType = "AM Log"
            End If
            iLookupCol = GetLookupColumn(sProdLine, sWbType, sInfoType)
            arrResultsCols = GetOutputCols(sProdLine, sWbType)
            For Each varVar2 In varVar.Worksheets
                If i = k Then
                    If i > 0 Then i = i - 1
                    arrOutput(4, i) = varVar.Name & " (" & varVar2.Name & ")"
                    arrOutput(0, i) = "Source"
                Else
                    i = i + 1 'puts a blank line before each source
                    ReDim Preserve arrOutput(4, i) 'redim final array
                    arrOutput(4, i) = varVar.Name & " (" & varVar2.Name & ")"
                    arrOutput(0, i) = "Source"
                End If
                i = i + 1
                k = i
                
                arrRng = varVar2.Range("A1:M2500").Value
                For j = 1 To 2500
                    If UCase(arrRng(j, iLookupCol)) Like "*" & UCase(sInfo) & "*" Then
                        ReDim Preserve arrOutput(4, i) 'redim final array
                        arrOutput(0, i) = arrRng(j, arrResultsCols(0))
                        If InStr(sWbType, "SN") > 0 Then 'use initial CO/date
                            arrOutput(1, i) = TrimCO(CStr(arrRng(j, arrResultsCols(1))), False)
                            arrOutput(3, i) = arrRng(j, arrResultsCols(2))
                        Else 'use latest CO/date
                            arrOutput(1, i) = TrimCO(CStr(arrRng(j, arrResultsCols(3))), False)
                            arrOutput(3, i) = arrRng(j, arrResultsCols(4))
                        End If
                        If Not arrOutput(1, i) Like "######" Then arrOutput(1, i) = ""
                        arrOutput(2, i) = arrRng(j, arrResultsCols(5))
                        arrOutput(4, i) = arrRng(j, arrResultsCols(6))
                        i = i + 1
                    End If
                Next
            Next
            For Each varVar2 In arrWerentOpen
                If UCase(varVar2) = UCase(varVar.Name) Then
                    varVar.Close savechanges:=False
                    Exit For
                End If
            Next
        End If
    Next
    
    ThisWorkbook.Activate
    If i > 0 Then 'some results were found
        If UCase(sInfoType) = "SERIAL" Then
            Range("A1").Value = "Results for machine " & """" & sInfo & """"
        Else
            Range("A1").Value = "Results for customer " & """" & sInfo & """"
        End If
        Range("A3:E" & i + 2).Value = WorksheetFunction.Transpose(arrOutput)
        If UCase(arrOutput(0, i - 1)) = "SOURCE" Then 'no results for that sheet (remove that row)
            Range("A" & i + 2 & ":E" & i + 2).ClearContents
        End If
    Else 'no results were found
        Range("A1").Value = "No results found for " & """" & sInfo & """"
    End If
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Exit Function
errhandler:
    Application.StatusBar = False
    MsgBox "Error in FillHistory function"
    Call ErrorRep("FillHistory", "Function", "N/A", Err.Number, Err.Description, "")
End Function


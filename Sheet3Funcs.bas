Attribute VB_Name = "Sheet3Funcs"
Option Explicit
 
Function GetLookups(sInputVal As String, sInputType As String, sProdLine As String) As String()
'fills in lookup formulas on sheet4 (for some reason VBA can't pull from closed sharepoint workbooks)
'(0)SN (1)COinitial (2)COiDate (3)COlate (4)COlDate (5)Customer (6)Model
'if sInputType is not SERIAL then output (0)SN, (1)-(6) are empty
    
    Dim i As Integer, ii As Integer, j As Integer, k As Integer, x As Integer
    Dim sInputCol As String, arrOutCols As Variant
    Dim arrOutputs(7) As String, arrOut(5) As String, arrSheetNames As Variant
    Dim lateCO As Long
    Dim rngInput As Range, rngInputType As Range, rngSerNum As Range
    Dim sWBpathFull As String, sWBpath As String, sWbName As String
    Dim sColLet As String, sInputTxt As String
    Dim colWBsheets As Collection, sSheetName As String
    Dim varVar As Variant, vTestCO As Variant
    Dim bInfoFound As Boolean
    Dim rngResult As Range, arrRows() As Integer
    
    On Error GoTo 0 'errhandler
    
    With Sheet3
        '''''''''hardcoded'''''''''''''
        If UCase(sProdLine) = "BURT" Then
            Set rngInput = .Range("M2")
            Set rngInputType = .Range("M3")
        ElseIf UCase(sProdLine) = "CARR" Then
            Set rngInput = .Range("N2")
            Set rngInputType = .Range("N3")
        Else
            Set rngInput = .Range("O2")
            Set rngInputType = .Range("O3")
        End If
        '''''''''''''''''''''''''''''''
        
        rngInput.Value = sInputVal
        rngInputType.Value = sInputType
        
        Application.DisplayAlerts = False
        
        i = WorksheetFunction.Match(sProdLine, .Range("A:A"), 0) 'first row of prod line
        ii = i
        Do While .Range("A" & ii + 1).Value = 0 And .Range("C" & ii + 1).Value > 0
            ii = ii + 1
            If ii > 100 Then Exit Do 'something went wrong
        Loop
        If UCase(sInputType) = "SERIAL" Then 'prioritize/check based on SN
            j = i - 1 + WorksheetFunction.Match(sInputType, .Range("B" & i & ":B" & ii), 0) 'first row of type
            k = j
            Do While .Range("A" & k + 1).Value = 0 And .Range("C" & k + 1).Value > 0
                k = k + 1
                If k > 100 Then Exit Do 'something went wrong
            Loop 'k is last row of type
            For i = j To k 'rows
                If .Cells(i, 4).Value Like "*" & sInputVal & "*" Then
                    If arrOutputs(0) = "" Or arrOutputs(0) = "0" Then arrOutputs(0) = .Cells(i, 4).Value 'don't want to compare
                    If arrOutputs(5) = "" Or arrOutputs(5) = "0" Then arrOutputs(5) = .Cells(i, 9).Value 'don't want to compare
                    If arrOutputs(6) = "" Or arrOutputs(6) = "0" Then arrOutputs(6) = .Cells(i, 10).Value 'don't want to compare
                    If .Cells(i, 5).Value Like "*######*" Then 'initial CO number
                        If arrOutputs(1) = "" Or arrOutputs(1) = "0" Then 'one hasn't been found yet
                            arrOutputs(1) = TrimCO(.Cells(i, 5).Value, False)
                            arrOutputs(2) = .Cells(i, 6).Value
                        Else 'one was already found -- shouldn't ever happen
                            If TrimCO(.Cells(i, 5).Value, False) > arrOutputs(1) Then
                                arrOutputs(1) = TrimCO(Sheet4.Cells(i, 5).Value, False)
                                arrOutputs(2) = Sheet4.Cells(i, 6).Value
                            End If
                        End If
                    End If
                    If .Cells(i, 7).Value Like "*######*" Then 'latest CO number
                        If arrOutputs(3) = "" Or arrOutputs(3) = "0" Then 'one hasn't been found yet
                            arrOutputs(3) = TrimCO(.Cells(i, 7).Value, False)
                            arrOutputs(4) = .Cells(i, 8).Value
                        Else 'one was already found
                            If TrimCO(.Cells(k, 7).Value, False) > arrOutputs(3) Then
                                arrOutputs(3) = TrimCO(.Cells(i, 7).Value, False)
                                arrOutputs(4) = Sheet4.Cells(i, 8).Value
                            End If
                        End If
                    End If
                End If
            Next i
        Else 'find SN -> find the rest
            If sInputType = "CO" Then 'CO -> SN
                x = WorksheetFunction.CountIf(.Range("E" & i & ":E" & ii), "*" & sInputVal & "*")
                If x = 0 And IsNumeric(sInputVal) Then x = WorksheetFunction.CountIf(.Range("E" & i & ":E" & ii), CDbl(sInputVal))
                If x > 0 Then 'initial CO
                    For j = i To ii
                        If .Range("E" & j).Value Like "*" & sInputVal & "*" Then
                            If .Range("D" & j).Value > 0 Then
                                arrOutputs(0) = .Range("D" & j).Value 'SN
                                arrOutputs(1) = TrimCO(.Range("E" & j).Value, False) 'CO Init
                                arrOutputs(2) = Format(.Range("F" & j).Value, "mm/dd/yyyy") 'date
                                Exit For
                            End If
                        End If
                    Next j
                End If
                If arrOutputs(0) = "" Or arrOutputs(0) = "0" Then 'check column G
                    x = WorksheetFunction.CountIf(.Range("G" & i & ":G" & ii), "*" & sInputVal & "*")
                    If x = 0 And IsNumeric(sInputVal) Then x = WorksheetFunction.CountIf(.Range("G" & i & ":G" & ii), CDbl(sInputVal))
                    If x > 0 Then 'latest CO
                        For j = i To ii
                            If .Range("G" & j).Value Like "*" & sInputVal & "*" Then
                                If .Range("D" & j).Value > 0 Then
                                    arrOutputs(0) = .Range("D" & j).Value 'SN
                                    arrOutputs(3) = TrimCO(.Range("G" & j).Value, False) 'CO Lat
                                    arrOutputs(4) = Format(.Range("H" & j).Value, "mm/dd/yyyy") 'date
                                    Exit For
                                End If
                            End If
                        Next j
                    End If
                End If
                If arrOutputs(0) <> "" Then
                    arrOutputs(5) = .Range("I" & j).Value
                    arrOutputs(6) = .Range("J" & j).Value
                End If
            Else 'CUST -> SN
                x = WorksheetFunction.CountIf(.Range("I" & i & ":I" & ii), "*" & sInputVal & "*")
                If x = 0 And IsNumeric(sInputVal) Then x = WorksheetFunction.CountIf(.Range("I" & i & ":I" & ii), CDbl(sInputVal))
                If x > 0 Then 'something was found
                    For j = i To ii
                        If UCase(.Range("I" & j).Value) Like "*" & UCase(sInputVal) & "*" Then 'check CO for recentness
                            vTestCO = TrimCO(.Range("E" & j).Value, False) 'initial CO
                            If IsNumeric(vTestCO) Then
                                If vTestCO > lateCO Then
                                    arrOutputs(0) = .Cells(j, 4).Value
                                    lateCO = vTestCO
                                End If
                            End If
                            vTestCO = TrimCO(.Range("G" & j).Value, False) 'latest CO
                            If IsNumeric(vTestCO) Then
                                If vTestCO > lateCO Then
                                    arrOutputs(0) = .Cells(j, 4).Value
                                    lateCO = vTestCO
                                End If
                            End If
                        End If
                    Next j
                End If
            End If
            If arrOutputs(0) <> "" And arrOutputs(0) <> "0" Then 'find the rest of the details
                rngInput.Value = arrOutputs(0) 'so the formulas search by SN
                For j = i To ii
                    If .Range("D" & j).Value = arrOutputs(0) Then
                        If .Range("E" & j).Value > 0 Then 'set initial CO, customer, model
                            arrOutputs(1) = .Range("E" & j).Value
                            If .Range("I" & j).Value > 0 Then arrOutputs(5) = .Range("G" & j).Value
                            If .Range("J" & j).Value > 0 Then arrOutputs(6) = .Range("G" & j).Value
                        End If
                        If .Range("F" & j).Value > 0 Then arrOutputs(2) = Format(.Range("F" & j).Value, "mm/dd/yyyy")
                        If .Range("G" & j).Value > 0 Then arrOutputs(3) = .Range("G" & j).Value
                        If .Range("H" & j).Value > 0 Then arrOutputs(4) = Format(.Range("H" & j).Value, "mm/dd/yyyy")
                        If arrOutputs(5) = "" Or arrOutputs(5) = "0" Then
                            If .Range("I" & j).Value > 0 Then arrOutputs(5) = .Range("I" & j).Value
                        End If
                        If arrOutputs(6) = "" Or arrOutputs(6) = "0" Then
                            If .Range("J" & j).Value > 0 Then arrOutputs(6) = .Range("J" & j).Value
                        End If
                    End If
                Next j
            End If
        End If
    
    End With

    arrOut(0) = arrOutputs(0)
    arrOut(1) = TrimCO(arrOutputs(1) & "(" & arrOutputs(2) & ")", True)
    arrOut(2) = TrimCO(arrOutputs(3) & "(" & arrOutputs(4) & ")", True)
    arrOut(3) = arrOutputs(5)
    arrOut(4) = arrOutputs(6)
    
    GetLookups = arrOut
    Exit Function
errhandler:
    MsgBox "Error in GetLookups function"
    Call ErrorRep("GetLookups", "Function", "N/A", Err.Number, Err.Description, "")
End Function

Public Function AddToHistory(sProdLine As String, sSearchTerm As String)
    Dim i As Integer, sColLet As String, bPresent As Boolean, iRow As Integer
    
    '''''hardcoded''''''
        If UCase(sProdLine) = "BURT" Then
            sColLet = "M"
        ElseIf UCase(sProdLine) = "CARR" Then
            sColLet = "N"
        Else
            sColLet = "O"
        End If
    ''''''''''''''''''''
    For i = 11 To 20
        If UCase(Sheet3.Range(sColLet & i).Value) = UCase(sSearchTerm) Then
            bPresent = True
            If i = 11 Then Exit Function 'same as last search
            iRow = i
            Exit For
        End If
    Next
    
    If Not bPresent Then iRow = 20
    
    For i = iRow To 12 Step -1
        Sheet3.Range(sColLet & i).Value = Sheet3.Range(sColLet & i - 1).Value
    Next
    Sheet3.Range(sColLet & 11).Value = sSearchTerm
    
End Function

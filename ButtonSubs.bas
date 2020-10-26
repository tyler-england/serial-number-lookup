Attribute VB_Name = "ButtonSubs"
Sub GoBack()
    If ActiveSheet.Name = Sheet4.Name Then
        Range("A1").Value = "Waiting for results..."
        Range("A3:K5000").ClearContents
    End If
    Sheet1.Visible = xlSheetVisible
    ActiveSheet.Visible = xlSheetHidden
    Range("A1").Select
End Sub
Sub SeeLinks()
    Sheet2.Visible = xlSheetVisible
    ActiveSheet.Visible = xlSheetHidden
    Range("T1").Select
    ActiveWindow.ScrollColumn = 1
End Sub
Sub SerNumHistory()
    Call GlobalVariables
    Sheet1.Calculate
    Dim i As Integer
    If sProdLine = "" Then
        For i = 1 To 3
            If Cells(2, i).Value > 0 Then
                sProdLine = Cells(1, i).Value
                Exit For
            End If
        Next
    End If
    If sProdLine = "" Then
        MsgBox "No serial number has been entered"
        Exit Sub
    End If
    If Range("A6").Value = 0 Then
        MsgBox "No serial number has been entered"
        Exit Sub
    End If
    Sheet4.Visible = xlSheetVisible
    ActiveSheet.Visible = xlSheetHidden
    Range("T1").Select
    ActiveWindow.ScrollColumn = 1
    
    Call FillHistory(Sheet1.Range("A6").Value, "SERIAL", sProdLine)
End Sub
Sub CustHistory()
    Call GlobalVariables
    Sheet1.Calculate
    'determine search term - user input? results of previous search?
    Dim i As Integer, j As Integer, sTerm As String, varVar As Variant
    
    For i = 1 To 3
        If UCase(Cells(1, i).Value) = UCase(sProdLine) Then
            sTerm = Cells(2, i).Value
            Exit For
        End If
    Next
    If sTerm = "" Then
        If sProdLine = "" Then 'shouldn't happen
            For i = 1 To 3
                If Cells(2, i).Value > 0 Then
                    sProdLine = Cells(1, i).Value
                    sTerm = Cells(2, i).Value
                    Exit For
                End If
            Next
        End If
        If sProdLine = "" Then
            MsgBox "No customer info has been entered"
            Exit Sub
        End If
    End If
    varVar = DetectInputType(sTerm, sProdLine)
    If varVar(0) <> "CUST" Then sTerm = Range("B8").Value 'use the customer output from the last search
    If sTerm = "" Then
        MsgBox "No customer info available for this project"
        Exit Sub
    End If
    Sheet4.Visible = xlSheetVisible
    ActiveSheet.Visible = xlSheetHidden
    Range("T1").Select
    ActiveWindow.ScrollColumn = 1
    Call FillHistory(sTerm, "CUST", sProdLine)
End Sub

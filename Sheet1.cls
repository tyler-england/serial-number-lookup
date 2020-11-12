VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public bDuringChange As Boolean, bInputRange As Boolean, rngEnd As Range
Private Sub Worksheet_Activate()
    bDuringChange = False
    Set rngEnd = Range("A2")
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo errhandler
    
    If Not bDuringChange Then
        Call Sheet1Vars 'sets ranges for burt,carr,mateer inputs
        Dim varVar As Variant
        For Each varVar In arrRngInput
            If Target.Cells(1).Address = varVar.Address Then
                bInputRange = True 'adjusted cell is in one of those ranges
                Exit For
            End If
        Next
    End If

    If bInputRange Then '
    
        bDuringChange = True 'prevents looping
        bInputRange = False 'prevents looping
        
        Set rngEnd = Target 'range to select after macro does stuff
        
        Dim sInput As String, sInputTyp As String, sTargAddr As String
        Dim lCoNum As Long, lSerNum As Long
        Dim bContinue As Boolean, bFast As Boolean, bLinks As Boolean, bLinksAll As Boolean
        Dim arrOutputInfo As Variant, arrLookups As Variant
        Dim sCO As String
        Dim i As Integer
        
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        sTargAddr = Target.Cells(1).Address 'required because merged cells for input ranges have diff. targets --> "cells(1)" is nec.
        
        If sTargAddr = rngBrt.Address Or sTargAddr = rngMat.Address Or sTargAddr = rngCar.Address Then 'equipment was edited
            Call FormatPlain(rngSN)
            Call FormatPlain(rngCoInit)
            Call FormatPlain(rngCOlat)
            Call FormatPlain(rngLayout)
            If Target.Cells(1).Value = 0 Then 'range input was cleared
                rngSN.Value = ""
                rngCoInit.Value = ""
                rngCOlat.Value = ""
                rngLayout.Value = ""
                rngCust.Value = ""
                rngModel.Value = ""
                bDuringChange = False
                rngEnd.Select
                Exit Sub
            End If
            
            Application.StatusBar = "Looking for primary details..."
            sProdLine = Target.Offset(-1, 0).Value
            
            varVar = DetectInputType(Target.Value, sProdLine) 'find out what was entered (SN, CO, cust, none)
            sInputTyp = varVar(0) 'SERIAL, CO, CUST, or NONE
            sInput = varVar(1) 'the data
            Call GlobalVariables
            
            Call AddToHistory(sProdLine, sInput)
            varVar = CheckFastLinks
            bFast = varVar(0)
            If varVar(1) Then 'see if all or local
                bLinks = varVar(1)
                bLinksAll = varVar(2) 'if false, then only vault/local links
            End If
            rngCOlat.Offset(-1, 0).Value = "Latest CO" 'required because sometimes it changes
            
            If bFast Then 'use lookup formulas to find results (last result may not actually be "latest")
                arrOutputInfo = GetLookups(sInput, sInputTyp, sProdLine)
            Else 'open each workbook and get the actual results
                arrOutputInfo = GivenInfo(sInput, sInputTyp, sProdLine)
            End If
            
            For i = 0 To UBound(arrOutputInfo)
                If arrOutputInfo(i) = "" Then arrOutputInfo(i) = "Not found"
            Next i
            
            rngSN.Value = arrOutputInfo(0)
            rngCoInit.Value = arrOutputInfo(1)
            rngCOlat.Value = arrOutputInfo(2)
            rngCust.Value = arrOutputInfo(3)
            rngModel.Value = arrOutputInfo(4)
            
            If UCase(rngCoInit.Value) Like "*NOT*FOUND*" And UCase(sProdLine) <> "CARR" Then
                rngLayout.Value = "Not found"
            ElseIf Not bLinks Then
                rngLayout.Value = "-"
            Else 'find layout #
                Call LinkEwsQuoteLayout(rngSN, rngCoInit, rngCOlat, rngLayout, sProdLine, bLinks, bLinksAll) 'tries to find the layout & link it
            End If
            Application.StatusBar = False
        ElseIf sTargAddr = rngPhonePer.Address Or sTargAddr = rngPhoneRm.Address Then 'telephone section was edited
            If Target.Cells(1).Value = 0 Then 'data was cleared
                rngPhoneName.Value = ""
                rngPhoneExt.Value = ""
                bDuringChange = False
                rngEnd.Select
                Exit Sub
            ElseIf Target.Address = rngPhonePer.Address Then 'clear room
                rngPhoneRm.Value = ""
            ElseIf Target.Address = rngPhoneRm.Address Then 'clear person
                rngPhonePer.Value = ""
            End If
            
            sInput = Target.Value
            If Target.Address = rngPhonePer.Address Then 'phone (person) input
                sInputTyp = "Name"
            ElseIf Target.Address = rngPhoneRm.Address Then 'phone (room) input
                sInputTyp = "Room"
            End If
            
            bContinue = PhoneNumber(sInput, sInputTyp, rngPhoneName, rngPhoneExt)
            If Not bContinue Then GoTo errhandler 'almost unnecessary - only avoids a double error
            
        End If
        bDuringChange = False
        rngEnd.Select
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
    End If
errhandler:
    Application.StatusBar = False
    bDuringChange = False
End Sub

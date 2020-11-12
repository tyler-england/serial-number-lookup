VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    If InStr(UCase(Application.UserName), "ENGLAND") > 0 Then ExportModules
End Sub


Private Sub Workbook_Open()
    ThisWorkbook.AutoSaveOn = False
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheet1.Visible = xlSheetVisible
    Dim i As Integer, j As Integer
    Call GlobalVariables
    For Each Sheet In ThisWorkbook.Sheets
        If Sheet.Name <> Sheet1.Name And Sheet.Visible = xlSheetVisible Then
            Sheet.Visible = xlSheetHidden
        End If
    Next
    Call Sheet1Vars 'required to set ranges
    'Call MapSharepoint(sSharePointLink)
    bLinksUpdated = False 'only updates links the first time, resets upon wkbk open
    Application.DisplayAlerts = True
End Sub

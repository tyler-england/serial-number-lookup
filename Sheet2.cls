VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    
    Call GlobalVariables
    
    If Not SharePointAccess Then
        Call MapSharepoint(sSharePointLink)
    End If
        
    If Not bLinksUpdated Then
        Application.StatusBar = "Making sure all the links are correct..."
        Application.ScreenUpdating = False
        
        Dim sSpDriveLet As String
        
        Dim rngEwsList As Range, rngEwsArchive As Range, rngBOMLoader As Range
        Dim rngBurtLog As Range, rngBurtAM As Range
        Dim rngCarrLog As Range, rngCentLog As Range, rngCarrAM As Range
        Dim rngMatLog As Range, rngMatPM As Range
        Dim rngPhoneCLW As Range, rngPhoneAKR As Range
        
        ''''''hardcoded'''''''''''''
        Set rngEwsList = Range("B2")
        Set rngBurtLog = Range("C2")
        Set rngCarrLog = Range("D2")
        Set rngMatLog = Range("E2")
        Set rngPhoneAKR = Range("F2")
        ''''''''''''''''''''''''''''
        Set rngEwsArchive = rngEwsList.Offset(1, 0)
        Set rngBOMLoader = rngEwsArchive.Offset(1, 0)
        Set rngBurtAM = rngBurtLog.Offset(1, 0)
        Set rngCentLog = rngCarrLog.Offset(1, 0)
        Set rngCarrAM = rngCentLog.Offset(1, 0)
        Set rngMatPM = rngMatLog.Offset(1, 0)
        Set rngPhoneCLW = rngPhoneAKR.Offset(1, 0)
        
        Call FillLink(rngEwsList, sEWSListFolder, "Online EWS List", "Folder")
        Call FillLink(rngEwsArchive, sEWSArchFolder, "EWS Archive", "Folder")
        Call FillLink(rngBOMLoader, FindMostCurrent(sBOMLoaderFolder), "BOM Loader")
        Call FillLink(rngBurtLog, sBurtLogPath, "Serial Number Log")
        Call FillLink(rngBurtAM, sBurtPathAM, "Aftermarket Log")
        Call FillLink(rngCarrLog, sCarrLogPath, "Carr SN Log")
        Call FillLink(rngCentLog, sCentLogPath, "Centritech SN Log")
        Call FillLink(rngCarrAM, sCarrPathAM, "Aftermarket Log")
        Call FillLink(rngMatLog, sMateerLogFolder, "Serial Number Log", "Mateer")
        Call FillLink(rngMatPM, sMateerPmPath, "Proj. Man. Report")
        Call FillLink(rngPhoneCLW, sPhonePathCLW, "Clearwater Ext's")
        Call FillLink(rngPhoneAKR, sPhonePathAKR, "Akron/Stow Ext's")
        
        If sSpDriveLet <> "" Then
            Call UnmapSharepoint(sSpDriveLet) 'unmap the sharepoint drive
            sSpDriveLet = "" 'clear so we don't try unmapping again
        End If
        
        bLinksUpdated = True
        Application.ScreenUpdating = True
        Application.StatusBar = False
    End If
    
    Exit Sub
errhandler:
    On Error Resume Next
    Application.StatusBar = False
    If sSpDriveLet <> "" Then Call UnmapSharepoint(sSpDriveLet)  'unmap the sharepoint drive
    MsgBox "Unable to update all the links"
End Sub


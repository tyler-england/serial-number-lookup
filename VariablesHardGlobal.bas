Attribute VB_Name = "VariablesHardGlobal"
Option Explicit
Public sEWSListFolder As String, sEWSArchFolder As String
Public sBOMLoaderPath As String, sBOMLoaderFolder As String
Public sPhonePathCLW As String, sPhonePathAKR As String
Public sMateerLogFolder As String, sMateerPmPath As String, sMateerLogPath As String
Public sBurtLogPath As String, sBurtPathAM As String
Public sCarrLogPath As String, sCentLogPath As String, sCarrPathAM As String
Public sSharePointLink As String, sSharePointMap As String
Public bLinksUpdated As Boolean 'not involved in globalvars... for links on sheet2
Public arrErrorEmails() As String, iNumMsgs As Integer, bNewMsg As Boolean 'for ErrorRep
Public rngMatSN As Range, rngMatPM1 As Range, rngMatPM2 As Range, rngMatPM3 As Range
Public rngCarrSN As Range, rngCentSN As Range, rngBurtSN As Range
Public sProdLine As String

Sub sptests()
    'Works but paths are listed as URLs
    Dim wb As Workbook
    Set wb = Workbooks.Open("https://bw1.sharepoint.com/:x:/r/sites/PSA/PSAENG/PSAGDI/Engineering%20Work%20Sheet%20EWS/EWS300000%20to%20EWS399999/EWS300000/EWS300000/EWS300000.xls")
    
    
'    'doesn't work
'    Dim folder As folder
'    Dim f As File
'    Dim fs As New FileSystemObject
'    Dim wb As Workbook
'
'    Set folder = fs.GetFolder("//bw1.sharepoint.com/sites/PSA/PSApjm/SitePages/Home.aspx")
'
'    For Each f In folder.Files
'       If f.Name Like "*" Then
'           Debug.Print f.Name
'       End If
'    Next f


'    'try if login gets complicated
'    'try FollowHyperlink

End Sub

Sub GlobalVariables()
    On Error Resume Next
    '''''''''''''''''''''hardcoded values''''''''''''''''
    sSharePointLink = "https://bw1.sharepoint.com/sites/PSA/PSAENG/PSAGDI/Engineering%20Work%20Sheet%20EWS/Forms/AllItems.aspx?ExplorerWindowUrl=%2Fsites%2FPSA%2FPSAENG%2FPSAGDI%2FEngineering%20Work%20Sheet%20EWS"
    sSharePointMap = "\\bw1.sharepoint.com@SSL\DavWWWRoot\sites\PSA\PSAENG\PSAGDI"
    sEWSListFolder = "\\bw1.sharepoint.com@SSL\DavWWWRoot\sites\PSA\PSAENG\PSAGDI\Engineering Work Sheet EWS"
    sEWSArchFolder = "\\bw1.sharepoint.com@SSL\DavWWWRoot\sites\PSA\PSAENG\PSAGDI\EWS Archive"
    sBOMLoaderFolder = "\\bw1.sharepoint.com@SSL\DavWWWRoot\sites\PSA\PSAMFG\Shared Documents\BOM Loader"
    sPhonePathCLW = "\\bw1.sharepoint.com@SSL\DavWWWRoot\sites\PSA\PSAngelus Home\PSA Phone Numbers\PSA Extensions Clearwater.xlsx"
    sPhonePathAKR = "\\bw1.sharepoint.com@SSL\DavWWWRoot\sites\PSA\PSAngelus Home\PSA Phone Numbers\PSA Akron-Stow Phone List.xlsx"
    sMateerLogFolder = "\\bw1.sharepoint.com@SSL\DavWWWRoot\sites\PSA\PSAPJM\Shared Documents\New Machine Orders\Akron"
    sMateerLogPath = "\\bw1.sharepoint.com@SSL\DavWWWRoot\sites\PSA\PSAPJM\Shared Documents\New Machine Orders\Akron\PSC Machine Serial Number Log - Filler Mateer.xlsx"
    sMateerPmPath = "\\PSACLW02\HOME\SHARED\MATEER\PROJECT MANAGEMENT\PM REPORT.XLSX"
    sBurtLogPath = "\\PSACLW02\RELEASED\MB SOP\BURT MK 6 PROGRAMS\MACHINE LISTS.XLSX"
    sBurtPathAM = "A"
    sCarrLogPath = "C:\EPDM\PSA_VAULT\CLW\CARR-CENT SERIAL NUMBER LOGS\CARR SERIAL NO LOG.XLS"
    sCentLogPath = "C:\EPDM\PSA_VAULT\CLW\CARR-CENT SERIAL NUMBER LOGS\CENTRITECH SERIAL NUMBER LOG.XLSX"
    sCarrPathAM = "A"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Sub ToggleFast()
Attribute ToggleFast.VB_ProcData.VB_Invoke_Func = "M\n14"
    Call Sheet1Vars
    If UCase(rngFast.Value) = "HARE" Then
        rngFast.Value = "TORTOISE"
        Exit Sub
    Else
        rngFast.Value = "HARE"
    End If
End Sub
Sub ToggleLinks()
Attribute ToggleLinks.VB_ProcData.VB_Invoke_Func = "L\n14"
    Call Sheet1Vars
    If UCase(rngLinks.Value) = "NO LINKS" Then
        rngLinks.Value = "LOC LINKS"
        Exit Sub
    ElseIf UCase(rngLinks.Value) = "LOC LINKS" Then
        rngLinks.Value = "ALL LINKS"
        If Not SharePointAccess Then Call MapSharepoint(sSharePointLink)
        Exit Sub
    Else
        rngLinks.Value = "NO LINKS"
    End If
End Sub

Function ExportModules() As Boolean
    Dim wbMacro As Workbook, varVar As Variant, bOpen As Boolean
    For Each varVar In Application.Workbooks
        If UCase(varVar.Name) = "MACROBOOK.XLSM" Then
            bOpen = True
            Set wbMacro = varVar
            Exit For
        End If
    Next
    If Not bOpen Then Set wbMacro = Workbooks.Open("\\PSACLW02\HOME\SHARED\MacroBook.xlsm")
    Application.Run "'" & wbMacro.Name & "'!ExportModules", ThisWorkbook
    If Not bOpen Then wbMacro.Close savechanges:=False
    ExportModules = True
End Function

Public Sub ErrorRep(rouName, rouType, curVal, errNum, errDesc, miscInfo)
    Dim wbMacro As Workbook, varVar As Variant, bOpen As Boolean
    bNewMsg = True 'default value
    If iNumMsgs > 0 Then 'at least one email has been generated already
        For Each varVar In arrErrorEmails 'see if there were any matches
            If UCase(varVar) = UCase(ThisWorkbook.Name & "-" & errNum) Then Exit Sub 'repeat message (this was already sent this session)
        Next
    End If
    For Each varVar In Application.Workbooks
        If UCase(varVar.Name) = "MACROBOOK.XLSM" Then
            bOpen = True
            Set wbMacro = varVar
            Exit For
        End If
    Next
    If Not bOpen Then Set wbMacro = Workbooks.Open("\\PSACLW02\HOME\SHARED\MacroBook.xlsm")
    Application.Run "'MacroBook.xlsm'!ErrorReport", rouName, rouType, curVal, errNum, errDesc, miscInfo
    If Not bOpen Then wbMacro.Close savechanges:=False
    iNumMsgs = iNumMsgs + 1
    ReDim Preserve arrErrorEmails(iNumMsgs)
    arrErrorEmails(iNumMsgs) = ThisWorkbook.Name & "-" & errNum
End Sub




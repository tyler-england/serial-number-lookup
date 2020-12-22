Attribute VB_Name = "MiscFuncs"
Option Explicit


Public Function Testlink(link As String) As Boolean
'tests online/sharepoint link
Dim RequestSharepoint As New WinHttpRequest
    
    On Error GoTo testlinkError
    RequestSharepoint.Open "HEAD", link
    RequestSharepoint.Send
    If RequestSharepoint.Status = 401 Or RequestSharepoint.Status = 200 Then
        Testlink = True
    Else
        Testlink = False
    End If
    Exit Function
    
testlinkError:
    Testlink = False
End Function
Function SearchVault(sSearchInp As String, Optional sExtPatt As String) As String
    
    Dim i As Integer
    Dim eVault As IEdmVault5, eSearch As IEdmSearch5, eResult As IEdmSearchResult5
    Dim eFolder As IEdmFolder5, eName As IEdmObject5
    Dim vaultName As String, sOutput As String, sResult As String, sSearchTerm As String
    On Error GoTo errhandler
    
    vaultName = "PSA_Vault" 'vault name
    Set eVault = New EdmVault5
    If Not eVault.IsLoggedIn Then 'log into vault
        Call eVault.LoginAuto(vaultName, 0)
    End If
    
    Set eSearch = eVault.CreateSearch 'set search
    eSearch.FindFiles = True
    eSearch.FindUnlockedFiles = True
    eSearch.FindLockedFiles = True
    eSearch.Recursive = True
    
    If sExtPatt <> "" Then
        If Left(sExtPatt, 1) <> "." Then sExtPatt = "." & sExtPatt
    End If
    
    sSearchTerm = sSearchInp & sExtPatt

    sSearchTerm = Replace(UCase(sSearchTerm), "*", "%") 'wildcard is %
    
    'eSearch.AddVariable "Part Number", sSearchTerm 'for use with Part Number
    eSearch.fileName = sSearchTerm
    
    Set eResult = eSearch.GetFirstResult 'results of search
    If Not eResult Is Nothing Then
        If InStr(UCase(eResult.Path), "VIRTUAL") > 0 Then
            Set eResult = eSearch.GetNextResult
        End If
        i = 0
        On Error Resume Next
        Do While Not UCase(Right(eResult.Path, Len(eResult.Path) - InStrRev(eResult.Path, "\"))) Like Replace(sSearchTerm, "%", "*") _
                Or UCase(Right(eResult.Path, Len(eResult.Path) - InStrRev(eResult.Path, "\"))) Like "*L#*" 'required for Carr layouts
            Set eResult = eSearch.GetNextResult
            i = i + 1
            If i > 15 Then 'in case there are no files that qualify
                Exit Do
            End If
        Loop
        If i < 15 And Not eResult Is Nothing Then sOutput = eResult.Path
        
    End If
    SearchVault = sOutput
errhandler:
End Function
Public Function GetLatest(sFilePath As String) As String
'gets latest version of an EPDM vault document
    Dim vaultName As String, i As Integer
    Dim eVault As IEdmVault5, eFile As IEdmFile8, eFolder As IEdmFolder5
    Dim ePos As IEdmPos5, arrFiles(0) As EdmSelItem, bchGet As IEdmBatchGet
    Dim vbAns As Variant

    On Error GoTo errhandler
    
    If InStr(UCase(sFilePath), "EPDM") = 0 Then
        GetLatest = sFilePath
        Exit Function
    End If
    GetLatest = ""
    
    vaultName = "PSA_Vault" 'vault name
    Set eVault = New EdmVault5
    If Not eVault.IsLoggedIn Then 'log into vault
        Call eVault.LoginAuto(vaultName, 0)
    End If
    
    Set eFile = eVault.GetFileFromPath(sFilePath)
    
    If eFile.CurrentVersion = eFile.GetLocalVersionNo(sFilePath) Then 'no GetLatest required
        GetLatest = sFilePath
        Exit Function
    Else 'check if user wants to overwrite
        vbAns = MsgBox("You don't have the latest version of the following file..." & vbCrLf & vbCrLf & _
                    Right(sFilePath, Len(sFilePath) - InStrRev(sFilePath, "\")) & vbCrLf & vbCrLf & _
                    "Do you want to get the latest version & overwrite the one on your local disk?", vbYesNo)
        If vbAns = vbNo Then 'abandon the GetLatest
            GetLatest = sFilePath
            Exit Function
        End If
    End If
    
    Set ePos = eFile.GetFirstFolderPosition 'req. for efolder
    Set eFolder = eFile.GetNextFolder(ePos)
    
    arrFiles(0).mlDocID = eFile.ID
    arrFiles(0).mlProjID = eFolder.ID
    
    Set bchGet = eVault.CreateUtility(EdmUtil_BatchGet)
    Call bchGet.AddSelection(eVault, arrFiles)
    Call bchGet.CreateTree(Application.hwnd, EdmGetCmdFlags.Egcf_Nothing)
    Call bchGet.GetFiles(Application.hwnd, Nothing)
    
    
    GetLatest = sFilePath
    Exit Function
errhandler:
    MsgBox "Unable to get file " & Right(sFilePath, Len(sFilePath) - InStrRev(sFilePath, "\")) & " from the Vault"
End Function
Function FindFullPath(sPath As String, sFileName As String, sExtPatt As String, Optional bFolder As Boolean) As String

    Dim FSO As New FileSystemObject
    Dim myFolder As Scripting.folder
    Dim mySubFolder As Scripting.folder
    Dim myFile As Scripting.File
    Dim sOut As String
    On Error GoTo errhandler
    Set myFolder = FSO.GetFolder(sPath)
    
    If Left(sExtPatt, 1) <> "." Then sExtPatt = "." & sExtPatt
    
    If bFolder Then
        For Each mySubFolder In myFolder.SubFolders
            If UCase(mySubFolder.Name) Like sFileName Then
                sOut = mySubFolder.Path & "\" & mySubFolder.Name
                Exit For
            End If
            If sOut <> "" Then Exit For
            sOut = FindFullPath(mySubFolder.Path, sFileName, sExtPatt, True)
        Next
    Else
        For Each mySubFolder In myFolder.SubFolders
            For Each myFile In mySubFolder.Files
                'If sOut <> "" Then Exit For
                If UCase(myFile.Name) Like UCase("*" & sFileName & "*" & sExtPatt) Then
                    sOut = myFile.Path 'Or do whatever you want with the file
                    Exit For
                End If
            Next
            If sOut <> "" Then Exit For
            sOut = FindFullPath(mySubFolder.Path, sFileName, sExtPatt)
        Next
    End If
    FindFullPath = sOut
    Exit Function
errhandler:
    Call ErrorRep("FindFullPath", "Function", sOut, Err.Number, Err.Description, "")
End Function

Function FindFullPathz(sUmbPath As String, sFileName As String, sExtPat As String) As String
'gets file path for file name
    Dim oFSO As Object
    Dim myFolder As Scripting.folder
    Dim mySubFolder As Scripting.folder
    Dim myFile As Scripting.File
    Dim sOut As String
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set myFolder = oFSO.GetFolder(sUmbPath)
    
    If Left(sExtPat, 1) <> "." Then sExtPat = "." & sExtPat
    
    For Each mySubFolder In myFolder.SubFolders
        For Each myFile In mySubFolder.Files
            If UCase(myFile.Name) Like UCase("*" & sFileName & "*" & sExtPat) Then
                Debug.Print "path:" & myFile.Path
                sOut = myFile.Path
                Exit For
            End If
        Next
        If sOut <> "" Then Exit For
        FindFullPath = FindFullPath(mySubFolder.Path, sFileName, sExtPat)
    Next
    FindFullPath = sOut
End Function
Function FindMostCurrent(sFolderPath As String, Optional sKeyword As String) As String
'returns most current file in directory (containing skeyword if given)
    
    Dim sOutFile As String, sTestFile As String, dLatestDate As Date, sSearchTerm As String
    On Error GoTo 0 'errhandler
    dLatestDate = 0
    If Right(sFolderPath, 1) <> "\" Then sFolderPath = sFolderPath & "\"
    sSearchTerm = sFolderPath & "*" & sKeyword & "*"
    
    On Error Resume Next
        sTestFile = Dir$(sSearchTerm)
        If Err.Number = 52 Then 'sharepoint needs mapping
            Dim sDriveLet As String
            sDriveLet = MapSharepoint(sSharePointLink)
            sTestFile = Dir$(sSearchTerm)
            Call UnmapSharepoint(sDriveLet)
        End If
    On Error GoTo 0 'errhandler
    
    If Right(sFolderPath, 1) <> "\" Then sFolderPath = sFolderPath & "\"
    Do While sTestFile <> "" 'do for all matching files in directory
        If FileDateTime(sFolderPath & sTestFile) > dLatestDate Then
            sOutFile = sTestFile
            dLatestDate = FileDateTime(sFolderPath & sTestFile)
        End If
        sTestFile = Dir$()
    Loop
    If sOutFile = "" Then 'no result
        FindMostCurrent = "-"
    Else
        FindMostCurrent = sFolderPath & sOutFile
    End If
    Exit Function
errhandler:
    MsgBox "Error in FindMostCurrent function"
    Call ErrorRep("FindMostCurrent", "Function", FindMostCurrent, Err.Number, Err.Description, "")
End Function

Function ListDrives() As String()
    Dim WQL As String
    Dim i As Integer
    Dim SrvEx As Variant, WMIObj As Variant, WMIObjEx As Variant
    Dim arrOut() As String
    WQL = "Select * From Win32_LogicalDisk"
    Set SrvEx = GetObject("winmgmts:root/CIMV2")
    Set WMIObj = SrvEx.ExecQuery(WQL)
    ReDim arrOut(0)
    For Each WMIObjEx In WMIObj
        ReDim Preserve arrOut(i)
        arrOut(i) = WMIObjEx.Name & WMIObjEx.ProviderName 'returns drive letter & ":" & path
        i = i + 1
    Next
    ListDrives = arrOut
End Function
Function GetLookupColumn(sProdLine As String, sLogType As String, sInputType As String) As Integer
    'On Error GoTo errhandler
'column with the info being used as the source of the lookup

    If InStr(sLogType, " ") > 0 Then sLogType = Trim(Left(sLogType, InStr(sLogType, " ")))
    Dim sColLet As Integer
    If UCase(sInputType) = "SERIAL" Then 'search by SN
        If UCase(sLogType) = "SN" Then 'SN log
            If UCase(sProdLine) = "BURT" Then 'burt log"
                sColLet = 4
            ElseIf UCase(sProdLine) = "CARR" Then  'carr log"
                sColLet = 1
            Else 'assume mateer
                sColLet = 2
            End If
        ElseIf UCase(sLogType) = "AM" Then 'aftermarket log
            If UCase(sProdLine) = "BURT" Then 'burt log
                sColLet = 13 'no such log
            ElseIf UCase(sProdLine) = "CARR" Then 'carr log
                sColLet = 13 'no such log
            Else 'assume mateer
                sColLet = 4
            End If
        Else 'assume "other"
            If UCase(sProdLine) = "BURT" Then 'burt log
                sColLet = 13 'no such log
            ElseIf UCase(sProdLine) = "CARR" Then 'carr log
                sColLet = 2
            Else 'assume mateer
                sColLet = 13 'no such log
            End If
        End If
    ElseIf UCase(sInputType) = "CO" Then 'search by CO
        If UCase(sLogType) = "SN" Then 'SN log
            If UCase(sProdLine) = "BURT" Then 'burt log
                sColLet = 2
            ElseIf UCase(sProdLine) = "CARR" Then 'carr log
                sColLet = 5
            Else 'assume mateer
                sColLet = 3
            End If
        ElseIf UCase(sLogType) = "AM" Then 'aftermarket log
            If UCase(sProdLine) = "BURT" Then 'burt log
                sColLet = 13 'no such log
            ElseIf UCase(sProdLine) = "CARR" Then 'carr log
                sColLet = 13 'no such log
            Else 'assume mateer
                sColLet = 3
            End If
        Else 'assume "other"
            If UCase(sProdLine) = "BURT" Then 'burt log
                sColLet = 13 'no such log
            ElseIf UCase(sProdLine) = "CARR" Then 'carr log
                sColLet = 3
            Else 'assume mateer
                sColLet = 13 'no such log
            End If
        End If
    Else 'assume customer
        If UCase(sLogType) = "SN" Then 'SN log
            If UCase(sProdLine) = "BURT" Then 'burt log
                sColLet = 3
            ElseIf UCase(sProdLine) = "CARR" Then 'carr log
                sColLet = 2
            Else 'assume mateer
                sColLet = 4
            End If
        ElseIf UCase(sLogType) = "AM" Then 'aftermarket log
            If UCase(sProdLine) = "BURT" Then 'burt log
                sColLet = 13 'no such log
            ElseIf UCase(sProdLine) = "CARR" Then 'carr log
                sColLet = 4 'no such log, but Centritech log is used
            Else 'assume mateer
                sColLet = 7
            End If
        Else 'assume "other"
            If UCase(sProdLine) = "BURT" Then 'burt log
                sColLet = 13 'no such log
            ElseIf UCase(sProdLine) = "CARR" Then 'carr log
                sColLet = 4
            Else 'assume mateer
                sColLet = 13 'no such log
            End If
        End If
    End If

    GetLookupColumn = sColLet
    Exit Function

errhandler:
    MsgBox "Error in GetLookupColumn function"
    Call ErrorRep("GetLookupColumn", "Function", "N/A", Err.Number, Err.Description, "")
End Function
Function GetOutputCols(sProdLine As String, sLogType As String) As Variant()
'columns with the info to be looked up

    'On Error GoTo errhandler
    If InStr(sLogType, " ") > 0 Then
        sLogType = UCase(Trim(Left(sLogType, InStr(sLogType, " "))))
    End If
    Dim arrOut(6) As Variant, varVar As Variant, i As Integer
    '(0)SN (1)COinitial (2)COiDate (3)COlate (4)COlDate (5)Customer (6)Model
    
    If UCase(sProdLine) = "BURT" Then 'burt
        If sLogType = "SN" Then
            varVar = Array(4, 2, 13, 13, 13, 3, 1)
        ElseIf sLogType = "AM" Then
            varVar = Array(13, 13, 13, 13, 13, 13, 13)
        Else
            varVar = Array(13, 13, 13, 13, 13, 13, 13)
        End If
    ElseIf UCase(sProdLine) = "CARR" Then 'carr
        If sLogType = "SN" Then
            varVar = Array(1, 5, 4, 13, 13, 2, 1)
'            ElseIf sLogType = "AM" Then 'no such log -> use Centritech log
'                varVar = Array(13, 13, 13, 13, 13, 13, 13)
        Else
            varVar = Array(2, 3, 5, 3, 5, 4, 1) '3 & 5 used for initial and latest
        End If
    Else ' assume mateer
        If sLogType = "SN" Then
            varVar = Array(2, 3, 6, 13, 13, 4, 5)
        ElseIf sLogType = "AM" Then
            varVar = Array(4, 13, 13, 3, 11, 7, 8)
        Else
            varVar = Array(13, 13, 13, 13, 13, 13, 13)
        End If
    End If
    
    For i = 0 To 6 'turn variant into array
        arrOut(i) = CInt(varVar(i))
    Next
    
    GetOutputCols = arrOut
    
    Exit Function
errhandler:
    MsgBox "Error in GetOutputCols function"
    Call ErrorRep("GetOutputCols", "Function", "N/A", Err.Number, Err.Description, "")
End Function

Function MapSharepoint(sSharePointLink As String) As String
'maps sharepoint to a drive, letting user bypass the need for credentials on desktop

    Dim sDriveLet As String, i As Integer, j As Integer, iStage As Integer
    Dim oMappedDrive As Scripting.drive
    Dim oFSO As New Scripting.FileSystemObject
    Dim oNetwork As New WshNetwork
    Dim oIE As InternetExplorer
    Dim oWinTest As Object, oShell As Object, oWindow As Object
    Dim varVar As Variant, arrMappedPaths As Variant
    Dim bIeOpen As Boolean, bMapped As Boolean, sTest As String
    
    'On Error GoTo errhandler
    iStage = 0
    Call GlobalVariables
    
    arrMappedPaths = ListDrives
    For Each varVar In arrMappedPaths
        If UCase(Right(varVar, Len(varVar) - 2)) = UCase(sSharePointMap) Then
            bMapped = True
            sDriveLet = Left(varVar, 1)
            Exit For
        End If
    Next

    If bMapped Then 'test if authentication is expired
        If Not SharePointAccess Then 'expired
            Call UnmapSharepoint(sDriveLet)
            bMapped = False
        End If
    End If

    If Not bMapped Then 'need to map to a drive letter
        bIeOpen = False
        Set oShell = CreateObject("Shell.Application").Windows()
        For Each varVar In oShell
            If (Not varVar Is Nothing) And varVar.Name = "Internet Explorer" Then
                Set oIE = varVar
                bIeOpen = True
            End If
        Next
        iStage = 1
        If oIE Is Nothing Then Set oIE = CreateObject("InternetExplorer.Application")
        Application.ScreenUpdating = True
        oIE.Visible = True
        
        Do While j < 5 And Not bMapped
        
            If bIeOpen Then
                oIE.Navigate sSharePointLink, CLng(2048) 'new tab
            Else
                oIE.Navigate sSharePointLink 'no new tab
            End If
            
           ' Application.Wait (Now + TimeValue("0:00:10")) 'necessary?
            
            For i = Asc("Z") To Asc("A") Step -1
                sDriveLet = Chr(i)
                If Not oFSO.DriveExists(sDriveLet) Then
                    Application.DisplayAlerts = True
                    Application.ScreenUpdating = True
                    oNetwork.MapNetworkDrive sDriveLet & ":", sSharePointMap
                    bMapped = True
                    Exit For
                End If
            Next

            For Each oWinTest In oShell 'close ssharepointmap explorer window
                If InStr(oWinTest.LocationURL, "sharepoint.com") Then oWinTest.Quit 'a bit broad but probably fine
            Next
            
            Debug.Print "map: " & bMapped & " (" & sDriveLet & ")"
            Debug.Print "access: " & SharePointAccess
            
            If Not SharePointAccess Then bMapped = False 'retry :(
            j = j + 1
            
        Loop
        
        iStage = 3
        If Not bIeOpen Then oIE.Quit
        iStage = 4
    End If
    
    MapSharepoint = sDriveLet
    
    Exit Function
errhandler:
    'MsgBox "Unable to map SharePoint drive (Error " & Err.Number & ")"
    Application.StatusBar = False
    Call ErrorRep("MapSharepoint", "Function", MapSharepoint, Err.Number, Err.Description, "stage: " & iStage)
End Function

Function UnmapSharepoint(sDriveLet As String)
'gets rid of sharepoint drive map
    Dim oMappedDrive As Scripting.drive
    Dim oFSO As New Scripting.FileSystemObject
    Dim oNetwork As New WshNetwork
    On Error Resume Next
    Set oMappedDrive = oFSO.GetDrive(sDriveLet)
    If Not oMappedDrive Is Nothing Then
        If oMappedDrive.IsReady Then
          oNetwork.RemoveNetworkDrive sDriveLet & ":"
        End If
        Set oMappedDrive = Nothing
  End If
End Function
Function SharePointAccess() As Boolean
    SharePointAccess = False
    On Error GoTo errhandler
    Dim str As String
    str = Dir$(sSharePointMap, vbDirectory)
    SharePointAccess = True
errhandler:
End Function
Function GetClosedWbValue(sPathName, sFileName, sSheetName, sRngName)
'Retrieves a value from a closed workbook
    On Error GoTo errhandler
    Dim sArg As String
    'Make sure the file exists
    If Right(sPathName, 1) <> "\" Then sPathName = sPathName & "\"
    If Dir(sPathName & sFileName) = "" Then
        GetClosedWbValue = ""
        Exit Function
    End If
    'Create the argument
    sArg = "'" & sPathName & "[" & sFileName & "]" & sSheetName & "'!" & Range(sRngName).Address(, , xlR1C1)
    'Execute XLM macro
    GetClosedWbValue = ExecuteExcel4Macro(sArg)
    Exit Function
errhandler:
End Function
Function GetClosedWbRange(sPathName, sFileName, sSheetName, sRngName)
'Retrieves a value from a closed workbook
    On Error GoTo errhandler
    Dim sArg As String
    'Make sure the file exists
    If Right(sPathName, 1) <> "\" Then sPathName = sPathName & "\"
    If Dir(sPathName & sFileName) = "" Then
        GetClosedWbRange = ""
        Exit Function
    End If
    'Create the argument
    sArg = "'" & sPathName & "[" & sFileName & "]" & sSheetName & "'!" & Range(sRngName).Range("A1").Address
    'Execute XLM macro
    GetClosedWbRange = ExecuteExcel4Macro(sArg)
    Exit Function
errhandler:
End Function

Function TrimCO(sCOnum As String, bDateIncl As Boolean) As String
    Dim arrFormats(1) As String, i As Integer, j As Integer, sCOdate As String
    
    arrFormats(0) = "######"
    arrFormats(1) = "######"
    If bDateIncl Then arrFormats(1) = "###### (##/##/##)"
    
    TrimCO = sCOnum 'on error this default gets returned
    sCOnum = Replace(Replace(Replace(sCOnum, vbCrLf, ""), vbCr, ""), vbLf, "")
    If Not (sCOnum Like arrFormats(0) Or sCOnum Like arrFormats(1)) Then 'fix CO number (initial)
        If sCOnum Like "*" & arrFormats(0) & "*" Then 'has extra info
            i = Len(sCOnum) - 5 'go from the end in case multiple CO's are there (get latest)
            Do While Not Trim(Mid(sCOnum, i, 6)) Like arrFormats(0)
                i = i - 1
                If i = 1 Then Exit Do
            Loop
            If bDateIncl Then 'include date, if possible
                If Len(sCOnum) - i > 4 Then
                    If sCOnum Like "*##/#/####*" Then 'day needs a 0
                        sCOnum = Left(sCOnum, InStr(sCOnum, "/")) & "0" & Right(sCOnum, Len(sCOnum) - InStr(sCOnum, "/"))
                    ElseIf sCOnum Like "*#/#/####*" Then 'month & day need a 0
                        sCOnum = Left(sCOnum, InStr(sCOnum, "/")) & "0" & Right(sCOnum, Len(sCOnum) - InStr(sCOnum, "/"))
                        sCOnum = Left(sCOnum, InStr(sCOnum, "/") - 2) & "0" & Right(sCOnum, Len(sCOnum) - InStr(sCOnum, "/") + 2)
                    ElseIf sCOnum Like "*#/##/####*" And Not sCOnum Like "*##/##/####*" Then 'month needs a 0
                        sCOnum = Left(sCOnum, InStr(sCOnum, "/") - 2) & "0" & Right(sCOnum, Len(sCOnum) - InStr(sCOnum, "/") + 2)
                    End If
                    If sCOnum Like "*##/##/####*" Then
                        j = 1
                        Do While Not Trim(Mid(sCOnum, j, 10)) Like "##/##/####"
                            j = j + 1
                            If j = Len(sCOnum) - 1 Then Exit Do
                        Loop
                        If Trim(Mid(sCOnum, j, 10)) Like "##/##/####" Then
                            sCOdate = "(" & Format(Trim(Mid(sCOnum, j, 10)), "mm/dd/yy") & ")"
                            If CDate(Trim(Mid(sCOnum, j, 10))) < CDate("1/1/1910") Then sCOdate = "" 'date is nonsensical/old
                        End If
                        sCOnum = Mid(sCOnum, i, 6) & " " & sCOdate
                    Else
                        sCOnum = Mid(sCOnum, i, 6)
                    End If
                Else
                    sCOnum = "Not found"
                End If
            Else 'don't include date, even if it's there
                If Len(sCOnum) - i > 4 Then 'when this is 5 instead of 4, it's no good (why??)
                    sCOnum = Mid(sCOnum, i, 6)
                Else
                    sCOnum = "Not found"
                End If
            End If
        Else 'no good, but may have date
            If bDateIncl Then
                If sCOnum Like "*##/##/####*" Then
                    If UBound(Split(sCOnum, "/")) < 3 Then 'no errant/random slashes in the string
                        sCOnum = Right(sCOnum, Len(sCOnum) - InStr(sCOnum, "/") - 2)
                        sCOnum = Left(sCOnum, InStr(sCOnum, "/"))
                        sCOnum = "? " & sCOnum
                    Else
                        sCOnum = "Not found"
                    End If
                Else
                    sCOnum = "Not found"
                End If
            End If
        End If
    End If
    TrimCO = sCOnum
errhandler:
End Function

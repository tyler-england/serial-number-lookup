Attribute VB_Name = "WhenOpen"
Option Explicit

'''''for NumLock'''''''
Private Type OSVERSIONINFO ' Declare Type for API call:
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128   '  Maintenance string for PSS usage
End Type

' API declarations:
Private Declare PtrSafe Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
 (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, _
  ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Declare PtrSafe Function GetKeyboardState Lib "user32" _
 (pbKeyState As Byte) As Long

Private Declare PtrSafe Function SetKeyboardState Lib "user32" _
 (lppbKeyState As Byte) As Long
 
Const VK_NUMLOCK = &H90
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2
Const VER_PLATFORM_WIN32_NT = 2
Const VER_PLATFORM_WIN32_WINDOWS = 1
'''''''''''''''''''''''''''
 
 'Mouse events
Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10

'Sleep
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Active Window
Public Declare PtrSafe Function GetForegroundWindow Lib "user32" () As Long
Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, _
    ByVal lpString As String, ByVal cch As Long) As Long
    
Type POINTAPI
    X_Pos As Long
    Y_Pos As Long
End Type

Sub RecurOpenTeams()
    Dim x As Integer, y As Integer
'    x = MousePos.X_Pos
'    y = MousePos.Y_Pos
    If Hour(Now) > 15 Then
        Call MousePosClick(2000, 10)
    Else
        Call OpenTeams
    End If
    x = CInt(Rnd() * 1000)
    y = CInt(Rnd() * 500)
    SetCursorPos x, y  'set mouse position -- random
    'Call MousePosClick(x, y)
    If Not IsNumLockOn Then ToggleNumLock
    Dim dTarget As Date
    If Hour(Now) > 5 And Hour(Now) < 7 And Minute(Now) > 35 Then 'prep for work
        dTarget = Now + (1 / 24 / 60) * 15 '15 min
        Debug.Print "work prep"
    ElseIf Hour(Now) > 6 And Hour(Now) < 16 Then 'work-time
        dTarget = Now + (1 / 24 / 60) * 10 '15 min
        Debug.Print "work time"
    ElseIf (Weekday(Date) = 6 And Hour(Now) > 15) Or Weekday(Date) = 7 Or Weekday(Date) = 1 Then 'late on Friday
        dTarget = Date - Weekday(Date) + vbMonday - 7 * (vbMonday <= Weekday(Date)) + TimeValue("06:45:00") '6:45 Monday morning
        Debug.Print "late Fri"
    ElseIf Hour(Now) > 15 Then 'late on M-R
        dTarget = Date + 1 + TimeValue("06:45:00") '6:45 next morning
        Call MousePosClick(x, y)
        Debug.Print "elif"
    Else
        Debug.Print "ELSE -- " & Format(Now, "HH:MM:SS")
    End If
    Debug.Print "Teams to open at: " & Format(dTarget, "MMM DD, HH:MM:SS")
    Application.OnTime dTarget, "RecurOpenTeams"
End Sub

Sub OpenTeams()
    On Error GoTo errhandler
    Dim sPath As String, iX As Integer, iY As Integer, sActive As String
    sActive = ActiveWindowTitle
    iX = 75
    iY = 25
    iX = iX + 10 * Rnd()
    iY = iY + 10 * Rnd()
    sPath = "C:\Users\englandt\AppData\Local\Microsoft\Teams\Update.exe --processStart " & """" & "Teams.exe" & """"
    Call Shell(sPath, vbNormalFocus)
    Application.Wait 1000
    Call MousePosClick(iX, iY) '1000 px from left, 50 px from top
    SendKeys " "
    AppActivate sActive
errhandler:
End Sub

Function IsNumLockOn() As Boolean
        Dim o As OSVERSIONINFO
        Const VK_NUMLOCK = &H90
        o.dwOSVersionInfoSize = Len(o)
        GetVersionEx o
        Dim keys(0 To 255) As Byte
        GetKeyboardState keys(0)
        IsNumLockOn = keys(VK_NUMLOCK)
End Function

Sub ToggleNumLock()
        Dim o As OSVERSIONINFO
        o.dwOSVersionInfoSize = Len(o)
        GetVersionEx o
        Dim keys(0 To 255) As Byte
        GetKeyboardState keys(0)
        If o.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then  '=====Win95
              keys(VK_NUMLOCK) = Abs(Not keys(VK_NUMLOCK))
              SetKeyboardState keys(0)
        ElseIf o.dwPlatformId = VER_PLATFORM_WIN32_NT Then   '=====WinNT
        'Simulate Key Press
          keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        'Simulate Key Release
          keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY _
             Or KEYEVENTF_KEYUP, 0
        End If
End Sub

Function MousePos() As POINTAPI
    Dim Hold As POINTAPI
    GetCursorPos Hold
    MousePos = Hold
End Function

Function MousePosClick(iX As Integer, iY As Integer)
    SetCursorPos iX, iY 'set mouse position
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0 'click
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0 'let click go
End Function

Function ActiveWindowTitle() As String
    Dim sWinText As String, lHWnd As Long, L As Long
    lHWnd = GetForegroundWindow
    sWinText = String(255, vbNullChar)
    L = GetWindowText(lHWnd, sWinText, 255)
    sWinText = Left(sWinText, InStr(1, sWinText, vbNullChar) - 1)
    ActiveWindowTitle = sWinText
End Function

Attribute VB_Name = "globals"
Option Explicit
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function GetCharWidth Lib "gdi32" Alias "GetCharWidthA" (ByVal hdc As Long, ByVal wFirstChar As Long, ByVal wLastChar As Long, lpBuffer As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Public Const WM_CLOSE = &H10
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Const SC_CLOSE = &HF060&
Private Const WM_SYSCOMMAND = &H112
Public Const CLR_INVALID = &HFFFF
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Declare Function TerminateProcess& Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long)
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'Public autorefreshtime As Integer
Public strGame As String
'Public strCmdLine As String
'Public strExePath As String
'Public ctr As Integer
'Public ping1 As Currency
'Public hptimer As Boolean
'Public l1so As Boolean
'Public l2so As Boolean
'Public l2sortcolumn As Integer
'Public l1sortcolumn As Integer

'Public ralph As Class1
'Public lv1 As ListView
'Public lv2 As ListView
'Public lv3 As ListView
'Public labels(1 To 5) As Label
'Public players As Integer
'Public maxplayers As Integer
'Public commands(1 To 4) As CommandButton
'Public Timer1 As Timer
'Public Timer2 As Timer
'Public winsock As winsock
'Public ucscalewidth As Integer

'Public Sub lv1add(i1 As Variant, i2 As Variant, i3 As Variant)
'        Dim itmx As ListItem
'        Set itmx = lv1.ListItems.Add(, , i1)
'        itmx.SubItems(1) = i2
'        itmx.SubItems(2) = i3
'        Set itmx = Nothing
'End Sub
'Public Sub lv2add(i1 As Variant, i2 As Variant)
'        Dim itmx As ListItem
'        Set itmx = lv2.ListItems.Add(, , i1)
'        itmx.SubItems(1) = i2
'        Set itmx = Nothing
'End Sub

'Public Sub evil(ByRef x As ListView) ', ByRef y As ListView, ByRef z As ListView, ByRef l1 As Label, ByRef l2 As Label, ByRef l3 As Label, ByRef l4 As Label, ByRef l5 As Label, _
ByRef c1 As CommandButton, ByRef c2 As CommandButton, ByRef c3 As CommandButton, ByRef c4 As CommandButton, _
ByRef t1 As Timer, ByRef t2 As Timer, ByRef wsock As winsock, ByRef ucsw As Integer)
'    Set lv1 = x
'    Set lv2 = y
'    Set lv3 = z
'    Set labels(1) = l1
'    Set labels(2) = l2
'    Set labels(3) = l3
'    Set labels(4) = l4
'    Set labels(5) = l5
'    Set commands(1) = c1
'    Set commands(2) = c2
'    Set commands(3) = c3
'    Set commands(4) = c4
'    Set Timer1 = t1
'    Set Timer2 = t2
'    Set winsock = wsock
'    ucscalewidth = ucsw
'End Sub
'Public Sub sort()
'        lv3.ListItems.Clear
'        lv1.Sorted = False
'        Dim i As Integer
'        Dim high As Integer
'        Dim j As Integer
'        Dim z As ListItem
'        Dim x As ListItems
'        Dim r As Integer
'        Dim temp1 As Integer
'        Dim temp2 As Integer
'        r = lv1.ListItems.Count
'        l1sortcolumn = 1
'        If l1so = 1 Then
'            lv1.SortOrder = lvwDescending
'            l1so = 1
'            While j < r
'                i = 1
'                high = 1
'                While i <= lv1.ListItems.Count
'                    temp1 = Val(lv1.ListItems.Item(i).SubItems(1))
'                    temp2 = Val(lv1.ListItems.Item(high).SubItems(1))
'
'                    If temp1 < temp2 Then
'                        high = i
'                    End If
'                    i = i + 1
'                Wend
'                Set z = lv3.ListItems.Add(, , lv1.ListItems.Item(high))
'                z.SubItems(1) = lv1.ListItems.Item(high).SubItems(1)
'                z.SubItems(2) = lv1.ListItems.Item(high).SubItems(2)
'                Set z = Nothing
'                lv1.ListItems.Remove (high)
'                j = j + 1
'            Wend
'        Else
'            lv1.SortOrder = lvwAscending
'            l1so = 1
'            While j < r
'                i = 1
'                high = 1
'                While i <= lv1.ListItems.Count
'                    temp1 = Val(lv1.ListItems.Item(i).SubItems(1))
'                    temp2 = Val(lv1.ListItems.Item(high).SubItems(1))
'                    If temp1 > temp2 Then
'                        high = i
'                    End If
'                    i = i + 1
'                Wend
'                Set z = lv3.ListItems.Add(, , lv1.ListItems.Item(high))
'                z.SubItems(1) = lv1.ListItems.Item(high).SubItems(1)
'                z.SubItems(2) = lv1.ListItems.Item(high).SubItems(2)
'                Set z = Nothing
'                lv1.ListItems.Remove (high)
'                j = j + 1
'            Wend
'        End If
'        i = 1
'        While i <= lv3.ListItems.Count
'            Set z = lv1.ListItems.Add(, , lv3.ListItems.Item(i))
'            z.SubItems(1) = lv3.ListItems.Item(i).SubItems(1)
'            z.SubItems(2) = lv3.ListItems.Item(i).SubItems(2)
'            Set z = Nothing
'            i = i + 1
'        Wend
'
'End Sub


Public Sub setgame(game As String)
    strGame = game
End Sub

Public Sub CloseWindowByHwnd(ByVal hwnd&)
'Nucleus
    Dim lPid                    As Long
    Dim lHp                     As Long

    SendMessageTimeout hwnd, WM_SYSCOMMAND, SC_CLOSE, 0, 0, 500, 0
    '
    ' If the window doesn't like gentle persuasion, bring out the nipple clamps to force it to close
    '
    If IsWindow(hwnd) Then
        Call GetWindowThreadProcessId(hwnd, lPid)
        lHp = OpenProcess(PROCESS_ALL_ACCESS, 0&, lPid)
        TerminateProcess lHp&, 0&
        CloseHandle lHp
    End If
    '
End Sub


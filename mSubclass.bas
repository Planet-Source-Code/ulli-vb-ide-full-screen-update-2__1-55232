Attribute VB_Name = "mSubclass"
Option Explicit

'This tries to hook the code pane window to prevent the user from closing it

Public VBInstance           As VBIDE.VBE
Attribute VBInstance.VB_VarDescription = "VB itself"
Public NewMenuButton        As Office.CommandBarButton

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SYSCOMMAND     As Long = &H112
Attribute WM_SYSCOMMAND.VB_VarDescription = "Windows message identifier"
Private Const SC_MAXIMIZE       As Long = 61488 '&H0F030
Private Const SC_KEYMENU        As Long = 61696 '&H0F100

Private hWndCodeWindow          As Long
Private PrevProcPtr             As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const IDX_WINDOWPROC    As Long = -4
Attribute IDX_WINDOWPROC.VB_VarDescription = "Pointer into window class properties"

Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Private Const MF_BYPOSITION     As Long = &H400
Attribute MF_BYPOSITION.VB_VarDescription = "Menu constant"
Private Const MF_GRAYED         As Long = 1
Attribute MF_GRAYED.VB_VarDescription = "Menu constant"
Private Const MenuCloseItem     As Long = 6
Attribute MenuCloseItem.VB_VarDescription = "Menu item position"

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCPAINT        As Long = &H85
Attribute WM_NCPAINT.VB_VarDescription = "Windows message identifier"

Private Function CodeWindowProc(ByVal hWnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    CodeWindowProc = 0
    'forward all messages to original destination
    'but consume the wm_syscommand message unless it's SC_MAXIMIZE
    If nMsg <> WM_SYSCOMMAND Or wParam = SC_MAXIMIZE Or wParam = SC_KEYMENU Then
        CodeWindowProc = CallWindowProc(PrevProcPtr, hWnd, nMsg, wParam, lParam)
    End If

End Function

Public Sub HookCodeWindow()

    With VBInstance
        On Error Resume Next
            If .ActiveWindow Is .ActiveCodePane.Window Then 'we have a code pane window
                With .MainWindow
                    'subclass the active codepane
                    hWndCodeWindow = FindWindowEx(FindWindowEx(.hWnd, 0&, "MDIClient", vbNullString), 0&, "VbaWindow", VBInstance.ActiveWindow.Caption)
                    If hWndCodeWindow Then 'we have an open code pane
                        PrevProcPtr = SetWindowLong(hWndCodeWindow, IDX_WINDOWPROC, AddressOf CodeWindowProc)
                        EnableMenuItem GetSystemMenu(.hWnd, False), MenuCloseItem, MF_BYPOSITION Or MF_GRAYED 'disable (gray) main window sysmenu close
                        SendMessage .hWnd, WM_NCPAINT, 1&, 0& 'repaint the frame and sysmenu
                    End If
                End With '.MAINWINDOW
            End If
        On Error GoTo 0
    End With 'VBINSTANCE

End Sub

Public Sub UnHookCodeWindow()

    If hWndCodeWindow Then
        SetWindowLong hWndCodeWindow, IDX_WINDOWPROC, PrevProcPtr
        With VBInstance.MainWindow
            EnableMenuItem GetSystemMenu(.hWnd, False), MenuCloseItem, MF_BYPOSITION 'enable main window sysmenu close
            SendMessage .hWnd, WM_NCPAINT, 1&, 0& 'repaint the frame and sysmenu
        End With 'VBINSTANCE.MAINWINDOW
    End If

End Sub

':) Ulli's VB Code Formatter V2.17.4 (2004-Aug-02 11:20) 29 + 45 = 74 Lines

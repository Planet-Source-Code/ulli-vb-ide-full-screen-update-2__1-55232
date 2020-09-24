VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} dZoomer 
   ClientHeight    =   11385
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   13080
   _ExtentX        =   23072
   _ExtentY        =   20082
   _Version        =   393216
   Description     =   "This Add-In  toggles your code pane window between full screen and normal screen display."
   DisplayName     =   "Ulli's Full Screen Display Add-In"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   SatName         =   "WPsWSaddin.dll"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "dZoomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'© 2003     UMGEDV GmbH  (umgedv@aol.com)
'
'Author     UMG (Ulli K. Muehlenweg)
'
'Title      VB6 IDE Full Screen Display
'
'           Toggles the VB IDE display between full screen and normal display.
'           Simply compile the .DLL into your VB folder and then use the
'           AddIns Manager to load this AddIn into VB. You will then find a new
'           menu item in the main window.
'
'How to     Press Alt+S (or your own accelerator key if you modified that) or click
'           on Main Menu Item 'Full Screen / Returm to Normal Screen'.
'
'           The AddIn works in MDI mode only.
'
'Credits    This code is based on a submission to PSC by WPsoftware®.
'
'**********************************************************************************
'Development History
'**********************************************************************************
'02Aug2004 Version 2.2.8     UMG
'
'Added Exception handling
'(see mExeption and cExeption)
'
'Added Simple Mixer Class (the crash sound is just too good to go by unnoticed,
'so now we switch sound and speakers on)
'
'Mofified loading and unloading to get the menu button installed and removed properly.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'29Jul2004 Version 2.1.4     UMG
'
'Added safeguards against closing codepanes or VB while in full screen mode
'(see mSubclass)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'29Apr2003 Version 1.0.0     UMG
'
'Prototype
'
'**********************************************************************************

Private WithEvents UllisMenuButton   As CommandBarEvents
Attribute UllisMenuButton.VB_VarHelpID = -1
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Const VK_LSHIFT     As Long = &HA0

Private Const SleepTime     As Long = 555
Attribute SleepTime.VB_VarDescription = "splah popup time"

Private Type WinProp
    Width       As Long
    Height      As Long
    State       As Long
    Visible     As Boolean
End Type
Private WinProps()          As WinProp
Attribute WinProps.VB_VarDescription = "properties of hidden windows"

Private Type ComProp
    Visible     As Boolean
End Type
Private ComProps()          As ComProp
Attribute ComProps.VB_VarDescription = "Command properties"
Private Const MainMenuBar   As Long = 1

Private WindowState(1 To 2) As Long 'windowstate of main window and of active code pane
Attribute WindowState.VB_VarDescription = "windowstate of active code pane"
Private Expand              As Boolean 'toggle --> when true then next click switches to full sreen
Attribute Expand.VB_VarDescription = "Toggles to indicate what the next menu click will do"
Private OptAll              As Boolean
Private i                   As Long
Attribute i.VB_VarDescription = "general use"

Private ExeptionHandler     As cException

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Const CaptionFull           As String = "Full &Screen" 'you may want to localize this text but be sure to use an accelerator key
Attribute CaptionFull.VB_VarDescription = "String constant"
Private Const CaptionNormalShort    As String = "Normal &Screen"
Private Const CaptionNormalLong     As String = "Return to Normal &Screen Display Mode" 'which does not clash with any other main menu item
Attribute CaptionNormalLong.VB_VarDescription = "String constant"
Private Const ToolTipTextNormal     As String = "Switch to Full Screen Display Mode "
Attribute ToolTipTextNormal.VB_VarDescription = "String constant"
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Function AcceleratorKeyIn(InText As String) As String

    i = InStr(InText, "&")
    If i Then
        AcceleratorKeyIn = "(Alt+" & Mid$(InText, i + 1, 1) & ")"
    End If

End Function

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)

  'Called by system when AddIn is connected/loaded

    Set VBInstance = Application
    If ConnectMode = ext_cm_External Then
        MsgBox "This Add-In must be activated from the Add-Ins menu."
      Else 'NOT CONNECTMODE...
        fSplash.Show
        With VBInstance
            If .DisplayModel = vbext_dm_SDI Then 'not in MDI mode
                fSplash.lblAbout = "Disabling Full Screen Add-In..."
              Else 'NOT VBINSTANCE.DISPLAYMODEL... 'NOT .DISPLAYMODEL...
                Set ExeptionHandler = New cException
                ExeptionHandler.Noisy = True
                'create menu button
                Set NewMenuButton = .CommandBars(1).Controls.Add(msoControlButton)
                With NewMenuButton
                    .Caption = CaptionFull
                    .ToolTipText = ToolTipTextNormal & AcceleratorKeyIn(CaptionFull)
                    .Style = msoButtonIconAndCaption 'msoButtonIcon 'msoButtonCaption
                    .BeginGroup = True
                    .State = msoButtonUp
                End With 'NewMenuButton
                SetMenuIcon NewMenuButton, fSplash.picMenuUp 'give it an icon
                Set UllisMenuButton = .Events.CommandBarEvents(NewMenuButton) 'hook events for this menu button
                Expand = True 'preset for first click on menu button
            End If
        End With 'VBINSTANCE
        DoEvents
        Sleep SleepTime
        fSplash.Hide
    End If

End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)

  'Called by system when AddIn is disconnected/unloaded manually from AddIns manager

    Set ExeptionHandler = Nothing
    If Not NewMenuButton Is Nothing Then 'in case we have no new menu item
        Set UllisMenuButton = Nothing
        NewMenuButton.Delete
    End If

End Sub

Private Sub SetMenuIcon(MenuButton As Office.CommandBarButton, Pic As PictureBox)

  Dim TmpStr As String

    With Clipboard
        TmpStr = .GetText
        .SetData Pic.Image
        MenuButton.PasteFace
        .Clear
        .SetText TmpStr
    End With 'CLIPBOARD

End Sub

Private Sub Toggle()

  'Toggles display mode

  Dim ActiveCaption     As String

    With VBInstance
        If .DisplayModel = vbext_dm_MDI And .CodePanes.Count Then
            If Expand Then 'to full screen
                mException.NeedsFallback = Expand
                OptAll = (GetAsyncKeyState(VK_LSHIFT) >= 0)
                WindowState(1) = .MainWindow.WindowState
                .MainWindow.WindowState = vbext_ws_Maximize
                ReDim WinProps(0 To .Windows.Count)
                ActiveCaption = .ActiveCodePane.Window.Caption
                'hide all windows except main and code window
                For i = 1 To .Windows.Count
                    With .Windows(i)
                        WinProps(i).Width = .Width
                        WinProps(i).Height = .Height
                        WinProps(i).Visible = .Visible
                        WinProps(i).State = .WindowState
                        If .Type <> vbext_wt_MainWindow And .Caption <> ActiveCaption Then
                            .Visible = False
                            DoEvents
                        End If
                    End With '.WINDOWS(I)
                Next i
                If OptAll Then
                    'hide all commandbars
                    ReDim ComProps(1 To .CommandBars.Count)
                    For i = 1 To .CommandBars.Count
                        ComProps(i).Visible = .CommandBars(i).Visible
                        On Error Resume Next
                            .CommandBars(i).Visible = False
                        On Error GoTo 0
                        DoEvents
                    Next i
                    'hide all main menu items
                    With .CommandBars(MainMenuBar)
                        For i = 1 To .Controls.Count
                            .Controls(i).Visible = False
                            DoEvents
                        Next i
                    End With '.COMMANDBARS(MAINMENUBAR)
                End If
                'maximize active code pane
                With .ActiveCodePane.Window
                    WindowState(2) = .WindowState
                    .WindowState = vbext_ws_Maximize
                    DoEvents
                End With '.ACTIVECODEPANE.WINDOW
                'set menu caption and state
                With NewMenuButton
                    If OptAll Then
                        .Caption = CaptionNormalLong
                        .ToolTipText = ""
                      Else 'OPTALL = FALSE/0
                        .Caption = CaptionNormalShort
                        .ToolTipText = CaptionNormalLong
                    End If
                    .State = msoButtonUp
                    .Visible = True
                End With 'NewMenuButton
                SetMenuIcon NewMenuButton, fSplash.picMenuDown
                NewMenuButton.Style = msoButtonIconAndCaption 'msoButtonIcon 'msoButtonCaption
                HookCodeWindow
                .ActiveCodePane.Window.SetFocus
              Else 'to normal screen 'Expand = FALSE/0
                'restore all main menu items
                If OptAll Then
                    With .CommandBars(MainMenuBar)
                        For i = 1 To .Controls.Count
                            .Controls(i).Visible = True
                        Next i
                    End With '.COMMANDBARS(MAINMENUBAR)
                End If
                'set menu caption and state
                With NewMenuButton
                    .Caption = CaptionFull
                    .ToolTipText = ToolTipTextNormal & AcceleratorKeyIn(CaptionFull)
                End With 'NewMenuButton
                SetMenuIcon NewMenuButton, fSplash.picMenuUp
                'restore comandbars in reverse order
                For i = .CommandBars.Count To 1 Step -1
                    On Error Resume Next 'just leave it if it don't work
                        .CommandBars(i).Visible = ComProps(i).Visible
                    On Error GoTo 0
                    DoEvents
                Next i
                .ActiveCodePane.Window.WindowState = vbext_ws_Normal 'to be able to restore others
                'restore hidden windows in reverse order
                On Error Resume Next
                    For i = .Windows.Count To 1 Step -1
                        With .Windows(i)
                            .Visible = WinProps(i).Visible
                            If WinProps(i).State <> vbext_ws_Minimize Then
                                .Width = WinProps(i).Width
                                .Height = WinProps(i).Height
                            End If
                        End With '.WINDOWS(I)
                    Next i
                On Error GoTo 0
                'restore windowstate of active code pane
                With .ActiveCodePane.Window
                    .WindowState = WindowState(2)
                    If WindowState(2) <> vbext_ws_Minimize Then
                        'reset focus to active codepane
                        .SetFocus
                    End If
                End With '.ACTIVECODEPANE.WINDOW
                With .MainWindow
                    If .WindowState = vbext_ws_Maximize Then 'still maximized
                        .WindowState = WindowState(1)
                    End If
                End With '.MAINWINDOW
                DoEvents
                Erase ComProps, WinProps
                UnHookCodeWindow
                'and finally menu button up
                With NewMenuButton
                    .Style = msoButtonIconAndCaption ' msoButtonIcon 'msoButtonCaption
                    Sleep 40
                    .State = msoButtonUp
                    DoEvents
                End With 'NEWMENUBUTTON
            End If
            Expand = Not Expand
        End If
    End With 'VBINSTANCE

End Sub

Private Sub UllisMenuButton_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

  'Called by VB on menu click

    With VBInstance
        Select Case 0
          Case .VBProjects.Count
            MsgBox "Cannot see any project. You must open a project first.", vbCritical
          Case .CodePanes.Count
            MsgBox "Cannot see any code pane. You must open a code pane first.", vbCritical
          Case Else
            Toggle
        End Select
    End With 'VBINSTANCE

End Sub

':) Ulli's VB Code Formatter V2.17.4 (2004-Aug-02 11:20) 80 + 223 = 303 Lines

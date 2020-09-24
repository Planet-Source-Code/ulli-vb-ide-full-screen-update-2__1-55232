Attribute VB_Name = "mException"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const EXCEPTION_MAXIMUM_PARAMETERS  As Long = 15
Private Const EXCEPTION_EXECUTE_HANDLER     As Long = 1
Private Const EXCEPTION_CONTINUE_EXECUTION  As Long = -1
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_NOSIZE                     As Long = &H1
Public Const SWP_NOMOVE                     As Long = &H2
Public Const SWP_NOACTIVATE                 As Long = &H10
Public Const SWP_FLAGS                      As Long = SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
Public Const TOPMOST                        As Long = -1

Private Type EXCEPTION_POINTERS
    pExceptionRecord    As Long 'pointer to an EXCEPTION_RECORD structure
    pContextRecord      As Long 'pointer to a CONTEXT structure
End Type

'not used, just for documentation
'Private Type EXCEPTION_RECORD
'    ExceptionCode       As Long ' + 0
'    ExceptionFlags      As Long ' + 4
'    pExceptionRecord    As Long ' + 8    'pointer to a nested EXCEPTION_RECORD structure
'    ExceptionAddress    As Long ' + 12
'    NumberParameters    As Long ' + 16
'    ExceptionInformation(0 To EXCEPTION_MAXIMUM_PARAMETERS) As Long
'End Type

Private myNoisy         As Boolean
Private IgnFurtheErrs   As Boolean
Public NeedsFallback   As Boolean

Private Function HexFormat(ErrNum As Long) As String

    HexFormat = Format$(Right$("00000000" & Hex$(ErrNum), 8), "<@@\-@@\-@@\-@@") & " (" & Format$(ErrNum, "#,0") & ")"

End Function

Public Function Interceptor(lpException As EXCEPTION_POINTERS) As Long

  Dim i As Long, j As Long

    If IgnFurtheErrs Then
        Interceptor = EXCEPTION_CONTINUE_EXECUTION
      Else 'IgnFurtheErrs = FALSE/0
        With lpException
            CopyMemory i, ByVal .pExceptionRecord + 0, 4 'exception code
            CopyMemory j, ByVal .pExceptionRecord + 12, 4 'exception address
        End With 'LPEXCEPTION
        With fException
            .txRef = HexFormat(i) & " / " & HexFormat(j)
            Do While .Visible
                SetWindowPos .hWnd, TOPMOST, 0&, 0&, 0&, 0&, SWP_FLAGS 'make sure the window stays on top
                DoEvents
            Loop
            i = Val(.Tag)
            If i And 1 Then 'continue
                Interceptor = EXCEPTION_CONTINUE_EXECUTION
              Else 'NOT I...
                TidyUp
                Interceptor = EXCEPTION_EXECUTE_HANDLER
            End If
            IgnFurtheErrs = CBool(i And 4)
        End With 'FEXCEPTION
        Unload fException
        Set fException = Nothing
    End If

End Function

Public Function Noisy(Optional nuNoisy As Variant) As Boolean

    Noisy = myNoisy
    If Not IsMissing(nuNoisy) Then
        myNoisy = CBool(nuNoisy)
    End If

End Function

Private Sub TidyUp()

  '*******************************************
  'insert code to tidy up before abortion here
  '*******************************************

    If NeedsFallback Then  'it is in maximized state
        Sleep 500
        NewMenuButton.Execute 'so return to normal state
        DoEvents
        Sleep 1000
    End If

End Sub

':) Ulli's VB Code Formatter V2.17.4 (2004-Aug-02 11:20) 32 + 64 = 96 Lines

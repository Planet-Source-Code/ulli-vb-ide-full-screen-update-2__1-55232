VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long
Private PrevProcPtr As Long

Private Sub Class_Initialize()

  'set interceptor active

    PrevProcPtr = SetUnhandledExceptionFilter(Far(AddressOf mException.Interceptor)) 'hihi - cheating VB

End Sub

Private Sub Class_Terminate()

  'set interceptor back to what it was

    SetUnhandledExceptionFilter PrevProcPtr

End Sub

Private Function Far(ByVal Pointer As Long) As Long

  'little helper to be able to use 'AddressOf' in a Class

    Far = Pointer

End Function

Public Property Let Noisy(nuNoisy As Boolean)

    mException.Noisy nuNoisy

End Property

Public Property Get Noisy() As Boolean

    Noisy = mException.Noisy

End Property

':) Ulli's VB Code Formatter V2.17.4 (2004-Aug-02 11:20) 4 + 38 = 42 Lines

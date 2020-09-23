VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMousePointer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' the CMouseCursor class
Private oldMousePointer As Variant

' enforce a new cursor
Sub SetCursor(Optional NewCursor As MousePointerConstants = vbHourglass)
    If IsEmpty(oldMousePointer) Then
        ' save the original cursor, but only the first
        ' time this method is invoked in this instance
        oldMousePointer = Screen.MousePointer
    End If
    Screen.MousePointer = NewCursor
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    ' restore original mouse cursor, if it changed
    If Not IsEmpty(oldMousePointer) Then
        Screen.MousePointer = oldMousePointer
    End If
End Sub


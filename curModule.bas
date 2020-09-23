Attribute VB_Name = "curModule"
Option Explicit






Private Type POINTAPI
        x As Long
        y As Long
End Type


Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long



Public Function CursorPos(x As Long, y As Long) As Boolean

    
    Dim pt As POINTAPI
    
    If GetCursorPos(pt) Then
        x = pt.x
        y = pt.y
        CursorPos = True
    End If

End Function


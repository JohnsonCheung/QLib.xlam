Attribute VB_Name = "MVb_Fs_Pth_Sfx"
Option Explicit
Public Const PthSep$ = "\"

Function HasPthSfx(A) As Boolean
HasPthSfx = LasChr(A) = PthSep
End Function
Function PthEnsSfx$(A)
If HasPthSfx(A) Then
    PthEnsSfx = A
Else
    PthEnsSfx = A & PthSep
End If
End Function

Function PthRmvSfx$(Pth)
PthRmvSfx = RmvSfx(Pth, PthSep)
End Function



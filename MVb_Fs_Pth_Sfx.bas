Attribute VB_Name = "MVb_Fs_Pth_Sfx"
Option Explicit
Public Const PthSep$ = "\"

Function HasPthSfx(Pth$) As Boolean
HasPthSfx = LasChr(Pth) = PthSep
End Function
Function EnsPthSfx$(Pth$)
If HasPthSfx(Pth) Then
    EnsPthSfx = Pth
Else
    EnsPthSfx = Pth & PthSep
End If
End Function

Function RmvPthSfx$(Pth$)
RmvPthSfx = RmvSfx(Pth, PthSep)
End Function



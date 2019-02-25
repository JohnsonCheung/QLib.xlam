Attribute VB_Name = "MVb_Seed"
Option Explicit

Function SeedExpand$(VblQQStr, Ny0)
'Seed is a VblQQ-String
Dim A$, J%, O$()
Dim Ny$()
Ny = CvNy(Ny0)
For J = 0 To UB(Ny)
    Push O, Replace(VblQQStr, "?", Ny(J))
Next
'SeedExpand = LineszVbl(Lines(O))
End Function

Private Sub Z_SeedExpand()
Dim VblQQStr, Ny0
'
VblQQStr = "Sub Tst?()|Dim A As New ?: A.Tst|End Sub"
Ny0 = "Xws Xwb Xfx Xrg"
GoSub Tst
Exit Sub
Tst:
    Act = SeedExpand(VblQQStr, Ny0)
    C
    Debug.Print Act
    Stop
    Return
End Sub

Private Sub Z()
Z_SeedExpand
MVb__Seed:
End Sub

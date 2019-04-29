Attribute VB_Name = "MVb_Seed"
Option Explicit

Function Expand$(Seed$, Ny0)
'Seed is a VblQQ-String
Dim A$, J%, O$()
Dim Ny$()
Ny = TermAy(Ny0)
A = RplVbl(Seed)
For J = 0 To UB(Ny)
    Push O, Replace(A, "?", Ny(J))
Next
Expand = JnCrLf(O)
End Function

Private Sub Z_Expand()
Dim VblQQStr$, Ny0
'
VblQQStr = "Sub Tst?()|Dim A As New ?: A.Tst|End Sub"
Ny0 = "Xws Xwb Xfx Xrg"
GoSub Tst
Exit Sub
Tst:
    Act = Expand(VblQQStr, Ny0)
    Debug.Print Act
    Stop
    C
    Debug.Print Act
    Stop
    Return
End Sub

Private Sub Z()
Z_Expand
MVb__Seed:
End Sub

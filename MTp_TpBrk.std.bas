Attribute VB_Name = "MTp_TpBrk"
Option Explicit
Public Type TpSec
    Nm As String
    GpAy() As Gp
End Type
Public Type TpBrk
    Er() As String
    RmkDic As New Dictionary
    SecAy() As TpSec
End Type

Function TpBrk(Tp$) As TpBrk
', OErLy$(), ORmkDic As Dictionary, Ny0, ParamArray OLyAp())
Dim O(), J%, U%
'O = ClnBrk1(ClnLy(SplitCrLf(A)), Ny0)
U = UB(O)
For J = 0 To U - 2
    'OLyAp(J) = O(J)
Next
'OErLy = O(U + 1)
'Set ORmkDic = O(U + 2)
End Function

Function LnxAy(Ly$()) As Lnx()
Dim J&, O() As Lnx
If Sz(Ly) = 0 Then Exit Function
For J = 0 To UB(Ly)
    PushObj O, Lnx(J, Ly(J))
Next
LnxAy = O
End Function
Function HasMajPfx(Ly$(), MajPfx$) As Boolean
Dim Cnt%, J%
For J = 0 To UB(Ly)
    If HasPfx(Ly(J), MajPfx) Then Cnt = Cnt + 1
Next
HasMajPfx = Cnt > (Sz(Ly) \ 2)
End Function


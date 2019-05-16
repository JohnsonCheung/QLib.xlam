Attribute VB_Name = "QTp_TpBrk"
Option Explicit
Private Const CMod$ = "MTp_TpBrk."
Private Const Asm$ = "QTp"
Type TpSec
    Nm As String
    Blk As Blk
End Type
Type TpBrk
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


Function HasMajPfx(Ly$(), MajPfx$) As Boolean
Dim Cnt%, J%
For J = 0 To UB(Ly)
    If HasPfx(Ly(J), MajPfx) Then Cnt = Cnt + 1
Next
HasMajPfx = Cnt > (Si(Ly) \ 2)
End Function


Attribute VB_Name = "QTp_TpBrk"
Option Explicit
Private Const CMod$ = "MTp_TpBrk."
Private Const Asm$ = "QTp"
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

Function LnxAyzT1(Ly$(), T1$) As Lnx()
Dim J&, O() As Lnx
For J = 0 To UB(Ly)
    If T1zS(Ly(J)) = T1 Then
        PushObj O, Lnx(J, Ly(J))
    End If
Next
LnxAyzT1 = O
End Function

Function LnxAyDic(Ly$()) As Dictionary
Set LnxAyDic = LnxAyDiczT1nn(Ly, AywDist(T1Sy(Ly)))
End Function

Function LnxAyDiczT1nn(Ly$(), T1nn$) As Dictionary
Dim T$, I
Set LnxAyDiczT1nn = New Dictionary
For Each I In TermSy(T1nn)
    T = I
    LnxAyDiczT1nn.Add T, LnxAyzT1(Ly, T)
Next
End Function

Function LnxAy(Ly$()) As Lnx()
Dim J&, O() As Lnx
If Si(Ly) = 0 Then Exit Function
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
HasMajPfx = Cnt > (Si(Ly) \ 2)
End Function


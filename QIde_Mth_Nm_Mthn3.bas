Attribute VB_Name = "QIde_Mth_Nm_Mthn3"
Option Explicit
Private Const CMod$ = "Mthn3."
Type Mthn3: Nm As String: ShtTy As String: ShtMdy As String: End Type

Function Mthn3(Nm, ShtMdy, ShtTy) As Mthn3
With Mthn3
    .Nm = Nm
    .ShtMdy = ShtMdy
    .ShtTy = ShtTy
End With
End Function

Function MthDnzN3(A As Mthn3)
With A
If .Nm = "" Then Exit Function
MthDnzN3 = JnDotAp(.ShtMdy, .Nm, .ShtTy)
End With
End Function

Function Mthn3zL(Lin) As Mthn3
Mthn3zL = ShfMthn3(CStr(Lin))
End Function
Function ShfMthn3(OLin$) As Mthn3
Dim M$: M = ShfShtMdy(OLin)
Dim T$: T = ShfShtMthTy(OLin):: If T = "" Then Exit Function
ShfMthn3 = Mthn3(ShfNm(OLin), M, T)
End Function


Function HitMthn3(A As Mthn3, B As WhMth) As Boolean
Select Case True
Case A.Nm = "":
Case True 'IsEmpWhMth(B)
    HitMthn3 = True
Case True ' _
    Not HitNm(A.Nm, B.WhNm), _
    Not HitShtMdy(A.ShtMdy, B.ShtMthMdyAy), _
    Not HitAy(A.ShtTy, B.ShtTyAy)
Case Else
    HitMthn3 = True
End Select
End Function



Function MthDnzMthn3$(A As Mthn3)
If A.Nm = "" Then Exit Function
MthDnzMthn3 = A.Nm & "." & A.ShtTy & "." & A.ShtMdy
End Function

Function RmvMthn3$(Lin)
Dim L$: L = Lin
RmvMthMdy L
If ShfMthTy(L) = "" Then Exit Function
If ShfNm(L) = "" Then Thw CSub, "Not as SrcLin", "Lin", Lin
RmvMthn3 = L
End Function
Sub DmpMthn3(A As Mthn3)
D FmtMthn3(A)
End Sub
Function FmtMthn3$(A As Mthn3)

End Function

Function Mthn3zDNm(MthDn$) As Mthn3
Dim Nm$, Ty$, Mdy$
If MthDn = "*Dcl" Then
    Nm = "*Dcl"
Else
    Dim B$(): B = SplitDot(MthDn)
    If Si(B) <> 3 Then
        Thw CSub, "Given MthDn SplitDot should be 3 elements", "NEle-SplitDot MthDn", Si(B), MthDn
    End If
    Dim ShtMdy$, ShtTy$
    AsgAp B, Nm, ShtTy, ShtMdy
    Ty = MthTyBySht(ShtTy)
    Mdy = ShtMthMdy(ShtMdy)
End If
Mthn3zDNm = Mthn3(Nm, Mdy, Ty)
End Function



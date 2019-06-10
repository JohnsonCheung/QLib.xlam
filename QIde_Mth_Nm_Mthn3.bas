Attribute VB_Name = "QIde_Mth_Nm_Mthn3"
Option Compare Text
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

Function Mthn3zL(Lin) As Mthn3
Mthn3zL = ShfMthn3(CStr(Lin))
End Function
Function ShfLHS$(OLin$)
Dim L$: L = OLin
Dim IsSet As Boolean: IsSet = ShfTermX(L, "Set")
Dim S$: If IsSet Then S = "Set "
Dim LHS$: LHS = ShfNm(L)
If ShfPfx(L, " = ") Then
    ShfLHS = S & LHS & " = "
    OLin = L
End If
End Function
Function ShfLRHS(OLin$) As Variant()
'If HasSubStr(OLin, "Then RplLin Md, DeLNewO") Then Stop
Dim L$: L = OLin
Dim LHS$: LHS = ShfLHS(L)
With Brk1(L, "'")
    Dim RHS$: RHS = .S1
    OLin = "'" & .S2
End With
ShfLRHS = Array(LHS, RHS)
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

Function RmvMthn3$(Lin)
Dim L$: L = Lin
RmvMthMdy L
If ShfMthTy(L) = "" Then Exit Function
If ShfNm(L) = "" Then Thw CSub, "Not as SrcLin", "Lin", Lin
RmvMthn3 = L
End Function
Function FmtMthn3$(A As Mthn3)
With A
FmtMthn3 = JnDotAp(.Nm, .ShtMdy, .ShtTy)
End With
End Function
Sub DmpMthn3(A As Mthn3)
D FmtMthn3(A)
End Sub



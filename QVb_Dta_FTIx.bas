Attribute VB_Name = "QVb_Dta_FTIx"
Type FCnt
    FmLno As Long
    Cnt As Long
End Type
Type FEIx
FmIx As Long
EIx As Long ' To loop For J=FmIx To EIx-1, Always ToIx <-- EIx-1
End Type
Type FEIxs: N As Long: Ay() As FEIx: End Type
Function FCnt(FmLno, Cnt) As FCnt
If FmLno <= 0 Then Exit Function
If Cnt <= 0 Then Exit Function
FCnt.FmLno = FmLno
FCnt.Cnt = Cnt
End Function
Function FCntzFEIx(A As FEIx) As FCnt
With A
    FCntzFEIx = FCnt(.FmIx + 1, LinCntzFEIx(A))
End With
End Function
Function FEIx(FmIx, EIx) As FEIx
If 0 > FmIx Then Exit Function
If 0 > EIx Then Exit Function
If FmIx > EIx Then Exit Function
FEIx.FmIx = FmIx
FEIx.EIx = EIx
End Function
Sub PushFEIx(O As FEIxs, M As FEIx)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub
Sub PushFEIxs(O As FEIxs, M As FEIxs)
Dim J&
For J = 0 To M.N - 1
    PushFEIx O, M.Ay(J)
Next
End Sub
Function AddFEIxs(A As FEIxs, B As FEIxs) As FEIxs
AddFEIxs = A
PushFEIxs A, B
End Function
Function SngFEIx(A As FEIx) As FEIxs
PushFEIx SngFEIx, A
End Function

Function BetFEIx(Ix, A As FEIx) As Boolean
If Ix < 0 Then Thw CSub, "Ix cannot be -ve", "Ix", Ix
If A.FmIx > U Then Exit Function
If A.EIx < U Then Exit Function
BetFEIx = True
End Function

Function CntzFEIx&(A As FEIx)
Dim O&
O = A.EIx - A.FmIx
If O < 0 Then Stop
CntzFEIx = O
End Function
Function FEIxzFC(FmIx, Cnt) As FEIx
FEIxzFC = FEIx(FmIx, FmIx + Cnt - 1)
End Function

Function IsEqFEIxs(A As FEIxs, B As FEIxs) As Boolean
If A.N <> B.N Then Exit Function
Dim J&
For J = 0 To A.N - 1
    If Not IsEqFEIx(A.Ay(J), B.Ay(J)) Then Exit Function
    J = J + 1
Next
IsEqFEIxs = True
End Function
Function IsFEIxzEmp(A As FEIx) As Boolean

End Function
Function IsFEIxsInOrd(A As FEIxs) As Boolean
Dim J%
For J = 0 To A.N - 1
    With FCntzFEIx(A.Ay(J))
        If .FmLno = 0 Then Exit Function
        If .Cnt = 0 Then Exit Function
        If .FmLno + .Cnt > FCntzFEIx(A.Ay(J + 1)).FmLno Then Exit Function
    End With
Next
IsFEIxsInOrd = True
End Function

Function Positive(N)
If N > 0 Then Positive = N
End Function
Function LinCntzFEIx&(A As FEIx)
LinCntzFEIx = Positive(A.EIx - A.FmIx)
End Function
Function LinCntzFEIxs&(A As FEIxs)
Dim J&, O&
For J = 0 To A.N - 1
    O = O + LinCntzFEIx(A.Ay(J))
Next
LinCntzFEIxs = O
End Function

Function IsEqFEIx(A As FEIx, B As FEIx) As Boolean
With A
    If .FmIx <> B.FmIx Then Exit Function
    If .EIx <> B.EIx Then Exit Function
End With
IsEqFEIx = True
End Function

Private Sub ZZ()
Dim A As Variant
Dim C As FEIx
IsEqFEIx C, C
End Sub


Attribute VB_Name = "QVb_FmCnt"
Option Explicit
Private Const CMod$ = "MVb_FmCnt."
Private Const Asm$ = "QVb"

Function IsEqFTIxAy(A() As FTIx, B() As FTIx) As Boolean
If Si(A) <> Si(B) Then Exit Function
Dim X, J&
For Each X In Itr(A)
    If Not FTIxIsEq(CvFTIx(X), B(J)) Then Exit Function
    J = J + 1
Next
IsEqFTIxAy = True
End Function

Function FTIxAyIsInOrd(A() As FTIx) As Boolean
Dim J%
For J = 0 To UB(A) - 1
    With A(J)
        If .FmNo = 0 Then Exit Function
        If .Cnt = 0 Then Exit Function
        If .FmNo + .Cnt > A(J + 1).FmNo Then Exit Function
    End With
Next
FTIxAyIsInOrd = True
End Function

Function FTIxAyLinCnt%(A() As FTIx)
Dim I, C%, O%
For Each I In A
    C = CvFTIx(I).Cnt
    If C > 0 Then O = O + C
Next
FTIxAyLinCnt = O
End Function

Function FTIxIsEq(A As FTIx, B As FTIx) As Boolean
With A
    If .FmNo <> B.FmNo Then Exit Function
    If .Cnt <> B.Cnt Then Exit Function
End With
FTIxIsEq = True
End Function

Private Sub ZZ()
Dim A As Variant
Dim B() As FTIx
Dim C As FTIx
CvFTIx A
FTIx A, A
IsEqFTIxAy B, B
FTIxAyIsInOrd B
FTIxAyLinCnt B
LyzFTIxAy B
FTIxIsEq C, C
End Sub

Private Sub Z()
End Sub

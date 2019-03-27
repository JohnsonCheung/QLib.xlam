Attribute VB_Name = "MDta_Srt"
Option Explicit
Dim A_SrtColIxAy%()
Dim A_IsDesAy() As Boolean
Function DrsSrt(A As Drs, Optional SrtByFF = "", Optional IsDes As Boolean) As Drs
Dim Fny$(): If SrtByFF = "" Then Fny = Sy(A.Fny()(0)) Else Fny = NyzNN(SrtByFF)
Set DrsSrt = Drs(A.Fny, DrySrt(A.Dry, IxAy(A.Fny, Fny), IsDes))
End Function

Function DrySrt(Dry(), ColIxAy%(), IsDesAy() As Boolean) As Variant()
If Si(IxAy) <> Si(IsDesAy) Then Thw CSub, "Si of ColIxAy and IsDesAy are dif", "Si-ColIxAy Si-IsDesAy", Si(IxAy) <> Si(IsDesAy)
If Si(IxAy) = 1 Then
    DrySrt = DrySrtzCol(Dry, IxAy(0), IsDesAy(0))
Else
    DrySrt = DrySrtzColIxAy(Dry, IxAy, IsDesAy)
End If
End Function

Function DtSrt(A As Dt, Optional SrtByFF = "", Optional IsDes As Boolean) As Dt
Set DtSrt = DtzDrs(DrsSrt(DrszDt(A), SrtByFF, IsDes), A.DtNm)
End Function

Function DrySrtzCol(Dry(), ColIx%, Optional IsDes As Boolean) As Variant()
Attribute DrySrtzCol.VB_Description = "12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789"
Dim Col: Col = ColzDry(Dry, ColIx)
Dim Ix&(): Ix = IxAyzAySrt(Col, IsDes)
Dim J%
For J = 0 To UB(Ix)
   Push DrySrtzCol, Dry(Ix(J))
Next
End Function

Private Function DrySrtzColIxAy(Dry(), SrtColIxAy%(), IsDesAy() As Boolean) As Variant()
Dim O(): O = Dry
A_SrtColIxAy = SrtColIxAy
A_IsDesAy = IsDesAy
DrySrtLH O, 0, UB(Dry)
DrySrtzColIxAy = O
End Function

Private Function IsGT(Dr1, Dr2) As Boolean
Dim J%
IsGT = True
For J = 0 To UB(A_ColIxAy)
    Ix = A_ColIxAy(J)
    If A_IsDesAy(J) Then
        If Dr1(Ix) < Dr2(Ix) Then Exit Function
    Else
        If Dr1(Ix) > Dr2(Ix) Then Exit Function
    End If
Next
IsGT = False
End Function

Private Function Partition&(ODry, L&, H&)
Dim Dr, I&, J&
Dr = ODry(L)
I = L
J = H
Dim Z&
Do
    Z = Z + 1
    If Z > 10000 Then Stop
    While IsGT(Dr, ODry(I))
        I = I + 1
    Wend
    
    While IsGT(ODry(J), Dr)
        J = J - 1
    Wend
    If I >= J Then
        Partition = J
        Exit Function
    End If

    Swap ODry(I), ODry(J)
Loop
End Function

Private Sub DrySrtLH(ODry, L&, H&)
If L >= H Then Exit Sub
Dim P&
P = Partition(ODry, L, H)
DrySrtLH ODry, L, P
DrySrtLH ODry, P + 1, H
End Sub

Attribute VB_Name = "MDta_Srt"
Option Explicit
Dim A_SrtColIxAy%()
Dim A_IsDesAy() As Boolean
Private Sub Asg(Fny$(), SrtByFF$, OColIxAy%(), OIsDesAy() As Boolean) 'SrtByFF may have - as sfx which means IsDes
Erase OColIxAy
Erase OIsDesAy
'-- SrtByFF is not given, use fst col
    If SrtByFF = "" Then
        PushI OColIxAy, 0
        PushI OIsDesAy, False
        Exit Sub
    End If

'-- ReSz O*
    Dim U%
    Dim SrtFny$(): SrtFny = NyzNN(SrtByFF)
    U = UB(SrtFny)
    If U = -1 Then Exit Sub
    ReDim OColIxAy(U)
    ReDim OIsDesAy(U)

'-- Set O*
    Dim F, J%
    For Each F In SrtFny
        If LasChr(F) = "-" Then
            OIsDesAy(J) = True
            OColIxAy(J) = IxzAy(Fny, RmvLasChr(F))
        Else
            OColIxAy(J) = IxzAy(Fny, F)
        End If
    Next
End Sub
Function DrsSrt(A As Drs, Optional SrtByFF$ = "") As Drs 'If SrtByFF is blank use fst col.
Dim Fny$(): Fny = NyzNN(SrtByFF)
Dim ColIxAy%(): ColIxAy = IntAyzLngAy(IxAy(A.Fny, Fny))
Dim IsDesAy() As Boolean
    Asg Fny, SrtByFF, _
        ColIxAy, IsDesAy
Set DrsSrt = Drs(A.Fny, DrySrt(A.Dry, ColIxAy, IsDesAy))
End Function

Function DrySrt(Dry(), ColIxAy%(), IsDesAy() As Boolean) As Variant()
If Si(ColIxAy) <> Si(IsDesAy) Then Thw CSub, "Si of ColIxAy and IsDesAy are dif", "Si-ColIxAy Si-IsDesAy", Si(ColIxAy) <> Si(IsDesAy)
If Si(ColIxAy) = 1 Then
    DrySrt = DrySrtzCol(Dry, ColIxAy(0), IsDesAy(0))
Else
    DrySrt = DrySrtzColIxAy(Dry, ColIxAy, IsDesAy)
End If
End Function

Function DtSrt(A As Dt, Optional SrtByFF$ = "") As Dt
Set DtSrt = DtzDrs(DrsSrt(DrszDt(A), SrtByFF), A.DtNm)
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
Dim J%, Ix%
IsGT = True
For J = 0 To UB(A_SrtColIxAy)
    Ix = A_SrtColIxAy(J)
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

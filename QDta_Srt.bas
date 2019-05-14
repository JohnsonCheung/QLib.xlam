Attribute VB_Name = "QDta_Srt"
Option Explicit
Private Const CMod$ = "MDta_Srt."
Private Const Asm$ = "QDta"
Dim A_SrtColIxy%()
Dim A_IsDesAy() As Boolean
Private Sub Asg(Fny$(), SrtByFF$, OColIxy&(), OIsDesAy() As Boolean) 'SrtByFF may have - as sfx which means IsDes
Erase OColIxy
Erase OIsDesAy
'-- SrtByFF is not given, use fst col
    If SrtByFF = "" Then
        PushI OColIxy, 0
        PushI OIsDesAy, False
        Exit Sub
    End If

'-- ReSz O*
    Dim U%
    Dim SrtFny$(): SrtFny = TermAy(SrtByFF)
    U = UB(SrtFny)
    If U = -1 Then Exit Sub
    ReDim OColIxy(U)
    ReDim OIsDesAy(U)

'-- Set O*
    Dim F$, I, J%
    For Each I In SrtFny
        F = I
        If LasChr(F) = "-" Then
            OIsDesAy(J) = True
            OColIxy(J) = IxzAy(Fny, RmvLasChr(F))
        Else
            OColIxy(J) = IxzAy(Fny, F)
        End If
    Next
End Sub
Function DrsSrt(A As Drs, Optional SrtByFF$ = "") As Drs 'If SrtByFF is blank use fst col.
Dim Fny$(): Fny = TermAy(SrtByFF)
Dim ColIxy&(): ColIxy = Ixy(A.Fny, Fny)
Dim IsDesAy() As Boolean
    Asg Fny, SrtByFF, _
        ColIxy, IsDesAy
DrsSrt = Drs(A.Fny, DrySrt(A.Dry, ColIxy, IsDesAy))
End Function

Function DrySrt(Dry(), ColIxy&(), IsDesAy() As Boolean) As Variant()
If Si(ColIxy) <> Si(IsDesAy) Then Thw CSub, "Si of ColIxy and IsDesAy are dif", "Si-ColIxy Si-IsDesAy", Si(ColIxy) <> Si(IsDesAy)
If Si(ColIxy) = 1 Then
    DrySrt = DrySrtzCol(Dry, ColIxy(0), IsDesAy(0))
Else
    DrySrt = DrySrtzColIxy(Dry, ColIxy, IsDesAy)
End If
End Function

Function DtSrt(A As Dt, Optional SrtByFF$ = "") As Dt
DtSrt = DtzDrs(DrsSrt(DrszDt(A), SrtByFF), A.DtNm)
End Function

Function DrySrtzCol(Dry(), C&, Optional IsDes As Boolean) As Variant()
Attribute DrySrtzCol.VB_Description = "12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789"
Dim Col(): Col = ColzDry(Dry, C)
Dim Ix&(): Ix = IxyzAySrt(Col, IsDes)
Dim J%
For J = 0 To UB(Ix)
   Push DrySrtzCol, Dry(Ix(J))
Next
End Function

Private Function DrySrtzColIxy(Dry(), SrtColIxy&(), IsDesAy() As Boolean) As Variant()
Dim O(): O = Dry
'A_SrtColIxy = SrtColIxy
'A_IsDesAy = IsDesAy
DrySrtLH O, 0, UB(Dry)
DrySrtzColIxy = O
End Function

Private Function IsGT(Dr1, Dr2) As Boolean
Dim J%, Ix%
IsGT = True
For J = 0 To UB(A_SrtColIxy)
    Ix = A_SrtColIxy(J)
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

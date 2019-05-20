Attribute VB_Name = "QDta_Srt"
Option Compare Text
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
Function SrtDrs(A As Drs, Optional SrtByFF$ = "") As Drs 'If SrtByFF is blank use fst col.
Dim Fny$(): Fny = TermAy(SrtByFF)
Dim ColIxy&(): ColIxy = Ixy(A.Fny, Fny)
Dim IsDesAy() As Boolean
    Asg Fny, SrtByFF, _
        ColIxy, IsDesAy
SrtDrs = Drs(A.Fny, SrtDry(A.Dry, ColIxy, IsDesAy))
End Function

Function SrtDry(Dry(), ColIxy&(), IsDesAy() As Boolean) As Variant()
If Si(ColIxy) <> Si(IsDesAy) Then Thw CSub, "Si of ColIxy and IsDesAy are dif", "Si-ColIxy Si-IsDesAy", Si(ColIxy) <> Si(IsDesAy)
If Si(ColIxy) = 1 Then
    SrtDry = SrtDryzCol(Dry, ColIxy(0), IsDesAy(0))
Else
    SrtDry = SrtDryzColIxy(Dry, ColIxy, IsDesAy)
End If
End Function

Function SrtDt(A As Dt, Optional SrtByFF$ = "") As Dt
SrtDt = DtzDrs(SrtDrs(DrszDt(A), SrtByFF), A.DtNm)
End Function

Function SrtDryzCol(Dry(), C&, Optional IsDes As Boolean) As Variant()
Attribute SrtDryzCol.VB_Description = "12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789"
Dim Col(): Col = ColzDry(Dry, C)
Dim Ix&(): Ix = IxyzSrtAy(Col, IsDes)
Dim J%
For J = 0 To UB(Ix)
   Push SrtDryzCol, Dry(Ix(J))
Next
End Function

Private Function SrtDryzColIxy(Dry(), SrtColIxy&(), IsDesAy() As Boolean) As Variant()
Dim O(): O = Dry
'A_SrtColIxy = SrtColIxy
'A_IsDesAy = IsDesAy
SrtDryLH O, 0, UB(Dry)
SrtDryzColIxy = O
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

Private Sub SrtDryLH(ODry, L&, H&)
If L >= H Then Exit Sub
Dim P&
P = Partition(ODry, L, H)
SrtDryLH ODry, L, P
SrtDryLH ODry, P + 1, H
End Sub

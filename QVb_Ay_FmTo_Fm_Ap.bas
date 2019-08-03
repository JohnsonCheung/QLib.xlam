Attribute VB_Name = "QVb_Ay_FmTo_Fm_Ap"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Ay_FmTo_Fm_Ap."
Private Const Asm$ = "QVb"

Function AvzAy(Ay) As Variant()
If IsAv(Ay) Then AvzAy = Ay: Exit Function
Dim I
For Each I In Itr(Ay)
    Push AvzAy, I
Next
End Function

Function Av(ParamArray ApOf_Itm_Or_Ay()) As Variant()
Dim Av1(): Av = ApOf_Itm_Or_Ay
Av1 = ApOf_Itm_Or_Ay
Av = AvzAyOfItmOrAy(Av1)
End Function

Function AvzAyOfItmOrAy(AyOfItmOrAy) As Variant()
Dim V: For Each V In Itr(AyOfItmOrAy)
    If IsArray(V) Then
        PushIAy AvzAyOfItmOrAy, V
    Else
        PushI AvzAyOfItmOrAy, V
    End If
Next
End Function

Function AvzAp(ParamArray ApOf_Itm_Or_Ay()) As Variant()
Dim Av(): Av = ApOf_Itm_Or_Ay
Av = ApOf_Itm_Or_Ay
AvzAp = AvzAyOfItmOrAy(Av)
End Function

Private Function SyzAv(AvOf_Itm_or_Ay()) As String()
Dim I: For Each I In Itr(AvOf_Itm_or_Ay)
    If IsArray(I) Then
        PushIAy SyzAv, I
    Else
        PushI SyzAv, I
    End If
Next
End Function

Function SyzAp(ParamArray ApOf_Itm_Or_Ay()) As String()
Dim Av(): Av = ApOf_Itm_Or_Ay
SyzAp = SyzAv(Av)
End Function

Function Sy(ParamArray ApOf_Itm_Or_Ay()) As String()
Dim Av(): Av = ApOf_Itm_Or_Ay
Sy = SyzAv(Av)
End Function

Function DteAy(ParamArray Ap()) As Date()
Dim Av(): Av = Ap
DteAy = IntozAy(DteAy, Av)
End Function

Function IntAyzLngAy(LngAp&()) As Integer()
Dim I
For Each I In Itr(LngAp)
    PushI IntAyzLngAy, I
Next
End Function
Function IntAySS(IntSS$) As Integer()
Dim I
For Each I In Itr(SyzSS(IntSS))
    PushI IntAySS, I
Next
End Function

Function IntAy(ParamArray Ap()) As Integer()
Dim Av(): Av = Ap
IntAy = IntozAy(EmpIntAy, Av)
End Function

Function LngAp(ParamArray Ap()) As Long()
Dim Av(): Av = Ap
LngAp = IntozAy(EmpLngAy, Av)
End Function

Function SngAy(ParamArray Ap()) As Single()
Dim Av(): Av = Ap
SngAy = IntozAy(SngAy, Av)
End Function

Function SyNB(ParamArray S_or_Sy()) As String()
Dim Av(): Av = S_or_Sy
Dim I
For Each I In Av
    Select Case True
    Case IsArray(I): PushNBAy SyNB, I
    Case Else: PushNB SyNB, I
    End Select
Next
End Function

Function IntAyzFT(F%, T%) As Integer()
Stop
Dim O%(): ReDim O(Abs(T - F))
IntAyzFT = IntozFT(O, F, T)
End Function

Private Function IntozFT(Into, F, T)
Dim O: O = Into
Dim S: S = IIf(T > F, 1, -1) ' Step
Dim V, I&: For V = F To T Step S
    O(I) = V
    I = I + 1
Next
IntozFT = O
End Function

Function LngAyzFT(F&, T&) As Long()
Dim O&(): ReDim O(T - F)
LngAyzFT = IntozFT(O, F, T)
End Function


Function SyzAy(Ay) As String()
SyzAy = IntozAy(EmpSy, Ay)
End Function


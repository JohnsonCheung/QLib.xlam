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
Function Av(ParamArray Ap()) As Variant()
Av = Ap
End Function
Function AvzAp(ParamArray Ap()) As Variant()
AvzAp = Ap
End Function

Private Function SyzAv(AvOf_Itm_or_Ay()) As String()
Dim I, Av()
Av = AvOf_Itm_or_Ay
For Each I In Itr(Av)
    Select Case True
    Case IsArray(I): PushIAy SyzAv, I
    Case Else: PushI SyzAv, I
    End Select
Next
End Function
Function SyzAp(ParamArray ApOf_Itm_or_Ay()) As String()
Dim Av(): Av = ApOf_Itm_or_Ay
SyzAp = SyzAv(Av)
End Function

Function Sy(ParamArray ApOf_Itm_or_Ay()) As String()
Dim Av(): Av = ApOf_Itm_or_Ay
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
Function IntAy(ParamArray Ap()) As Integer()
Dim Av(): Av = Ap
IntAy = IntozAy(EmpIntAy, Av)
End Function

Function LngSeq(Fm&, ToLng&) As Long()

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

'=========================================================
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

'=========================================================

Function SyzAy(Ay) As String()
SyzAy = IntozAy(EmpSy, Ay)
End Function



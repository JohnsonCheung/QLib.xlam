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
Function SyNonBlank(ParamArray Ap()) As String()
Dim Av(): Av = Ap: SyNonBlank = RmvBlankzAy(SyzAv(Av))
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


Function LngAp(ParamArray Ap()) As Long()
Dim Av(): Av = Ap
LngAp = IntozAy(EmpLngAy, Av)
End Function

Function SngAy(ParamArray Ap()) As Single()
Dim Av(): Av = Ap
SngAy = IntozAy(SngAy, Av)
End Function

Function SyNoBlank(ParamArray S_or_Sy()) As String()
Dim Av(): Av = S_or_Sy
Dim I
For Each I In Av
    Select Case True
    Case IsSy(I): PushNonBlankAy SyNoBlank, CvSy(I)
    Case IsStr(I): PushNonBlank SyNoBlank, CStr(I)
    Case Else: Thw CSub, "Itm must be S or Sy", "TypeName(Itm)", TypeName(I)
    End Select
Next
End Function

'=========================================================

Function IntAyzFmTo(FmInt%, ToInt%) As Integer()
Dim O%(), I&, V%
ReDim O(ToInt - FmInt)
For V = FmInt To ToInt
    O(I) = V
    I = I + 1
Next
IntAyzFmTo = O
End Function
Function LngAyzFmTo(FmLng&, ToLng&) As Long()
Dim O&(), I&, V&
ReDim O(ToLng - FmLng)
For V = FmLng To ToLng
    O(I) = V
    I = I + 1
Next
LngAyzFmTo = O
End Function

'=========================================================

Function SyzAy(Ay) As String()
SyzAy = IntozAy(EmpSy, Ay)
End Function



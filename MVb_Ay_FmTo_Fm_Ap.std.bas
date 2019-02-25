Attribute VB_Name = "MVb_Ay_FmTo_Fm_Ap"
Option Explicit
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
Function SyzAp(ParamArray Ap()) As String()
Dim Av(): Av = Ap: SyzAp = SyzAv(Av)
Dim O$(), I
For Each I In Av
    PushI Sy, I
Next
End Function
Function Sy(ParamArray Ap()) As String()
Dim Av(): Av = Ap: Sy = SyzAv(Av)
End Function

Function DteAy(ParamArray Ap()) As Date()
Dim Av(): Av = Ap
DteAy = IntozAy(DteAy, Av)
End Function

Function IntAy(ParamArray Ap()) As Integer()
Dim Av(): Av = Ap
IntAy = IntozAy(EmpIntAy, Av)
End Function


Function LngAy(ParamArray Ap()) As Long()
Dim Av(): Av = Ap
LngAy = IntozAy(LngAy, Av)
End Function

Function SngAy(ParamArray Ap()) As Single()
Dim Av(): Av = Ap
SngAy = IntozAy(SngAy, Av)
End Function

Function SyNoBlank(ParamArray Itm_or_AyAp()) As String()
Dim Av(): Av = Itm_or_AyAp
Dim I
For Each I In Av
    If IsArray(I) Then
        PushNonBlankSy SyNoBlank, CvSy(I)
    Else
        PushNonBlankStr SyNoBlank, I
    End If
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

Function SyzAv(Av() As Variant) As String()
SyzAv = IntozAy(EmpSy, Av)
End Function


Function SyzAy(Ay) As String()
SyzAy = IntozAy(EmpSy, Ay)
End Function

Function SyzAyNonBlank(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushNonBlankStr SyzAyNonBlank, I
Next
End Function


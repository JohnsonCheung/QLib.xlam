Attribute VB_Name = "QVb_Wrd_Pos"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Wrd_Pos."
Private Const Asm$ = "QVb"

Function WrdLblLinPos$(WrdPos%(), OFmNo&)
Dim O$(), A$, B$, W%, J%
If Si(WrdPos) = 0 Then Exit Function
PushNB O, Space(WrdPos(0) - 1)
For J = 0 To UB(WrdPos) - 1
    A = OFmNo
    W = WrdPos(J + 1) - WrdPos(J)
    If W > Len(A) Then
        A = AlignL(A, W)
        If Len(A) <> W Then Stop
    Else
        A = Space(W)
    End If
    PushI O, A
    OFmNo = OFmNo + 1
Next
A = OFmNo
PushI O, A
WrdLblLinPos = Jn(O)
End Function
Function WrdLblLin(Lin, OFmNo&)
WrdLblLin = WrdLblLinPos(WrdPosAy(Lin), OFmNo)
End Function

Function WrdPosAy(Lin) As Integer()
Dim J%, LasIsSpc As Boolean, CurIsSpc As Boolean
LasIsSpc = True
For J = 1 To Len(Lin)
    CurIsSpc = Mid(Lin, J, 1) = " "
    Select Case True
    Case CurIsSpc And LasIsSpc
    Case CurIsSpc:          LasIsSpc = True
    Case LasIsSpc:          PushI WrdPosAy, J
                            LasIsSpc = False
    Case Else
    End Select
Next
End Function
Function WrdLblLinPairLno(Lin, Lno, LnoWdt, OFmNo&) As String()
Dim O$(): O = WrdLblLinPair(Lin, OFmNo)
O(0) = Space(LnoWdt) & " : " & O(0)
'O(1) = AlignL(Lno, LnoWdt) & " : " & O(1)
WrdLblLinPairLno = O
End Function
Function WrdLblLinPair(Lin, OFmNo&) As String()
PushI WrdLblLinPair, WrdLblLin(Lin, OFmNo)
PushI WrdLblLinPair, Lin
End Function
Function WrdLblLy(Ly$(), OFmNo&) As String()
Dim J&, LnoWdt%, A$
A = UB(Ly)
LnoWdt = Len(A)
For J = 1 To UB(Ly)
    PushIAy WrdLblLy, WrdLblLinPairLno(Ly(J), J, LnoWdt, OFmNo)
Next
End Function


Private Sub Z_WrdLblLin()
Dim Lin, FmNo&
GoSub T0
Exit Sub
T0:
    FmNo = 2
    '               10        20        30        40        50        60
    '      123456789 123456789 123456789 123456789 123456789 123456789 123456789
    Lin = "Lbl01 Lbl02 Lbl03    Lbl04 Lbl05 Lbl06 Lbl07 Lbl08 Lbl09 Lbl10"
    Ept = "2     3     4        5     6     7     8     9     10    11"
    GoTo Tst
Tst:
    Act = WrdLblLin(Lin, FmNo)
    C
    Return
End Sub
Private Sub Z_WrdPosAy()
Dim Lin
GoSub T0
Exit Sub
T0:
    '               10        20        30        40        50        60
    '      123456789 123456789 123456789 123456789 123456789 123456789 123456789
    Lin = "Lbl01 Lbl02 Lbl03    Lbl04 Lbl05 Lbl06 Lbl07 Lbl08 Lbl09 Lbl10"
    Ept = IntAy(1, 7, 13, 22, 28, 34, 40, 46, 52, 58)
    GoTo Tst
Tst:
    Act = WrdPosAy(Lin)
    C
    Return
End Sub

Private Sub Z_WrdLblLy()
Dim Fm&: Fm = 1
Brw WrdLblLy(SrczP(CPj), Fm)
End Sub


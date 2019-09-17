Attribute VB_Name = "MxWrd"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxWrd."
Const WrdPatn$ = "[a-zA-Z][a-zA-Z0-9_]*"
Sub Z_DiWrdqCnt()
Dim A As Dictionary: Set A = DiWrdqCnt(JnCrLf(SrczP(CPj)))
Set A = SrtDic(A)
BrwDic A
End Sub

Sub Z_WrdStszL()
Debug.Print WrdSts(SrczP(CPj))
End Sub

Function WrdStszL$(Lines)
WrdStszL = WrdSts(SplitCrLf(Lines))
End Function

Function WrdSts$(Ly$())
Dim W&, D&, Sy$(), B&, L&, S$
S = JnCrLf(Ly)
Sy = WrdAy(S)
W = Si(Sy)
D = Si(AwDist(Sy))
L = Si(Ly)
B = Len(S)
WrdSts = WrdSts_(B, L, W, D)
End Function

Function WrdSts_$(B&, L&, W&, D&)
Dim BB As String * 9: RSet BB = B
Dim LL As String * 9: RSet LL = L
Dim WW As String * 9: RSet WW = W
Dim DD As String * 9: RSet DD = D
WrdSts_ = FmtQQ("Len            : ?|Lines          : ?|Words          : ?|Distinct Words : ?", BB, LL, WW, DD)
End Function

Function NWrd&(S)
NWrd = Si(WrdAy(S))
End Function

Function NDistWrd&(S)
NDistWrd = Si(AwDist(WrdAy(S)))
End Function

Function DiWrdqCnt(S) As Dictionary
Set DiWrdqCnt = DiKqCnt(WrdAy(S))
End Function

Function WrdAset(S) As Aset
Set WrdAset = AsetzAy(WrdAy(S))
End Function

Function CvMch(A) As IMatch
Set CvMch = A
End Function

Function FstWrdAsetP() As Aset
Set FstWrdAsetP = New Aset
Dim L: For Each L In Itr(AeVbRmk(SrczP(CPj)))
    FstWrdAsetP.PushItm FstWrd(L)
Next
End Function

Function FstWrd$(S)
Static X As RegExp
If IsNothing(X) Then Set X = Rx(WrdPatn)
Dim A As MatchCollection: Set A = X.Execute(S)
Select Case A.Count
Case 0: Exit Function
Case Else: FstWrd = CvMch(A(0)).Value
End Select
End Function
Function NNmChrRx() As RegExp
Dim O As New RegExp
Set O = Rx("\W")
End Function

Function RplNNmChr$(S)
'Ret :S #Rpl-Non-NmChr-To-Spc# @@
RplNNmChr = NNmChrRx.Replace(S, " ")
End Function
Function AmRplNNmChr(Ay) As String()
Dim L: For Each L In Itr(Ay)
    PushI AmRplNNmChr, RplNNmChr(L)
Next
End Function
Private Sub Z_AmRplNNmChr()
Brw AmRplNNmChr(SrczP(CPj))
End Sub
Function WrdAy(S) As String()
':Wrd: :S #Wrd# !Fst-Let is FstNmChr and Rest is NmChr.  All Non-NmChr will be replace to Spc and Take SyzSS
Static X As RegExp
If IsNothing(X) Then Set X = Rx(WrdPatn, IsGlobal:=True, MultiLine:=True)
Dim I As Match: For Each I In X.Execute(S)
    PushI WrdAy, I.Value
Next
End Function

Function Wrdss$(S)
Wrdss = JnSpc(WrdAy(S))
End Function

Function WrdssAy(Sy$()) As String()
Dim S: For Each S In Itr(Sy)
    PushI WrdssAy, Wrdss(S)
Next
End Function


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


Sub Z_WrdLblLin()
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
Sub Z_WrdPosAy()
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

Sub Z_WrdLblLy()
Dim Fm&: Fm = 1
Brw WrdLblLy(SrczP(CPj), Fm)
End Sub

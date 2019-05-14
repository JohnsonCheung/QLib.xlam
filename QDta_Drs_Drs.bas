Attribute VB_Name = "QDta_Drs_Drs"
Option Explicit
Const Asm$ = "QDta"
Const NS$ = "Dta.Ds"
Private Const CMod$ = "BDrs."
Type Drs: Fny() As String: Dry() As Variant: End Type
Type Drss: N As Integer: Ay() As Drs: End Type
Enum EmCnt
    EiCntAll
    EiCntDup
    EiCntSng
End Enum
Enum EmCntSrtOpt
    eNoSrt
    eSrtByCnt
    eSrtByItm
End Enum
Function DrszFF(FF$, Dry()) As Drs
DrszFF = Drs(TermAy(FF), Dry)
End Function
Function Drs(Fny$(), Dry()) As Drs
With Drs
    .Fny = Fny
    .Dry = Dry
End With
End Function

Function DrsAddCol(A As Drs, ColNm$, CnstBrk) As Drs
DrsAddCol = Drs(CvSy(AddAyItm(A.Fny, ColNm)), DryAddColzC(A.Dry, CnstBrk))
End Function

Function DrsAddIxCol(A As Drs, HidIxCol As Boolean) As Drs
If HidIxCol Then
    DrsAddIxCol = A
    Exit Function
End If

Dim Fny$()
    Fny = AyInsEle(A.Fny, "Ix")
Dim Dry()
    Dim J&, I, Dr
    For Each I In Itr(A.Dry)
        Dr = AyInsEle(I, J): J = J + 1
        Push Dry, Dr
    Next
DrsAddIxCol = Drs(Fny, Dry)
End Function

Function AvDrsC(A As Drs, C) As Variant()
AvDrsC = IntoDrsC(Array(), A, C)
End Function

Function IntoDrsC(Into, A As Drs, C)
Dim O, Ix%, Dry(), Dr
Ix = IxzAy(A.Fny, C): If Ix = -1 Then Stop
O = Into
Erase O
Dry = A.Dry
If Si(Dry) = 0 Then IntoDrsC = O: Exit Function
For Each Dr In Dry
    Push O, Dr(Ix)
Next
IntoDrsC = O
End Function
Function DmpRec(A As Drs)
D FmtRec(A)
End Function
Function FmtRec(A As Drs) As String()
Dim Fny$(), Dr, N&, Ix&, Fny$()
Fny = AlignLzAy(A.Fny)
For Each Dr In Itr(A.Dry)
    PushIAy FmtRec, FmtRec_FmAlignedFny_AndDr(Fny, Dr, Ix, N)
    Ix = Ix + 1
Next
End Function
Function IxOfUStr$(Ix&, U&)

End Function
Function FmtReczFnyDr(Fny$(), Dr, Optional Ix& = -1, Optional N& = -1)
FmtReczFnyDr = FmtRec_FmAlignedFny_AndDr(AlignLzAy(Fny), Dr, Ix, N)
End Function

Function FmtRec_FmAlignedFny_AndDr(AlignedFny$(), Dr, Optional Ix& = -1, Optional N& = -1)
PushNonBlank FmtReczFnyDr, IxOfUStr(Ix, U)

End Function
Sub DmpDrs(A As Drs, Optional MaxColWdt% = 100, Optional BrkColNm$)
DmpAy FmtDrs(A, MaxColWdt, BrkColNm$)
End Sub

Function DrpCny(A As Drs, CC$) As Drs
DrpCny = SelDrsCC(A, MinusAy(A.Fny, Ny(CC)))
End Function

Function SelDrsCC(A As Drs, CC$) As Drs
Const CSub$ = CMod & "SelDrsCC"
Dim OFny$(): OFny = TermAy(CC)
If Not IsAySub(A.Fny, OFny) Then Thw CSub, "Given FF has some field not in Drs.Fny", "CC Drs.Fny", CC, A.Fny
Dim ODry()
    Dim IAy&()
    IAy = Ixy(A.Fny, OFny)
    ODry = DrySelColIxy(A.Dry, IAy)
SelDrsCC = Drs(OFny, ODry)
End Function
Function DrySelColIxy(Dry(), Ixy&()) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
    PushI DrySelColIxy, AywIxy(Dr, Ixy)
Next
End Function

Function DrsInsCV(A As Drs, C$, V) As Drs
DrsInsCV = Drs(CvSy(AyInsEle(A.Fny, C)), DryInsColzV(A.Dry, V, IxzAy(A.Fny, C)))
End Function

Function DrsInsCVAft(A As Drs, C$, V, AftFldNm$) As Drs
DrsInsCVAft = DrsInsCVIsAftFld(A, C, V, True, AftFldNm)
End Function

Function DrsInsCVBef(A As Drs, C$, V, BefFldNm$) As Drs
DrsInsCVBef = DrsInsCVIsAftFld(A, C, V, False, BefFldNm)
End Function

Private Function DrsInsCVIsAftFld(A As Drs, C$, V, IsAft As Boolean, FldNm$) As Drs
Dim Fny$(), Dry(), Ix&, Fny1$()
Fny = A.Fny
Ix = IxzAy(Fny, C): If Ix = -1 Then Stop
If IsAft Then
    Ix = Ix + 1
End If
Fny1 = AyInsEle(Fny, FldNm, CLng(Ix))
Dry = DryInsColzV(A.Dry, V, Ix)
DrsInsCVIsAftFld = Drs(Fny1, Dry)
End Function

Function IsEqDrs(A As Drs, B As Drs) As Boolean
If Not IsEqAy(A.Fny, B.Fny) Then Exit Function
If Not IsEqDry(A.Dry, B.Dry) Then Exit Function
IsEqDrs = True
End Function

Sub BrwCnt(Ay, Optional Opt As EmCnt)
Brw FmtCntDic(CntDic(Ay, Opt))
End Sub
Function DicItmWdt%(A As Dictionary)
Dim I, O%
For Each I In A.Items
    O = Max(Len(I), O)
Next
DicItmWdt = O
End Function
Private Function CntLyzCntDic(CntDic As Dictionary, CntWdt%) As String()
Dim K
For Each K In CntDic.Keys
    PushI CntLyzCntDic, AlignR(CntDic(K), CntWdt) & " " & K
Next
End Function
Function CntLy(Ay, Optional Opt As EmCnt, Optional SrtOpt As EmCntSrtOpt, Optional IsDesc As Boolean) As String()
Dim D As Dictionary: Set D = CntDic(Ay, Opt)
Dim K
Dim W%: W = DicItmWdt(D)
Dim O$()
Select Case SrtOpt
Case eNoSrt
    CntLy = CntLyzCntDic(D, W)
Case eSrtByCnt
    CntLy = QSrt1(CntLyzCntDic(D, W), IsDesc)
Case eSrtByItm
    CntLy = CntLyzCntDic(SrtDic(D, IsDesc), W)
Case Else
    Thw CSub, "Invalid SrtOpt", "SrtOpt", SrtOpt
End Select
End Function

Function NColzDrs%(A As Drs)
NColzDrs = Max(Si(A.Fny), NColzDry(A.Dry))
End Function

Function NRowDrs&(A As Drs)
NRowDrs = Si(A.Dry)
End Function

Function DrwIxy(Dr(), Ixy&())
Dim U&: U = MaxAy(Ixy)
Dim O: O = Dr
If UB(O) < U Then
    ReDim Preserve O(U)
End If
DrwIxy = AywIxy(O, Ixy)
End Function
Function SelCol(Dry(), Ixy&()) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
    PushI SelCol, AywIxy(Dr, Ixy)
Next
End Function
Function ReOrdCol(A As Drs, BySubFF$) As Drs
Dim SubFny$(): SubFny = TermAy(BySubFF)
Dim OFny$(): OFny = AyReOrd(A.Fny, SubFny)
Dim IAy&(): IAy = Ixy(A.Fny, OFny)
Dim ODry(): ODry = SelCol(A.Dry, IAy)
ReOrdCol = Drs(OFny, ODry)
End Function

Function NRowzColEv&(A As Drs, ColNm$, EqVal)
NRowzColEv = NRowzInDryzColEv(A.Dry, IxzAy(A.Fny, ColNm), EqVal)
End Function

Function SqzDrs(A As Drs) As Variant()
Dim NC&, NR&, Dry(), Fny$()
    Fny = A.Fny
    Dry = A.Dry
    NC = Max(NColzDry(Dry), Si(Fny))
    NR = Si(Dry)
Dim O()
ReDim O(1 To 1 + NR, 1 To NC)
Dim C&, R&, Dr
    For C = 1 To Si(Fny)
        O(1, C) = Fny(C - 1)
    Next
    For R = 1 To NR
        Dr = Dry(R - 1)
        For C = 1 To Min(Si(Dr), NC)
            O(R + 1, C) = Dr(C - 1)
        Next
    Next
SqzDrs = O
End Function

Function SyzDrsC(A As Drs, ColNm$) As String()
SyzDrsC = IntoDrsC(EmpSy, A, ColNm)
End Function

Sub PushDrs(O As Drss, M As Drs)
With O
    ReDim Preserve .Ay(.N)
    .Ay(.N) = M
    .N = .N + 1
End With
End Sub
Sub ApdDrs(O As Drs, M As Drs)
If Not IsEqAy(O.Fny, M.Fny) Then Thw CSub, "Fny are dif", "O.Fny M.Fny", O.Fny, M.Fny
Dim UO&, UM&, U&, J&
UO = UB(O.Dry)
UM = UB(M.Dry)
U = UO + UM + 1
ReDim Preserve O.Dry(U)
For J = UO + 1 To U
    O.Dry(J) = M.Dry(J - UO - 1)
Next
End Sub

Private Sub ZZ_GpDicDKG()
Dim Act As Dictionary, Dry(), Dr1, Dr2, Dr3
Dr1 = Array("A", , 1)
Dr2 = Array("A", , 2)
Dr3 = Array("B", , 3)
Dry = Array(Dr1, Dr2, Dr3)
Set Act = DryGpDic(Dry, IntAy(0), 2)
Ass Act.Count = 2
Ass IsEqAy(Act("A"), Array(1, 2))
Ass IsEqAy(Act("B"), Array(3))
Stop
End Sub

Private Sub ZZ_CntDiczDrs()
Dim Drs As Drs, Dic As Dictionary
'Drs = Vbe_Mth12Drs(CVbe)
Set Dic = CntDiczDrs(Drs, "Nm")
BrwDic Dic
End Sub

Private Sub ZZ_SelDrs()
BrwDrs SelDrs(SampDrs1, "A B D")
End Sub

Private Property Get Z_FmtDrs()
GoTo ZZ
ZZ:
DmpAy FmtDrs(SampDrs1)
End Property

Private Sub ZZ()
Dim A As Variant
Dim B()
Dim C As Drs
Dim D$
Dim E%
Dim F$()
DrsAddCol C, D, A
DrsAddCol C, D, A
AddColzValIdzCntzDrs C, D, D
BrwDrs C, E, D, D
DtzDrs C, D
DrsInsCV C, D, A
End Sub


Function DrsAddCC(A As Drs, FF$, C1, C2) As Drs
Dim Fny$(), Dry()
Fny = AddAy(A.Fny, TermAy(FF))
Dry = DryAddColzCC(A.Dry, C1, C2)
DrsAddCC = Drs(Fny, Dry)
End Function


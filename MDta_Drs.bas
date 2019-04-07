Attribute VB_Name = "MDta_Drs"
Option Explicit
Const CMod$ = "MDta_Drs."
Enum eCntOpt
    eAllCnt
    eDupCnt
    eSngCnt
End Enum
Enum eCntSrtOpt
    eNoSrt
    eSrtByCnt
    eSrtByItm
End Enum
Function CvDrs(A) As Drs
Set CvDrs = A
End Function

Function Drs(FF, Dry()) As Drs
Dim O As New Drs
Set Drs = O.Init(NyzNN(FF), Dry)
End Function

Function DrsAddCol(A As Drs, ColNm$, ConstVal) As Drs
Set DrsAddCol = Drs(AyAddItm(A.Fny, ColNm), DryAddColzC(A.Dry, ConstVal))
End Function

Function DrsAddIxCol(A As Drs, HidIxCol As Boolean) As Drs
If HidIxCol Then
    Set DrsAddIxCol = A
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
Set DrsAddIxCol = Drs(Fny, Dry)
End Function


Function IsDrs(A) As Boolean
IsDrs = TypeName(A) = "Drs"
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

Sub DmpDrs(A As Drs, Optional MaxColWdt% = 100, Optional BrkColNm$)
DmpAy FmtDrs(A, MaxColWdt, BrkColNm$)
End Sub

Function DrsDrpCC(A As Drs, CC) As Drs
Set DrsDrpCC = DrsSelCC(A, AyMinus(A.Fny, Ny(CC)))
End Function

Function DrsSelCC(A As Drs, CC) As Drs
Const CSub$ = CMod & "DrsSelCC"
Dim OFny$(): OFny = NyzNN(CC)
If Not IsAySub(A.Fny, OFny) Then Thw CSub, "Given FF has some field not in Drs.Fny", "CC Drs.Fny", CC, A.Fny
Dim ODry()
    Dim IAy&()
    IAy = IxAy(A.Fny, OFny)
    ODry = DrySelColIxAy(A.Dry, IAy)
Set DrsSelCC = Drs(OFny, ODry)
End Function
Function DrySelColIxAy(Dry(), IxAy&()) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
    PushI DrySelColIxAy, AywIxAy(Dr, IxAy)
Next
End Function
Function DtDrsDtnm(A As Drs, DtNm$) As Dt
Set DtDrsDtnm = Dt(DtNm, A.Fny, A.Dry)
End Function

Function DrsInsCV(A As Drs, C$, V) As Drs
Set DrsInsCV = Drs(AyInsEle(A.Fny, C), DryInsColzV(A.Dry, V, IxzAy(A.Fny, C)))
End Function

Function DrsInsCVAft(A As Drs, C$, V, AftFldNm$) As Drs
Set DrsInsCVAft = DrsInsCVIsAftFld(A, C, V, True, AftFldNm)
End Function

Function DrsInsCVBef(A As Drs, C$, V, BefFldNm$) As Drs
Set DrsInsCVBef = DrsInsCVIsAftFld(A, C, V, False, BefFldNm)
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
Set DrsInsCVIsAftFld = Drs(Fny1, Dry)
End Function

Function IsEqDrs(A As Drs, B As Drs) As Boolean
If Not IsEqAy(A.Fny, B.Fny) Then Exit Function
If Not IsEqDry(A.Dry, B.Dry) Then Exit Function
IsEqDrs = True
End Function

Sub BrwCnt(Ay, Optional IgnCas As Boolean, Optional Opt As eCntOpt)
Brw FmtCntDic(CntDic(Ay, IgnCas, Opt))
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
Function CntLy(Ay, Optional Opt As eCntOpt, Optional SrtOpt As eCntSrtOpt, Optional IsDesc As Boolean, Optional IgnCas As Boolean) As String()
Dim D As Dictionary: Set D = CntDic(Ay, IgnCas, Opt)
Dim K
Dim W%: W = DicItmWdt(D)
Dim O$()
Select Case SrtOpt
Case eNoSrt
    CntLy = CntLyzCntDic(D, W)
Case eSrtByCnt
    CntLy = AyQSrt(CntLyzCntDic(D, W), IsDesc)
Case eSrtByItm
    CntLy = CntLyzCntDic(DicSrt(D, IsDesc), W)
Case Else
    Thw CSub, "Invalid SrtOpt", "SrtOpt", SrtOpt
End Select
End Function
Function CntDic(Ay, Optional IgnCas As Boolean, Optional Opt As eCntOpt) As Dictionary
Dim O As New Dictionary, I, C&
If IgnCas Then O.CompareMode = TextCompare
For Each I In Itr(Ay)
    If O.Exists(I) Then
        O(I) = O(I) + 1
    Else
        O.Add I, 1
    End If
Next
Dim K
Select Case Opt
Case eCntOpt.eDupCnt
    Set CntDic = New Dictionary
    For Each K In O.Keys
        C = O(K)
        If C > 1 Then CntDic.Add K, C
    Next
Case eCntOpt.eSngCnt
    Set CntDic = New Dictionary
    For Each K In O.Keys
        C = O(K)
        If C = 1 Then CntDic.Add K, C
    Next
Case Else: Set CntDic = O
End Select
End Function

Function CntDiczDrs(A As Drs, C$) As Dictionary
Set CntDiczDrs = CntDic(ColzDrs(A, C))
End Function
Function NColzDrs%(A As Drs)
NColzDrs = Max(Si(A.Fny), NColzDry(A.Dry))
End Function

Function NRowDrs&(A As Drs)
NRowDrs = Si(A.Dry)
End Function

Function DrwIxAy(Dr, IxAy)
Dim U&: U = MaxAy(IxAy)
Dim O: O = Dr
If UB(O) < U Then
    ReDim Preserve O(U)
End If
DrwIxAy = AywIxAy(O, IxAy)
End Function
Function DrySelIxAy(Dry(), IxAy) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
    PushI DrySelIxAy, DrwIxAy(Dr, IxAy)
Next
End Function
Function DrsReOrdBy(A As Drs, BySubFF) As Drs
Dim SubFny$(): SubFny = NyzNN(BySubFF)
Dim OFny$(): OFny = AyReOrd(A.Fny, SubFny)
Dim IAy&(): IAy = IxAy(A.Fny, OFny)
Dim ODry(): ODry = DrySelIxAy(A.Dry, IAy)
Set DrsReOrdBy = Drs(OFny, ODry)
End Function

Function NRowDrsCEv&(A As Drs, ColNm$, EqVal)
NRowDrsCEv = NRowDryCEv(A.Dry, IxzAy(A.Fny, ColNm), EqVal)
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

Function SyDrsC(A As Drs, ColNm) As String()
SyDrsC = IntoDrsC(EmpSy, A, ColNm)
End Function

Sub PushDrs(O As Drs, A As Drs)
If IsNothing(O) Then
    Set O = A
    Exit Sub
End If
If IsNothing(A) Then Exit Sub
If Not IsEq(O.Fny, A.Fny) Then Stop
Set O = Drs(O.Fny, CvAy(AyAddAp(O.Dry, A.Dry)))
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
'Set Drs = Vbe_Mth12Drs(CurVbe)
Set Dic = CntDiczDrs(Drs, "Nm")
BrwDic Dic
End Sub

Private Sub ZZ_DrsSel()
BrwDrs DrsSel(SampDrs1, "A B D")
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
CvDrs A
DrsAddCol C, D, A
DrsAddCol C, D, A
AddColzValIdzCntzDrs C, D, D
BrwDrs C, E, D, D
DrsSelCC C, A
DtDrsDtnm C, D
DrsInsCV C, D, A
PushDrs C, C
End Sub

Private Sub Z()
End Sub
Function DrsAddCC(A As Drs, FF, C1, C2) As Drs
Dim Fny$(), Dry()
Fny = AyAdd(A.Fny, NyzNN(FF))
Dry = DryAddColzCC(A.Dry, C1, C2)
Set DrsAddCC = Drs(Fny, Dry)
End Function


Attribute VB_Name = "MDta_Drs"
Option Explicit
Const CMod$ = "MDta_Drs."

Function CvDrs(A) As DRs
Set CvDrs = A
End Function

Function DRs(FF, Dry()) As DRs
Dim O As New DRs
Set DRs = O.Init(FnyzFF(FF), Dry)
End Function

Function DrsAddCol(A As DRs, ColNm$, ConstVal) As DRs
Set DrsAddCol = DRs(AyAddItm(A.Fny, ColNm), DryAddCol(A.Dry, ConstVal))
End Function

Function DrsAddIxCol(A As DRs, HidIxCol As Boolean) As DRs
If HidIxCol Then
    Set DrsAddIxCol = A
    Exit Function
End If

Dim Fny$()
    Fny = AyInsItm(A.Fny, "Ix")
Dim Dry()
    Dim J&, I, Dr
    For Each I In Itr(A.Dry)
        Dr = AyInsItm(I, J): J = J + 1
        Push Dry, Dr
    Next
Set DrsAddIxCol = DRs(Fny, Dry)
End Function


Function IsDrs(A) As Boolean
IsDrs = TypeName(A) = "Drs"
End Function

Function AvDrsC(A As DRs, C) As Variant()
AvDrsC = IntoDrsC(Array(), A, C)
End Function

Function IntoDrsC(Into, A As DRs, C)
Dim O, Ix%, Dry(), Dr
Ix = IxzAy(A.Fny, C): If Ix = -1 Then Stop
O = Into
Erase O
Dry = A.Dry
If Sz(Dry) = 0 Then IntoDrsC = O: Exit Function
For Each Dr In Dry
    Push O, Dr(Ix)
Next
IntoDrsC = O
End Function

Sub DmpDrs(A As DRs, Optional MaxColWdt% = 100, Optional BrkColNm$)
DmpAy FmtDrs(A, MaxColWdt, BrkColNm$)
End Sub

Function DrsDrpCC(A As DRs, CC) As DRs
Set DrsDrpCC = DrsSelCC(A, AyMinus(A.Fny, Ny(CC)))
End Function

Function DrsSelCC(A As DRs, CC) As DRs
Const CSub$ = CMod & "DrsSelCC"
Dim OFny$(): OFny = FnyzFF(CC)
If Not IsAySub(A.Fny, OFny) Then Thw CSub, "Given FF has some field not in Drs.Fny", "CC Drs.Fny", CC, A.Fny
Dim ODry()
    Dim IAy&()
    IAy = IxAy(A.Fny, OFny)
    ODry = DrySelColIxAy(A.Dry, IAy)
Set DrsSelCC = DRs(OFny, ODry)
End Function
Function DrySelColIxAy(Dry(), IxAy) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
    PushI DrySelColIxAy, AywIxAy(Dr, IxAy)
Next
End Function
Function DtDrsDtnm(A As DRs, DtNm$) As DT
Set DtDrsDtnm = DT(DtNm, A.Fny, A.Dry)
End Function

Function DrsInsCV(A As DRs, C$, V) As DRs
Set DrsInsCV = DRs(AyInsItm(A.Fny, C), DryInsV(A.Dry, V, IxzAy(A.Fny, C)))
End Function

Function DrsInsCVAft(A As DRs, C$, V, AftFldNm$) As DRs
Set DrsInsCVAft = DrsInsCVIsAftFld(A, C, V, True, AftFldNm)
End Function

Function DrsInsCVBef(A As DRs, C$, V, BefFldNm$) As DRs
Set DrsInsCVBef = DrsInsCVIsAftFld(A, C, V, False, BefFldNm)
End Function

Private Function DrsInsCVIsAftFld(A As DRs, C$, V, IsAft As Boolean, FldNm$) As DRs
Dim Fny$(), Dry(), Ix&, Fny1$()
Fny = A.Fny
Ix = IxzAy(Fny, C): If Ix = -1 Then Stop
If IsAft Then
    Ix = Ix + 1
End If
Fny1 = AyInsItm(Fny, FldNm, CLng(Ix))
Dry = DryInsV(A.Dry, V, Ix)
Set DrsInsCVIsAftFld = DRs(Fny1, Dry)
End Function

Function IsEqDrs(A As DRs, B As DRs) As Boolean
If Not IsEqAy(A.Fny, B.Fny) Then Exit Function
If Not IsEqDry(A.Dry, B.Dry) Then Exit Function
IsEqDrs = True
End Function

Function CntDic(Ay, Optional IgnCas As Boolean) As Dictionary
Dim O As New Dictionary, I
If IgnCas Then O.CompareMode = TextCompare
For Each I In Itr(Ay)
    If O.Exists(I) Then
        O(I) = O(I) + 1
    Else
        O.Add I, 1
    End If
Next
Set CntDic = O
End Function

Function CntDiczDrs(A As DRs, C$) As Dictionary
Set CntDiczDrs = CntDic(ColzDrs(A, C))
End Function
Function NColzDrs%(A As DRs)
NColzDrs = Max(Sz(A.Fny), NColDry(A.Dry))
End Function

Function NRowDrs&(A As DRs)
NRowDrs = Sz(A.Dry)
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
Function DrsReOrdBy(A As DRs, BySubFF) As DRs
Dim SubFny$(): SubFny = FnyzFF(BySubFF)
Dim OFny$(): OFny = AyReOrd(A.Fny, SubFny)
Dim IAy&(): IAy = IxAy(A.Fny, OFny)
Dim ODry(): ODry = DrySelIxAy(A.Dry, IAy)
Set DrsReOrdBy = DRs(OFny, ODry)
End Function

Function NRowDrsCEv&(A As DRs, ColNm$, EqVal)
NRowDrsCEv = NRowDryCEv(A.Dry, IxzAy(A.Fny, ColNm), EqVal)
End Function

Function SqzDrs(A As DRs) As Variant()
Dim NC&, NR&, Dry(), Fny$()
    Fny = A.Fny
    Dry = A.Dry
    NC = Max(NColDry(Dry), Sz(Fny))
    NR = Sz(Dry)
Dim O()
ReDim O(1 To 1 + NR, 1 To NC)
Dim C&, R&, Dr
    For C = 1 To Sz(Fny)
        O(1, C) = Fny(C - 1)
    Next
    For R = 1 To NR
        Dr = Dry(R - 1)
        For C = 1 To Min(Sz(Dr), NC)
            O(R + 1, C) = Dr(C - 1)
        Next
    Next
SqzDrs = O
End Function

Function SyDrsC(A As DRs, ColNm) As String()
SyDrsC = IntoDrsC(EmpSy, A, ColNm)
End Function
Function PrpNy(PP) As String()
PrpNy = FnyzFF(PP) 'Stop '
End Function

Sub PushDrs(O As DRs, A As DRs)
If IsNothing(O) Then
    Set O = A
    Exit Sub
End If
If IsNothing(A) Then Exit Sub
If Not IsEq(O.Fny, A.Fny) Then Stop
Set O = DRs(O.Fny, CvAy(AyAddAp(O.Dry, A.Dry)))
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
Dim DRs As DRs, Dic As Dictionary
'Set Drs = Vbe_Mth12Drs(CurVbe)
Set Dic = CntDiczDrs(DRs, "Nm")
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
Dim C As DRs
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
Function DrsAddCC(A As DRs, FF, C1, C2) As DRs
Dim Fny$(), Dry()
Fny = AyAdd(A.Fny, CvNy(FF))
Dry = DryAddCC(A.Dry, C1, C2)
Set DrsAddCC = DRs(Fny, Dry)
End Function


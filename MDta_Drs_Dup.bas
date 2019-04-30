Attribute VB_Name = "MDta_Drs_Dup"
Option Explicit
Function DrswDup(A As Drs, FF$) As Drs
DrswDup = DrswRowIxAy(A, RowIxAyOfDupRow(A, FF))
End Function

Function DrseDup(A As Drs, FF$) As Drs
Dim RowIxAy&(): RowIxAy = RowIxAyOfDupRow(A, FF)
DrseDup = DrseRowIxAy(A, RowIxAy)
End Function

Private Function RowIxAyOfDupRow(A As Drs, FF$) As Long()
Dim Fny$(): Fny = TermAy(FF)
If Si(Fny) = 1 Then
    RowIxAyOfDupRow = IxAyzDup(ColzDrs(A, Fny(0)))
    Exit Function
End If
Dim ColIxAy&(): ColIxAy = IxAy(A.Fny, Fny, ThwNotFnd:=True)
Dim Dry(): Dry = DrySel(A.Dry, ColIxAy)
RowIxAyOfDupRow = RowIxAyzOfDupzDry(Dry)
End Function

Private Function RowIxAyzOfDupzDry(Dry()) As Long()
Dim DupD(): DupD = DrywDup(Dry)
Dim Dr, Ix&, O&()
For Each Dr In Dry
    If HasDr(DupD, Dr) Then PushI O, Ix
    Ix = Ix + 1
Next
If Si(O) < Si(DupD) * 2 Then Stop
RowIxAyzOfDupzDry = O
End Function

Function DrywDup(Dry()) As Variant()
If Si(Dry) = 0 Then Exit Function
Dim Dr
For Each Dr In DryGpCnt(Dry)
    If Dr(0) > 1 Then
        PushI DrywDup, AyeFstEle(Dr)
    End If
Next
End Function
Function DrywDist(Dry()) As Variant()
If Si(Dry) = 0 Then Exit Function
Dim Dr
For Each Dr In DryGpCnt(Dry)
    PushI DrywDist, Dr
Next
End Function

Function DryGpCnt(Dry()) As Variant()
#If True Then
    DryGpCnt = DryGpCntzSlow(Dry)
#Else
    DryGpCnt = DryGpCntzQuick(Dry)
#End If
End Function

Private Function DryGpCntzQuick(Dry()) As Variant()
End Function

Private Function DryGpCntzSlow(Dry()) As Variant()
If Si(Dry) = 0 Then Exit Function
Dim OKeyDry(), OCnt&(), Dr
    Dim LasIx&: LasIx = Si(Dry(0))
    Dim J&
    For Each Dr In Dry
        If J Mod 500 = 0 Then Debug.Print "DryGpCntzSlow"
        If J Mod 50 = 0 Then Debug.Print J;
        J = J + 1
        With IxOptzDryDr(OKeyDry, Dr)
            Select Case .Som
            Case True: OCnt(.Lng) = OCnt(.Lng) + 1
            Case Else: PushI OKeyDry, Dr: PushI OCnt, 1
            End Select
        End With
    Next
    If Si(OKeyDry) <> Si(OCnt) Then Thw CSub, "Si Diff", "OKeyDry-Si OCnt-Si", Si(OKeyDry), Si(OCnt)
For J = 0 To UB(OCnt)
    PushI DryGpCntzSlow, AyAdd(Array(OCnt(J)), OKeyDry(J)) '<===========
Next
End Function

Private Function IxOptzDryDr(Dry(), Dr) As LngOpt
Dim IDr, Ix&
For Each IDr In Itr(Dry)
    If IsEqAy(IDr, Dr) Then IxOptzDryDr = SomLng(Ix): Exit Function
    Ix = Ix + 1
Next
End Function
Private Sub Z_DrswDup()
Dim A As Drs, FF$, Act As Drs
GoSub T0
Exit Sub
T0:
    A = DrszFF("A B C", Av(Av(1, 2, "xxx"), Av(1, 2, "yyyy"), Av(1, 2), Av(1), Av(Empty, 2)))
    FF = "A B"
    GoTo Tst
Tst:
    Act = DrswDup(A, FF)
    DmpDrs Act
    Return
End Sub
'======================================================================
Private Function RowIxAyzOfDupzDryColIx(Dry(), ColIx&) As Long()
Dim D As New Dictionary, FstIx&, V, O As New Rel, Ix&, I
For Ix = 0 To UB(Dry)
    V = Dry(Ix)(ColIx)
    If D.Exists(V) Then
        O.PushParChd V, D(V)
        O.PushParChd V, Ix
    Else
        D.Add V, Ix
    End If
Next
For Each I In O.SetOfPar.Itms
    PushIAy RowIxAyzOfDupzDryColIx, O.ParChd(I).Av
Next
End Function

Private Sub Z_RowIxAyzOfDupzDryColIx()
Dim Dry(), ColIx&, Act&(), Ept&()
GoSub T0
Exit Sub
T0:
    ColIx = 0
    Dry = Array(Array(1, 2, 3, 4), Array(1, 2, 3), Array(2, 4, 3))
    Ept = LngAy(0, 1)
    GoTo Tst
Tst:
    Act = RowIxAyzOfDupzDryColIx(Dry, ColIx)
    If Not IsEqAy(Act, Ept) Then Stop
    C
    Return
End Sub

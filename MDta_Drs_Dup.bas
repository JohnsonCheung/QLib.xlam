Attribute VB_Name = "MDta_Drs_Dup"
Option Explicit
Function DrswDup(A As Drs, FF) As Drs
Set DrswDup = DrswRowIxAy(A, RowIxAyzDupzDrs(A, FF))
End Function

Function DrseDup(A As Drs, FF) As Drs
Dim RowIxAy&(): RowIxAy = RowIxAyzDupzDrs(A, FF)
Set DrseDup = DrseRowIxAy(A, RowIxAy)
End Function

Private Function RowIxAyzDupzDrs(A As Drs, FF) As Long()
Dim Fny$(): Fny = NyzNN(FF)
If Si(Fny) = 1 Then
    RowIxAyzDupzDrs = IxAyzDup(ColzDrs(A, Fny(0)))
    Exit Function
End If
Dim ColIxAy&(): ColIxAy = IxAy(A.Fny, Fny, ThwNotFnd:=True)
Dim Dry(): Dry = DrySel(A.Dry, ColIxAy)
RowIxAyzDupzDrs = RowIxAyzDupzDry(Dry)
End Function

Private Function RowIxAyzDupzDry(Dry()) As Long()
Dim DupD(): DupD = DryzDup(Dry)
Dim Dr, Ix&, O&()
For Each Dr In Dry
    If HasDr(DupD, Dr) Then PushI O, Ix
    Ix = Ix + 1
Next
If Si(O) < Si(DupD) * 2 Then Stop
RowIxAyzDupzDry = O
End Function
Private Function DryzDup(Dry()) As Variant()
If Si(Dry) = 0 Then Exit Function
Dim Dr
For Each Dr In GpCntDry(Dry)
    If Dr(0) > 1 Then
        PushI DryzDup, AyeFstEle(Dr)
    End If
Next
End Function
Function GpCntDry(Dry()) As Variant()
#If True Then
    GpCntDry = GpCntDryzSlow(Dry)
#Else
    GpCntDry = GpCntDryzQuick(Dry)
#End If
End Function
Private Function GpCntDryzQuick(Dry()) As Variant()
End Function

Private Function GpCntDryzSlow(Dry()) As Variant()
If Si(Dry) = 0 Then Exit Function
Dim OKeyDry(), OCnt&(), Dr
    Dim LasIx&: LasIx = Si(Dry(0))
    Dim J&
    For Each Dr In Dry
        If J Mod 50 = 0 Then Debug.Print J;
        If J Mod 500 = 0 Then Debug.Print
        J = J + 1
        With IxzDryDr(OKeyDry, Dr)
            Select Case .Som
            Case True: OCnt(.Lng) = OCnt(.Lng) + 1
            Case Else: PushI OKeyDry, Dr: PushI OCnt, 1
            End Select
        End With
    Next
    If Si(OKeyDry) <> Si(OCnt) Then Thw CSub, "Si Diff", "OKeyDry-Si OCnt-Si", Si(OKeyDry), Si(OCnt)
For J = 0 To UB(OCnt)
    PushI GpCntDryzSlow, AyAdd(Array(OCnt(J)), OKeyDry(J)) '<===========
Next
End Function

Private Function IxzDryDr(Dry(), Dr) As LngRslt
Dim IDr, Ix&
For Each IDr In Itr(Dry)
    If IsEqAy(IDr, Dr) Then IxzDryDr = LngRslt(Ix): Exit Function
    Ix = Ix + 1
Next
End Function
Private Sub Z_DrswDup()
Dim A As Drs, FF$, Act As Drs
GoSub T0
Exit Sub
T0:
    Set A = Drs("A B C", Av(Av(1, 2, "xxx"), Av(1, 2, "yyyy"), Av(1, 2), Av(1), Av(Empty, 2)))
    FF = "A B"
    GoTo Tst
Tst:
    Set Act = DrswDup(A, FF)
    DmpDrs Act
    Return
End Sub
'======================================================================
Private Function RowIxAyzDupzDryColIx(Dry(), ColIx&) As Long()
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
For Each I In O.Par.Itms
    PushIAy RowIxAyzDupzDryColIx, O.ParChd(I).Av
Next
End Function

Private Sub Z_RowIxAyzDupzDryColIx()
Dim Dry(), ColIx&, Act&(), Ept&()
GoSub T0
Exit Sub
T0:
    ColIx = 0
    Dry = Array(Array(1, 2, 3, 4), Array(1, 2, 3), Array(2, 4, 3))
    Ept = LngAy(0, 1)
    GoTo Tst
Tst:
    Act = RowIxAyzDupzDryColIx(Dry, ColIx)
    If Not IsEqAy(Act, Ept) Then Stop
    C
    Return
End Sub

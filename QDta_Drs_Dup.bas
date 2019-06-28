Attribute VB_Name = "QDta_Drs_Dup"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_DDup."
Private Const Asm$ = "QDta"
Function SelDrsExlCC(A As Drs, ExlCCLik$) As Drs
Dim LikC
For Each LikC In SyzSS(ExlCCLik)
'    MinusAy(
Next
End Function
Function DeDup(A As Drs, FF$) As Drs
Dim Rxy&(): Rxy = RxyzDup(A, FF)
DeDup = DeRxy(A, Rxy)
End Function

Function DwCeqC(A As Drs, CC$) As Drs
Dim Dr, C1&, C2&
AsgIx A, CC, C1, C2
For Each Dr In Itr(A.Dy)
    If Dr(C1) = Dr(C2) Then
        PushI DwCeqC.Dy, Dr
    End If
Next
DwCeqC.Fny = A.Fny
End Function

Function DeCeqC(A As Drs, CC$) As Drs
Dim Dr, C1&, C2&
AsgIx A, CC, C1, C2
For Each Dr In Itr(A.Dy)
    If Dr(C1) <> Dr(C2) Then
        PushI DeCeqC.Dy, Dr
    End If
Next
DeCeqC.Fny = A.Fny
End Function

Function RxyzDup(A As Drs, FF$) As Long()
Dim Fny$(): Fny = TermAy(FF)
If Si(Fny) = 1 Then
    RxyzDup = IxyzDup(ColzDrs(A, Fny(0)))
    Exit Function
End If
Dim ColIxy&(): ColIxy = Ixy(A.Fny, Fny)
Dim Dy(): Dy = SelDy(A.Dy, ColIxy)
RxyzDup = RxyzDupDy(Dy)
End Function

Private Function RxyzDupDy(Dy()) As Long()
Dim DupD(): DupD = DywDup(Dy)
Dim Dr, Ix&, O&()
For Each Dr In Dy
    If HasDr(DupD, Dr) Then PushI O, Ix
    Ix = Ix + 1
Next
If Si(O) < Si(DupD) * 2 Then Stop
RxyzDupDy = O
End Function

Function DywDup(Dy()) As Variant()
If Si(Dy) = 0 Then Exit Function
Dim Dr
For Each Dr In GRxyzCyCnt(Dy)
    If Dr(0) > 1 Then
        PushI DywDup, AeFstEle(Dr)
    End If
Next
End Function

Function GRxyzCyCnt(Dy()) As Variant()
#If True Then
    GRxyzCyCnt = GRxyzCyCntzSlow(Dy)
#Else
    GRxyzCyCnt = GRxyzCyCntzQuick(Dy)
#End If
End Function

Private Function GRxyzCyCntzQuick(Dy()) As Variant()
End Function

Private Function GRxyzCyCntzSlow(Dy()) As Variant()
If Si(Dy) = 0 Then Exit Function
Dim OKeyDy(), OCnt&(), Dr
    Dim LasIx&: LasIx = Si(Dy(0))
    Dim J&
    For Each Dr In Dy
        If J Mod 500 = 0 Then Debug.Print "GRxyzCyCntzSlow"
        If J Mod 50 = 0 Then Debug.Print J;
        J = J + 1
        With IxOptzDyDr(OKeyDy, Dr)
            Select Case .Som
            Case True: OCnt(.Lng) = OCnt(.Lng) + 1
            Case Else: PushI OKeyDy, Dr: PushI OCnt, 1
            End Select
        End With
    Next
    If Si(OKeyDy) <> Si(OCnt) Then Thw CSub, "Si Diff", "OKeyDy-Si OCnt-Si", Si(OKeyDy), Si(OCnt)
For J = 0 To UB(OCnt)
    PushI GRxyzCyCntzSlow, AddAy(Array(OCnt(J)), OKeyDy(J)) '<===========
Next
End Function

Private Function IxOptzDyDr(Dy(), Dr) As LngOpt
Dim IDr, Ix&
For Each IDr In Itr(Dy)
    If IsEqAy(IDr, Dr) Then IxOptzDyDr = SomLng(Ix): Exit Function
    Ix = Ix + 1
Next
End Function
Private Sub Z_DwDup()
Dim A As Drs, FF$, Act As Drs
GoSub T0
Exit Sub
T0:
    A = DrszFF("A B C", Av(Av(1, 2, "xxx"), Av(1, 2, "eyey"), Av(1, 2), Av(1), Av(Empty, 2)))
    FF = "A B"
    GoTo Tst
Tst:
    Act = DwDup(A, FF)
    DmpDrs Act
    Return
End Sub
'======================================================================
Private Function RxyzDupDyColIx(Dy(), ColIx&) As Long()
Dim D As New Dictionary, FstIx&, V, O As New Rel, Ix&, I
For Ix = 0 To UB(Dy)
    V = Dy(Ix)(ColIx)
    If D.Exists(V) Then
        O.PushParChd V, D(V)
        O.PushParChd V, Ix
    Else
        D.Add V, Ix
    End If
Next
For Each I In O.SetOfPar.Itms
    PushIAy RxyzDupDyColIx, O.ParChd(I).Av
Next
End Function

Private Sub Z_RxyzDupDyColIx()
Dim Dy(), ColIx&, Act&(), Ept&()
GoSub T0
Exit Sub
T0:
    ColIx = 0
    Dy = Array(Array(1, 2, 3, 4), Array(1, 2, 3), Array(2, 4, 3))
    Ept = LngAp(0, 1)
    GoTo Tst
Tst:
    Act = RxyzDupDyColIx(Dy, ColIx)
    If Not IsEqAy(Act, Ept) Then Stop
    C
    Return
End Sub

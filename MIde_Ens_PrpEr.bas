Attribute VB_Name = "MIde_Ens_PrpEr"
Option Explicit
Const CMod$ = "MIde_Ens_PrpEr."
Private Const LinOfLblX$ = "X: Debug.Print CSub & "".PrpEr["" &  Err.Description & ""]"""
Public Type TopRmkLyAaaMthLy
    TopRmkLy() As String
    MthLy() As String
End Type
Private Function HasLinExitAaaLblX(MthLy$(), LinOfExit) As Boolean
Dim U&: U = UB(MthLy): If U < 2 Then Exit Function
If MthLy(U - 1) <> LinOfLblX Then Exit Function
If MthLy(U - 2) <> MthExitLin(MthLy(0)) Then Exit Function
End Function

Private Function InsLinExitAaaLblX(MthLy$(), LinOfExit$) As String()
InsLinExitAaaLblX = AyInsAyAt(MthLy, Sy(LinOfExit, LinOfLblX), UB(MthLy))
End Function

Private Function EnsLinExitAaaLblX(MthLy$(), LinOfExit$) As String()
If HasLinExitAaaLblX(MthLy, LinOfExit) Then EnsLinExitAaaLblX = MthLy: Exit Function
Dim O$():
O = AyeEle(MthLy, LinOfExit)
O = AyeEle(MthLy, LinOfLblX)
EnsLinExitAaaLblX = InsLinExitAaaLblX(O, LinOfExit)
End Function

Function AyEnsEle(Ay, NeedDlt As Boolean, NeedIns As Boolean, NewEle, DltIx&, InsIx&)
Dim O
Select Case True
Case NeedDlt And NeedIns: Asg NewEle, O(InsIx)
Case NeedDlt:             O = AyeEleAt(Ay, DltIx)
Case NeedIns:             O = AyInsEle(Ay, InsIx)
End Select
End Function
Private Function RmvOnErGoNonX(MthLy$()) As String()
Dim J&, I&
For J = 0 To UB(MthLy)
    If HasPfx(MthLy(J), "On Error Goto") Then
        For I = J + 1 To UB(MthLy)
            PushI RmvOnErGoNonX, MthLy(I)
        Next
        Exit Function
    End If
    PushI RmvOnErGoNonX, MthLy(J)
Next
End Function

Private Function InsOnErGoX(MthLy$()) As String()
InsOnErGoX = AyInsEle(MthLy, "On Error Goto X", NxtSrcIx(MthLy))
End Function

Private Function EnsLinzOnEr(MthLy$()) As String()
Dim O$()
O = RmvOnErGoNonX(MthLy)
EnsLinzOnEr = InsOnErGoX(O)
End Function
Function MthEndLin$(MthLin)
MthEndLin = MthXXXLin(MthLin, "End")
End Function
Function MthExitLin$(MthLin)
MthExitLin = MthXXXLin(MthLin, "Exit")
End Function
Private Function MthXXXLin$(MthLin, XXX$)
Dim X$: X = MthKd(MthLin): If X = "" Then Thw CSub, "Given Lin is not MthLin", "Lin", MthLin
MthXXXLin = XXX & " " & X
End Function

Private Function IxOfExit&(MthLy$())
Dim J&, LinOfExit$
LinOfExit = "Exit " & MthKd(MthLy(0))
For J = 0 To UB(MthLy)
    If MthLy(J) = LinOfExit Then IxOfExit = J: Exit Function
Next
IxOfExit = -1
End Function

Private Function IxOfLblX&(MthLy$())
Dim J&, L$
For J = 0 To UB(MthLy)
    If HasPfx(MthLy(J), "X: Debug.Print") Then IxOfLblX = J: Exit Function
Next
IxOfLblX = -1
End Function

Private Function IxOfOnEr&(PurePrpLy$())
Dim J&
For J = 0 To UB(PurePrpLy)
    If HasPfx(PurePrpLy(J), "On Error Goto X") Then IxOfOnEr = J: Exit Function
Next
IxOfOnEr = -1
End Function

Private Sub Z_EnsPrpOnerzSrc()
Dim Src$()
Const TstId% = 2
GoSub ZZ
'GoSub T1
Exit Sub
T1:
    Src = TstLy(TstId, CSub, 1, "Pm-Src", IsEdt:=False)
    Ept = TstLy(TstId, CSub, 1, "Ept", IsEdt:=False)
    GoTo Tst
    Return
Tst:
    Act = EnsPrpOnErzSrc(Src)
    C
    Return
ZZ:
    Src = CurSrc
    Vc EnsPrpOnErzSrc(Src)
    Return
End Sub
Private Function EnsPrpOnErzSrc(Src$()) As String()
If Si(Src) = 0 Then Exit Function
Dim D As Dictionary: Set D = MthDic(Src, WiTopRmk:=True)
Dim MthNm, MthLin$, MthLyWiTopRmk$()
Dim O$(): Erase O
If D.Exists("*Dcl") Then
    PushIAy O, SplitCrLf(D("*Dcl"))
    D.Remove "*Dcl"
End If

For Each MthNm In D.Keys
    MthLyWiTopRmk = SplitCrLf(D(MthNm))
    PushI O, ""
    With TopRmkLyAaaMthLy(MthLyWiTopRmk)
        PushIAy O, .TopRmkLy
        Dim MthLy$(): MthLy = .MthLy
    End With
    MthLin = ContLin(MthLy)
    If IsPurePrpLin(MthLin) Then
        PushIAy O, EnsOnEr(MthLy)
    Else
        PushIAy O, MthLy
    End If
Next
If Si(O) = 0 Then Thw CSub, "O$() should have something"
EnsPrpOnErzSrc = O
End Function

Function TopRmkLyAaaMthLy(MthLyWiTopRmk$()) As TopRmkLyAaaMthLy
Dim Lin, J&, TopRmkLy$(), MthLy$()
For Each Lin In MthLyWiTopRmk
    If FstChr(Lin) = "'" Then
        PushI TopRmkLy, Lin
    Else
        MthLy = AywFmIx(MthLyWiTopRmk, J)
        TopRmkLyAaaMthLy.TopRmkLy = TopRmkLy
        TopRmkLyAaaMthLy.MthLy = MthLy
        Exit Function
    End If
Next
Thw CSub, "MthLyWiTopRmk is invalid, it does not have non remark line", "MthLyWiTopRmk", MthLyWiTopRmk
End Function

Sub Z_EnsprpOnErDiczPj()
BrwDic EnsprpOnErDiczPj(CurPj)
End Sub

Function EnsprpOnErDiczPj(A As VBProject) As Dictionary
Dim N$, Old$, S$(), C As VBComponent
Dim O As New Dictionary
For Each C In A.VBComponents
    S = Src(C.CodeModule)
    If Si(S) > 0 Then
        Old = JnCrLf(S)
        N = JnCrLf(EnsPrpOnErzSrc(S))
        If Old <> N Then
            O.Add C.Name, N
        End If
    End If
Next
Set EnsprpOnErDiczPj = O
End Function

Private Sub Z_EnsOnEr()
Const TstId% = 1
Const CSub$ = CMod & "Z_EnsOnEr"
Dim MthLy$(), Cas$, IsEdt As Boolean
GoSub T1
GoSub T2
Exit Sub
T1:
    Cas = "1"
    WrtTstPth TstId, CSub
    MthLy = TstLy(TstId, CSub, Cas, "Pm-MthLy", IsEdt:=False)
    Ept = TstLy(TstId, CSub, Cas, "Ept", IsEdt:=False)
    GoTo Tst
T2:
    Cas = "2"
    WrtTstPth TstId, CSub
    MthLy = TstLy(TstId, CSub, Cas, "Pm-MthLy", IsEdt:=False)
    Ept = TstLy(TstId, CSub, Cas, "Ept", IsEdt:=False)
    GoTo Tst
Tst:
    Act = EnsOnEr(MthLy)
    C
    Return
End Sub

Private Sub Z_IsSngLinMth()
Dim MthLy$()
GoSub T1
Exit Sub
T1:
    MthLy = Sy("Function AA():End Function")
    Ept = True
    GoTo Tst
Tst:
    Act = IsSngLinMth(MthLy)
    C
    Return
End Sub

Function IsSngLinMth(MthLy$()) As Boolean
If Si(MthLy) <> 1 Then Exit Function
IsSngLinMth = HasSubStr(MthLy(0), "End " & MthKd(MthLy(0)))
End Function

Private Function EnsOnEr(MthLy$()) As String()
If IsSngLinMth(MthLy) Then EnsOnEr = MthLy: Exit Function
Dim O$(), LinOfExit$: LinOfExit = MthExitLin(MthLy(0))
O = EnsLinExitAaaLblX(MthLy, LinOfExit)
EnsOnEr = EnsLinzOnEr(O)
End Function

Private Function RmvPrpOnErzPurePrpLy(PurePrpLy$()) As String()
Dim L&(): L = LngAy( _
IxOfExit(PurePrpLy), _
IxOfOnEr(PurePrpLy), _
IxOfLblX(PurePrpLy))
RmvPrpOnErzPurePrpLy = CvSy(AyeIxAy(PurePrpLy, L))
End Function
Private Function RmvPrpOnzSrc(Src$()) As String()

End Function
Private Sub RmvPrpOnErzMd(A As CodeModule)
RplMd A, JnCrLf(RmvPrpOnzSrc(Src(A)))
End Sub

Sub RmvPrpOnErOfMd()
RmvPrpOnErzMd CurMd
End Sub

Sub EnsPrpOnErzMd(A As CodeModule)
RplMd A, JnCrLf(EnsPrpOnErzSrc(Src(A)))
End Sub

Sub EnsPrpOnErOfMd()
EnsPrpOnErzMd CurMd
End Sub





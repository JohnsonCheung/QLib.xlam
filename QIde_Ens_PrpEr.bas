Attribute VB_Name = "QIde_Ens_PrpEr"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Ens_PrpEr."
Private Const LinOfLblX$ = "X: Debug.Print CSub & "".PrpEr["" &  Err.Description & ""]"""
Private Function HasLinExitAndLblX(MthLy$(), LinOfExit) As Boolean
Dim U&: U = UB(MthLy): If U < 2 Then Exit Function
If MthLy(U - 1) <> LinOfLblX Then Exit Function
If MthLy(U - 2) <> MthExitLin(MthLy(0)) Then Exit Function
End Function

Private Function InsLinExitAndLblX(MthLy$(), LinOfExit$) As String()
InsLinExitAndLblX = AyInsAyAt(MthLy, Sy(LinOfExit, LinOfLblX), UB(MthLy))
End Function

Private Function EnsLinExitAndLblX(MthLy$(), LinOfExit$) As String()
If HasLinExitAndLblX(MthLy, LinOfExit) Then EnsLinExitAndLblX = MthLy: Exit Function
Dim O$():
O = AyeEle(MthLy, LinOfExit)
O = AyeEle(MthLy, LinOfLblX)
EnsLinExitAndLblX = InsLinExitAndLblX(O, LinOfExit)
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
InsOnErGoX = AyInsEle(MthLy, "On Error Goto X", NxtIxzSrc(MthLy))
End Function

Private Function EnsLinzOnEr(MthLy$()) As String()
Dim O$()
O = RmvOnErGoNonX(MthLy)
EnsLinzOnEr = InsOnErGoX(O)
End Function

Function MthExitLin$(MthLin)
MthExitLin = MthXXXLin(MthLin, "Exit")
End Function
Private Function MthXXXLin(MthLin, XXX$)
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

Private Sub Z_EnsPrpOnerzS()
Dim Src$()
Const TstId& = 2
GoSub ZZ
'GoSub T1
Exit Sub
T1:
    Src = TstLy(TstId, CSub, 1, "Pm-Src", IsEdt:=False)
    Ept = TstLy(TstId, CSub, 1, "Ept", IsEdt:=False)
    GoTo Tst
    Return
Tst:
    Act = EnsPrpOnErzS(Src)
    C
    Return
ZZ:
    Src = CSrc
    Vc EnsPrpOnErzS(Src)
    Return
End Sub
Function EnsPrpOnErzS(Src$()) As String()

End Function
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

Private Sub RmvPrpOnErzM(M As CodeModule, Optional Rpt As EmRpt)
'Dim L&(): L = LngAp( _
IxOfExit(PurePrpLy), _
IxOfOnEr(PurePrpLy), _
IxOfLblX(PurePrpLy))
'RmvPrpOnErzPurePrpLy = CvSy(AyeIxy(PurePrpLy, L))
End Sub

Sub RmvPrpOnErM()
RmvPrpOnErzM CMd
End Sub

Sub EnsPrpOnErzM(M As CodeModule)
'RplMd A, JnCrLf(EnsPrpOnErzS(Src(A)))
End Sub

Sub EnsPrpOnErM()
EnsPrpOnErzM CMd
End Sub

Private Sub ZZZ()
QIde_Ens_PrpEr:
End Sub

Attribute VB_Name = "MxEnsPrpEr"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxEnsPrpEr."
Const LinOfLblX$ = "X: Debug.Print CSub & "".PrpEr["" &  Err.Description & ""]"""
Function HasLinExitAndLblX(Mthly$(), LinOfExit) As Boolean
Dim U&: U = UB(Mthly): If U < 2 Then Exit Function
If Mthly(U - 1) <> LinOfLblX Then Exit Function
If Mthly(U - 2) <> MthExitLin(Mthly(0)) Then Exit Function
End Function

Function InsLinExitAndLblX(Mthly$(), LinOfExit$) As String()
InsLinExitAndLblX = InsAy(Mthly, Sy(LinOfExit, LinOfLblX), UB(Mthly))
End Function

Function EnsLinExitAndLblX(Mthly$(), LinOfExit$) As String()
If HasLinExitAndLblX(Mthly, LinOfExit) Then EnsLinExitAndLblX = Mthly: Exit Function
Dim O$():
O = AeEle(Mthly, LinOfExit)
O = AeEle(Mthly, LinOfLblX)
EnsLinExitAndLblX = InsLinExitAndLblX(O, LinOfExit)
End Function

Function RmvOnErGoNonX(Mthly$()) As String()
Dim J&, I&
For J = 0 To UB(Mthly)
    If HasPfx(Mthly(J), "On Error Goto") Then
        For I = J + 1 To UB(Mthly)
            PushI RmvOnErGoNonX, Mthly(I)
        Next
        Exit Function
    End If
    PushI RmvOnErGoNonX, Mthly(J)
Next
End Function

Function InsOnErGoX(Mthly$()) As String()
InsOnErGoX = InsEle(Mthly, "On Error Goto X", NxtIxzSrc(Mthly))
End Function

Function EnsLinzOnEr(Mthly$()) As String()
Dim O$()
O = RmvOnErGoNonX(Mthly)
EnsLinzOnEr = InsOnErGoX(O)
End Function

Function MthExitLin$(MthLin)
MthExitLin = MthXXXLin(MthLin, "Exit")
End Function
Function MthXXXLin(MthLin, XXX$)
Dim X$: X = MthKd(MthLin): If X = "" Then Thw CSub, "Given Lin is not MthLin", "Lin", MthLin
MthXXXLin = XXX & " " & X
End Function

Function IxOfExit&(Mthly$())
Dim J&, LinOfExit$
LinOfExit = "Exit " & MthKd(Mthly(0))
For J = 0 To UB(Mthly)
    If Mthly(J) = LinOfExit Then IxOfExit = J: Exit Function
Next
IxOfExit = -1
End Function

Function IxOfLblX&(Mthly$())
Dim J&, L$
For J = 0 To UB(Mthly)
    If HasPfx(Mthly(J), "X: Debug.Print") Then IxOfLblX = J: Exit Function
Next
IxOfLblX = -1
End Function

Function IxOfOnEr&(PurePrpLy$())
Dim J&
For J = 0 To UB(PurePrpLy)
    If HasPfx(PurePrpLy(J), "On Error Goto X") Then IxOfOnEr = J: Exit Function
Next
IxOfOnEr = -1
End Function

Sub Z_EnsPrpOnerzS()
Dim Src$()
Const TstId& = 2
GoSub Z
'GoSub T1
Exit Sub
T1:
    Src = TstLy(TstId, CSub, 1, "Pm-Src", IsEdt:=False)
    Ept = TstLy(TstId, CSub, 1, "Ept", IsEdt:=False)
    GoTo Tst
    Return
Tst:
    Act = EnsPrpOnEoS(Src)
    C
    Return
Z:
    Src = CSrc
    Vc EnsPrpOnEoS(Src)
    Return
End Sub

Sub Z_IsMthLinSngL()
Dim L
GoSub T1
Exit Sub
T1:
    L = Sy("Function AA():End Function")
    Ept = True
    GoTo Tst
Tst:
    Act = IsMthLinSngL(L)
    C
    Return
End Sub
Sub RmvPrpOnEoM(M As CodeModule, Optional Upd As EmUpd)
'Dim L&(): L = LngAp( _
IxOfExit(PurePrpLy), _
IxOfOnEr(PurePrpLy), _
IxOfLblX(PurePrpLy))
'RmvPrpOnEoPurePrpLy = CvSy(AeIxy(PurePrpLy, L))
End Sub

Sub RmvPrpOnErM()
RmvPrpOnEoM CMd
End Sub

Sub EnsPrpOnEoM(M As CodeModule)
'RplMd A, JnCrLf(EnsPrpOnEoS(Src(A)))
End Sub

Function EnsPrpOnEoS(Src$()) As String()
'RplMd A, JnCrLf(EnsPrpOnEoS(Src(A)))
End Function

Sub EnsPrpOnErM()
EnsPrpOnEoM CMd
End Sub


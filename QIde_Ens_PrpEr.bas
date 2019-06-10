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

Private Function SrcOptOfEnsprpOnEr(Src$()) As LyOpt ' Ret None is no change in Src
If Si(Src) = 0 Then Exit Function
Dim D As Dictionary: 'Set D = MthDic(Src)
Dim Mthn, MthLin, MthLyWiTopRmk$()
Dim O$(): Erase O
If D.Exists("*Dcl") Then
    PushIAy O, SplitCrLf(D("*Dcl"))
    D.Remove "*Dcl"
End If

For Each Mthn In D.Keys
    MthLyWiTopRmk = SplitCrLf(D(Mthn))
    PushI O, ""
'    With TopRmkLyAndMthLy(MthLyWiTopRmk)
'        PushIAy O, .TopRmkLy
'        Dim MthLy$(): MthLy = .MthLy
'    End With
'    MthLin = ContLin(MthLy)
    If IsPurePrpLin(MthLin) Then
'        PushIAy O, MthLyOfEnsonEr(MthLy)
    Else
'        PushIAy O, MthLy
    End If
Next
If Si(O) = 0 Then Thw CSub, "O$() should have something"
SrcOptOfEnsprpOnEr = SomLy(O)
End Function

Private Function EnsPrpOnErzS(Src$()) As String()
If Si(Src) = 0 Then Exit Function
Dim D As Dictionary: 'Set D = MthDic(Src)
Dim Mthn, MthLin, MthLy$()
Dim O$(): Erase O
If D.Exists("*Dcl") Then
    PushIAy O, SplitCrLf(D("*Dcl"))
    D.Remove "*Dcl"
End If

For Each Mthn In D.Keys
    MthLy = SplitCrLf(D(Mthn))
    PushI O, ""
'    With TopRmkLyAndMthLy(MthLyWiTopRmk)
'        PushIAy O, .TopRmkLy
        'Dim MthLy$() ': MthLy = .MthLy
'    End With
'    MthLin = ContLin(MthLy)
    If IsPurePrpLin(MthLin) Then
'        PushIAy O, MthLyOfEnsonEr(MthLy)
    Else
        PushIAy O, MthLy
    End If
Next
If Si(O) = 0 Then Thw CSub, "O$() should have something"
EnsPrpOnErzS = O
End Function

Private Sub Z_EnsprpOnErDiczP()
BrwDic EnsprpOnErDiczP(CPj)
End Sub

Function EnsprpOnErDicP() As Dictionary
Set EnsprpOnErDicP = EnsprpOnErDiczP(CPj)
End Function

Function EnsprpOnErDiczP(P As VBProject) As Dictionary
Dim N$, Old$, S$(), C As VBComponent
Dim O As New Dictionary
For Each C In P.VBComponents
    S = Src(C.CodeModule)
    If Si(S) > 0 Then
        Old = JnCrLf(S)
        N = JnCrLf(EnsPrpOnErzS(S))
        If Old <> N Then
            O.Add C.Name, N
        End If
    End If
Next
Set EnsprpOnErDiczP = O
End Function

Private Sub Z_MthLyOfEnsOnEr()
Const TstId& = 1
Const CSub$ = CMod & "Z_MthLyOfEnsOnEr"
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
'    Act = MthLyOfEnsonEr(MthLy)
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

Private Function MthLyRsltOfEnsonEr(MthLy$()) As LyRslt
'If IsSngLinMth(MthLy) Then MthLyOfEnsonEr = MthLy: Exit Function
Dim LinOfExit$: LinOfExit = MthExitLin(MthLy(0))
Dim O$(): O = EnsLinExitAndLblX(MthLy, LinOfExit)
'MthLyRsltOfEnsonEr = EnsLinzOnEr(O)
End Function

Private Function RmvPrpOnErzPurePrpLy(PurePrpLy$()) As String()
Dim L&(): L = LngAp( _
IxOfExit(PurePrpLy), _
IxOfOnEr(PurePrpLy), _
IxOfLblX(PurePrpLy))
RmvPrpOnErzPurePrpLy = CvSy(AyeIxy(PurePrpLy, L))
End Function
Private Function RmvPrpOnzS(Src$()) As String()

End Function
Private Sub RmvPrpOnErzMd(M As CodeModule)
'RplMd A, JnCrLf(RmvPrpOnzS(Src(A)))
End Sub

Sub RmvPrpOnErM()
RmvPrpOnErzMd CMd
End Sub

Sub EnsPrpOnErzMd(M As CodeModule)
'RplMd A, JnCrLf(EnsPrpOnErzS(Src(A)))
End Sub

Sub EnsPrpOnErM()
EnsPrpOnErzMd CMd
End Sub

Private Sub ZZZ()
QIde_Ens_PrpEr:
End Sub



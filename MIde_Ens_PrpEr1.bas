Attribute VB_Name = "MIde_Ens_PrpEr1"
Option Explicit
Const CMod$ = "MIde_Ens_PrpEr."

Private Sub EnsLinzExit(OMthLy$())
Const CSub$ = CMod & "EnsLinzExit"
Dim L&
L = IxOfInsExit(OMthLy)
If L = 0 Then Exit Sub
OMthLy = AyInsItm(OMthLy, "Exit Property", L)
End Sub

Private Sub EnsLinLblX(OMthLy$())
Const CSub$ = CMod & "EnsLinzLblX"
Dim E$, Ix&, ActLblXLin$, EndPrpLno&
E = LinzLblX
Ix = IxOfLblX(OMthLy)
If Ix <> 0 Then
    ActLblXLin = OMthLy(Ix)
End If
If E <> ActLblXLin Then
    If Ix = 0 Then
        EndPrpLno = UB(OMthLy)
        If EndPrpLno = 0 Then Stop
        'OMthLy.InsertLines EndPrpLno, E
        Inf CSub, "Inserted [at] with [line]", EndPrpLno, E
    Else
        'OMthLy.ReplaceLine L, E
        Inf CSub, "Replaced [at] with [line]", Ix, E
    End If
End If
End Sub

Private Sub EnsLinzOnEr(OMthLy$())
Const CSub$ = CMod & "EnsLinzOnEr"
Dim L&
'L = IxOfOnEr(A)
If L <> 0 Then Exit Sub
'A.InsertLines PrpLinix + 1, "On Error Goto X"
'If Trc Then Msg CSub, "Exit Property is inserted [at]", L
End Sub

Private Function IxOfExit&(MthLy$())
Dim J&
For J = 0 To UB(MthLy)
    If MthLy(J) = "Exit Property" Then IxOfExit = J: Exit Function
Next
IxOfExit = -1
End Function

Private Function IxOfInsExit&(MthLy$())
'If IxOfExit(A) <> 0 Then Exit Function
Dim L%
'L = IxOfLblX(A)
If L = 0 Then Stop
IxOfInsExit = L
End Function

Private Function LinzLblX$()
LinzLblX = "X: Debug.Print CSub & "".PrpEr["" &  Err.Description & ""]"""
End Function

Private Function IxOfLblX&(MthLy$())
Dim J&, L$
'For J = PrpLinix + 1 To A.CountOfLines
'    L = A.Lines(J, 1)
    If HasPfx(L, "X: Debug.Print") Then IxOfLblX = J: Exit Function
    If HasPfx(L, "End Property") Then Exit Function
'Next
Stop
End Function

Private Function IxOfOnEr&(MthLy$())
Dim J&
For J = 0 To UB(MthLy)
    If HasPfx(MthLy(J), "On Error Goto X") Then IxOfOnEr = J: Exit Function
Next
IxOfOnEr = -1
End Function
Private Sub Z_SrcEnsPrpOnEr()
Dim Src$()
GoSub ZZ
GoSub T1
Exit Sub
T1:
    Return
Tst:
    Act = SrcEnsPrpOnEr(Src)
    C
    Return
ZZ:
    Vc SrcEnsPrpOnEr(Src)
    Return
End Sub
Private Function SrcEnsPrpOnEr(Src$()) As String()
Dim D As Dictionary: Set D = MthDic(Src)
Dim MthNm
For Each MthNm In D.Keys
    Dim MthLy$(): MthLy = SplitCrLf(D(MthNm))
    If MthNm = "*Dcl" Then
        PushIAy SrcEnsPrpOnEr, MthLy
    Else
        PushIAy SrcEnsPrpOnEr, MthLyEnsPrpOnEr(MthLy)
    End If
Next
End Function

Function SyPairOfTopRmkOooMthLy(MthLyWiTopRmk$()) As SyPair
Dim Lin, J&, OTopRmk$(), OMthLy$()
For Each Lin In MthLyWiTopRmk
    If FstChr(Lin) = "'" Then
        PushI OTopRmk, Lin
    Else
        OMthLy = AywFmIx(MthLyWiTopRmk, J)
'        SyPairOfTopRmkAdMthLy = SyPair(OTopRmk, OMthLy)
        Exit Function
    End If
Next
Thw CSub, "MthLyWiTopRmk is invalid, it does not have non remark line", "MthLyWiTopRmk", MthLyWiTopRmk
End Function

Private Function MthLyEnsPrpOnEr(MthLyWiTopRmk$()) As String()
Dim T$(), Ly$()
    With SyPairOfTopRmkOooMthLy(MthLyWiTopRmk): T = .Sy1: Ly = .Sy2: End With
If HasSubStr(Ly(0), "End Property") Then MthLyEnsPrpOnEr = MthLyWiTopRmk: Exit Function
MthLyEnsPrpOnEr = T
EnsLinLblX Ly
EnsLinzExit Ly
EnsLinzOnEr Ly
PushIAy MthLyEnsPrpOnEr, Ly
End Function
Private Function MthLyRmvPrpOnEr(MthLy$()) As String()
Dim L&(): L = LngAy( _
IxOfExit(MthLy), _
IxOfOnEr(MthLy), _
IxOfLblX(MthLy))
MthLyRmvPrpOnEr = CvSy(AyeIxAy(L, L))
End Function
Private Function RmvPrpOnErzSrc(Src$()) As String()

End Function
Private Sub RmvPrpOnErzMd(A As CodeModule)
MdRpl A, JnCrLf(RmvPrpOnErzSrc(Src(A)))
End Sub

Sub RmvPrpOnErOfMd()
'RmvPrpOnErzMd CurMd
End Sub

Sub EnsPrpOnErzMd(A As CodeModule)
MdRpl A, JnCrLf(SrcEnsPrpOnEr(Src(A)))
End Sub

Sub EnsPrpOnErOfMd1()
EnsPrpOnErzMd CurMd
End Sub

Private Sub Z_EnsPrpOnErzMd()
'EnsPrpOnErzMd ZZMd
End Sub




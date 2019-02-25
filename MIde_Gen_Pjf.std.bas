Attribute VB_Name = "MIde_Gen_Pjf"
Option Explicit
Function DistPth$(SrcPth)
ThwNotSrcPth SrcPth
DistPth = PthEns(SiblingPth(SrcPth, "Dist"))
End Function

Function DistFba$(SrcPth)
DistFba = DistPth(SrcPth) & FbaFnzSrcPth(SrcPth)
End Function

Function DistFxa$(SrcPth)
DistFxa = DistPth(SrcPth) & FxaFnzSrcPth(SrcPth)
End Function

Private Sub Z_DistPjf()
Dim Pth
For Each Pth In SrcPthAyzExpgzInst
    Debug.Print Pth
    Debug.Print DistFba(Pth)
    Debug.Print DistFxa(Pth)
    Debug.Print
Next
End Sub

Private Function PjFnzSrcPth$(SrcPth)
Dim R$: R$ = SrcRoot(SrcPth): If R = "" Then Exit Function
PjFnzSrcPth = RmvExt(Fdr(R))
End Function

Private Function FxaFnzSrcPth$(SrcPth)
FxaFnzSrcPth = RplExt(PjFnzSrcPth(SrcPth), FxaExt)
End Function

Private Function SrcRoot$(SrcPth)
Dim P$: P$ = ParPth(SrcPth)
If IsSrcRoot(P) Then SrcRoot = P: Exit Function
Dim F$: F$ = Fdr(P)
If Not IsDteTimStr(F) Then Exit Function
Dim P1$: P1$ = ParPth(P)
If IsSrcRoot(P1) Then SrcRoot = P1
End Function

Private Function IsSrcRoot(Pth) As Boolean
If Not IsPth(Pth) Then Exit Function
Dim F$: F$ = Fdr(Pth)
Select Case True
Case Ext(F) <> ".Src", Not IsPjf(RmvExt(F))
Case Else: IsSrcRoot = True
End Select
End Function

Private Sub ThwNotSrcRoot(Pth)
If Not IsSrcRoot(Pth) Then Err.Raise 1, , "Not SourceRoot(" & Pth & ")"
End Sub

Private Function FbaFnzSrcPth$(SrcPth)
FbaFnzSrcPth = RplExt(PjFnzSrcPth(SrcPth), FbaExt)
End Function

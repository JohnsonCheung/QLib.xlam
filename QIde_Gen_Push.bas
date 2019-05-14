Attribute VB_Name = "QIde_Gen_Push"
Option Explicit
Private Const CMod$ = "MIde_Gen_Push."
Private Const Asm$ = "QIde"
Sub GenPushMd()
GenPushzMd CMd
End Sub
Sub GenPushPj()
GenPushzP CPj
End Sub

Private Sub GenPushzMd(A As CodeModule)
Dim Gen$(): Gen = TyNyzGen(A) 'TyNy need to generate Push
Dim Dlt$(): Dlt = TyNyzDlt(A) ' TyNy need to delete
EnsMth A, MthDic(Gen)
'RmvMth A, MthnyzDltTyNy(Dlt)
End Sub

Sub EnsMth(A As CodeModule, MthDic As Dictionary)

End Sub
Function TynzLin$(Lin)
Dim L$: L = RmvMdy(Lin)
If Not ShfTermTy(L) Then Exit Function
TynzLin = Nm(L)
End Function

Function TyNyzM(A As CodeModule) As String()
TynyzM = TynyzS(DclLyzMd
For Each L In DclLy Itr(A)
    PushNonBlank TyNyzM, TynzLin(L)
Next
End Function

Private Function TyNyzDlt(A As CodeModule) As String()

End Function
Private Sub GenPushzP(P As VBProject)
Dim C As VBComponent
For Each C In P.VBComponents
    GenPushzMd C.CodeModule
Next
End Sub

Private Function MthDic(TyNyzGen$()) As Dictionary
End Function

Private Function MthnyzDltTyNy(TyNyzDlt$()) As String()
End Function


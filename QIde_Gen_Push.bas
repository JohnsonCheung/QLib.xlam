Attribute VB_Name = "QIde_Gen_Push"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Gen_Push."
Private Const Asm$ = "QIde"
Sub GenPushMd()
GenPushzMd CMd
End Sub
Sub GenPushPj()
GenPushzP CPj
End Sub

Private Sub GenPushzMd(M As CodeModule)
Dim Gen$(): 'Gen = TynyzGen(A) 'Tyny need to generate Push
Dim Dlt$(): Dlt = TynyzDlt(M) ' Tyny need to delete
EnsMth M, MthDic(Gen)
'RmvMth A, MthnyzDltTyny(Dlt)
End Sub

Sub EnsMth(M As CodeModule, MthDic As Dictionary)

End Sub
Function TynzLin$(Lin)
Dim L$: L = RmvMdy(Lin)
If Not ShfTermTy(L) Then Exit Function
TynzLin = Nm(L)
End Function

Function TynyzM(M As CodeModule) As String()
TynyzM = TynyzS(DclLyzM(M))
Dim L
For Each L In Itr(M)
    PushNonBlank TynyzM, TynzLin(L)
Next
End Function

Private Function TynyzDlt(M As CodeModule) As String()

End Function
Private Sub GenPushzP(P As VBProject)
Dim C As VBComponent
For Each C In P.VBComponents
    GenPushzMd C.CodeModule
Next
End Sub

Private Function MthDic(TynyzGen$()) As Dictionary
End Function

Private Function MthnyzDltTyny(TynyzDlt$()) As String()
End Function


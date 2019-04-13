Attribute VB_Name = "MIde_Gen_Push"
Option Explicit
Sub GenPushMd()
GenPushzMd CurMd
End Sub
Sub GenPushPj()
GenPushzPj CurPj
End Sub

Private Sub GenPushzMd(A As CodeModule)
Dim Gen$(): Gen = TyNyzGen(A) 'TyNy need to generate Push
Dim Dlt$(): Dlt = TyNyzDlt(A) ' TyNy need to delete
EnsMdMth A, MthDic(Gen)
RmvMth A, MthNyzDltTyNy(Dlt)
End Sub

Sub EnsMdMth(A As CodeModule, MthDic As Dictionary)

End Sub
Private Function TyNyzGen(A As CodeModule) As String()

End Function

Private Function TyNyzDlt(A As CodeModule) As String()

End Function
Private Sub GenPushzPj(A As VBProject)
Dim C As VBComponent
For Each C In A.VBComponents
    GenPushzMd C.CodeModule
Next
End Sub

Private Function MthDic(TyNyzGen$()) As Dictionary
End Function

Private Function MthNyzDltTyNy(TyNyzDlt$()) As String()
End Function


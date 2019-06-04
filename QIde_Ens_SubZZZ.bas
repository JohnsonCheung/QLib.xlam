Attribute VB_Name = "QIde_Ens_SubZZZ"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Ens_SubZZZ."
Private Const Asm$ = "QIde"
Sub EnsSubZZzP()
EnsSubZZZzP CPj
EnsPrvZzP CPj
End Sub

Sub EnsSubZZZMd()
EnsSubZZZzMd CMd 'Ensure Sub Z()
EnsPrvZzMd CMd 'Ensure all Z_XX() as Private
End Sub

Private Function SubZZZEptzMd$(M As CodeModule)
SubZZZEptzMd = SubZEptzMd(M) & vbCrLf & vbCrLf & SubZZEpt(M)
End Function

Private Sub EnsSubZZZzMd(M As CodeModule)
Ept = SubZZZEptzMd(M)
'Act = SubZZZzMd(A)
If Act = Ept Then Exit Sub
'Brw Ept
'Stop
'CmprLines Act, Ept, "Act-SubZ & SubZZ", "Ept"
RmvMth M, "Z"
RmvMth M, "ZZ"
If Ept <> "" Then
    ApdLines M, vbCrLf & Ept
End If
End Sub

Private Sub EnsSubZZZzP(P As VBProject)
Dim C As VBComponent
For Each C In P.VBComponents
    Debug.Print C.Name
    EnsSubZZZzMd C.CodeModule
Next
End Sub


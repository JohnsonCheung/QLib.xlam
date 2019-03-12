Attribute VB_Name = "MIde_Ens_SubZZZ"
Option Explicit
Sub EnsSubZZZPj()
EnsSubZZZzPj CurPj
EnsPrvZzPj CurPj
End Sub

Sub EnsSubZZZMd()
EnsSubZZZzMd CurMd 'Ensure Sub Z()
EnsPrvZzMd CurMd 'Ensure all Z_XX() as Private
End Sub

Private Function SubZZZEptzMd$(A As CodeModule)
SubZZZEptzMd = SubZEptzMd(A) & vbCrLf & vbCrLf & SubZZEpt(A)
End Function

Private Sub EnsSubZZZzMd(A As CodeModule)
Ept = SubZZZEptzMd(A)
'Act = SubZZZzMd(A)
If Act = Ept Then Exit Sub
'Brw Ept
'Stop
'CmpLines Act, Ept, "Act-SubZ & SubZZ", "Ept"
RmvMdMth A, "Z"
RmvMdMth A, "ZZ"
If Ept <> "" Then
    AppLines A, vbCrLf & Ept
End If
End Sub

Private Sub EnsSubZZZzPj(A As VBProject)
Dim C As VBComponent
For Each C In A.VBComponents
    Debug.Print C.Name
    EnsSubZZZzMd C.CodeModule
Next
End Sub


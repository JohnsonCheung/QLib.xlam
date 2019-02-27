Attribute VB_Name = "MIde_Gen_Pjf_Bas"
Option Explicit
Sub LoadBas(A As VBProject)
Dim Ay$(): Ay = BasFfnAy(SrcPth(A))
Dim BasItm
For Each BasItm In Itr(Ay)
    A.VBComponents.Import BasItm
Next
End Sub

Private Function BasFfnAy(SrcPth) As String()
Dim Ffn
For Each Ffn In FfnAy(SrcPth)
    If IsBasFfn(Ffn) Then
        PushI BasFfnAy, Ffn
    End If
Next
End Function
Private Function IsBasFfn(Ffn) As Boolean
Select Case True
Case HasSfx(Ffn, ".std.bas"), HasSfx(Ffn, ".cls.bas"): IsBasFfn = True
End Select
End Function


Attribute VB_Name = "MIde_Gen_Pjf_Bas"
Option Explicit
Sub LoadBas(DistPj As VBProject)
Dim BasItm
For Each BasItm In Itr(BasFfnAy(SrcpzDistPj(DistPj)))
    DistPj.VBComponents.Import BasItm
Next
End Sub

Private Function BasFfnAy(Srcp$) As String()
Dim Ffn$, I
For Each I In FfnSy(Srcp)
    Ffn = I
    If IsBasFfn(Ffn) Then
        PushI BasFfnAy, Ffn
    End If
Next
End Function
Private Function IsBasFfn(Ffn$) As Boolean
Select Case True
Case HasSfx(Ffn$, ".bas"): IsBasFfn = True
End Select
End Function


Attribute VB_Name = "QIde_Gen_Pjf_Bas"
Option Explicit
Private Const CMod$ = "MIde_Gen_Pjf_Bas."
Private Const Asm$ = "QIde"
Sub LoadBas(A As VBProject, Srcp$)
Dim BasItm
For Each BasItm In Itr(BasFfnAy(Srcp))
    A.VBComponents.Import BasItm
Next
End Sub

Private Function BasFfnAy(Srcp$) As String()
Dim Ffn$, I
For Each I In Itr(FfnSy(Srcp))
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

